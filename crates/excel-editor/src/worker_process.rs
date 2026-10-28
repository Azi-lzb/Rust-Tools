//! 受控 worker 进程生命周期：启动、超时与整棵进程树的跨平台清理。
//!
//! - Windows：子进程以 CREATE_SUSPENDED 启动，先加入专用 Job Object
//!   （KILL_ON_JOB_CLOSE）再恢复主线程，关闭“子进程在受控前逃逸”的竞争窗口；
//!   超时或异常时 TerminateJobObject 终止整棵进程树，Job 句柄关闭兜底。
//! - Unix：子进程通过 `process_group(0)` 建立独立进程组，超时时
//!   `kill(-pgid, SIGKILL)` 终止整组，再回收直接子进程。
//!
//! 两者都只作用于本次派生的 worker 及其可控行生子进程，不会影响主程序
//! 或用户其他 Python 进程。

use std::io;
use std::process::{Child, Command, ExitStatus};
use std::time::{Duration, Instant};

/// Drop 时若子进程仍在运行，先终止进程树再回收，最多等待该时长。
const DROP_WAIT: Duration = Duration::from_secs(3);

pub(crate) struct WorkerProcess {
    child: Child,
    platform: PlatformGuard,
}

impl std::fmt::Debug for WorkerProcess {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WorkerProcess")
            .field("pid", &self.child.id())
            .finish()
    }
}

enum PlatformGuard {
    #[cfg(windows)]
    Job { handle: windows_sys::Win32::Foundation::HANDLE },
    #[cfg(unix)]
    ProcessGroup,
    /// 仅测试：terminate 恒失败，用于验证清理错误进入诊断消息。
    #[cfg(test)]
    AlwaysFails,
}

impl WorkerProcess {
    /// 以受控方式启动 worker：进程树约束在 spawn 与首次执行之间建立。
    pub(crate) fn spawn(mut command: Command) -> io::Result<Self> {
        #[cfg(windows)]
        {
            use std::os::windows::io::AsRawHandle;
            use std::os::windows::process::CommandExt;
            // CREATE_SUSPENDED：保证在子进程执行任何指令前完成入 Job。
            command.creation_flags(CREATE_SUSPENDED);
            let mut child = command.spawn()?;
            let handle = match create_kill_on_close_job() {
                Ok(handle) => handle,
                Err(error) => {
                    // Job 建不出来时不能留下一个不受控的子进程。
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(error);
                }
            };
            let assigned = unsafe {
                windows_sys::Win32::System::JobObjects::AssignProcessToJobObject(
                    handle,
                    child.as_raw_handle() as windows_sys::Win32::Foundation::HANDLE,
                )
            };
            if assigned == 0 {
                let error = io::Error::last_os_error();
                unsafe {
                    windows_sys::Win32::System::JobObjects::TerminateJobObject(handle, 1);
                    windows_sys::Win32::Foundation::CloseHandle(handle);
                }
                let _ = child.wait();
                return Err(error);
            }
            // 子进程已受控，恢复主线程；恢复失败必须终止而不是留下悬挂进程。
            if let Err(error) = resume_child_threads(child.id()) {
                unsafe {
                    windows_sys::Win32::System::JobObjects::TerminateJobObject(handle, 1);
                }
                let _ = child.wait();
                return Err(error);
            }
            Ok(Self {
                child,
                platform: PlatformGuard::Job { handle },
            })
        }
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            // 让 worker 成为新进程组的组长（pgid == 子进程 pid）。
            command.process_group(0);
            let child = command.spawn()?;
            Ok(Self {
                child,
                platform: PlatformGuard::ProcessGroup,
            })
        }
    }

    pub(crate) fn id(&self) -> u32 {
        self.child.id()
    }

    pub(crate) fn try_wait(&mut self) -> io::Result<Option<ExitStatus>> {
        self.child.try_wait()
    }

    pub(crate) fn wait(&mut self) -> io::Result<ExitStatus> {
        self.child.wait()
    }

    /// 终止整棵受控进程树（不等待；调用方随后 `wait` 回收直接子进程）。
    /// 返回终止动作本身的错误（例如句柄失效），供诊断记录。
    pub(crate) fn terminate_tree(&mut self) -> io::Result<()> {
        match &self.platform {
            #[cfg(windows)]
            PlatformGuard::Job { handle } => {
                let ok = unsafe {
                    windows_sys::Win32::System::JobObjects::TerminateJobObject(*handle, 1)
                };
                if ok == 0 {
                    return Err(io::Error::last_os_error());
                }
                Ok(())
            }
            #[cfg(unix)]
            PlatformGuard::ProcessGroup => {
                let pgid = self.child.id() as i32;
                let result = unsafe { libc::kill(-pgid, libc::SIGKILL) };
                if result != 0 {
                    let error = io::Error::last_os_error();
                    if error.raw_os_error() == Some(libc::ESRCH) {
                        // 整组已不存在（例如直接子进程恰好先退出）；
                        // 保证直接子进程被终止即可。
                        return self.child.kill();
                    }
                    return Err(error);
                }
                Ok(())
            }
            #[cfg(test)]
            PlatformGuard::AlwaysFails => Err(io::Error::other("注入的清理失败")),
        }
    }
}

#[cfg(test)]
impl WorkerProcess {
    /// 以普通方式启动子进程并挂接“终止恒失败”的守卫：
    /// 模拟 terminate_tree 返回错误的场景（回收路径仍可用）。
    pub(crate) fn spawn_with_failing_termination(mut command: Command) -> io::Result<Self> {
        let child = command.spawn()?;
        Ok(Self {
            child,
            platform: PlatformGuard::AlwaysFails,
        })
    }
}

impl Drop for WorkerProcess {
    fn drop(&mut self) {
        let started = Instant::now();
        loop {
            match self.try_wait() {
                Ok(Some(_)) => return,
                Ok(None) if started.elapsed() < DROP_WAIT => {
                    let _ = self.terminate_tree();
                    std::thread::sleep(Duration::from_millis(10));
                }
                // 终止失败或等待超时：交由平台兜底（Windows 的
                // KILL_ON_JOB_CLOSE 在句柄关闭时杀树；Unix 上 SIGKILL
                // 已发出，无法终止的进程不属于本模块可处置范围）。
                _ => return,
            }
        }
    }
}

#[cfg(windows)]
const CREATE_SUSPENDED: u32 = 0x0000_0004;

/// 建立带 KILL_ON_JOB_CLOSE 的 Job：句柄关闭（含主进程崩溃）时整树被杀。
#[cfg(windows)]
fn create_kill_on_close_job() -> io::Result<windows_sys::Win32::Foundation::HANDLE> {
    use windows_sys::Win32::System::JobObjects::{
        JobObjectExtendedLimitInformation, SetInformationJobObject,
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION, JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
    };
    unsafe {
        let handle = windows_sys::Win32::System::JobObjects::CreateJobObjectW(
            std::ptr::null(),
            std::ptr::null(),
        );
        if handle.is_null() {
            return Err(io::Error::last_os_error());
        }
        let mut limits: JOBOBJECT_EXTENDED_LIMIT_INFORMATION = std::mem::zeroed();
        limits
            .BasicLimitInformation
            .LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        let ok = SetInformationJobObject(
            handle,
            JobObjectExtendedLimitInformation,
            &limits as *const _ as *const core::ffi::c_void,
            std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
        );
        if ok == 0 {
            let error = io::Error::last_os_error();
            windows_sys::Win32::Foundation::CloseHandle(handle);
            return Err(error);
        }
        Ok(handle)
    }
}

/// 恢复以 CREATE_SUSPENDED 启动的子进程：枚举其全部线程并逐个 ResumeThread。
/// 进程刚启动时通常只有一个线程；全部恢复成功才算受控启动完成。
#[cfg(windows)]
fn resume_child_threads(child_pid: u32) -> io::Result<()> {
    use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
    use windows_sys::Win32::System::Diagnostics::ToolHelp::{
        CreateToolhelp32Snapshot, Thread32First, Thread32Next, TH32CS_SNAPTHREAD, THREADENTRY32,
    };
    use windows_sys::Win32::System::Threading::{
        OpenThread, ResumeThread, THREAD_SUSPEND_RESUME,
    };
    unsafe {
        let snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
        if snapshot == INVALID_HANDLE_VALUE || snapshot.is_null() {
            return Err(io::Error::last_os_error());
        }
        let mut entry: THREADENTRY32 = std::mem::zeroed();
        entry.dwSize = std::mem::size_of::<THREADENTRY32>() as u32;
        let mut resumed = false;
        let mut result = Ok(());
        if Thread32First(snapshot, &mut entry) != 0 {
            loop {
                if entry.th32OwnerProcessID == child_pid {
                    let thread = OpenThread(THREAD_SUSPEND_RESUME, 0, entry.th32ThreadID);
                    if thread.is_null() || thread == INVALID_HANDLE_VALUE {
                        result = Err(io::Error::last_os_error());
                        break;
                    }
                    // ResumeThread 返回恢复前的挂起计数；-1 表示失败。
                    let previous = ResumeThread(thread);
                    windows_sys::Win32::Foundation::CloseHandle(thread);
                    if previous == u32::MAX {
                        result = Err(io::Error::last_os_error());
                        break;
                    }
                    resumed = true;
                }
                if Thread32Next(snapshot, &mut entry) == 0 {
                    break;
                }
            }
        } else {
            result = Err(io::Error::last_os_error());
        }
        CloseHandle(snapshot);
        if result.is_ok() && !resumed {
            // 找不到任何属于子进程的线程：进程可能在恢复前已退出。
            // 让调用方用 try_wait 判断；此处按可等待的正常路径放行。
            return Ok(());
        }
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::process::Stdio;

    /// 测试用受控 Python：优先隔离测试运行时环境变量，其次仓库 dist 运行时。
    fn test_python() -> Option<std::path::PathBuf> {
        if let Some(path) = std::env::var_os("RUST_WORKSPACE_OPENPYXL_PYTHON") {
            let path = std::path::PathBuf::from(path);
            if path.is_file() {
                return Some(path);
            }
        }
        if let Some(path) = std::env::var_os("PYTOOLS_OPENPYXL_PYTHON") {
            let path = std::path::PathBuf::from(path);
            if path.is_file() {
                return Some(path);
            }
        }
        let workspace = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()?
            .parent()?
            .to_path_buf();
        let dist_python = if cfg!(windows) {
            workspace.join("dist/python/python.exe")
        } else {
            workspace.join("dist/python/python3")
        };
        dist_python.is_file().then_some(dist_python)
    }

    fn hang_command(python: &std::path::Path) -> Command {
        let mut command = Command::new(python);
        command.arg("-I").arg("-c").arg("import time; time.sleep(120)");
        command.stdin(Stdio::null());
        command
    }

    fn process_alive(pid: u32) -> bool {
        #[cfg(windows)]
        {
            use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
            use windows_sys::Win32::System::Threading::{
                OpenProcess, WaitForSingleObject, PROCESS_QUERY_LIMITED_INFORMATION,
            };
            // WaitForSingleObject 需要句柄带 SYNCHRONIZE 权限。
            const PROCESS_SYNCHRONIZE: u32 = 0x0010_0000;
            const WAIT_TIMEOUT: u32 = 0x102;
            unsafe {
                let handle =
                    OpenProcess(PROCESS_SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, 0, pid);
                if handle.is_null() || handle == INVALID_HANDLE_VALUE {
                    // 打不开句柄一般表示进程不存在（ACCESS_DENIED 极罕见，
                    // 测试环境中以不存在处理）。
                    return false;
                }
                let wait = WaitForSingleObject(handle, 0);
                CloseHandle(handle);
                wait == WAIT_TIMEOUT
            }
        }
        #[cfg(unix)]
        {
            // kill(pid, 0)：ESRCH = 不存在；EPERM = 存在但无权限。
            let result = unsafe { libc::kill(pid as i32, 0) };
            result == 0 || std::io::Error::last_os_error().raw_os_error() == Some(libc::EPERM)
        }
    }

    #[test]
    fn worker_exits_normally_and_is_reaped() {
        let Some(python) = test_python() else {
            eprintln!("SKIP: no local python runtime for worker_process tests");
            return;
        };
        let mut command = Command::new(&python);
        command.arg("-I").arg("-c").arg("import sys; sys.exit(0)");
        command.stdin(Stdio::null());
        let mut worker = WorkerProcess::spawn(command).expect("spawn controlled worker");
        assert!(process_alive(worker.id()), "worker must be alive right after spawn");
        let mut exited = false;
        let deadline = Instant::now() + Duration::from_secs(30);
        while Instant::now() < deadline {
            if worker.try_wait().expect("try_wait").is_some() {
                exited = true;
                break;
            }
            std::thread::sleep(Duration::from_millis(10));
        }
        assert!(exited, "a normal worker must exit on its own");
    }

    #[test]
    fn timeout_terminates_tree_and_descendants() {
        let Some(python) = test_python() else {
            eprintln!("SKIP: no local python runtime for worker_process tests");
            return;
        };
        let dir = std::env::temp_dir().join(format!(
            "excel-editor-tree-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let pid_file = dir.join("descendant.pid");
        let script = dir.join("spawn_hang_worker.py");
        std::fs::write(
            &script,
            "import os, subprocess, sys, time\n\
             child = subprocess.Popen([sys.executable, '-I', '-c', 'import time; time.sleep(120)'])\n\
             open(os.environ['TREE_TEST_PID_FILE'], 'w').write(str(child.pid))\n\
             time.sleep(120)\n",
        )
        .unwrap();
        let mut command = Command::new(&python);
        command.arg("-I").arg(&script);
        command.env("TREE_TEST_PID_FILE", &pid_file);
        command.stdin(Stdio::null());

        let mut worker = WorkerProcess::spawn(command).expect("spawn tree test worker");
        // 给测试 worker 足够时间启动并写出孙进程 pid。
        let deadline = Instant::now() + Duration::from_secs(20);
        let mut descendant_pid = None;
        while Instant::now() < deadline {
            if let Ok(text) = std::fs::read_to_string(&pid_file) {
                if let Ok(pid) = text.trim().parse::<u32>() {
                    descendant_pid = Some(pid);
                    break;
                }
            }
            std::thread::sleep(Duration::from_millis(20));
        }
        let descendant_pid = descendant_pid.expect("worker must record its child pid in time");
        assert!(
            process_alive(descendant_pid),
            "descendant must be alive before termination"
        );

        let terminate_result = worker.terminate_tree();
        assert!(terminate_result.is_ok(), "terminate_tree failed: {terminate_result:?}");
        let status = worker.wait().expect("wait after terminate_tree");
        assert!(!status.success(), "terminated worker must not report success");
        let deadline = Instant::now() + Duration::from_secs(10);
        while process_alive(descendant_pid) && Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(20));
        }
        assert!(
            !process_alive(descendant_pid),
            "descendant {descendant_pid} survived the process-tree termination"
        );
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn spawn_failure_is_reported_without_side_effects() {
        let missing = if cfg!(windows) {
            std::path::PathBuf::from(r"C:\__pytools_missing_python__.exe")
        } else {
            std::path::PathBuf::from("/__pytools_missing_python__")
        };
        let command = hang_command(&missing);
        let error = WorkerProcess::spawn(command).expect_err("missing python must fail to spawn");
        assert!(
            error.kind() == io::ErrorKind::NotFound,
            "unexpected spawn error: {error}"
        );
    }
}
