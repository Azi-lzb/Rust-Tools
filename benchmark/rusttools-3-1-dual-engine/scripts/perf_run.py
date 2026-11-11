# -*- coding: utf-8 -*-
"""Wall time + CPU time + peak working set for a process tree (no psutil).

Uses ctypes Toolhelp snapshots + GetProcessTimes/GetProcessMemoryInfo to sample
the whole tree (exe + spawned python workers) every 150ms. Children that exit
mid-run keep their accumulated deltas.

Usage:
  python -I scripts/perf_run.py <exe> <config> <openpyxl|rust> <sink.jsonl> [reps]
"""
import ctypes
import ctypes.wintypes as wt
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[3]
DIST_PYTHON = REPO / "dist" / "python" / "python.exe"

kernel32 = ctypes.windll.kernel32
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
TH32CS_SNAPPROCESS = 0x2


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wt.DWORD),
        ("cntUsage", wt.DWORD),
        ("th32ProcessID", wt.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wt.DWORD),
        ("cntThreads", wt.DWORD),
        ("th32ParentProcessID", wt.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wt.DWORD),
        ("szExeFile", wt.WCHAR * 260),
    ]


class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
    _fields_ = [
        ("cb", wt.DWORD),
        ("PageFaultCount", wt.DWORD),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
    ]


kernel32.CreateToolhelp32Snapshot.restype = wt.HANDLE
kernel32.OpenProcess.restype = wt.HANDLE
psapi = ctypes.windll.psapi


def snapshot_processes():
    handle = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if handle == wt.HANDLE(-1):
        return {}
    entry = PROCESSENTRY32W()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
    result = {}
    if kernel32.Process32FirstW(handle, ctypes.byref(entry)):
        while True:
            result[entry.th32ProcessID] = entry.th32ParentProcessID
            if not kernel32.Process32NextW(handle, ctypes.byref(entry)):
                break
    kernel32.CloseHandle(handle)
    return result


def tree_cpu_and_peak(root_pid, last_times):
    """Return (cpu_seconds_delta_sum, max_peak_ws_bytes_this_sample)."""
    procs = snapshot_processes()
    # children map for the current snapshot
    children = {}
    for pid, ppid in procs.items():
        children.setdefault(ppid, set()).add(pid)
    tree, queue = set(), [root_pid]
    while queue:
        pid = queue.pop()
        if pid in tree:
            continue
        tree.add(pid)
        queue.extend(children.get(pid, ()))
    cpu_delta = 0.0
    peak_here = 0
    for pid in tree:
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            continue
        created, exited, kernel_t, user_t = (ctypes.c_ulonglong() for _ in range(4))
        if kernel32.GetProcessTimes(handle, ctypes.byref(created), ctypes.byref(exited),
                                    ctypes.byref(kernel_t), ctypes.byref(user_t)):
            total = (kernel_t.value + user_t.value) / 1e7
            prev = last_times.get(pid, 0.0)
            if total >= prev:
                cpu_delta += total - prev
            last_times[pid] = total
        pmc = PROCESS_MEMORY_COUNTERS()
        pmc.cb = ctypes.sizeof(PROCESS_MEMORY_COUNTERS)
        if psapi.GetProcessMemoryInfo(handle, ctypes.byref(pmc), pmc.cb):
            peak_here = max(peak_here, pmc.PeakWorkingSetSize)
        kernel32.CloseHandle(handle)
    return cpu_delta, peak_here


def run_once(exe, config, engine):
    env = dict(os.environ)
    env.update({
        "PYTOOLS_CONFIG_PATH": str(Path(config).resolve()),
        "PYTOOLS_OPENPYXL_PYTHON": str(DIST_PYTHON),
        "PYTOOLS_OPENPYXL_WORKER": str(REPO / "dist" / "python" / "openpyxl_worker.py"),
    })
    payload = {"openpyxl": "3\n1\n0\n0\n", "rust": "3\n2\n0\n0\n"}[engine].encode()
    last_times = {}
    cpu_seconds = 0.0
    peak_ws = 0
    started = time.perf_counter()
    proc = subprocess.Popen(
        [str(exe)], stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL, env=env, cwd=str(Path(exe).resolve().parent),
    )
    # 必须写入并关闭 stdin，否则子进程会一直等输入而挂住。
    try:
        proc.stdin.write(payload)
        proc.stdin.close()
    except (BrokenPipeError, OSError):
        pass
    # 采样必须轻量：ctypes 快照在进程数多时会拖慢甚至阻塞，
    # 因此采样间隔放宽到 250ms，且单次快照失败立即跳过。
    while proc.poll() is None:
        try:
            delta, peak = tree_cpu_and_peak(proc.pid, last_times)
            cpu_seconds += delta
            peak_ws = max(peak_ws, peak)
        except Exception:
            pass
        time.sleep(0.25)
    try:
        delta, peak = tree_cpu_and_peak(proc.pid, last_times)
        cpu_seconds += delta
        peak_ws = max(peak_ws, peak)
    except Exception:
        pass
    wall_ms = int((time.perf_counter() - started) * 1000)
    return {"exit": proc.returncode, "wall_ms": wall_ms,
            "cpu_seconds": round(cpu_seconds, 3),
            "peak_working_set_mb": round(peak_ws / (1024 * 1024), 1)}


def main():
    exe, config, engine, sink = sys.argv[1], Path(sys.argv[2]), sys.argv[3], Path(sys.argv[4])
    reps = int(sys.argv[5]) if len(sys.argv) > 5 else 5
    for rep in range(1, reps + 1):
        record = {"config": config.name, "engine": engine, "rep": rep,
                  **run_once(exe, config, engine)}
        with open(sink, "a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")
        print(engine, rep, "wall", record["wall_ms"], "cpu", record["cpu_seconds"],
              "peakMB", record["peak_working_set_mb"], "exit", record["exit"])


if __name__ == "__main__":
    main()
