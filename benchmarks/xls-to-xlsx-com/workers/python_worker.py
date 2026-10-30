"""Synthetic-workload-only Excel COM worker for benchmark scheme B."""

from __future__ import annotations

import gc
import json
import sys
import threading
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client


def emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def owned_excel_pid(app, baseline: set[int]) -> tuple[int, object]:
    _, pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))
    if not pid or pid in baseline or int(app.Workbooks.Count) != 0:
        raise RuntimeError(
            f"Refusing non-owned Excel instance PID={pid}, workbooks={app.Workbooks.Count}"
        )
    handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    return int(pid), handle


def convert_one(app, source: str, target: str, index: int) -> dict:
    result = {"source": source, "target": target, "index": index, "status": "failed"}
    workbook = None
    start = time.perf_counter()
    try:
        phase = time.perf_counter()
        workbook = app.Workbooks.Open(source, 0, True)
        result["open_ms"] = (time.perf_counter() - phase) * 1000
        phase = time.perf_counter()
        workbook.SaveAs(target, FileFormat=51)
        result["save_as_ms"] = (time.perf_counter() - phase) * 1000
        phase = time.perf_counter()
        workbook.Close(False)
        workbook = None
        result["close_ms"] = (time.perf_counter() - phase) * 1000
        result["status"] = "succeeded"
    except Exception as error:  # keep processing later inputs after one file fails
        result["error"] = f"Open/SaveAs {source}: {type(error).__name__}: {error}"
        if workbook is not None:
            try:
                phase = time.perf_counter()
                workbook.Close(False)
                result["close_ms"] = (time.perf_counter() - phase) * 1000
            except Exception as close_error:
                result["close_error"] = f"{type(close_error).__name__}: {close_error}"
            workbook = None
    finally:
        gc.collect()
        result["total_ms"] = (time.perf_counter() - start) * 1000
    return result


def main() -> int:
    # READY excludes Excel creation and COM apartment setup from worker startup.
    emit(
        {
            "event": "worker_ready",
            "worker_pid": __import__("os").getpid(),
            "sta_thread_id": threading.get_ident(),
        }
    )
    request = json.loads(sys.stdin.buffer.read().decode("utf-8"))
    timings: dict[str, float] = {}
    app = None
    excel_pid = None
    excel_handle = None
    apartment = False
    files = []
    fatal = None
    try:
        phase = time.perf_counter()
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        apartment = True
        timings["sta_init_ms"] = (time.perf_counter() - phase) * 1000

        phase = time.perf_counter()
        baseline = {int(pid) for pid in request["baseline_pids"]}
        app = win32com.client.DispatchEx(request.get("excel_progid", "Excel.Application"))
        excel_pid, excel_handle = owned_excel_pid(app, baseline)
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        configuration_book = app.Workbooks.Add()
        try:
            app.Calculation = -4135  # xlCalculationManual
            app.CalculateBeforeSave = False
            if int(app.Calculation) != -4135:
                raise RuntimeError("Excel did not enter manual calculation mode")
        finally:
            configuration_book.Close(False)
            configuration_book = None
        timings["excel_create_config_ms"] = (time.perf_counter() - phase) * 1000
        emit({"event": "excel_ready", "excel_pid": excel_pid})

        for index, entry in enumerate(request["files"]):
            files.append(convert_one(app, entry["input"], entry["output"], index))
            if (index + 1) % 10 == 0:
                emit({"event": "progress", "completed": index + 1, "total": len(request["files"])})
    except Exception as error:
        fatal = f"{type(error).__name__}: {error}"
    finally:
        if app is not None and excel_pid is not None:
            phase = time.perf_counter()
            try:
                app.Quit()
            except Exception as error:
                fatal = fatal or f"Excel Quit failed: {type(error).__name__}: {error}"
            timings["excel_quit_ms"] = (time.perf_counter() - phase) * 1000
        app = None
        gc.collect()
        if apartment:
            phase = time.perf_counter()
            pythoncom.CoUninitialize()
            timings["sta_release_ms"] = (time.perf_counter() - phase) * 1000
        excel_exit_confirmed = None
        if excel_handle is not None:
            phase = time.perf_counter()
            excel_exit_confirmed = (
                win32event.WaitForSingleObject(excel_handle, 30000)
                == win32event.WAIT_OBJECT_0
            )
            timings["excel_exit_wait_ms"] = (time.perf_counter() - phase) * 1000
            win32api.CloseHandle(excel_handle)

    emit(
        {
            "event": "worker_result",
            "worker_pid": __import__("os").getpid(),
            "sta_thread_id": threading.get_ident(),
            "excel_pid": excel_pid,
            "excel_instance_count": int(excel_pid is not None),
            "excel_exit_confirmed": excel_exit_confirmed,
            "timings": timings,
            "files": files,
            "fatal_error": fatal,
        }
    )
    return 0 if fatal is None and excel_exit_confirmed is not False else 2


if __name__ == "__main__":
    raise SystemExit(main())
