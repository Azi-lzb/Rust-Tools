#!/usr/bin/env python3
"""Batch-scoped LibreOffice UNO calculator for disposable XLSX copies.

This worker performs only formula calculation and save-back. Rust owns the
business logic and reads the resulting cached values through Calamine.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any


PROTOCOL_VERSION = 1


def _property(uno: Any, name: str, value: Any) -> Any:
    item = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


class LibreOfficeSession:
    def __init__(self) -> None:
        import uno

        self.uno = uno
        self.profile = tempfile.TemporaryDirectory(prefix="rusttools-lo-calc-profile-")
        profile_path = Path(self.profile.name).resolve()
        pipe_name = f"rusttools_calc_{os.getpid()}_{uuid.uuid4().hex}"
        profile_url = uno.systemPathToFileUrl(str(profile_path))
        self.office: subprocess.Popen[bytes] | None = None
        office_binary = shutil.which("soffice") or shutil.which("libreoffice")
        if not office_binary:
            raise RuntimeError("LibreOffice executable not found (expected soffice or libreoffice)")

        accept = f"--accept=pipe,name={pipe_name};urp;StarOffice.ServiceManager"
        self.office = subprocess.Popen(
            [
                office_binary,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile_url}",
                accept,
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        connection = f"uno:pipe,name={pipe_name};urp;StarOffice.ComponentContext"
        last_error: Exception | None = None
        deadline = time.monotonic() + 30.0
        while time.monotonic() < deadline:
            if self.office.poll() is not None:
                raise RuntimeError(
                    f"LibreOffice exited during startup (status {self.office.returncode})"
                )
            try:
                context = resolver.resolve(connection)
                self.desktop = context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", context
                )
                break
            except Exception as error:  # UNO pipe is not ready yet.
                last_error = error
                time.sleep(0.2)
        else:
            raise RuntimeError(f"could not connect to LibreOffice UNO: {last_error}")

    def recalculate(self, source_path: str) -> dict[str, Any]:
        path = Path(source_path).resolve(strict=True)
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"calculation copy must be .xlsx or .xlsm: {path.name}")

        url = self.uno.systemPathToFileUrl(str(path))
        options = (
            _property(self.uno, "Hidden", True),
            _property(self.uno, "ReadOnly", False),
            _property(self.uno, "MacroExecutionMode", 0),
            _property(self.uno, "UpdateDocMode", 3),
        )
        document = self.desktop.loadComponentFromURL(url, "_blank", 0, options)
        if document is None:
            raise RuntimeError(f"LibreOffice did not open calculation copy: {path.name}")

        try:
            document.enableAutomaticCalculation(True)
            document.calculateAll()
            document.store()
        finally:
            try:
                document.close(True)
            except Exception:
                pass

        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError("LibreOffice calculation did not leave a non-empty saved copy")
        return {"ok": True, "application": "LibreOffice Calc"}

    def close(self) -> None:
        try:
            self.desktop.terminate()
        except Exception:
            pass
        if self.office is not None:
            try:
                self.office.wait(timeout=8)
            except subprocess.TimeoutExpired:
                self.office.terminate()
                try:
                    self.office.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    self.office.kill()
                    self.office.wait()
        self.profile.cleanup()


def _write(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main() -> int:
    if "--version" in sys.argv:
        print(f"libreoffice_calc_worker protocol={PROTOCOL_VERSION}")
        return 0

    try:
        session = LibreOfficeSession()
    except Exception as error:
        _write({"ok": False, "stage": "start LibreOffice calculation session", "error": str(error)})
        return 2

    _write({"ok": True, "ready": True, "protocol": PROTOCOL_VERSION})
    try:
        for line in sys.stdin:
            try:
                request = json.loads(line)
                if request.get("op") == "shutdown":
                    _write({"ok": True, "stopped": True})
                    return 0
                if request.get("op") != "recalculate":
                    raise ValueError(f"unsupported operation: {request.get('op')!r}")
                response = session.recalculate(str(request["source_copy"]))
                _write(response)
            except Exception as error:
                _write({"ok": False, "stage": "recalculate and save workbook copy", "error": str(error)})
    finally:
        session.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
