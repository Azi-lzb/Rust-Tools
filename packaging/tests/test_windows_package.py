"""Packaging-only regressions using fake files in an isolated temporary repo."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
import zipfile

SCRIPT = Path(__file__).resolve().parents[1] / "windows-package.ps1"


@unittest.skipUnless(os.name == "nt", "Windows PowerShell packaging")
class PackageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="rusttools-package-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        subprocess.run(["git", "init", "-b", "main", str(self.root)], check=True, capture_output=True)
        for name in ("dist/python", "archives", "staging"):
            (self.root / name).mkdir(parents=True)
        for name in ("config.xlsx", "README.md", "dist/RustTools.exe", "dist/python/runtime-marker.txt"):
            (self.root / name).write_bytes(b"synthetic packaging fixture")

    def package(self, *args):
        return subprocess.run([
            shutil.which("pwsh") or "powershell", "-NoProfile", "-File", str(SCRIPT),
            "-ProjectRoot", str(self.root), "-StageRoot", str(self.root / "staging"),
            "-ArchiveDirectory", str(self.root / "archives"), *args,
        ], capture_output=True, text=True, encoding="utf-8", errors="replace")

    def test_includes_optional_self_contained_worker_without_pdb(self):
        worker = self.root / "dist/workers/aspose"
        worker.mkdir(parents=True)
        for name in ("AsposeWorker.exe", "AsposeWorker.dll", "AsposeWorker.runtimeconfig.json", "Aspose.Cells.dll", "coreclr.dll", "AsposeWorker.pdb"):
            (worker / name).write_bytes(b"synthetic")
        result = self.package()
        self.assertEqual(result.returncode, 0, result.stderr)
        archives = list((self.root / "archives").glob("*TEST-main*.zip"))
        self.assertEqual(len(archives), 1)
        with zipfile.ZipFile(archives[0]) as archive:
            names = {name.replace("\\", "/") for name in archive.namelist()}
            self.assertIn("workers/aspose/AsposeWorker.exe", names)
            self.assertIn("config/config.xlsx", names)
            self.assertFalse(any(name.endswith(".pdb") for name in names))
        self.assertEqual(list((self.root / "staging").iterdir()), [])

    def test_incomplete_worker_refused_without_archive(self):
        (self.root / "dist/workers/aspose").mkdir(parents=True)
        result = self.package()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Incomplete self-contained worker", result.stderr)
        self.assertEqual(list((self.root / "archives").iterdir()), [])
        self.assertEqual(list((self.root / "staging").iterdir()), [])

    def test_release_requires_acceptance_and_main(self):
        self.assertNotEqual(self.package("-PackageKind", "release").returncode, 0)
        subprocess.run(["git", "-C", str(self.root), "symbolic-ref", "HEAD", "refs/heads/feature/mock"], check=True)
        self.assertNotEqual(self.package("-PackageKind", "release", "-ReleaseAccepted").returncode, 0)
        self.assertEqual(list((self.root / "archives").iterdir()), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
