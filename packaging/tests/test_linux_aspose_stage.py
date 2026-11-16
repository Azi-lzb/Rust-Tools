"""Guard/parser regressions; real ELF/worker acceptance is a separate smoke."""
import os
from pathlib import Path
import subprocess
import tempfile
import unittest


SCRIPT = Path(__file__).resolve().parents[1] / "aspose/stage-linux-workers.sh"


@unittest.skipUnless(os.name == "posix", "Linux staging script")
class LinuxAsposeStageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "source"
        self.app = self.root / "application"
        self.app.mkdir()
        self.tools = self.root / "tools"
        self.tools.mkdir()
        objdump = self.tools / "objdump"
        objdump.write_text('#!/bin/sh\nif [ "$1" = "-f" ]; then echo "architecture: ${TEST_ARCH:-i386:x86-64}"; else echo "GLIBC_${TEST_GLIBC:-2.28}"; fi\n')
        objdump.chmod(0o755)
        for directory, name in (("aspose", "AsposeWorker"), ("aspose-words", "AsposeWordsWorker")):
            location = self.source / directory
            location.mkdir(parents=True)
            for leaf in (name, "libcoreclr.so", "libhostfxr.so", "libSkiaSharp.so"):
                (location / leaf).write_bytes(b"\x7fELFparser-fixture")
            (location / (name + ".dll")).write_bytes(b"managed-fixture")
            (location / (name + ".runtimeconfig.json")).write_text("{}")

    def run_stage(self, **variables):
        env = dict(os.environ, PATH=str(self.tools) + os.pathsep + os.environ["PATH"], **variables)
        return subprocess.run(["bash", str(SCRIPT), str(self.source), str(self.app)],
                              capture_output=True, text=True, env=env, timeout=30)

    def test_glibc_below_and_at_ceiling_are_accepted(self):
        for version in ("2.27", "2.28", "2.28.0"):
            with self.subTest(version=version):
                self.app = self.root / ("application-" + version)
                self.app.mkdir()
                result = self.run_stage(TEST_GLIBC=version)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertTrue((self.app / "workers/aspose/AsposeWorker").stat().st_mode & 0o111)
                self.assertEqual((self.app / "workers/aspose/libcoreclr.so").stat().st_mode & 0o777, 0o644)
                self.assertEqual((self.app / "workers/aspose/AsposeWorker").stat().st_mode & 0o777, 0o755)
                self.assertTrue((self.app / "workers/MANIFEST.sha256").is_file())

    def test_newer_unknown_glibc_and_wrong_arch_fail_before_copy(self):
        for variables in ({"TEST_GLIBC": "2.29"}, {"TEST_GLIBC": "unknown"}, {"TEST_ARCH": "aarch64"}):
            with self.subTest(variables=variables):
                self.assertNotEqual(self.run_stage(**variables).returncode, 0)
                self.assertFalse((self.app / "workers").exists())

    def test_incomplete_and_windows_workers_are_rejected(self):
        skia = self.source / "aspose/libSkiaSharp.so"
        skia.unlink()
        self.assertNotEqual(self.run_stage().returncode, 0)
        skia.write_bytes(b"\x7fELFparser-fixture")
        (self.source / "aspose/AsposeWorker.exe").write_bytes(b"MZ")
        self.assertNotEqual(self.run_stage().returncode, 0)
        self.assertFalse((self.app / "workers").exists())

    def test_existing_destination_is_never_overwritten(self):
        target = self.app / "workers"
        target.mkdir()
        marker = target / "existing"
        marker.write_bytes(b"preserve")
        self.assertNotEqual(self.run_stage().returncode, 0)
        self.assertEqual(marker.read_bytes(), b"preserve")

    def test_symlinks_are_rejected(self):
        (self.source / "aspose/external").symlink_to(self.root)
        self.assertNotEqual(self.run_stage().returncode, 0)
        self.assertFalse((self.app / "workers").exists())


if __name__ == "__main__":
    unittest.main()
