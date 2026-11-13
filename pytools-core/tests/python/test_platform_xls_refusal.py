"""Replay a synthetic WPS/Excel BIFF fixture without needing COM on Linux."""
import importlib.util
import os
from pathlib import Path
import shutil
import tempfile
import unittest

spec = importlib.util.spec_from_file_location(
    "boundaries", Path(__file__).with_name("test_general_boundaries.py")
)
boundaries = importlib.util.module_from_spec(spec)
spec.loader.exec_module(boundaries)


@unittest.skipUnless(os.environ.get("RUSTTOOLS_OFFICE_FIXTURE_DIR") and os.environ.get("RUSTTOOLS_BOUNDARY_CLI"),
                     "requires a synthetic BIFF fixture and current CLI")
class XlsRefusal(unittest.TestCase):
    def test_existing_biff_is_refused_without_changing_inputs(self):
        fixture = Path(os.environ["RUSTTOOLS_OFFICE_FIXTURE_DIR"])
        with tempfile.TemporaryDirectory() as folder:
            d = Path(folder)
            source, target = d / "source.xlsx", d / "target.xls"
            for name in ("source.xlsx", "target.xls"):
                shutil.copyfile(fixture / name, d / name)
            self.assertEqual(target.read_bytes()[:8], bytes.fromhex("d0cf11e0a1b11ae1"))
            case = boundaries.Boundaries()
            config = case.config(d, target)
            before = {p: boundaries.digest(p) for p in (source, target, config)}
            text = case.cli(config, "2\n3\n0\n", "replayed-xls-refusal")
            self.assertIn("既有 XLS 去重追加已安全停用", text)
            for path, expected in before.items():
                self.assertEqual(boundaries.digest(path), expected)
            self.assertEqual(list(d.glob("*.xls")), [target])
            self.assertFalse(any("candidate" in p.name.lower() for p in d.rglob("*")))


if __name__ == "__main__":
    unittest.main(verbosity=2)
