"""Linux-side no-overwrite syscall diagnostic; writes synthetic .bin files only."""
import ctypes
import errno
import json
from pathlib import Path
import sys

root = Path(sys.argv[1]).resolve()
assert root.is_relative_to(Path("/mnt/f/RustBuild/pytools/release-push-20261116"))
source, target = root / "syscall-source.bin", root / "syscall-target.bin"
assert not source.exists() and not target.exists()
source.write_bytes(b"synthetic syscall source")
target.write_bytes(b"synthetic preexisting target")
libc = ctypes.CDLL(None, use_errno=True)
rename = libc.renameat2
rename.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
rename.restype = ctypes.c_int
result = rename(-100, bytes(source), -100, bytes(target), 1)
code = ctypes.get_errno()
assert result == -1 and code == errno.EEXIST, (result, code)
assert source.read_bytes() == b"synthetic syscall source"
assert target.read_bytes() == b"synthetic preexisting target"
absent = root / "syscall-new-target.bin"
assert not absent.exists()
result = rename(-100, bytes(source), -100, bytes(absent), 1)
code = ctypes.get_errno()
assert result == -1 and code == errno.EINVAL, (result, code)
assert source.read_bytes() == b"synthetic syscall source"
assert target.read_bytes() == b"synthetic preexisting target"
assert not absent.exists()
print(json.dumps(dict(conflict_errno=errno.EEXIST, absent_target_result=result, absent_target_errno=code,
                      source_unchanged=True, target_unchanged=True, new_target_absent=True)))
