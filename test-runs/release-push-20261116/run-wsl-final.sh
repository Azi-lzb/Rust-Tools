#!/usr/bin/env bash
set -eu
run=/mnt/f/RustBuild/pytools/release-push-20261116
# Retire only this task's malformed inline launcher, never Cargo or another job.
python3 - <<'PY'
import os, signal
from pathlib import Path
matches=[]
for entry in Path('/proc').iterdir():
    if not entry.name.isdigit(): continue
    try: args=(entry/'cmdline').read_bytes().split(b'\0')
    except OSError: continue
    if len(args)>2 and args[0].endswith(b'bash') and args[1]==b'-c' and b'RUSTTOOLS_TEST_LABEL=wsl-final-workspace' in args[2] and b'release-push-20261116' in args[2]:
        matches.append(int(entry.name))
assert len(matches)<=1, matches
for pid in matches:
    os.kill(pid,signal.SIGTERM)
print('Retired task-only inline launcher:',matches,flush=True)
PY
while [ ! -f "$run/wsl-workspace.exit" ]; do sleep 2; done
cp "$run/wsl-start.txt" "$run/wsl-initial-start.txt"
cp "$run/wsl-end.txt" "$run/wsl-initial-end.txt"
export RUSTTOOLS_TEST_LABEL=wsl-final-workspace
export RUSTTOOLS_TEST_STAGE="$(cat "$run/wsl-stage.txt")"
exec bash /mnt/c/Users/AZI/Desktop/VibeCode/Rust-Tools/test-runs/release-push-20261116/run-wsl.sh
