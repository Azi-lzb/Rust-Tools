#!/bin/bash
set -u
cd ~/RustTools-src
cargo test --workspace --no-fail-fast > wsl-test.log 2>&1
RC=$?
OK=$(grep -cE "test result: ok" wsl-test.log)
PASSED=$(grep -oE "test result: ok\. [0-9]+" wsl-test.log | awk "{s+=\$4} END {print s}")
FAILED=$(grep -oE "[0-9]+ failed" wsl-test.log | awk "{s+=\$1} END {print s}")
echo "ok-binaries: $OK"
echo "passed: $PASSED"
echo "failed: $FAILED"
grep -E "\.\.\. FAILED" wsl-test.log | sort -u | head -4
echo "exit: $RC"
