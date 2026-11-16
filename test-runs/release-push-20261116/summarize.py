"""Summarize recorded execution results; absent completion is never a pass."""
import json
from pathlib import Path
import re

RUN = Path("F:/RustBuild/pytools/release-push-20261116")
results = {}
for platform in ("windows", "wsl", "docker"):
    final_label = f"{platform}-final-workspace"
    label = final_label if (RUN / f"{final_label}.log").exists() else f"{platform}-workspace"
    log = RUN / f"{label}.log"
    status = RUN / f"{label}.exit"
    text = log.read_text(encoding="utf-8-sig", errors="replace") if log.exists() else ""
    groups = re.findall(r"test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;", text)
    results[platform] = {
        "log": log.name,
        "complete": status.exists(),
        "exit": int(status.read_text(encoding="utf-8-sig").strip()) if status.exists() else None,
        "groups": len(groups),
        "passed": sum(int(g[0]) for g in groups),
        "failed": sum(int(g[1]) for g in groups),
        "ignored": sum(int(g[2]) for g in groups),
        "failed_lines": [line for line in text.splitlines() if " ... FAILED" in line or line.startswith("error:")],
        "ignored_lines": [line for line in text.splitlines() if " ... ignored" in line],
    }
(RUN / "workspace-results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({k: {f: v for f, v in value.items() if not f.endswith("_lines")} for k, value in results.items()}, ensure_ascii=False))
