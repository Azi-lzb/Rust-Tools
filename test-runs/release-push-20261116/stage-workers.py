"""Stage self-contained evaluation workers; exclude symbols and deduplicate runtime files."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys

root = Path(sys.argv[1]).resolve()
root.mkdir(exist_ok=False, parents=True)
sources = {
    "aspose": Path("F:/RustBuild/pytools/aspose-worker-selfcontained/publish"),
    "aspose-words": Path("F:/RustBuild/pytools/aspose-words-worker-selfcontained/publish"),
}
seen = {}
records = []
for name, source in sources.items():
    for original in sorted(source.rglob("*")):
        if not original.is_file() or original.suffix.lower() == ".pdb":
            continue
        digest = hashlib.sha256(original.read_bytes()).hexdigest()
        target = root / name / original.relative_to(source)
        target.parent.mkdir(exist_ok=True, parents=True)
        # Only new build output is linked. No source or business file is moved.
        if digest in seen:
            try:
                os.link(seen[digest], target)
            except OSError:
                shutil.copyfile(original, target)
        else:
            shutil.copyfile(original, target)
            seen[digest] = target
        assert hashlib.sha256(target.read_bytes()).hexdigest() == digest
        records.append({"path": target.relative_to(root).as_posix(), "sha256": digest, "bytes": target.stat().st_size})
(root / "MANIFEST.json").write_text(json.dumps(records, indent=2), encoding="utf-8")
print(json.dumps({"files": len(records), "bytes": sum(v["bytes"] for v in records), "unique_bytes": sum(p.stat().st_size for p in seen.values())}))
