"""Real isolated Aspose Modify, with a post-response fault-injection proxy.

The proxy changes one saved literal before Rust validates the candidate. All
evaluation content is retained. This exercises the Rust-side validation policy,
not Aspose's internal failure handling; no production route is changed.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess

import openpyxl


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--worker", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    repo = Path(__file__).resolve().parents[3]
    root = Path(args.output).resolve()
    spec = importlib.util.spec_from_file_location("aspose_review_fixture", repo / "tools/AsposeWorker/test_menu_matrix.py")
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    command, stdin, _, _ = helper.prepare(root, "5-2", True)
    source, config = root / "source.xlsx", root / "config.xlsx"
    protected = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, config)}
    proxy = root / "fault-proxy.py"
    proxy.write_text(
        "import subprocess,sys,json,os\nfrom pathlib import Path\nimport openpyxl\n"
        f"child=subprocess.Popen([{str(Path(args.worker).resolve())!r},'--stdio'],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=sys.stderr,text=True,encoding='utf-8')\n"
        "def emit(message):\n"
        " if 'pid' in message: message['pid']=os.getpid()\n"
        " print(json.dumps(message,ensure_ascii=False),flush=True)\n"
        "try:\n"
        " hello=json.loads(child.stdout.readline());emit(hello)\n"
        " for line in sys.stdin:\n"
        "  request=json.loads(line);child.stdin.write(line);child.stdin.flush()\n"
        "  reply=json.loads(child.stdout.readline())\n"
        "  if request.get('method')=='Modify' and reply.get('ok'):\n"
        "   output=Path(request['candidate']);book=openpyxl.load_workbook(output)\n"
        "   book.worksheets[0]['A1']='CORRUPTED NON-TARGET HEADER';book.save(output);book.close()\n"
        "  emit(reply)\n"
        "finally:\n"
        " child.stdin.close();child.wait(timeout=30)\n", encoding="utf-8")
    env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
    env.update(PYTHONIOENCODING="utf-8", LOCALAPPDATA=str(root / "state"), XDG_STATE_HOME=str(root / "state"),
               RUSTTOOLS_ASPOSE_DOTNET=os.sys.executable, RUSTTOOLS_ASPOSE_WORKER_DLL=str(proxy))
    proc = subprocess.run([str(Path(args.exe).resolve()), *command], input=stdin, cwd=root, env=env,
                          capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
    (root / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
    (root / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
    for path, before in protected.items(): assert hashlib.sha256(path.read_bytes()).hexdigest() == before
    reviews = list(root.glob("*_待核对_业务校验不符_*.xlsx"))
    assert len(reviews) == 1, proc.stdout + proc.stderr
    book = openpyxl.load_workbook(reviews[0])
    try:
        assert "Evaluation Warning" in book.sheetnames
        assert book.worksheets[0]["A1"].value == "CORRUPTED NON-TARGET HEADER"
        assert book.worksheets[0]["A2"].value == "标准机构"
    finally: book.close()
    audit = json.loads(Path(str(reviews[0]) + ".run.json").read_text(encoding="utf-8"))
    assert audit["business_validation_passed"] is False
    assert audit["original_target_replaced"] is False
    assert audit["status"] == "review_required"
    assert "待核对副本" in proc.stderr
    result = dict(status="PASS", engine="real Aspose via test fault proxy", original_and_config_unchanged=True,
                  validation_passed=False, review_copy=str(reviews[0]), evaluation_retained=True)
    (root / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
