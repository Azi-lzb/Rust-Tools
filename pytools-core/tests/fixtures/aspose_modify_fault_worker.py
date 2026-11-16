"""Failure injection only. Never an Aspose implementation or acceptance substitute."""
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as E

sys.stdin.reconfigure(encoding="utf-8")
sys.stdout.reconfigure(encoding="utf-8")

def emit(value):
    print(json.dumps(value), flush=True)

emit(dict(ready=True, protocol_version=1, engine="Aspose.Cells", pid=os.getpid()))
for line in sys.stdin:
    q = json.loads(line)
    with open(os.environ["ASPOSE_FAULT_RECORD"], "a", encoding="utf-8", newline="\n") as record:
        record.write(q["method"] + "\n")
    if os.environ.get("ASPOSE_FAULT_MODE") == "crash":
        with open(q["candidate"], "xb") as f:
            f.write(b"partial candidate")
        sys.exit(2)
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(q["input"]) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    workbook = E.fromstring(parts["xl/workbook.xml"])
    sheets = workbook.find("{" + ns + "}sheets")
    names = [s.attrib["name"] for s in sheets]
    E.SubElement(sheets, "{" + ns + "}sheet", {"name": "Evaluation Warning", "sheetId": "99", "{" + rel + "}id": "rIdEval"})
    parts["xl/workbook.xml"] = E.tostring(workbook)
    relationships = E.fromstring(parts["xl/_rels/workbook.xml.rels"])
    E.SubElement(relationships, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {"Id": "rIdEval", "Type": rel + "/worksheet", "Target": "worksheets/eval.xml"})
    parts["xl/_rels/workbook.xml.rels"] = E.tostring(relationships)
    types = E.fromstring(parts["[Content_Types].xml"])
    E.SubElement(types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override", {"PartName": "/xl/worksheets/eval.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
    parts["[Content_Types].xml"] = E.tostring(types)
    parts["xl/worksheets/eval.xml"] = ('<worksheet xmlns="' + ns + '"><sheetData/></worksheet>').encode()
    with zipfile.ZipFile(q["candidate"], "x") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    # Deliberately omit the requested edit while returning structurally valid output.
    emit(dict(ok=True, request_id=q["request_id"], candidate=q["candidate"], evaluation_output=True,
              source_sheets=names, output_sheets=names + ["Evaluation Warning"],
              source_active_index=0, output_active_index=0, processing_ms=1, load_ms=0,
              transform_ms=1, save_ms=0, pid=os.getpid(), attempts=1, draining=False))
