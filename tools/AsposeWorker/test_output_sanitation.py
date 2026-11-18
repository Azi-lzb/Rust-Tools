"""Synthetic real-SDK checks for production XLSX output sanitation.

Uses the direct .NET Worker only; does not execute Rust or modify dist.
"""
import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
from test_raw_sdk_archive import run, request


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def collision_cases(dotnet, dll, root):
    root.mkdir()
    notice = "Evaluation Only. Created with Aspose.Cells for .NET. Copyright 2003 - 2026 Aspose Pty Ltd."
    marker = "Evaluation Warning"

    def plan(kind, **fields):
        return dict(op="edit_plan_operation", operation=dict(operation=kind, **fields))

    fixtures = [
        ("empty", marker, None, True, "Calculate", [], marker, marker + " (1)"),
        ("data", marker, 17, True, "Calculate", [], marker, marker + " (1)"),
        ("similar", marker, notice + " (user copy)", True, "Convert", [], marker, marker + " (1)"),
        ("identical", marker, notice, True, "Calculate", [], marker, marker + " (1)"),
        ("case", "evaluation warning", 17, True, "Calculate", [], "evaluation warning", marker),
        ("near", marker + " backup", 17, True, "Calculate", [], marker + " backup", marker),
        ("create", "User", None, True, "Modify", [plan("create_sheet", name=marker)], "User", marker + " (1)"),
        ("delete", marker, 17, False, "Modify", [plan("delete_sheet", name=marker)], "Data", marker),
        ("rename", "User", 17, True, "Modify", [plan("rename_sheet", old_name="User", new_name=marker)], marker, marker + " (1)"),
    ]
    tasks, snapshots = [], []
    for label, name, value, reference, method, operations, active, appended in fixtures:
        source = root / (label + "-source.xlsx")
        book = openpyxl.Workbook()
        book.active.title = "Data"
        book["Data"]["A1"] = 7
        user = book.create_sheet(name)
        if value is not None:
            user["A5"] = value
        ref = book.create_sheet("Ref")
        if reference:
            ref["A1"] = ("='" + name + "'!A5") if " " in name else "=User!A5"
        book.active = 1
        book.save(source)
        book.close()
        tasks.append(request(source, root / (label + ".xlsx"), label, method, operations))
        snapshots.append(digest(source))
    responses = run(dotnet, dll, tasks, root, root / "raw")
    outcomes = []
    for fixture, task, reply, sha_before in zip(fixtures, tasks, responses, snapshots):
        label, name, value, reference, _, _, active, appended = fixture
        assert reply["ok"], (label, reply)
        assert digest(Path(task["input"])) == sha_before
        raw = reply["sdk_save_archive"]
        assert digest(Path(raw["path"])) == raw["sha256"]
        saved = openpyxl.load_workbook(raw["path"])
        final = openpyxl.load_workbook(task["candidate"])
        business_name = marker if label == "rename" else name
        expected = ["Data", business_name, "Ref"]
        if label == "create":
            expected.append(marker)
        if label == "delete":
            expected.remove(marker)
        # Readers may invent a suffix when raw XML has names differing only
        # by case. Verify the real SDK roster and active index directly.
        with zipfile.ZipFile(raw["path"]) as package:
            workbook = ET.fromstring(package.read("xl/workbook.xml"))
        raw_names = [s.attrib["name"] for s in workbook.findall("./{*}sheets/{*}sheet")]
        raw_active = int(workbook.find("./{*}bookViews/{*}workbookView").attrib["activeTab"])
        assert raw_names == expected + [appended] and raw_active == len(expected)
        assert final.sheetnames == expected and final.active.title == active
        assert reply["output_sanitation"]["removed_sheets"] == [appended]
        if label != "delete":
            assert saved[business_name]["A5"].value == final[business_name]["A5"].value == value
        if label == "create":
            assert final[marker]["A5"].value is None
        if reference:
            assert final["Ref"]["A1"].value == (f"='{business_name}'!A5" if " " in business_name else "=User!A5")
        assert final["Data"]["A1"].value == 7
        outcomes.append(dict(case=label, input_sheet=name, raw=raw_names, raw_active=appended, output=final.sheetnames, active=active))
        saved.close()
        final.close()
    return outcomes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dotnet", required=True)
    parser.add_argument("--dll", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    source = root / "source.xlsx"
    book = openpyxl.Workbook()
    book.active.title = "Data"
    book["Data"]["A1"] = 7
    book["Data"]["A2"] = 3
    book["Data"]["B1"] = "=SUM(A1:A2)"
    book["Data"]["A1"].comment = Comment("synthetic business comment", "test")
    book.create_sheet("Blank")
    book.create_sheet("Calc")["A1"] = "=Data!B1+1"
    book.defined_names.add(DefinedName("BusinessRange", attr_text="Data!$A$1:$A$2"))
    book.active = 2
    book.save(source)
    book.close()
    before = digest(source)
    archive = root / "evidence/raw-sdk-save"
    def plan(kind, **fields):
        return dict(op="edit_plan_operation", operation=dict(operation=kind, **fields))
    tasks = [request(source, root / "calculate.xlsx", "calc"),
             request(source, root / "convert.xlsx", "convert", "Convert"),
             request(source, root / "modify.xlsx", "modify", "Modify", [dict(op="set_value", sheet="Data", cell="C3", value=dict(type="text", value="updated"))]),
             request(source, root / "renamed.xlsx", "rename", "Modify", [plan("rename_sheet", old_name="Calc", new_name="Renamed"), plan("move_sheet", sheet="Renamed", new_index=0)]),
             request(source, root / "deleted.xlsx", "delete", "Modify", [plan("delete_sheet", name="Calc")]),
             request(root / "calculate.xlsx", root / "chain.xlsx", "chain"),
             request(source, root / "calculate.xlsx", "existing")]
    dynamic = root / "dynamic.xlsx"
    book = openpyxl.load_workbook(source)
    book["Data"]["C1"] = '=INDIRECT("A1")'
    book.save(dynamic)
    book.close()
    dynamic_hash = digest(dynamic)
    tasks.append(request(dynamic, root / "unsafe-reference.xlsx", "unsafe-reference"))
    responses = run(args.dotnet, args.dll, tasks, root, archive)
    cases = []
    for task, response in zip(tasks[:6], responses[:6]):
        assert response["ok"], response
        sanitation = response["output_sanitation"]
        assert sanitation["validated"] and sanitation["removed_sheets"] == ["Evaluation Warning"]
        assert sanitation["removed_count"] == 1
        raw = response["sdk_save_archive"]
        assert raw["sha256"] == digest(Path(raw["path"]))
        original = openpyxl.load_workbook(raw["path"])
        assert "Evaluation Warning" in original.sheetnames
        original.close()
        result = openpyxl.load_workbook(task["candidate"])
        if "rename" in task["request_id"]:
            names, active = ["Renamed", "Data", "Blank"], "Renamed"
        elif "delete" in task["request_id"]:
            names, active = ["Data", "Blank"], "Data"
        else:
            names, active = ["Data", "Blank", "Calc"], "Calc"
        assert result.sheetnames == names and result.active.title == active
        assert result["Data"]["B1"].value == "=SUM(A1:A2)"
        assert result["Data"]["A1"].comment.text == "synthetic business comment"
        assert result.defined_names["BusinessRange"].attr_text == "Data!$A$1:$A$2"
        if "modify" in task["request_id"]:
            assert result["Data"]["C3"].value == "updated"
        result.close()
        with zipfile.ZipFile(task["candidate"]) as package:
            assert b"Evaluation Only. Created with Aspose.Cells" not in package.read("xl/sharedStrings.xml")
        if task["method"] in ("Calculate", "Convert"):
            cache = openpyxl.load_workbook(task["candidate"], data_only=True)
            assert cache["Data"]["B1"].value == 10 and cache["Calc"]["A1"].value == 11
            cache.close()
        cases.append(dict(case=task["request_id"], status="PASS", sheets=names, active=active))
    assert all(not row["ok"] for row in responses[6:]), responses[6:]
    assert not (root / "unsafe-reference.xlsx").exists()
    assert digest(source) == before and digest(dynamic) == dynamic_hash
    # A failed post-Save cleanup keeps its raw evidence, without publishing a candidate.
    assert len(list(archive.glob("*.xlsx"))) == 7
    assert not list(root.glob("*.sanitation-*.xlsx"))
    log = [json.loads(line) for line in (root / "stderr.jsonl").read_text("utf-8").splitlines()]
    assert sum(row.get("stage") == "output_sanitation" and row.get("status") == "validated" for row in log) == 6
    assert sum(row.get("stage") == "output_sanitation" and row.get("error_code") == "OUTPUT_SANITATION_FAILED" for row in log) == 1
    collisions = collision_cases(args.dotnet, args.dll, root / "collisions")
    result = dict(status="PASS", success_cases=cases, collision_cases=collisions, protected_rejections=2,
                  source_unchanged=True, raw_sdk_archives_unchanged=True,
                  real_business_acceptance="NOT_TESTED", linux="NOT_TESTED")
    (root / "RESULT.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
