#!/usr/bin/env python3
"""Object-level OOXML judge for the native PivotTable/PivotChart task."""
from __future__ import annotations

import csv
import importlib.util
import json
import math
import os
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"m": MAIN, "c": CHART}

CRITERIA = [
    {"id": "R001", "description": "The workbook opens and contains every required reporting sheet.", "weight": 1, "type": "positive", "dimension": "file_usability", "method": "deterministic", "method_params": {}},
    {"id": "R002", "description": "Program_Data preserves all eight source events in an Excel Table named ProgramEventsTable covering exactly A3:F11.", "weight": 4, "type": "positive", "dimension": "source_table", "method": "ooxml", "method_params": {}},
    {"id": "R003", "description": "A genuine PivotCache binds exactly to ProgramEventsTable/Program_Data!A3:F11 with the six source fields in their exact identities and indexes.", "weight": 4, "type": "positive", "dimension": "pivot_source_binding", "method": "ooxml", "method_params": {}},
    {"id": "R004", "description": "The native PivotTable uses Region field index 1 as rows and Program field index 2 as columns.", "weight": 4, "type": "positive", "dimension": "pivot_axes", "method": "ooxml", "method_params": {}},
    {"id": "R005", "description": "Quarter field index 3 is a page filter with the selected item resolving exactly to 2024Q2.", "weight": 4, "type": "positive", "dimension": "pivot_filter", "method": "ooxml", "method_params": {}},
    {"id": "R006", "description": "Data fields are native SUM measures over Participants index 4 and Spend index 5.", "weight": 4, "type": "positive", "dimension": "pivot_measures", "method": "ooxml", "method_params": {}},
    {"id": "R007", "description": "The PivotCache is refreshable and configured to refresh when the workbook opens.", "weight": 3, "type": "positive", "dimension": "refresh_semantics", "method": "ooxml", "method_params": {}},
    {"id": "R008", "description": "GETPIVOTDATA-linked KPI cells have Excel-calculated cached values 132 participants, $396,000 spend, and 50 North Outreach participants.", "weight": 4, "type": "positive", "dimension": "aggregation_correctness", "method": "ooxml", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R009", "description": "A clustered-column chart part contains a pivotSource relationship naming ProgramDeliveryPivot; a regular range chart does not qualify.", "weight": 4, "type": "positive", "dimension": "pivot_chart_binding", "method": "ooxml", "method_params": {}},
    {"id": "P001", "description": "Critical integrity penalty for any fake/formula Pivot substitute, stale source, wrong field/filter/measure, absent refresh, wrong KPI, or non-Pivot chart.", "weight": -16, "type": "penalty", "dimension": "integrity", "method": "ooxml", "method_params": {}},
]

TASK = {
    "task_id": "P15-B-PUBLIC-PIVOT-001",
    "pass_threshold": 0.7,
    "required_sheets": ["Program_Data", "Pivot_Specification", "KPI_Summary", "Pivot_Chart_Requirement", "Checks"],
    "critical_criteria": ["R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009"],
    "criteria": CRITERIA,
}


def task_root():
    return Path(__file__).resolve().parents[1]


def split_context(split):
    contract_path = task_root() / "tests" / ("confirm/contract.json" if split == "confirm" else "private_contract.json")
    contract = json.loads(contract_path.read_text())
    return {
        "split": split,
        "contract": contract,
        "source_root": task_root() / contract["input_files_dir"],
        "oracle_path": task_root() / contract["oracle"],
    }


def load_oracle(context):
    spec = importlib.util.spec_from_file_location(f"pivot_oracle_{context['split']}", context["oracle_path"])
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module.recompute()


def truthy(value):
    return str(value).lower() in {"1", "true", "yes"}


def close(actual, expected, tolerance=0.01):
    return isinstance(actual, (int, float)) and math.isfinite(actual) and abs(actual - expected) <= tolerance


def xml(archive, name):
    return ET.fromstring(archive.read(name))


def source_records_match(workbook, source_root):
    with (source_root / "program_events.csv").open(newline="") as handle:
        expected = [(record["event_id"], record["region"], record["program"], record["quarter"], int(record["participants"]), int(record["spend"])) for record in csv.DictReader(handle)]
    actual = []
    for row in range(4, 12):
        actual.append(tuple(workbook["Program_Data"].cell(row=row, column=column).value for column in range(1, 7)))
    return actual == expected


def table_object_ok(archive):
    for name in archive.namelist():
        if not name.startswith("xl/tables/table") or not name.endswith(".xml"):
            continue
        root = xml(archive, name)
        if root.get("name") == "ProgramEventsTable" and root.get("displayName") == "ProgramEventsTable" and root.get("ref") == "A3:F11":
            columns = [node.get("name") for node in root.findall("./m:tableColumns/m:tableColumn", NS)]
            if columns == ["Event ID", "Region", "Program", "Quarter", "Participants", "Spend"]:
                return True
    return False


def cache_details(archive):
    candidates = sorted(name for name in archive.namelist() if name.startswith("xl/pivotCache/pivotCacheDefinition") and name.endswith(".xml"))
    for name in candidates:
        root = xml(archive, name)
        worksheet_source = root.find("./m:cacheSource/m:worksheetSource", NS)
        if worksheet_source is None:
            continue
        named_source = worksheet_source.get("name") == "ProgramEventsTable"
        ranged_source = worksheet_source.get("sheet") == "Program_Data" and worksheet_source.get("ref") == "A3:F11"
        fields = root.findall("./m:cacheFields/m:cacheField", NS)
        field_names = [field.get("name") for field in fields]
        if named_source or ranged_source:
            return {
                "name": name,
                "root": root,
                "source_ok": True,
                "fields": fields,
                "field_names": field_names,
                "refresh_on_load": truthy(root.get("refreshOnLoad")),
                "enable_refresh": root.get("enableRefresh") is None or truthy(root.get("enableRefresh")),
            }
    return None


def selected_page_value(table_root, cache_fields, field_index, page_field):
    pivot_fields = table_root.findall("./m:pivotFields/m:pivotField", NS)
    if field_index >= len(pivot_fields) or field_index >= len(cache_fields):
        return None
    shared = []
    shared_items = cache_fields[field_index].find("./m:sharedItems", NS)
    if shared_items is not None:
        for item in list(shared_items):
            shared.append(item.get("v"))
    try:
        page_item_index = int(page_field.get("item", "0"))
    except ValueError:
        return None
    pivot_items = pivot_fields[field_index].findall("./m:items/m:item", NS)
    cache_index = page_item_index
    if 0 <= page_item_index < len(pivot_items):
        mapped = pivot_items[page_item_index].get("x")
        if mapped is not None:
            try:
                cache_index = int(mapped)
            except ValueError:
                return None
    if 0 <= cache_index < len(shared):
        return shared[cache_index]
    return None


def pivot_table_details(archive, cache):
    candidates = sorted(name for name in archive.namelist() if name.startswith("xl/pivotTables/pivotTable") and name.endswith(".xml"))
    for name in candidates:
        root = xml(archive, name)
        if root.get("name") != "ProgramDeliveryPivot":
            continue
        row_indexes = [int(node.get("x")) for node in root.findall("./m:rowFields/m:field", NS) if node.get("x") is not None]
        column_indexes = [int(node.get("x")) for node in root.findall("./m:colFields/m:field", NS) if node.get("x") is not None]
        page_fields = root.findall("./m:pageFields/m:pageField", NS)
        data_fields = root.findall("./m:dataFields/m:dataField", NS)
        page = next((node for node in page_fields if node.get("fld") == "3"), None)
        selected = selected_page_value(root, cache["fields"], 3, page) if page is not None else None
        measures = {}
        for node in data_fields:
            try:
                field_index = int(node.get("fld"))
            except (TypeError, ValueError):
                continue
            # OOXML defines `sum` as the default dataField subtotal, so native
            # Excel may legally omit the attribute for a SUM measure.
            measures[field_index] = {"subtotal": (node.get("subtotal") or "sum").lower(), "name": node.get("name") or ""}
        return {
            "name": name,
            "root": root,
            "row_indexes": row_indexes,
            "column_indexes": column_indexes,
            "page_fields": page_fields,
            "selected_quarter": selected,
            "measures": measures,
        }
    return None


def pivot_chart_ok(archive):
    for name in archive.namelist():
        if not name.startswith("xl/charts/chart") or not name.endswith(".xml"):
            continue
        root = xml(archive, name)
        pivot_source = root.find("./c:pivotSource", NS)
        if pivot_source is None:
            continue
        pivot_name = pivot_source.find("./c:name", NS)
        bar_chart = root.find("./c:chart/c:plotArea/c:barChart", NS)
        bar_direction = bar_chart.find("./c:barDir", NS) if bar_chart is not None else None
        grouping = bar_chart.find("./c:grouping", NS) if bar_chart is not None else None
        if (
            pivot_name is not None
            and "ProgramDeliveryPivot" in (pivot_name.text or "")
            and bar_direction is not None
            and bar_direction.get("val") == "col"
            and grouping is not None
            and grouping.get("val") == "clustered"
        ):
            return True
    return False


def kpi_values_ok(candidate, oracle):
    formulas = openpyxl.load_workbook(candidate, data_only=False, read_only=False)
    cached = openpyxl.load_workbook(candidate, data_only=True, read_only=False)
    expected = [oracle["q2_participants"], oracle["q2_spend"], oracle.get("focus_participants", oracle.get("north_outreach"))]
    for row, value in zip(range(4, 7), expected):
        formula = formulas["KPI_Summary"][f"B{row}"].value
        cached_value = cached["KPI_Summary"][f"B{row}"].value
        normalized_formula = formula.upper().replace("'", "") if isinstance(formula, str) else ""
        if (
            not isinstance(formula, str)
            or not formula.startswith("=")
            or "GETPIVOTDATA" not in normalized_formula
            or "PIVOT_REPORT!$A$3" not in normalized_formula
        ):
            return False
        if not close(cached_value, value):
            return False
    return True


def score(criteria):
    positive = sum(row["weight"] for row in CRITERIA if row["type"] == "positive")
    earned = sum(row["weight"] * criteria.get(row["id"], 0.0) for row in CRITERIA if row["type"] == "positive")
    penalty = sum(abs(row["weight"]) for row in CRITERIA if row["type"] == "penalty" and criteria.get(row["id"], 0.0) > 0)
    normalized = max(0.0, (earned - penalty) / positive)
    if criteria.get("P001", 0.0) > 0 or any(criteria.get(criterion, 0.0) < 1.0 for criterion in TASK["critical_criteria"]):
        normalized = min(normalized, 0.69)
    return round(normalized, 6)


def evaluate(candidate, split="dev"):
    criteria = {row["id"]: 0.0 for row in CRITERIA}
    failures = []
    if not candidate.exists() or candidate.stat().st_size == 0:
        return criteria, ["OUTPUT_MISSING"], "TASK_INVALID"
    try:
        workbook = openpyxl.load_workbook(candidate, data_only=False, read_only=False)
    except Exception as exc:
        return criteria, [f"MALFORMED_XLSX:{type(exc).__name__}"], "TASK_INVALID"
    missing = [sheet for sheet in TASK["required_sheets"] if sheet not in workbook.sheetnames]
    if missing:
        return criteria, [f"MISSING_SHEET:{sheet}" for sheet in missing], "TASK_INVALID"
    criteria["R001"] = 1.0
    context = split_context(split)
    try:
        with zipfile.ZipFile(candidate) as archive:
            table_ok = source_records_match(workbook, context["source_root"]) and table_object_ok(archive)
            criteria["R002"] = float(table_ok)
            if not table_ok:
                failures.append("SOURCE_TABLE_OR_RECORD_MISMATCH")

            cache = cache_details(archive)
            expected_fields = ["Event ID", "Region", "Program", "Quarter", "Participants", "Spend"]
            cache_ok = cache is not None and cache["field_names"] == expected_fields
            criteria["R003"] = float(cache_ok)
            if cache is None:
                failures.append("MISSING_NATIVE_PIVOT_CACHE")
                criteria["P001"] = 1.0
                return criteria, sorted(set(failures)), "TASK_INVALID"
            if not cache_ok:
                failures.append("PIVOT_CACHE_SOURCE_OR_FIELD_IDENTITY_MISMATCH")

            table = pivot_table_details(archive, cache)
            if table is None:
                failures.append("MISSING_NATIVE_PROGRAM_DELIVERY_PIVOT")
                criteria["P001"] = 1.0
                return criteria, sorted(set(failures)), "TASK_INVALID"

            axes_ok = table["row_indexes"] == [1] and 2 in table["column_indexes"] and all(index in (2, -2) for index in table["column_indexes"])
            criteria["R004"] = float(axes_ok)
            if not axes_ok:
                failures.append("PIVOT_AXIS_FIELD_INDEX_MISMATCH")

            filter_ok = len(table["page_fields"]) == 1 and table["page_fields"][0].get("fld") == "3" and table["selected_quarter"] == "2024Q2"
            criteria["R005"] = float(filter_ok)
            if not filter_ok:
                failures.append("QUARTER_FILTER_NOT_EXACTLY_2024Q2")

            measures = table["measures"]
            measure_ok = set(measures) == {4, 5} and all(measures[index]["subtotal"] == "sum" for index in (4, 5)) and "participant" in measures[4]["name"].lower() and "spend" in measures[5]["name"].lower()
            criteria["R006"] = float(measure_ok)
            if not measure_ok:
                failures.append("SUM_MEASURE_FIELD_OR_FUNCTION_MISMATCH")

            refresh_ok = cache["refresh_on_load"] and cache["enable_refresh"]
            criteria["R007"] = float(refresh_ok)
            if not refresh_ok:
                failures.append("PIVOT_CACHE_REFRESH_ON_OPEN_MISSING")

            oracle = load_oracle(context)
            kpi_ok = kpi_values_ok(candidate, oracle)
            criteria["R008"] = float(kpi_ok)
            if not kpi_ok:
                failures.append("GETPIVOTDATA_OR_EXCEL_CACHED_ORACLE_MISMATCH")

            chart_ok = pivot_chart_ok(archive)
            criteria["R009"] = float(chart_ok)
            if not chart_ok:
                failures.append("PIVOTCHART_BINDING_OR_CLUSTERED_COLUMN_MISSING")

            criteria["P001"] = 0.0 if all(criteria.get(criterion, 0.0) == 1.0 for criterion in TASK["critical_criteria"]) else 1.0
            status = "NATIVE_OBJECT_CHECKED" if criteria["P001"] == 0.0 else "TASK_INVALID"
            return criteria, sorted(set(failures)), status
    except zipfile.BadZipFile:
        return criteria, ["MALFORMED_XLSX:BadZipFile"], "TASK_INVALID"
    except Exception as exc:
        return criteria, [f"OOXML_EVALUATION_ERROR:{type(exc).__name__}:{exc}"], "TASK_INVALID"


def parse_cli():
    split = os.environ.get("P15_EVAL_SPLIT", "dev").strip().lower()
    candidate = None
    arguments = iter(sys.argv[1:])
    for argument in arguments:
        if argument == "--split": split = next(arguments, "")
        elif argument.startswith("--split="): split = argument.split("=", 1)[1]
        elif candidate is None: candidate = Path(argument)
        else: raise ValueError(f"UNEXPECTED_ARGUMENT:{argument}")
    if split not in {"dev", "confirm"}: raise ValueError(f"INVALID_SPLIT:{split}")
    return candidate or Path("/app/output/answer.xlsx"), split


def main():
    candidate, split = parse_cli()
    criteria, failures, status = evaluate(candidate, split)
    total = score(criteria)
    payload = {
        "task_id": TASK["task_id"],
        "split": split,
        "candidate": str(candidate),
        "status": status,
        "blocker": None if status == "NATIVE_OBJECT_CHECKED" else "NATIVE_OBJECT_VALIDATION_FAILED",
        "normalized_score": total,
        "pass": total >= TASK["pass_threshold"] and status == "NATIVE_OBJECT_CHECKED",
        "criterion_scores": criteria,
        "failure_codes": failures,
        "stderr": [],
    }
    root = Path(os.environ.get("P15_VERIFIER_LOG_DIR", "/logs/verifier"))
    try:
        root.mkdir(parents=True, exist_ok=True)
        (root / "report.json").write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
        (root / "reward.txt").write_text(str(total) + "\n")
    except OSError:
        pass
    print(json.dumps(payload, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
