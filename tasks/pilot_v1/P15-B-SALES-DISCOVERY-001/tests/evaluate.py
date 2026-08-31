#!/usr/bin/env python3
"""Deterministic, task-specific judge for P15-B-SALES-DISCOVERY-001."""
from __future__ import annotations

import ast
import csv
import importlib.util
import json
import math
import os
import re
import statistics
import sys
from pathlib import Path

import openpyxl

CRITERIA = [
    {"id": "R001", "description": "A readable workbook contains every requested operating-review sheet.", "weight": 1, "type": "positive", "dimension": "file_usability", "method": "deterministic", "method_params": {}},
    {"id": "R002", "description": "The source-selection record identifies the only approved, complete 2024Q2 schema-1.1 release.", "weight": 3, "type": "positive", "dimension": "source_selection", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R003", "description": "Consolidated_Data contains exactly the six records from the selected release and no mixed-release records.", "weight": 4, "type": "positive", "dimension": "data_completeness", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R004", "description": "All four operating KPIs are formula-driven and equal the independent source replay.", "weight": 3, "type": "positive", "dimension": "metric_correctness", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R005", "description": "Changing the selected REV-100 bookings input by +100 flows to total bookings by exactly +100.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Consolidated_Data!E4"}},
    {"id": "R006", "description": "Changing the selected REV-101 recognized-revenue input by +100 flows to total and South revenue by exactly +100.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Consolidated_Data!F5"}},
    {"id": "R007", "description": "Regional record counts and recognized-revenue closures are formulas linked to the six consolidated records.", "weight": 3, "type": "positive", "dimension": "traceability", "method": "deterministic", "method_params": {}},
    {"id": "R008", "description": "Every release-registry source field is preserved exactly from release_registry.csv.", "weight": 3, "type": "positive", "dimension": "change_locality", "method": "deterministic", "method_params": {"source": "data/input_files/release_registry.csv"}},
    {"id": "R009", "description": "The chosen-source reason, four required-region statuses, and evidence labels remain reviewable.", "weight": 2, "type": "positive", "dimension": "auditability", "method": "deterministic", "method_params": {}},
    {"id": "P001", "description": "Penalty for any invalid selection, mixed source, hardcoded/broken KPI chain, unlinked coverage, or registry alteration.", "weight": -7, "type": "penalty", "dimension": "integrity", "method": "deterministic", "method_params": {}},
]

TASK = {
    "task_id": "P15-B-SALES-DISCOVERY-001",
    "pass_threshold": 0.7,
    "required_sheets": ["Release_Registry", "Selected_Sources", "Consolidated_Data", "KPI_Summary", "Coverage_Checks"],
    "formula_cells": ["KPI_Summary!B4", "KPI_Summary!B5", "KPI_Summary!B6", "KPI_Summary!B7"],
    "coverage_formula_cells": [f"Coverage_Checks!{column}{row}" for row in range(4, 8) for column in ("B", "C")],
    "critical_criteria": ["R002", "R003", "R004", "R005", "R006", "R007", "R008"],
    "criteria": CRITERIA,
}

REF = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?\$?([A-Z]+)\$?(\d+)")
RANGE = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)")

def cell_key(sheet, cell): return f"{sheet}!{cell}"

def col_index(column):
    value = 0
    for char in column: value = value * 26 + ord(char) - 64
    return value

def col_name(index):
    value = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        value = chr(65 + remainder) + value
    return value

def task_root(): return Path(__file__).resolve().parents[1]

def split_context(split):
    contract_path = task_root() / "tests" / ("confirm/contract.json" if split == "confirm" else "private_contract.json")
    contract = json.loads(contract_path.read_text())
    drivers = {
        "bookings_source_cell": "Consolidated_Data!E4",
        "bookings_delta": 100,
        "revenue_source_cell": "Consolidated_Data!F5",
        "revenue_delta": 100,
    }
    drivers.update(contract.get("drivers", {}))
    return {
        "split": split,
        "contract": contract,
        "source_root": task_root() / contract["input_files_dir"],
        "oracle_path": task_root() / contract["oracle"],
        "drivers": drivers,
    }

def load_oracle(context):
    spec = importlib.util.spec_from_file_location(f"sales_oracle_{context['split']}", context["oracle_path"])
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module.recompute()

class FormulaEngine:
    def __init__(self, workbook, overrides=None):
        self.workbook = workbook; self.overrides = overrides or {}; self.memo = {}; self.stack = set()

    def value(self, sheet, cell):
        key = cell_key(sheet, cell)
        if key in self.overrides: return self.overrides[key]
        if key in self.memo: return self.memo[key]
        if key in self.stack: raise ValueError(f"CIRCULAR_REFERENCE:{key}")
        if sheet not in self.workbook.sheetnames: raise ValueError(f"MISSING_SHEET:{sheet}")
        self.stack.add(key)
        raw = self.workbook[sheet][cell].value
        result = self.formula(raw[1:], sheet) if isinstance(raw, str) and raw.startswith("=") else raw
        self.stack.remove(key); self.memo[key] = result
        return result

    def range_values(self, sheet, c1, r1, c2, r2):
        return [self.value(sheet, f"{col_name(c)}{r}") for r in range(int(r1), int(r2) + 1) for c in range(col_index(c1), col_index(c2) + 1)]

    def formula(self, expression, current_sheet):
        expression = expression.replace("^", "**").replace("<>", "!=")
        def outside(pattern, replacement, text):
            pieces = re.split(r'("(?:[^"\\]|\\.)*")', text)
            return "".join(piece if piece.startswith('"') else pattern.sub(replacement, piece) for piece in pieces)
        def range_replace(match):
            sheet = match.group(1) or match.group(2) or current_sheet
            return repr(self.range_values(sheet, match.group(3), match.group(4), match.group(5), match.group(6)))
        expression = outside(RANGE, range_replace, expression)
        def ref_replace(match):
            sheet = match.group(1) or match.group(2) or current_sheet
            return repr(self.value(sheet, f"{match.group(3)}{match.group(4)}"))
        expression = outside(REF, ref_replace, expression)
        expression = re.sub(r"(?<![<>=!])=(?!=)", "==", expression)
        return self.safe_eval(ast.parse(expression, mode="eval").body)

    def safe_eval(self, node):
        if isinstance(node, ast.Constant): return node.value
        if isinstance(node, ast.List): return [self.safe_eval(item) for item in node.elts]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            value = self.safe_eval(node.operand); return -value if isinstance(node.op, ast.USub) else value
        if isinstance(node, ast.BinOp):
            left, right = self.safe_eval(node.left), self.safe_eval(node.right)
            operations = {ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b, ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b, ast.Pow: lambda a, b: a**b}
            for cls, operation in operations.items():
                if isinstance(node.op, cls): return operation(left, right)
            raise ValueError("UNSUPPORTED_OPERATOR")
        if isinstance(node, ast.Compare):
            left = self.safe_eval(node.left)
            for operator, comparator in zip(node.ops, node.comparators):
                right = self.safe_eval(comparator)
                ok = left == right if isinstance(operator, ast.Eq) else left != right if isinstance(operator, ast.NotEq) else left > right if isinstance(operator, ast.Gt) else left >= right if isinstance(operator, ast.GtE) else left < right if isinstance(operator, ast.Lt) else left <= right if isinstance(operator, ast.LtE) else False
                if not ok: return False
                left = right
            return True
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            args = [self.safe_eval(item) for item in node.args]; name = node.func.id.upper()
            values = [value for group in args for value in (group if isinstance(group, list) else [group]) if value is not None]
            if name == "SUM": return sum(values)
            if name == "AVERAGE": return statistics.mean(values)
            if name == "ABS": return abs(args[0])
            if name == "IF": return args[1] if args[0] else args[2]
            if name == "AND": return all(args)
            if name == "COUNTIF":
                if len(args) != 2 or not isinstance(args[0], list): raise ValueError("INVALID_COUNTIF_ARGUMENTS")
                return sum(1 for value in args[0] if value == args[1])
            if name == "SUMIF":
                if len(args) not in {2, 3} or not isinstance(args[0], list): raise ValueError("INVALID_SUMIF_ARGUMENTS")
                sum_values = args[0] if len(args) == 2 else args[2]
                if not isinstance(sum_values, list) or len(sum_values) != len(args[0]): raise ValueError("INVALID_SUMIF_RANGE")
                return sum(value for criterion_value, value in zip(args[0], sum_values) if criterion_value == args[1] and isinstance(value, (int, float)))
        raise ValueError(f"UNSUPPORTED_FORMULA_NODE:{ast.dump(node)}")

def close(actual, expected, tolerance=0.01):
    return isinstance(actual, (int, float)) and math.isfinite(actual) and abs(actual - expected) <= tolerance

def formula_present(workbook, address):
    sheet, cell = address.split("!", 1)
    return sheet in workbook.sheetnames and isinstance(workbook[sheet][cell].value, str) and workbook[sheet][cell].value.startswith("=")

def norm(value): return round(value, 6) if isinstance(value, float) else value

def rows(workbook, sheet, start, end, columns, engine=None):
    values = []
    for row in range(start, end + 1):
        record = tuple(norm(engine.value(sheet, f"{col_name(column)}{row}") if engine else workbook[sheet].cell(row=row, column=column).value) for column in columns)
        if any(value not in (None, "") for value in record): values.append(record)
    return values

def registry_matches_source(workbook, source_root):
    fields = ("release_id", "file", "period", "published", "supersedes", "schema", "coverage")
    with (source_root / "release_registry.csv").open(newline="") as handle:
        expected = [tuple(record[field] for field in fields) for record in csv.DictReader(handle)]
    actual = [tuple("" if workbook["Release_Registry"].cell(row=row, column=column).value is None else str(workbook["Release_Registry"].cell(row=row, column=column).value) for column in range(1, 8)) for row in range(4, 8)]
    return actual == expected

def task_checks(workbook, engine, expected, context):
    checks, failures = {}, []
    selection = tuple(norm(workbook["Selected_Sources"].cell(row=4, column=column).value) for column in range(1, 5))
    expected_selection = tuple(expected["selection"])
    coverage_text = str(selection[3] or "").strip().casefold()
    expected_coverage = str(expected_selection[3] or "").strip().casefold()
    coverage_ok = coverage_text == expected_coverage or (
        expected_coverage == "complete" and ("complete" in coverage_text or "4/4" in coverage_text)
    )
    selection_ok = selection[:3] == expected_selection[:3] and coverage_ok
    checks["R002"] = float(selection_ok)
    if not selection_ok: failures.append("INVALID_RELEASE_SELECTION")

    actual_rows = rows(workbook, "Consolidated_Data", 4, 12, range(1, 8))
    expected_rows = [tuple(row) for row in expected["rows"]]
    data_ok = len(actual_rows) == len(expected_rows) and set(actual_rows) == set(expected_rows)
    checks["R003"] = float(data_ok)
    if not data_ok: failures.append("SELECTED_RECORD_SET_MISMATCH")

    formulas_ok = all(formula_present(workbook, address) for address in TASK["formula_cells"])
    try:
        metrics_ok = close(engine.value("KPI_Summary", "B4"), expected["bookings"]) and close(engine.value("KPI_Summary", "B5"), expected["revenue"]) and close(engine.value("KPI_Summary", "B6"), expected["closed_won"]) and close(engine.value("KPI_Summary", "B7"), expected["south_revenue"])
    except Exception as exc:
        metrics_ok = False; failures.append(f"KPI_EVALUATION_FAILED:{type(exc).__name__}")
    checks["R004"] = float(formulas_ok and metrics_ok)

    try:
        source = context["drivers"]["bookings_source_cell"]
        sheet, cell = source.split("!", 1)
        baseline = workbook[sheet][cell].value
        delta = context["drivers"]["bookings_delta"]
        perturb = FormulaEngine(workbook, {source: baseline + delta})
        bookings_dynamic = close(perturb.value("KPI_Summary", "B4"), expected["bookings"] + delta)
    except Exception as exc:
        bookings_dynamic = False; failures.append(f"BOOKINGS_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R005"] = float(bookings_dynamic)

    try:
        source = context["drivers"]["revenue_source_cell"]
        sheet, cell = source.split("!", 1)
        baseline = workbook[sheet][cell].value
        delta = context["drivers"]["revenue_delta"]
        perturb = FormulaEngine(workbook, {source: baseline + delta})
        revenue_dynamic = close(perturb.value("KPI_Summary", "B5"), expected["revenue"] + delta) and close(perturb.value("KPI_Summary", "B7"), expected["south_revenue"] + delta)
    except Exception as exc:
        revenue_dynamic = False; failures.append(f"REVENUE_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R006"] = float(revenue_dynamic)

    coverage_formula_ok = all(formula_present(workbook, address) for address in TASK["coverage_formula_cells"])
    try:
        expected_coverage = {region: (len([record for record in expected_rows if record[1] == region]), sum(record[5] for record in expected_rows if record[1] == region)) for region in ("North", "South", "East", "West")}
        actual_coverage = {workbook["Coverage_Checks"].cell(row=row, column=1).value: (engine.value("Coverage_Checks", f"B{row}"), engine.value("Coverage_Checks", f"C{row}")) for row in range(4, 8)}
        coverage_ok = actual_coverage == expected_coverage
    except Exception as exc:
        coverage_ok = False; failures.append(f"COVERAGE_CLOSURE_FAILED:{type(exc).__name__}")
    checks["R007"] = float(coverage_formula_ok and coverage_ok)

    protected_ok = registry_matches_source(workbook, context["source_root"])
    checks["R008"] = float(protected_ok)
    if not protected_ok: failures.append("REGISTRY_SOURCE_FIELDS_CHANGED")

    reason = workbook["Selected_Sources"]["E4"].value
    statuses = [workbook["Coverage_Checks"].cell(row=row, column=4).value for row in range(4, 8)]
    evidence = [workbook["KPI_Summary"].cell(row=row, column=4).value for row in range(4, 8)]
    acceptable_statuses = {"required", "pass", "passed", "complete", "ok"}
    statuses_ok = all(str(value or "").strip().casefold() in acceptable_statuses for value in statuses)
    evidence_ok = all(isinstance(value, str) and len(value.strip()) >= 8 for value in evidence)
    checks["R009"] = float(isinstance(reason, str) and len(reason.strip()) >= 20 and statuses_ok and evidence_ok)

    checks["P001"] = 0.0 if all(checks.get(criterion, 0.0) == 1.0 for criterion in TASK["critical_criteria"]) else 1.0
    return checks, failures

def evaluate(candidate, split="dev"):
    criteria = {row["id"]: 0.0 for row in CRITERIA}
    if not candidate.exists() or candidate.stat().st_size == 0: return criteria, ["OUTPUT_MISSING"]
    try: workbook = openpyxl.load_workbook(candidate, data_only=False, read_only=False)
    except Exception as exc: return criteria, [f"MALFORMED_XLSX:{type(exc).__name__}"]
    missing = [sheet for sheet in TASK["required_sheets"] if sheet not in workbook.sheetnames]
    if missing: return criteria, [f"MISSING_SHEET:{sheet}" for sheet in missing]
    criteria["R001"] = 1.0
    failures = []
    try:
        context = split_context(split)
        checks, task_failures = task_checks(workbook, FormulaEngine(workbook), load_oracle(context), context)
        criteria.update(checks); failures.extend(task_failures)
    except Exception as exc: failures.append(f"SEMANTIC_EVALUATION_ERROR:{type(exc).__name__}:{exc}")
    return criteria, sorted(set(failures))

def score(criteria):
    positive = sum(row["weight"] for row in CRITERIA if row["type"] == "positive")
    earned = sum(row["weight"] * criteria.get(row["id"], 0.0) for row in CRITERIA if row["type"] == "positive")
    penalty = sum(abs(row["weight"]) for row in CRITERIA if row["type"] == "penalty" and criteria.get(row["id"], 0.0) > 0)
    normalized = max(0.0, (earned - penalty) / positive)
    if criteria.get("P001", 0.0) > 0 or any(criteria.get(criterion, 0.0) < 1.0 for criterion in TASK["critical_criteria"]): normalized = min(normalized, 0.69)
    return round(normalized, 6)

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
    criteria, failures = evaluate(candidate, split); total = score(criteria)
    payload = {"task_id": TASK["task_id"], "split": split, "candidate": str(candidate), "status": "SCORED", "normalized_score": total, "pass": total >= TASK["pass_threshold"], "criterion_scores": criteria, "failure_codes": failures, "stderr": []}
    log_root = Path(os.environ.get("P15_VERIFIER_LOG_DIR", "/logs/verifier"))
    try:
        log_root.mkdir(parents=True, exist_ok=True); (log_root / "report.json").write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n"); (log_root / "reward.txt").write_text(str(total) + "\n")
    except OSError: pass
    print(json.dumps(payload, indent=2, sort_keys=True))

if __name__ == "__main__": main()
