#!/usr/bin/env python3
"""Deterministic task-specific judge for P15-B-OPS-CLEAN-JOIN-001."""
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
    {"id": "R001", "description": "A readable workbook contains every requested operations sheet.", "weight": 1, "type": "positive", "dimension": "file_usability", "method": "deterministic", "method_params": {}},
    {"id": "R002", "description": "Clean_Data classifies each raw order with the required normalization and duplicate policy.", "weight": 4, "type": "positive", "dimension": "data_cleaning", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R003", "description": "Joined_Data contains exactly the valid one-to-one joins with correct product, region, units, unit cost, and extended cost.", "weight": 4, "type": "positive", "dimension": "join_correctness", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R004", "description": "Valid-order units and product costs flow through formula-linked clean, join, and extended-cost cells.", "weight": 3, "type": "positive", "dimension": "lineage", "method": "deterministic", "method_params": {}},
    {"id": "R005", "description": "Changing Raw_Orders ORD-100 units from 10 to 11 changes its extended cost and total by exactly $150.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Raw_Orders!D4"}},
    {"id": "R006", "description": "Changing Product_Master P-200 unit cost from $420 to $500 changes ORD-101 and the total by exactly $400.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Product_Master!D5"}},
    {"id": "R007", "description": "The four task-specific exceptions are complete and summary counts/totals close to their underlying records through formulas.", "weight": 3, "type": "positive", "dimension": "exception_integrity", "method": "deterministic", "method_params": {}},
    {"id": "R008", "description": "Every raw-order, product-master, and location-master source field is preserved exactly.", "weight": 3, "type": "positive", "dimension": "change_locality", "method": "deterministic", "method_params": {}},
    {"id": "R009", "description": "Checks link to the calculated summary and report the valid count, exception count, and total cost.", "weight": 2, "type": "positive", "dimension": "auditability", "method": "deterministic", "method_params": {}},
    {"id": "P001", "description": "Penalty for a cleaning, join, source-lineage, exception-closure, or protected-source failure.", "weight": -7, "type": "penalty", "dimension": "integrity", "method": "deterministic", "method_params": {}},
]

TASK = {
    "task_id": "P15-B-OPS-CLEAN-JOIN-001",
    "pass_threshold": 0.7,
    "required_sheets": ["Raw_Orders", "Product_Master", "Location_Master", "Clean_Data", "Joined_Data", "Exceptions", "Summary", "Checks"],
    "formula_cells": ["Clean_Data!D4", "Clean_Data!D5", "Clean_Data!D6", "Joined_Data!D4", "Joined_Data!D5", "Joined_Data!D6", "Joined_Data!E4", "Joined_Data!E5", "Joined_Data!E6", "Joined_Data!F4", "Joined_Data!F5", "Joined_Data!F6", "Summary!B4", "Summary!B5", "Summary!B6", "Summary!B7"],
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
        "units_order_id": "ORD-100",
        "raw_units_source_cell": "Raw_Orders!D4",
        "units_new_value": 11,
        "units_delta_cost": 150,
        "cost_order_id": "ORD-101",
        "product_cost_source_cell": "Product_Master!D5",
        "product_cost_new_value": 500,
        "product_cost_delta": 400,
        "summary_region": "Northeast",
    }
    drivers.update(contract.get("drivers", {}))
    return {"split": split, "contract": contract, "source_root": task_root() / contract["input_files_dir"], "oracle_path": task_root() / contract["oracle"], "drivers": drivers}


def load_oracle(context):
    global oracle_module
    spec = importlib.util.spec_from_file_location(f"task_oracle_{context['split']}", context["oracle_path"])
    oracle_module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(oracle_module)
    return oracle_module.recompute()


def excel_equal(left, right):
    if left in (None, "") and right in (None, ""):
        return True
    if isinstance(left, str) and isinstance(right, str):
        return left.casefold() == right.casefold()
    return left == right


def excel_criteria_match(value, criteria):
    operator = "="
    expected = criteria
    if isinstance(criteria, str):
        match = re.match(r"^(<=|>=|<>|!=|=|<|>)(.*)$", criteria)
        if match:
            operator, expected = match.groups()
        stripped = expected.strip()
        try:
            expected = float(stripped)
        except (TypeError, ValueError):
            expected = stripped
    if operator in {"=", "=="}:
        return excel_equal(value, expected)
    if operator in {"<>", "!="}:
        return not excel_equal(value, expected)
    try:
        if operator == "<": return value < expected
        if operator == "<=": return value <= expected
        if operator == ">": return value > expected
        if operator == ">=": return value >= expected
    except TypeError:
        return False
    raise ValueError(f"UNSUPPORTED_CRITERIA_OPERATOR:{operator}")


def as_range(value, function_name):
    if not isinstance(value, list):
        raise ValueError(f"{function_name}_REQUIRES_RANGE")
    return value


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
            operations = {ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b, ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b, ast.Pow: lambda a, b: a ** b}
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
            if name == "SUMIFS":
                if len(args) < 3 or len(args) % 2 == 0:
                    raise ValueError("SUMIFS_ARGUMENTS")
                sum_range = as_range(args[0], name)
                criteria_pairs = [
                    (as_range(args[index], name), args[index + 1])
                    for index in range(1, len(args), 2)
                ]
                if any(len(criteria_range) != len(sum_range) for criteria_range, _ in criteria_pairs):
                    raise ValueError("SUMIFS_RANGE_LENGTH")
                return sum(
                    value
                    for row, value in enumerate(sum_range)
                    if isinstance(value, (int, float))
                    and all(excel_criteria_match(criteria_range[row], criterion) for criteria_range, criterion in criteria_pairs)
                )
            if name == "COUNTA":
                return sum(value not in (None, "") for value in values)
            if name == "XLOOKUP":
                if len(args) != 3:
                    raise ValueError("XLOOKUP_ARGUMENTS")
                lookup_range = as_range(args[1], name)
                return_range = as_range(args[2], name)
                if len(lookup_range) != len(return_range):
                    raise ValueError("XLOOKUP_RANGE_LENGTH")
                for index, value in enumerate(lookup_range):
                    if excel_equal(value, args[0]):
                        return return_range[index]
                raise ValueError("XLOOKUP_NOT_FOUND")
            if name == "TRIM":
                if len(args) != 1:
                    raise ValueError("TRIM_ARGUMENTS")
                return re.sub(r" +", " ", str(args[0]).strip())
            if name == "UPPER":
                if len(args) != 1:
                    raise ValueError("UPPER_ARGUMENTS")
                return str(args[0]).upper()
            if name == "AVERAGE": return statistics.mean(values)
            if name == "ABS": return abs(args[0])
            if name == "IF": return args[1] if args[0] else args[2]
            if name == "AND": return all(args)
        raise ValueError(f"UNSUPPORTED_FORMULA_NODE:{ast.dump(node)}")


def close(actual, expected, tolerance=0.01):
    return isinstance(actual, (int, float)) and math.isfinite(actual) and abs(actual - expected) <= tolerance


def formula_present(workbook, address):
    sheet, cell = address.split("!", 1)
    return sheet in workbook.sheetnames and isinstance(workbook[sheet][cell].value, str) and workbook[sheet][cell].value.startswith("=")


def norm(value):
    if isinstance(value, float): return round(value, 6)
    return value


def category_token(value):
    if isinstance(value, bool):
        return "yes" if value else "no"
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().casefold()).strip("_")


def normalize_yes_no(value):
    token = category_token(value)
    return {
        "1": "yes", "true": "yes", "y": "yes", "yes": "yes",
        "0": "no", "false": "no", "n": "no", "no": "no",
    }.get(token, token)


def normalize_disposition(value):
    token = category_token(value)
    return {
        "valid": "join",
        "joined": "join",
        "include": "join",
        "included_in_joined_data": "join",
        "exception_duplicate_excluded": "exception_exact_duplicate",
        "exception_non_numeric_units": "exception_invalid_units",
    }.get(token, token)


def normalize_clean_record(record):
    values = list(record)
    for index in (4, 5, 6):
        values[index] = normalize_yes_no(values[index])
    values[7] = normalize_disposition(values[7])
    return tuple(values)


def rows(workbook, sheet, start, end, columns):
    values = []
    for row in range(start, end + 1):
        record = tuple(norm(workbook[sheet].cell(row=row, column=column).value) for column in columns)
        if any(value not in (None, "") for value in record): values.append(record)
    return values


def find_row(workbook, sheet, identifier, start=4, end=20):
    for row in range(start, end + 1):
        if workbook[sheet].cell(row=row, column=1).value == identifier:
            return row
    raise ValueError(f"MISSING_RECORD:{sheet}:{identifier}")


def source_tables_match(workbook, root):
    with (root / "raw_orders.csv").open(newline="") as handle:
        expected_raw = []
        for record in csv.DictReader(handle):
            units = int(record["units"]) if record["units"].isdigit() else record["units"]
            expected_raw.append((record["order_id"], record["product_code"], record["location_code"], units, record["order_date"]))
    with (root / "product_master.csv").open(newline="") as handle:
        expected_products = [(record["product_code"], record["product"], record["category"], int(record["unit_cost"])) for record in csv.DictReader(handle)]
    with (root / "location_master.csv").open(newline="") as handle:
        expected_locations = [(record["location_code"], record["region"], record["country"]) for record in csv.DictReader(handle)]
    actual_raw = rows(workbook, "Raw_Orders", 4, 10, range(1, 6))
    actual_products = rows(workbook, "Product_Master", 4, 6, range(1, 5))
    actual_locations = rows(workbook, "Location_Master", 4, 6, range(1, 4))
    return actual_raw == expected_raw and actual_products == expected_products and actual_locations == expected_locations


def task_checks(workbook, engine, oracle, context):
    checks, failures = {}, []
    expected = oracle
    clean_actual = []
    try:
        for row in range(4, 14):
            if workbook["Clean_Data"].cell(row=row, column=1).value in (None, ""):
                continue
            clean_actual.append(tuple(norm(engine.value("Clean_Data", f"{col_name(column)}{row}")) for column in range(1, 9)))
        clean_ok = {
            record[0]: normalize_clean_record(record)[1:]
            for record in clean_actual
        } == {
            record[0]: normalize_clean_record(record)[1:]
            for record in expected["clean"]
        }
    except Exception as exc:
        clean_ok = False
        failures.append(f"CLEAN_EVALUATION_FAILED:{type(exc).__name__}")
    checks["R002"] = 1.0 if clean_ok else 0.0
    if not clean_ok: failures.append("CLEAN_RECORDS_MISMATCH")

    joined_actual = []
    try:
        for row in range(4, 14):
            if workbook["Joined_Data"].cell(row=row, column=1).value in (None, ""): continue
            joined_actual.append(tuple(norm(engine.value("Joined_Data", f"{col_name(column)}{row}")) for column in range(1, 7)))
        joined_ok = {record[0]: record[1:] for record in joined_actual} == {record[0]: tuple(record[1:]) for record in expected["joined"]}
    except Exception as exc:
        joined_ok = False; failures.append(f"JOIN_EVALUATION_FAILED:{type(exc).__name__}")
    checks["R003"] = 1.0 if joined_ok else 0.0

    formulas_ok = all(formula_present(workbook, address) for address in TASK["formula_cells"])
    checks["R004"] = 1.0 if formulas_ok else 0.0
    try:
        order_id = context["drivers"]["units_order_id"]
        row = find_row(workbook, "Joined_Data", order_id)
        expected_record = next(record for record in expected["joined"] if record[0] == order_id)
        new_units = context["drivers"]["units_new_value"]
        perturb = FormulaEngine(workbook, {context["drivers"]["raw_units_source_cell"]: new_units})
        units_dynamic = close(perturb.value("Joined_Data", f"F{row}"), new_units * expected_record[4]) and close(perturb.value("Summary", "B5"), expected["total"] + context["drivers"]["units_delta_cost"])
    except Exception as exc:
        units_dynamic = False; failures.append(f"RAW_UNITS_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R005"] = 1.0 if units_dynamic else 0.0

    try:
        order_id = context["drivers"]["cost_order_id"]
        row = find_row(workbook, "Joined_Data", order_id)
        expected_record = next(record for record in expected["joined"] if record[0] == order_id)
        new_cost = context["drivers"]["product_cost_new_value"]
        perturb = FormulaEngine(workbook, {context["drivers"]["product_cost_source_cell"]: new_cost})
        cost_dynamic = close(perturb.value("Joined_Data", f"F{row}"), expected_record[3] * new_cost) and close(perturb.value("Summary", "B5"), expected["total"] + context["drivers"]["product_cost_delta"])
    except Exception as exc:
        cost_dynamic = False; failures.append(f"MASTER_COST_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R006"] = 1.0 if cost_dynamic else 0.0

    exceptions_actual = {(row[0], category_token(row[1])) for row in rows(workbook, "Exceptions", 4, 12, range(1, 3))}
    exceptions_ok = exceptions_actual == {(row[0], category_token(row[1])) for row in expected["exceptions"]}
    try:
        summary_region_total = expected.get("summary_region_total", expected.get("northeast"))
        summary_ok = close(engine.value("Summary", "B4"), len(expected["joined"])) and close(engine.value("Summary", "B5"), expected["total"]) and close(engine.value("Summary", "B6"), len(expected["exceptions"])) and close(engine.value("Summary", "B7"), summary_region_total)
    except Exception as exc:
        summary_ok = False; failures.append(f"SUMMARY_CLOSURE_FAILED:{type(exc).__name__}")
    checks["R007"] = 1.0 if exceptions_ok and summary_ok else 0.0

    protected_ok = source_tables_match(workbook, context["source_root"])
    checks["R008"] = 1.0 if protected_ok else 0.0
    if not protected_ok: failures.append("PROTECTED_SOURCE_TABLE_CHANGED")

    check_formulas = all(formula_present(workbook, f"Checks!B{row}") for row in range(4, 7))
    try:
        checks_ok = close(engine.value("Checks", "B4"), len(expected["joined"])) and close(engine.value("Checks", "B5"), len(expected["exceptions"])) and close(engine.value("Checks", "B6"), expected["total"])
    except Exception as exc:
        checks_ok = False; failures.append(f"CHECK_LINKAGE_FAILED:{type(exc).__name__}")
    checks["R009"] = 1.0 if check_formulas and checks_ok else 0.0

    checks["P001"] = 0.0 if all(checks.get(criterion, 0.0) == 1.0 for criterion in TASK["critical_criteria"]) else 1.0
    return checks, failures


def evaluate(candidate, split="dev"):
    criteria = {row["id"]: 0.0 for row in TASK["criteria"]}
    failures = []
    if not candidate.exists() or candidate.stat().st_size == 0: return criteria, ["OUTPUT_MISSING"]
    try: workbook = openpyxl.load_workbook(candidate, data_only=False, read_only=False)
    except Exception as exc: return criteria, [f"MALFORMED_XLSX:{type(exc).__name__}"]
    for sheet in TASK["required_sheets"]:
        if sheet not in workbook.sheetnames: failures.append(f"MISSING_SHEET:{sheet}")
    if failures: return criteria, failures
    criteria["R001"] = 1.0
    try:
        context = split_context(split)
        checks, task_failures = task_checks(workbook, FormulaEngine(workbook), load_oracle(context), context)
        criteria.update(checks); failures.extend(task_failures)
    except Exception as exc:
        failures.append(f"SEMANTIC_EVALUATION_ERROR:{type(exc).__name__}:{exc}")
    return criteria, sorted(set(failures))


def score(criteria):
    positive = sum(row["weight"] for row in TASK["criteria"] if row["type"] == "positive")
    earned = sum(row["weight"] * criteria.get(row["id"], 0.0) for row in TASK["criteria"] if row["type"] == "positive")
    penalty = sum(abs(row["weight"]) for row in TASK["criteria"] if row["type"] == "penalty" and criteria.get(row["id"], 0.0) > 0)
    normalized = max(0.0, (earned - penalty) / positive)
    if criteria.get("P001", 0.0) > 0 or any(criteria.get(criterion, 0.0) < 1.0 for criterion in TASK["critical_criteria"]):
        normalized = min(normalized, 0.69)
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
