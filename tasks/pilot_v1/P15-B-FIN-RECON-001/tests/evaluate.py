#!/usr/bin/env python3
"""Deterministic task-specific judge for P15-B-FIN-RECON-001."""
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
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl

CRITERIA = [
    {"id": "R001", "description": "A readable workbook contains every requested reconciliation sheet.", "weight": 1, "type": "positive", "dimension": "file_usability", "method": "deterministic", "method_params": {}},
    {"id": "R002", "description": "All ten ledger/subledger records use the correct transaction-date FX rate and formula-linked USD normalization.", "weight": 4, "type": "positive", "dimension": "fx_normalization", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R003", "description": "The four matched invoices use correct source amounts, the approved adjustment, differences, and the $1 tolerance status.", "weight": 4, "type": "positive", "dimension": "reconciliation_correctness", "method": "deterministic", "method_params": {"tolerance_usd": 1.0}},
    {"id": "R004", "description": "Variance_Bridge and Final_Reconciliation are formula-linked and close to the independent matched/unmatched replay.", "weight": 3, "type": "positive", "dimension": "reconciliation_closure", "method": "deterministic", "method_params": {}},
    {"id": "R005", "description": "Changing the 2024-06-30 EUR source rate from 1.07 to 1.08 flows to both INV-101 records and both bridge totals.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "FX_Rates!C6"}},
    {"id": "R006", "description": "Changing approved ADJ-105 from -$50 to -$40 flows to the adjusted amount, difference, and OUT_OF_TOLERANCE status.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Adjustment_Evidence!C4"}},
    {"id": "R007", "description": "Ledger-only and subledger-only exceptions remain distinct and their amounts/counts close to Final_Reconciliation formulas.", "weight": 3, "type": "positive", "dimension": "exception_integrity", "method": "deterministic", "method_params": {}},
    {"id": "R008", "description": "Every ledger, subledger, FX-rate, and adjustment-evidence source field is preserved exactly.", "weight": 3, "type": "positive", "dimension": "change_locality", "method": "deterministic", "method_params": {}},
    {"id": "R009", "description": "Checks expose zero matched difference, two exceptions, and zero final closure residual through formulas.", "weight": 2, "type": "positive", "dimension": "auditability", "method": "deterministic", "method_params": {}},
    {"id": "P001", "description": "Penalty for wrong FX lineage, false/tolerance-breaking matches, omitted adjustment, dropped exception, broken closure, or source alteration.", "weight": -7, "type": "penalty", "dimension": "integrity", "method": "deterministic", "method_params": {}},
]

TASK = {
    "task_id": "P15-B-FIN-RECON-001",
    "pass_threshold": 0.7,
    "required_sheets": ["Ledger", "Subledger", "FX_Rates", "Adjustment_Evidence", "Normalized_Records", "Matched_Items", "Unmatched_Items", "Variance_Bridge", "Final_Reconciliation", "Checks"],
    "formula_cells": [
        *[f"Normalized_Records!{column}{row}" for row in range(4, 14) for column in ("F", "G", "H")],
        *[f"Matched_Items!{column}{row}" for row in range(4, 8) for column in ("B", "C", "D", "E", "F", "G")],
        "Unmatched_Items!C4", "Unmatched_Items!C5",
        *[f"Variance_Bridge!B{row}" for row in range(4, 9)],
        *[f"Final_Reconciliation!B{row}" for row in range(4, 8)],
        "Checks!B4", "Checks!B5", "Checks!B6",
    ],
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
        "fx_source_cell": "FX_Rates!C6",
        "fx_invoice": "INV-101",
        "fx_new_value": 1.08,
        "adjustment_source_cell": "Adjustment_Evidence!C4",
        "adjustment_invoice": "INV-105",
        "adjustment_new_value": -40,
        "adjusted_subledger": 260,
        "adjustment_difference": -10,
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
        try:
            raw = self.workbook[sheet][cell].value
            result = self.formula(raw[1:], sheet) if isinstance(raw, str) and raw.startswith("=") else raw
            self.memo[key] = result
            return result
        finally:
            self.stack.discard(key)

    @staticmethod
    def formula_scalar(value):
        if isinstance(value, datetime): return value.date().isoformat()
        if isinstance(value, date): return value.isoformat()
        return value

    def range_values(self, sheet, c1, r1, c2, r2):
        return [self.formula_scalar(self.value(sheet, f"{col_name(c)}{r}")) for r in range(int(r1), int(r2) + 1) for c in range(col_index(c1), col_index(c2) + 1)]

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
            return repr(self.formula_scalar(self.value(sheet, f"{match.group(3)}{match.group(4)}")))
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
            if name == "SUM": return sum(value for value in values if isinstance(value, (int, float)))
            if name == "SUMIF":
                if len(args) not in {2, 3}:
                    raise ValueError("SUMIF_ARGUMENTS")
                criteria_range = as_range(args[0], name)
                sum_range = as_range(args[2], name) if len(args) == 3 else criteria_range
                if len(criteria_range) != len(sum_range):
                    raise ValueError("SUMIF_RANGE_LENGTH")
                return sum(
                    value
                    for value, candidate in zip(sum_range, criteria_range)
                    if isinstance(value, (int, float)) and excel_criteria_match(candidate, args[1])
                )
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
            if name == "COUNTIF":
                if len(args) != 2:
                    raise ValueError("COUNTIF_ARGUMENTS")
                criteria_range = as_range(args[0], name)
                return sum(excel_criteria_match(value, args[1]) for value in criteria_range)
            if name == "COUNTA": return sum(value not in (None, "") for value in values)
            if name == "AVERAGE": return statistics.mean(values)
            if name == "MAX": return max(values)
            if name == "MIN": return min(values)
            if name == "ABS": return abs(args[0])
            if name == "IF": return args[1] if args[0] else args[2]
            if name == "AND": return all(args)
            if name == "ROUND": return round(args[0], int(args[1]))
            if name == "LEFT": return str(args[0])[: int(args[1])]
            if name == "RIGHT": return str(args[0])[-int(args[1]) :]
            if name == "MID": return str(args[0])[int(args[1]) - 1 : int(args[1]) - 1 + int(args[2])]
            if name == "DATE": return f"{int(args[0]):04d}-{int(args[1]):02d}-{int(args[2]):02d}"
            if name == "TEXT":
                if len(args) != 2:
                    raise ValueError("TEXT_ARGUMENTS")
                return str(args[0])
        raise ValueError(f"UNSUPPORTED_FORMULA_NODE:{ast.dump(node)}")


def close(actual, expected, tolerance=0.01):
    return isinstance(actual, (int, float)) and math.isfinite(actual) and abs(actual - expected) <= tolerance


def formula_present(workbook, address):
    sheet, cell = address.split("!", 1)
    return sheet in workbook.sheetnames and isinstance(workbook[sheet][cell].value, str) and workbook[sheet][cell].value.startswith("=")


def norm(value):
    if isinstance(value, float): return round(value, 6)
    return value


def normalize_reconciliation_status(value):
    if not isinstance(value, str):
        return value
    text = unicodedata.normalize("NFKC", value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    if text in {"matched", "reconciled", "within tolerance", "within 1 tolerance"}:
        return "MATCHED"
    if text in {
        "out of tolerance",
        "outside tolerance",
        "out of 1 tolerance",
        "outside 1 tolerance",
        "variance exceeds tolerance",
    }:
        return "OUT_OF_TOLERANCE"
    return value


def rows(workbook, sheet, start, end, columns):
    values = []
    for row in range(start, end + 1):
        record = tuple(norm(workbook[sheet].cell(row=row, column=column).value) for column in columns)
        if any(value not in (None, "") for value in record): values.append(record)
    return values


def find_row(workbook, sheet, identifier, id_column=1, start=4, end=30):
    for row in range(start, end + 1):
        if workbook[sheet].cell(row=row, column=id_column).value == identifier:
            return row
    raise ValueError(f"MISSING_RECORD:{sheet}:{identifier}")


def source_tables_match(workbook, root):
    def records(filename):
        with (root / filename).open(newline="") as handle:
            return list(csv.DictReader(handle))
    expected_ledger = [(r["record_id"], r["date"], r["invoice"], r["currency"], float(r["original_amount"])) for r in records("ledger.csv")]
    expected_subledger = [(r["record_id"], r["date"], r["invoice"], r["currency"], float(r["original_amount"])) for r in records("subledger.csv")]
    expected_fx = [(r["date"], r["currency"], float(r["usd_per_unit"])) for r in records("fx_rates.csv")]
    expected_adjustments = [(r["adjustment_id"], r["invoice"], float(r["amount_usd"]), r["reason"]) for r in records("adjustments.csv")]
    actual_ledger = rows(workbook, "Ledger", 4, 8, range(1, 6))
    actual_subledger = rows(workbook, "Subledger", 4, 8, range(1, 6))
    actual_fx = rows(workbook, "FX_Rates", 4, 6, range(1, 4))
    actual_adjustments = rows(workbook, "Adjustment_Evidence", 4, 10, range(1, 5))
    return actual_ledger == expected_ledger and actual_subledger == expected_subledger and actual_fx == expected_fx and actual_adjustments == expected_adjustments


def task_checks(workbook, engine, oracle, context):
    checks, failures = {}, []
    expected = oracle
    try:
        normalized = []
        for row in range(4, 20):
            if workbook["Normalized_Records"].cell(row=row, column=1).value in (None, ""):
                continue
            normalized.append(tuple(norm(engine.value("Normalized_Records", f"{col_name(column)}{row}")) for column in range(1, 9)))
        normalized_ok = {(record[0], record[1]): record[2:] for record in normalized} == {(record[0], record[1]): tuple(norm(value) for value in record[2:]) for record in expected["normalized"]}
    except Exception as exc:
        normalized_ok = False; failures.append(f"FX_REPLAY_FAILED:{type(exc).__name__}")
    normalized_formulas = all(formula_present(workbook, f"Normalized_Records!{column}{row}") for row in range(4, 14) for column in ("F", "G", "H"))
    checks["R002"] = 1.0 if normalized_ok and normalized_formulas else 0.0

    try:
        matched = []
        tolerance_ok = True
        for row in range(4, 14):
            if workbook["Matched_Items"].cell(row=row, column=1).value in (None, ""):
                continue
            record = tuple(norm(engine.value("Matched_Items", f"{col_name(column)}{row}")) for column in range(1, 8))
            matched.append(record[:6])
            tolerance_ok = tolerance_ok and close(record[5], 0.0, 1.0) and normalize_reconciliation_status(record[6]) == "MATCHED"
        matched_ok = {record[0]: record[1:] for record in matched} == {record[0]: tuple(norm(value) for value in record[1:]) for record in expected["matched"]}
    except Exception as exc:
        matched_ok = tolerance_ok = False; failures.append(f"MATCH_REPLAY_FAILED:{type(exc).__name__}")
    matched_formulas = all(formula_present(workbook, f"Matched_Items!{column}{row}") for row in range(4, 8) for column in ("B", "C", "D", "E", "F", "G"))
    checks["R003"] = 1.0 if matched_ok and tolerance_ok and matched_formulas else 0.0

    bridge_formulas = all(formula_present(workbook, f"Variance_Bridge!B{row}") for row in range(4, 9)) and all(formula_present(workbook, f"Final_Reconciliation!B{row}") for row in range(4, 8))
    try:
        bridge_ok = close(engine.value("Variance_Bridge", "B4"), expected["ledger_total"]) and close(engine.value("Variance_Bridge", "B5"), expected["matched_total"]) and close(engine.value("Variance_Bridge", "B6"), expected["ledger_only"]) and close(engine.value("Variance_Bridge", "B7"), expected["subledger_only"]) and close(engine.value("Variance_Bridge", "B8"), expected["investigation"])
        final_ok = close(engine.value("Final_Reconciliation", "B4"), len(expected["matched"])) and close(engine.value("Final_Reconciliation", "B5"), expected["ledger_only"]) and close(engine.value("Final_Reconciliation", "B6"), expected["subledger_only"]) and close(engine.value("Final_Reconciliation", "B7"), expected["investigation"])
    except Exception as exc:
        bridge_ok = final_ok = False; failures.append(f"RECONCILIATION_CLOSURE_FAILED:{type(exc).__name__}")
    checks["R004"] = 1.0 if bridge_formulas and bridge_ok and final_ok else 0.0

    try:
        invoice = context["drivers"]["fx_invoice"]
        ledger_row = next(row for row in range(4, 20) if workbook["Normalized_Records"].cell(row=row, column=1).value == "Ledger" and workbook["Normalized_Records"].cell(row=row, column=3).value == invoice)
        subledger_row = next(row for row in range(4, 20) if workbook["Normalized_Records"].cell(row=row, column=1).value == "Subledger" and workbook["Normalized_Records"].cell(row=row, column=3).value == invoice)
        matched_row = find_row(workbook, "Matched_Items", invoice)
        source = context["drivers"]["fx_source_cell"]
        source_sheet, source_cell = source.split("!", 1)
        baseline_rate = workbook[source_sheet][source_cell].value
        new_rate = context["drivers"]["fx_new_value"]
        original = engine.value("Normalized_Records", f"F{ledger_row}")
        expected_usd = original * new_rate
        delta = original * (new_rate - baseline_rate)
        perturb = FormulaEngine(workbook, {source: new_rate})
        fx_dynamic = close(perturb.value("Normalized_Records", f"H{ledger_row}"), expected_usd) and close(perturb.value("Normalized_Records", f"H{subledger_row}"), expected_usd) and close(perturb.value("Matched_Items", f"F{matched_row}"), 0) and close(perturb.value("Variance_Bridge", "B4"), expected["ledger_total"] + delta) and close(perturb.value("Variance_Bridge", "B5"), expected["matched_total"] + delta)
    except Exception as exc:
        fx_dynamic = False; failures.append(f"FX_SOURCE_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R005"] = 1.0 if fx_dynamic else 0.0

    try:
        row = find_row(workbook, "Matched_Items", context["drivers"]["adjustment_invoice"])
        new_adjustment = context["drivers"]["adjustment_new_value"]
        perturb = FormulaEngine(workbook, {context["drivers"]["adjustment_source_cell"]: new_adjustment})
        adjustment_dynamic = close(perturb.value("Matched_Items", f"D{row}"), new_adjustment) and close(perturb.value("Matched_Items", f"E{row}"), context["drivers"]["adjusted_subledger"]) and close(perturb.value("Matched_Items", f"F{row}"), context["drivers"]["adjustment_difference"]) and normalize_reconciliation_status(perturb.value("Matched_Items", f"G{row}")) == "OUT_OF_TOLERANCE"
    except Exception as exc:
        adjustment_dynamic = False; failures.append(f"ADJUSTMENT_SOURCE_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R006"] = 1.0 if adjustment_dynamic else 0.0

    unmatched = {
        (
            workbook["Unmatched_Items"].cell(row=row, column=1).value,
            workbook["Unmatched_Items"].cell(row=row, column=2).value,
            norm(engine.value("Unmatched_Items", f"C{row}")),
        )
        for row in range(4, 9)
        if workbook["Unmatched_Items"].cell(row=row, column=1).value not in (None, "")
    }
    expected_unmatched = {(row[0], row[1], norm(row[2])) for row in expected["unmatched"]}
    exceptions_ok = unmatched == expected_unmatched
    exception_formulas = all(formula_present(workbook, address) for address in ("Unmatched_Items!C4", "Unmatched_Items!C5", "Final_Reconciliation!B4", "Final_Reconciliation!B5", "Final_Reconciliation!B6", "Final_Reconciliation!B7"))
    checks["R007"] = 1.0 if exceptions_ok and exception_formulas and final_ok else 0.0

    protected_ok = source_tables_match(workbook, context["source_root"])
    checks["R008"] = 1.0 if protected_ok else 0.0
    if not protected_ok: failures.append("PROTECTED_FINANCE_SOURCE_CHANGED")

    check_formulas = all(formula_present(workbook, f"Checks!B{row}") for row in range(4, 7))
    try:
        check_values_ok = close(engine.value("Checks", "B4"), 0) and close(engine.value("Checks", "B5"), len(expected["unmatched"])) and close(engine.value("Checks", "B6"), 0)
    except Exception as exc:
        check_values_ok = False; failures.append(f"CHECK_LINKAGE_FAILED:{type(exc).__name__}")
    checks["R009"] = 1.0 if check_formulas and check_values_ok else 0.0

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
