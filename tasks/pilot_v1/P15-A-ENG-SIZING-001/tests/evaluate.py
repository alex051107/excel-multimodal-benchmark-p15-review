#!/usr/bin/env python3
"""Deterministic semantic judge for P15-A-ENG-SIZING-001."""
from __future__ import annotations

import ast
import importlib.util
import json
import math
import os
import re
import statistics
import sys
from pathlib import Path

import openpyxl

TASK = {
  "task_id": "P15-A-ENG-SIZING-001",
  "pass_threshold": 0.7,
  "required_sheets": [
    "Inputs",
    "Unit_Conversions",
    "Calculations",
    "Equipment_Selection",
    "Checks"
  ],
  "protected": {
    "Inputs!B4": 42,
    "Inputs!B7": 150,
    "Equipment_Selection!A4": "P-120",
    "Equipment_Selection!B4": 45,
    "Equipment_Selection!C4": 25,
    "Equipment_Selection!D4": 11,
    "Equipment_Selection!A5": "P-180",
    "Equipment_Selection!B5": 58,
    "Equipment_Selection!C5": 36,
    "Equipment_Selection!D5": 18.5,
    "Equipment_Selection!A6": "P-250",
    "Equipment_Selection!B6": 80,
    "Equipment_Selection!C6": 48,
    "Equipment_Selection!D6": 30
  },
  "formula_cells": [
    "Unit_Conversions!C4",
    "Calculations!B7",
    "Calculations!B11",
    "Equipment_Selection!E5",
    "Equipment_Selection!B7"
  ],
  "criteria": [
    {"id":"R001","description":"The workbook opens and contains every required sizing sheet.","weight":1,"type":"positive","dimension":"file_usability","method":"deterministic","method_params":{"implemented_check":"required_sheets"}},
    {"id":"R002","description":"Flow, diameter, and cross-sectional-area conversions agree with the independent SI replay.","weight":3,"type":"positive","dimension":"unit_conversion","method":"deterministic","method_params":{"implemented_check":"si_conversions","oracle":"split_contract"}},
    {"id":"R003","description":"Velocity, velocity head, friction head, and total dynamic head agree with the hydraulic replay.","weight":3,"type":"positive","dimension":"hydraulic_head","method":"deterministic","method_params":{"implemented_check":"hydraulic_head_chain","oracle":"split_contract"}},
    {"id":"R004","description":"Hydraulic and shaft power agree with the independent replay.","weight":2,"type":"positive","dimension":"power_sizing","method":"deterministic","method_params":{"implemented_check":"power_chain","oracle":"split_contract"}},
    {"id":"R005","description":"Safety-adjusted minimum flow and motor rating agree with the independent replay.","weight":2,"type":"positive","dimension":"safety_sizing","method":"deterministic","method_params":{"implemented_check":"safety_adjusted_requirements","oracle":"split_contract"}},
    {"id":"R006","description":"Each catalog row's eligibility status matches all three sizing constraints.","weight":3,"type":"positive","dimension":"catalog_eligibility","method":"deterministic","method_params":{"implemented_check":"catalog_eligibility","oracle":"split_contract"}},
    {"id":"R007","description":"The selected pump is the first eligible catalog option.","weight":3,"type":"positive","dimension":"constraint_selection","method":"deterministic","method_params":{"implemented_check":"selected_pump","oracle":"split_contract"}},
    {"id":"R008","description":"Hydraulic sizing and pump selection respond correctly to the declared design-flow perturbation.","weight":2,"type":"positive","dimension":"native_recalculation","method":"deterministic","method_params":{"implemented_check":"flow_perturbation","split_weights":{"dev":2,"confirm":3}}},
    {"id":"R009","description":"Every declared engineering output carries the required unit label.","weight":1,"type":"positive","dimension":"unit_labels","method":"deterministic","method_params":{"implemented_check":"unit_labels"}},
    {"id":"R010","description":"All declared workbook check rows evaluate to PASS.","weight":1,"type":"positive","dimension":"model_checks","method":"deterministic","method_params":{"implemented_check":"check_rows"}},
    {"id":"R011","description":"The declared design inputs and catalog capacity values remain unchanged.","weight":2,"type":"positive","dimension":"change_locality","method":"deterministic","method_params":{"implemented_check":"input_catalog_protection"}},
    {"id":"R012","description":"The hydraulic head, motor requirement, eligibility, and selection respond correctly to a distinct internal-diameter perturbation while design flow and catalog values remain invariant.","weight":1,"type":"positive","dimension":"diameter_recalculation","method":"deterministic","method_params":{"implemented_check":"diameter_perturbation","applies_to":["dev"]}},
    {"id":"P001","description":"Penalty for altered inputs, missing safety mechanics, invalid eligibility, or hardcoded selection.","weight":-6,"type":"penalty","dimension":"sizing_integrity","method":"deterministic","method_params":{"implemented_check":"sizing_integrity_penalty"}}
  ]
}
REF = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?\$?([A-Z]+)\$?(\d+)")
RANGE = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)")


def cell_key(sheet, cell):
    return f"{sheet}!{cell}"


def col_index(column):
    value = 0
    for char in column:
        value = value * 26 + ord(char) - 64
    return value


def col_name(index):
    text = ""
    while index:
        index, rem = divmod(index - 1, 26)
        text = chr(65 + rem) + text
    return text


TASK_ROOT = Path(__file__).resolve().parents[1]


def requested_split():
    split = os.environ.get("P15_EVAL_SPLIT", "dev").strip().lower()
    for index, argument in enumerate(sys.argv[1:]):
        if argument == "--split" and index + 2 <= len(sys.argv) - 1:
            split = sys.argv[index + 2].strip().lower()
        elif argument.startswith("--split="):
            split = argument.split("=", 1)[1].strip().lower()
    if split not in {"dev", "confirm"}:
        raise ValueError(f"unsupported split: {split}")
    return split


ACTIVE_SPLIT = requested_split()


def load_contract():
    relative = "tests/confirm/contract.json" if ACTIVE_SPLIT == "confirm" else "tests/private_contract.json"
    return json.loads((TASK_ROOT / relative).read_text())


def load_oracle(contract):
    oracle_path = TASK_ROOT / contract["oracle"]
    source_path = TASK_ROOT / "metadata/oracle_recompute.py"
    if ACTIVE_SPLIT == "dev" and source_path.exists() and source_path.read_bytes() != oracle_path.read_bytes():
        raise RuntimeError("DEV_ORACLE_COPY_OUT_OF_SYNC")
    spec = importlib.util.spec_from_file_location(f"task_oracle_{ACTIVE_SPLIT}", oracle_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module.recompute(), module


def _betacf(a, b, x):
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d, h = 1.0, 1.0 - qab * x / qap, 0.0
    if abs(d) < 3e-14: d = 3e-14
    d = 1.0 / d
    h = d
    for m in range(1, 201):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < 3e-14: d = 3e-14
        c = 1.0 + aa / c
        if abs(c) < 3e-14: c = 3e-14
        d = 1.0 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < 3e-14: d = 3e-14
        c = 1.0 + aa / c
        if abs(c) < 3e-14: c = 3e-14
        d = 1.0 / d; delta = d * c; h *= delta
        if abs(delta - 1.0) < 3e-12: break
    return h


def _betai(a, b, x):
    if x <= 0: return 0.0
    if x >= 1: return 1.0
    factor = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x))
    return factor * _betacf(a, b, x) / a if x < (a + 1) / (a + b + 2) else 1 - factor * _betacf(b, a, 1 - x) / b


def student_t_two_tail(value, df):
    x = df / (df + value * value)
    return _betai(df / 2, 0.5, x)


def student_t_inv_two_tail(alpha, df):
    lo, hi = 0.0, 20.0
    for _ in range(90):
        mid = (lo + hi) / 2
        if student_t_two_tail(mid, df) > alpha: lo = mid
        else: hi = mid
    return (lo + hi) / 2


class ExcelFormulaError(ValueError):
    """An Excel value error that IFERROR may legitimately handle."""


def excel_equal(left, right):
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=1e-12, abs_tol=1e-12)
    if isinstance(left, str) or isinstance(right, str):
        return " ".join(str(left or "").strip().casefold().split()) == " ".join(str(right or "").strip().casefold().split())
    return left == right


def excel_criteria_match(value, criterion):
    if not isinstance(criterion, str):
        return excel_equal(value, criterion)
    match = re.fullmatch(r"\s*(<=|>=|<>|=|<|>)?\s*(.*?)\s*", criterion)
    operator, operand_text = match.groups() if match else (None, criterion)
    try:
        candidate = float(value)
        operand = float(operand_text)
    except (TypeError, ValueError):
        candidate = " ".join(str(value or "").strip().casefold().split())
        operand = " ".join(operand_text.strip().casefold().split())
    if operator in (None, "="):
        return candidate == operand
    if operator == "<>":
        return candidate != operand
    if operator == "<":
        return candidate < operand
    if operator == "<=":
        return candidate <= operand
    if operator == ">":
        return candidate > operand
    if operator == ">=":
        return candidate >= operand
    return False


class FormulaEngine:
    def __init__(self, workbook, overrides=None):
        self.workbook = workbook
        self.overrides = overrides or {}
        self.memo = {}
        self.stack = set()

    def value(self, sheet, cell):
        key = cell_key(sheet, cell)
        if key in self.overrides:
            return self.overrides[key]
        if key in self.memo:
            return self.memo[key]
        if key in self.stack:
            raise ValueError(f"CIRCULAR_REFERENCE:{key}")
        if sheet not in self.workbook.sheetnames:
            raise ValueError(f"MISSING_SHEET:{sheet}")
        self.stack.add(key)
        raw = self.workbook[sheet][cell].value
        if isinstance(raw, str) and raw.startswith("="):
            result = self.formula(raw[1:], sheet)
        else:
            result = raw
        self.stack.remove(key)
        self.memo[key] = result
        return result

    def range_values(self, sheet, c1, r1, c2, r2):
        values = []
        for r in range(int(r1), int(r2) + 1):
            for c in range(col_index(c1), col_index(c2) + 1):
                values.append(self.value(sheet, f"{col_name(c)}{r}"))
        return values

    def formula(self, expression, current_sheet):
        expression = expression.replace("^", "**").replace("<>", "!=")
        expression = re.sub(r"\bSTDEV\.S\b", "STDEV_S", expression, flags=re.I)
        expression = re.sub(r"\bT\.DIST\.2T\b", "T_DIST_2T", expression, flags=re.I)
        expression = re.sub(r"\bT\.DIST\.RT\b", "T_DIST_RT", expression, flags=re.I)
        expression = re.sub(r"\bT\.INV\.2T\b", "T_INV_2T", expression, flags=re.I)
        expression = re.sub(r"\bPI\(\)", "PI()", expression, flags=re.I)

        def replace_range(match):
            sheet = match.group(1) or match.group(2) or current_sheet
            return repr(self.range_values(sheet, match.group(3), match.group(4), match.group(5), match.group(6)))

        def sub_outside_strings(pattern, replacement, value):
            parts = re.split(r'("(?:[^"\\\\]|\\\\.)*")', value)
            return "".join(part if part.startswith('"') else pattern.sub(replacement, part) for part in parts)

        expression = sub_outside_strings(RANGE, replace_range, expression)

        def replace_ref(match):
            sheet = match.group(1) or match.group(2) or current_sheet
            return repr(self.value(sheet, f"{match.group(3)}{match.group(4)}"))

        expression = sub_outside_strings(REF, replace_ref, expression)
        expression = re.sub(r"(?<![<>=!])=(?!=)", "==", expression)
        tree = ast.parse(expression, mode="eval")
        return self.safe_eval(tree.body)

    def safe_eval(self, node):
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name) and node.id.upper() in {"TRUE", "FALSE"}:
            return node.id.upper() == "TRUE"
        if isinstance(node, ast.List):
            return [self.safe_eval(v) for v in node.elts]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            value = self.safe_eval(node.operand)
            return -value if isinstance(node.op, ast.USub) else value
        if isinstance(node, ast.BinOp):
            left, right = self.safe_eval(node.left), self.safe_eval(node.right)
            operators = {ast.Add: lambda a,b:a+b, ast.Sub: lambda a,b:a-b, ast.Mult: lambda a,b:a*b, ast.Div: lambda a,b:a/b, ast.Pow: lambda a,b:a**b}
            for cls, fun in operators.items():
                if isinstance(node.op, cls): return fun(left, right)
            raise ValueError("UNSUPPORTED_OPERATOR")
        if isinstance(node, ast.Compare):
            left = self.safe_eval(node.left)
            for op, comp in zip(node.ops, node.comparators):
                right = self.safe_eval(comp)
                good = (left == right if isinstance(op, ast.Eq) else left != right if isinstance(op, ast.NotEq) else left >= right if isinstance(op, ast.GtE) else left <= right if isinstance(op, ast.LtE) else left > right if isinstance(op, ast.Gt) else left < right if isinstance(op, ast.Lt) else False)
                if not good: return False
                left = right
            return True
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            name = node.func.id.upper()
            if name == "IF":
                if len(node.args) not in (2, 3):
                    raise ValueError("IF_REQUIRES_TWO_OR_THREE_ARGUMENTS")
                if self.safe_eval(node.args[0]):
                    return self.safe_eval(node.args[1])
                return self.safe_eval(node.args[2]) if len(node.args) == 3 else False
            if name == "IFERROR":
                if len(node.args) != 2:
                    raise ValueError("IFERROR_REQUIRES_TWO_ARGUMENTS")
                try:
                    return self.safe_eval(node.args[0])
                except (ExcelFormulaError, ArithmeticError):
                    return self.safe_eval(node.args[1])
            args = [self.safe_eval(a) for a in node.args]
            values = [v for group in args for v in (group if isinstance(group, list) else [group]) if v is not None]
            numeric_values = [v for v in values if isinstance(v, (int, float))]
            if name == "SUM": return sum(numeric_values)
            if name == "AVERAGE": return statistics.mean(numeric_values)
            if name == "COUNT": return len(numeric_values)
            if name == "STDEV_S": return statistics.stdev(numeric_values)
            if name == "SQRT": return math.sqrt(args[0])
            if name == "ABS": return abs(args[0])
            if name == "ROUND": return round(args[0], int(args[1]))
            if name == "PI": return math.pi
            if name == "AND": return all(args)
            if name == "OR": return any(args)
            if name == "NOT": return not args[0]
            if name == "POWER": return args[0] ** args[1]
            if name == "PRODUCT": return math.prod(numeric_values)
            if name == "MAX": return max(numeric_values)
            if name == "MIN": return min(numeric_values)
            if name == "MATCH":
                if len(args) not in {2, 3}:
                    raise ValueError("MATCH_REQUIRES_TWO_OR_THREE_ARGUMENTS")
                candidates = args[1] if isinstance(args[1], list) else [args[1]]
                match_type = args[2] if len(args) == 3 else 1
                if match_type != 0:
                    raise ValueError("ONLY_EXACT_MATCH_MODE_IS_SUPPORTED")
                for index, candidate in enumerate(candidates, start=1):
                    if excel_equal(candidate, args[0]):
                        return index
                raise ExcelFormulaError("MATCH_NOT_FOUND")
            if name == "INDEX":
                if len(args) not in {2, 3}:
                    raise ValueError("INDEX_REQUIRES_TWO_OR_THREE_ARGUMENTS")
                candidates = args[0] if isinstance(args[0], list) else [args[0]]
                if len(args) == 3 and int(args[2]) != 1:
                    raise ValueError("ONLY_SINGLE_COLUMN_INDEX_IS_SUPPORTED")
                row = int(args[1])
                if row < 1 or row > len(candidates):
                    raise ExcelFormulaError("INDEX_OUT_OF_RANGE")
                return candidates[row - 1]
            if name == "COUNTIF":
                if len(args) != 2:
                    raise ValueError("COUNTIF_REQUIRES_TWO_ARGUMENTS")
                candidates = args[0] if isinstance(args[0], list) else [args[0]]
                return sum(1 for candidate in candidates if excel_criteria_match(candidate, args[1]))
            if name == "NA": return None
            if name == "T_DIST_2T":
                return student_t_two_tail(args[0], args[1])
            if name == "T_DIST_RT":
                return student_t_two_tail(args[0], args[1]) / 2
            if name == "T_INV_2T":
                return student_t_inv_two_tail(args[0], args[1])
        raise ValueError(f"UNSUPPORTED_FORMULA_NODE:{ast.dump(node)}")


def close(actual, expected, tolerance):
    return isinstance(actual, (int, float)) and math.isfinite(actual) and abs(actual - expected) <= tolerance


def formula_present(workbook, address):
    sheet, cell = address.split("!", 1)
    return sheet in workbook.sheetnames and isinstance(workbook[sheet][cell].value, str) and workbook[sheet][cell].value.startswith("=")


def normalized_text(value):
    return " ".join(str(value or "").strip().casefold().split())


def eligibility_status(value):
    aliases = {
        "eligible": "eligible",
        "meets criteria": "eligible",
        "meets all constraints": "eligible",
        "pass": "eligible",
        "ineligible": "ineligible",
        "not eligible": "ineligible",
        "does not meet criteria": "ineligible",
        "does not meet all constraints": "ineligible",
        "fail": "ineligible",
    }
    return aliases.get(normalized_text(value))


def normalized_unit(value):
    return normalized_text(value).replace("\u00b2", "2").replace("\u00b3", "3").replace("^", "").replace(" ", "")


def task_checks(workbook, engine, oracle, contract, oracle_module):
    checks, failures = {}, []
    def validate_group(criterion_id, expected):
        ok = True
        for address, target in expected.items():
            if not formula_present(workbook, address):
                ok = False
                failures.append(f"FORMULA_REQUIRED:{address}")
            sheet, cell = address.split("!", 1)
            try:
                observed = engine.value(sheet, cell)
            except Exception as exc:
                observed = None
                failures.append(f"SIZING_EVALUATION_FAILED:{address}:{type(exc).__name__}")
            if not close(observed, target, 0.01):
                ok = False
                failures.append(f"SIZING_VALUE_MISMATCH:{address}")
        checks[criterion_id] = 1.0 if ok else 0.0
        return ok

    conversion_ok = validate_group("R002", {
        "Unit_Conversions!C4": oracle["flow_m3s"],
        "Unit_Conversions!C5": oracle["diameter_m"],
        "Unit_Conversions!C6": oracle["area_m2"],
    })
    head_ok = validate_group("R003", {
        "Calculations!B4": oracle["velocity"], "Calculations!B5": oracle["velocity_head"],
        "Calculations!B6": oracle["friction_head"], "Calculations!B7": oracle["tdh"],
    })
    power_ok = validate_group("R004", {
        "Calculations!B8": oracle["hydraulic_power"], "Calculations!B9": oracle["shaft_power"],
    })
    safety_ok = validate_group("R005", {
        "Calculations!B10": oracle["minimum_flow"], "Calculations!B11": oracle["minimum_motor"],
    })

    eligibility = [engine.value("Equipment_Selection", f"E{row}") for row in range(4, 7)]
    eligibility_formula_ok = all(formula_present(workbook, f"Equipment_Selection!E{row}") for row in range(4, 7))
    eligibility_ok = eligibility_formula_ok and [eligibility_status(value) for value in eligibility] == [eligibility_status(value) for value in oracle["eligibility"]]
    checks["R006"] = 1.0 if eligibility_ok else 0.0
    selected = engine.value("Equipment_Selection", "B7")
    selection_ok = formula_present(workbook, "Equipment_Selection!B7") and selected == oracle["selected_pump"]
    checks["R007"] = 1.0 if selection_ok else 0.0

    perturbations = {row["name"]: row for row in contract["perturbations"]}

    def perturbation_matches(case):
        replay = FormulaEngine(workbook, case["overrides"])
        for address, target in {**case["expected"], **case["invariants"]}.items():
            sheet, cell = address.split("!", 1)
            observed = replay.value(sheet, cell)
            if isinstance(target, (int, float)):
                if not close(observed, target, 0.01):
                    return False
            elif observed != target:
                return False
        if "expected_eligibility" in case:
            observed = [replay.value("Equipment_Selection", f"E{row}") for row in range(4, 7)]
            if [eligibility_status(value) for value in observed] != [eligibility_status(value) for value in case["expected_eligibility"]]:
                return False
        return True

    flow_case = perturbations["design_flow"]
    try:
        if "expected" in flow_case:
            dynamic_ok = perturbation_matches(flow_case)
        else:
            flow_value = flow_case["overrides"]["Inputs!B4"]
            dynamic_engine = FormulaEngine(workbook, flow_case["overrides"])
            expected_dynamic = oracle_module.recompute(flow_lps=flow_value)
            dynamic_ok = close(dynamic_engine.value("Calculations", "B10"), expected_dynamic["minimum_flow"], 0.01) and dynamic_engine.value("Equipment_Selection", "B7") == expected_dynamic["selected_pump"]
    except Exception as exc:
        dynamic_ok = False
        failures.append(f"FLOW_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R008"] = 1.0 if dynamic_ok else 0.0

    diameter_ok = True
    checks["R012"] = 0.0
    if ACTIVE_SPLIT == "dev":
        try:
            diameter_ok = perturbation_matches(perturbations["internal_diameter"])
        except Exception as exc:
            diameter_ok = False
            failures.append(f"DIAMETER_PERTURBATION_FAILED:{type(exc).__name__}")
        checks["R012"] = 1.0 if diameter_ok else 0.0
    expected_units = {
        "Unit_Conversions!D4": "m³/s", "Unit_Conversions!D5": "m", "Unit_Conversions!D6": "m²",
        "Calculations!C4": "m/s", "Calculations!C5": "m", "Calculations!C6": "m",
        "Calculations!C7": "m", "Calculations!C8": "kW", "Calculations!C9": "kW",
        "Calculations!C10": "L/s", "Calculations!C11": "kW",
    }
    units_ok = all(normalized_unit(workbook[sheet][cell].value) == normalized_unit(value) for key, value in expected_units.items() for sheet, cell in [key.split("!")])
    checks["R009"] = 1.0 if units_ok else 0.0
    check_ok = all(normalized_text(engine.value("Checks", f"D{row}")) == "pass" for row in (4, 5, 6))
    checks["R010"] = 1.0 if check_ok else 0.0
    protected_ok = all(workbook[sheet][cell].value == value for key, value in contract["protected"].items() for sheet, cell in [key.split("!", 1)])
    checks["R011"] = 1.0 if protected_ok else 0.0
    checks["P001"] = 0.0 if all((conversion_ok, head_ok, power_ok, safety_ok, eligibility_ok, selection_ok, dynamic_ok, diameter_ok, protected_ok)) else 1.0
    return checks, failures


def evaluate(candidate):
    criteria = {row["id"]: 0.0 for row in TASK["criteria"]}
    failures = []
    if not candidate.exists() or candidate.stat().st_size == 0:
        return criteria, ["OUTPUT_MISSING"]
    try:
        workbook = openpyxl.load_workbook(candidate, data_only=False, read_only=False)
    except Exception as exc:
        return criteria, [f"MALFORMED_XLSX:{type(exc).__name__}"]
    contract = load_contract()
    for sheet in contract["required_sheets"]:
        if sheet not in workbook.sheetnames:
            failures.append(f"MISSING_SHEET:{sheet}")
    if failures:
        return criteria, failures
    criteria["R001"] = 1.0
    try:
        oracle, oracle_module = load_oracle(contract)
        checks, task_failures = task_checks(workbook, FormulaEngine(workbook), oracle, contract, oracle_module)
        criteria.update(checks)
        failures.extend(task_failures)
    except Exception as exc:
        failures.append(f"SEMANTIC_EVALUATION_ERROR:{type(exc).__name__}:{exc}")
    return criteria, sorted(set(failures))


def score(criteria):
    active = [row for row in TASK["criteria"] if ACTIVE_SPLIT in row.get("method_params", {}).get("applies_to", [ACTIVE_SPLIT])]
    criterion_weight = lambda row: row.get("method_params", {}).get("split_weights", {}).get(ACTIVE_SPLIT, row["weight"])
    positive = sum(criterion_weight(row) for row in active if row["type"] == "positive")
    earned = sum(criterion_weight(row) * criteria.get(row["id"], 0.0) for row in active if row["type"] == "positive")
    penalty = sum(abs(row["weight"]) for row in TASK["criteria"] if row["type"] == "penalty" and criteria.get(row["id"], 0.0) > 0)
    return max(0.0, round((earned - penalty) / positive, 6))


def main():
    candidate = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/app/output/answer.xlsx")
    criteria, failures = evaluate(candidate)
    total = score(criteria)
    payload = {"task_id": TASK["task_id"], "split": ACTIVE_SPLIT, "status": "SCORED", "candidate": str(candidate), "normalized_score": total, "pass": total >= TASK["pass_threshold"], "criterion_scores": criteria, "failure_codes": failures, "stderr": []}
    log_root = Path(os.environ.get("P15_VERIFIER_LOG_DIR", "/logs/verifier"))
    try:
        log_root.mkdir(parents=True, exist_ok=True)
        (log_root / "report.json").write_text(json.dumps(payload, indent=2, sort_keys=True) + "\\n")
        (log_root / "reward.txt").write_text(str(total) + "\\n")
    except OSError:
        pass
    print(json.dumps(payload, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
