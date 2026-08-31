#!/usr/bin/env python3
"""Deterministic semantic judge for P15-A-POLICY-EIA-001."""
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
  "task_id": "P15-A-POLICY-EIA-001",
  "pass_threshold": 0.7,
  "required_sheets": [
    "Generation_Data",
    "Policy_Assumptions",
    "Scenario_Model",
    "Policy_Results",
    "Checks"
  ],
  "protected": {
    "Generation_Data!D4": 675000,
    "Generation_Data!D5": 1802000,
    "Generation_Data!D6": 425000,
    "Generation_Data!D7": 238000,
    "Generation_Data!D8": 4000000,
    "Generation_Data!E4": 1.0,
    "Generation_Data!E5": 0.4
  },
  "formula_cells": [
    "Scenario_Model!C5",
    "Scenario_Model!C6",
    "Scenario_Model!C7",
    "Scenario_Model!C8",
    "Scenario_Model!C9",
    "Policy_Results!B4",
    "Policy_Results!B5",
    "Policy_Results!B6"
  ],
  "criteria": [
    {
      "id": "R001",
      "description": "A readable scenario workbook opens and retains all requested sheets.",
      "weight": 1,
      "type": "positive",
      "dimension": "file_usability",
      "method": "deterministic",
      "method_params": {}
    },
    {
      "id": "R002",
      "description": "Policy coal generation applies the documented displacement rate to the protected source value.",
      "weight": 3,
      "type": "positive",
      "dimension": "policy_coal_generation",
      "method": "deterministic",
      "method_params": {
        "implemented_check": "Scenario_Model!C5 formula and independent replay value"
      }
    },
    {
      "id": "R003",
      "description": "Policy wind and solar generation apply their documented uplift schedules.",
      "weight": 2,
      "type": "positive",
      "dimension": "renewable_schedule",
      "method": "deterministic",
      "method_params": {"implemented_check": "Scenario_Model!C6:C7 formulas and independent replay values"}
    },
    {
      "id": "R004",
      "description": "Gas generation closes the policy demand balance and the results sheet reports the gas change.",
      "weight": 3,
      "type": "positive",
      "dimension": "gas_balancing",
      "method": "deterministic",
      "method_params": {"implemented_check": "Scenario_Model!C8 and Policy_Results!B6 formulas and replay values"}
    },
    {
      "id": "R005",
      "description": "Policy emissions match the independent replay and remain formula-linked.",
      "weight": 3,
      "type": "positive",
      "dimension": "policy_emissions",
      "method": "deterministic",
      "method_params": {"implemented_check": "Scenario_Model!C9 formula and replay value"}
    },
    {
      "id": "R006",
      "description": "The results sheet calculates absolute emissions reduction from the two scenarios.",
      "weight": 2,
      "type": "positive",
      "dimension": "emissions_reduction",
      "method": "deterministic",
      "method_params": {"implemented_check": "Policy_Results!B4 formula and replay value"}
    },
    {
      "id": "R007",
      "description": "The results sheet calculates emissions-intensity reduction from the two scenarios.",
      "weight": 2,
      "type": "positive",
      "dimension": "intensity_reduction",
      "method": "deterministic",
      "method_params": {"implemented_check": "Policy_Results!B5 formula and replay value"}
    },
    {
      "id": "R008",
      "description": "Changing the coal-displacement assumption updates coal, gas, policy emissions, and emissions reduction to the replayed values.",
      "weight": 2,
      "type": "positive",
      "dimension": "dynamic_recalculation",
      "method": "deterministic",
      "method_params": {"implemented_check": "coal-displacement perturbation propagation", "split_weights": {"dev": 2, "confirm": 3}}
    },
    {
      "id": "R009",
      "description": "Base and policy modeled generation each balance to their corresponding demand.",
      "weight": 2,
      "type": "positive",
      "dimension": "demand_balance",
      "method": "deterministic",
      "method_params": {"implemented_check": "Scenario_Model B4:B8 and C4:C8 balance equations"}
    },
    {
      "id": "R010",
      "description": "Emissions use the protected coal and gas factors through workbook formulas.",
      "weight": 2,
      "type": "positive",
      "dimension": "factor_linkage",
      "method": "deterministic",
      "method_params": {"implemented_check": "Scenario_Model!B9:C9 formula and factor-linked replay values"}
    },
    {
      "id": "R011",
      "description": "Historical generation, demand, and emissions-factor inputs remain unchanged.",
      "weight": 2,
      "type": "positive",
      "dimension": "change_locality",
      "method": "deterministic",
      "method_params": {"implemented_check": "Generation_Data D4:E8 protected values"}
    },
    {
      "id": "R012",
      "description": "A distinct policy-demand-growth perturbation updates demand, balancing gas, emissions, intensity, and reported results while preserving source generation and the coal/renewable policy schedule.",
      "weight": 1,
      "type": "positive",
      "dimension": "demand_growth_recalculation",
      "method": "deterministic",
      "method_params": {"implemented_check": "demand-growth perturbation propagation", "applies_to": ["dev"]}
    },
    {
      "id": "P001",
      "description": "Penalty for overwriting protected source values, hardcoding core scenario outputs, or breaking either demand balance.",
      "weight": -6,
      "type": "penalty",
      "dimension": "file_integrity",
      "method": "deterministic",
      "method_params": {"implemented_check": "critical conjunction of R002:R011"}
    }
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


def intensity_reporting_direction(workbook, actual, expected_reduction):
    """Return +1 for a positive reduction or -1 for a signed policy-minus-base change."""
    if not formula_present(workbook, "Policy_Results!B5"):
        return None
    label = workbook["Policy_Results"]["A5"].value
    normalized_label = re.sub(r"[^a-z0-9]+", " ", str(label or "").casefold()).strip()
    label_words = set(normalized_label.split())
    if "intensity" not in label_words:
        return None
    signed_label = bool(label_words & {"change", "delta", "difference", "variation"})
    reduction_label = bool(label_words & {"reduction", "decrease", "improvement", "saving", "savings"})
    contradictory_increase = bool(label_words & {"increase", "increased", "rise", "rising", "growth", "gain"})
    if contradictory_increase or not (signed_label or reduction_label):
        return None
    if close(actual, expected_reduction, 1e-9):
        return 1.0
    if signed_label and close(actual, -expected_reduction, 1e-9):
        return -1.0
    return None


def task_checks(workbook, engine, oracle, oracle_module, contract):
    checks, failures = {}, []

    def observed(address):
        sheet, cell = address.split("!", 1)
        try:
            return engine.value(sheet, cell)
        except Exception as exc:
            failures.append(f"POLICY_FORMULA_EVALUATION_FAILED:{address}:{type(exc).__name__}")
            return None

    for address in contract["formula_cells"]:
        if not formula_present(workbook, address):
            failures.append(f"FORMULA_REQUIRED:{address}")

    coal_ok = formula_present(workbook, "Scenario_Model!C5") and close(
        observed("Scenario_Model!C5"), oracle["policy"]["coal"], 0.5
    )
    checks["R002"] = 1.0 if coal_ok else 0.0

    renewables_ok = all(
        formula_present(workbook, address) and close(observed(address), target, 0.5)
        for address, target in {
            "Scenario_Model!C6": oracle["policy"]["wind"],
            "Scenario_Model!C7": oracle["policy"]["solar"],
        }.items()
    )
    checks["R003"] = 1.0 if renewables_ok else 0.0

    gas_ok = all(
        formula_present(workbook, address) and close(observed(address), target, 0.5)
        for address, target in {
            "Scenario_Model!C8": oracle["policy"]["gas"],
            "Policy_Results!B6": oracle["gas_change"],
        }.items()
    )
    checks["R004"] = 1.0 if gas_ok else 0.0

    emissions_ok = formula_present(workbook, "Scenario_Model!C9") and close(
        observed("Scenario_Model!C9"), oracle["policy_emissions"], 0.5
    )
    checks["R005"] = 1.0 if emissions_ok else 0.0

    reduction_ok = formula_present(workbook, "Policy_Results!B4") and close(
        observed("Policy_Results!B4"), oracle["emissions_reduction"], 0.5
    )
    checks["R006"] = 1.0 if reduction_ok else 0.0

    intensity_actual = observed("Policy_Results!B5")
    intensity_direction = intensity_reporting_direction(workbook, intensity_actual, oracle["intensity_reduction"])
    intensity_ok = intensity_direction is not None
    checks["R007"] = 1.0 if intensity_ok else 0.0

    perturbations = {row["name"]: row for row in contract["perturbations"]}

    def perturbation_matches(case, reporting_direction=1.0):
        replay = FormulaEngine(workbook, case["overrides"])
        for address, target in {**case["expected"], **case["invariants"]}.items():
            sheet, cell = address.split("!", 1)
            observed_value = replay.value(sheet, cell)
            if address == "Policy_Results!B5":
                target *= reporting_direction
            tolerance = 1e-9 if abs(target) < 1 else 0.5
            if not close(observed_value, target, tolerance):
                return False
        return True

    try:
        coal_case = perturbations["coal_displacement"]
        if "expected" in coal_case:
            dynamic_ok = perturbation_matches(coal_case)
        else:
            perturbed_policy = dict(oracle_module.POLICY)
            perturbed_policy["coal_displacement"] = coal_case["value"]
            perturbed_oracle = oracle_module.recompute(policy_case=perturbed_policy)
            perturbed_engine = FormulaEngine(workbook, {coal_case["address"]: coal_case["value"]})
            dynamic_ok = close(perturbed_engine.value("Policy_Results", "B4"), perturbed_oracle["emissions_reduction"], 0.5)
    except Exception as exc:
        dynamic_ok = False
        failures.append(f"COAL_DISPLACEMENT_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R008"] = 1.0 if dynamic_ok else 0.0

    demand_growth_ok = True
    checks["R012"] = 0.0
    if ACTIVE_SPLIT == "dev":
        try:
            demand_growth_ok = intensity_direction is not None and perturbation_matches(
                perturbations["demand_growth"], intensity_direction
            )
        except Exception as exc:
            demand_growth_ok = False
            failures.append(f"DEMAND_GROWTH_PERTURBATION_FAILED:{type(exc).__name__}")
        checks["R012"] = 1.0 if demand_growth_ok else 0.0

    try:
        base_balance = engine.value("Scenario_Model", "B4") - sum(
            engine.value("Scenario_Model", f"B{row}") for row in (5, 6, 7, 8)
        )
        policy_balance = engine.value("Scenario_Model", "C4") - sum(
            engine.value("Scenario_Model", f"C{row}") for row in (5, 6, 7, 8)
        )
    except Exception as exc:
        base_balance = policy_balance = None
        failures.append(f"POLICY_BALANCE_EVALUATION_FAILED:{type(exc).__name__}")
    balance_ok = close(base_balance, 0.0, 0.1) and close(policy_balance, 0.0, 0.1)
    checks["R009"] = 1.0 if balance_ok else 0.0

    factor_ok = all(
        formula_present(workbook, address) and close(observed(address), target, 0.5)
        for address, target in {
            "Scenario_Model!B9": oracle["base"]["emissions"],
            "Scenario_Model!C9": oracle["policy"]["emissions"],
        }.items()
    )
    checks["R010"] = 1.0 if factor_ok else 0.0

    protected_ok = all(
        workbook[sheet][cell].value == value
        for key, value in contract["protected"].items()
        for sheet, cell in [key.split("!", 1)]
    )
    checks["R011"] = 1.0 if protected_ok else 0.0

    critical_ok = all(
        checks[criterion] == 1.0
        for criterion in ("R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009", "R010", "R011")
    )
    critical_ok = critical_ok and demand_growth_ok
    checks["P001"] = 0.0 if critical_ok else 1.0
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
    for sheet in TASK["required_sheets"]:
        if sheet not in workbook.sheetnames:
            failures.append(f"MISSING_SHEET:{sheet}")
    if failures:
        return criteria, failures
    criteria["R001"] = 1.0
    try:
        contract = load_contract()
        oracle, oracle_module = load_oracle(contract)
        checks, task_failures = task_checks(workbook, FormulaEngine(workbook), oracle, oracle_module, contract)
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
