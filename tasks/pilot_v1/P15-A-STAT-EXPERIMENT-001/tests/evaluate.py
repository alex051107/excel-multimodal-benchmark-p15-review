#!/usr/bin/env python3
"""Deterministic semantic judge for P15-A-STAT-EXPERIMENT-001."""
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
  "task_id": "P15-A-STAT-EXPERIMENT-001",
  "pass_threshold": 0.7,
  "required_sheets": [
    "Data",
    "Plan",
    "Analysis",
    "Results",
    "Visualization",
    "Checks"
  ],
  "protected": {
    "Data!B4": 0.7,
    "Data!C13": 3.4,
    "Data!E4": "Yes"
  },
  "formula_cells": [
    "Data!D4",
    "Analysis!B5",
    "Analysis!B10",
    "Results!B8"
  ],
  "criteria": [
    {"id":"R001","description":"The workbook opens and contains every required analysis sheet.","weight":1,"type":"positive","dimension":"file_usability","method":"deterministic","method_params":{"implemented_check":"required_sheets"}},
    {"id":"R002","description":"Every included paired-difference row is formula-linked and numerically correct.","weight":2,"type":"positive","dimension":"paired_rows","method":"deterministic","method_params":{"implemented_check":"paired_differences","oracle":"split_contract"}},
    {"id":"R003","description":"Paired sample size and mean difference agree with the independent analysis.","weight":2,"type":"positive","dimension":"paired_location","method":"deterministic","method_params":{"implemented_check":"paired_n_mean","oracle":"split_contract"}},
    {"id":"R004","description":"Paired SD, SE, t statistic, and degrees of freedom agree with the independent analysis.","weight":3,"type":"positive","dimension":"paired_dispersion","method":"deterministic","method_params":{"implemented_check":"paired_sd_se_t_df","oracle":"split_contract"}},
    {"id":"R005","description":"The two-sided paired-test p-value uses the declared method and agrees with the independent analysis.","weight":3,"type":"positive","dimension":"two_sided_p_value","method":"deterministic","method_params":{"implemented_check":"two_sided_p_value","oracle":"split_contract"}},
    {"id":"R006","description":"Both 95% confidence-interval bounds agree with the independent paired analysis.","weight":2,"type":"positive","dimension":"confidence_interval","method":"deterministic","method_params":{"implemented_check":"confidence_interval","oracle":"split_contract"}},
    {"id":"R007","description":"The planned paired sample size equals the declared normal-approximation result.","weight":2,"type":"positive","dimension":"power_planning","method":"deterministic","method_params":{"implemented_check":"planned_sample_size","oracle":"split_contract"}},
    {"id":"R008","description":"The chart remains bound to all paired-difference rows through formula-linked visualization cells.","weight":2,"type":"positive","dimension":"chart_binding","method":"deterministic","method_params":{"implemented_check":"paired_chart_binding"}},
    {"id":"R009","description":"The p-value and paired-difference visualization respond correctly to the declared measurement-value perturbation.","weight":2,"type":"positive","dimension":"native_recalculation","method":"deterministic","method_params":{"implemented_check":"measurement_perturbation","split_weights":{"dev":2,"confirm":3}}},
    {"id":"R010","description":"The decision summary applies the declared two-sided 0.05 rule to the calculated p-value.","weight":2,"type":"positive","dimension":"decision_consistency","method":"deterministic","method_params":{"implemented_check":"decision_summary"}},
    {"id":"R011","description":"All source observations and inclusion flags match the split-specific starting workbook.","weight":2,"type":"positive","dimension":"source_protection","method":"deterministic","method_params":{"implemented_check":"source_pairing_protection"}},
    {"id":"R012","description":"Excluding one prespecified QC-flagged pair updates n, inference, and decision while preserving its raw measurements, planned sample size, and full-source chart binding.","weight":1,"type":"positive","dimension":"qc_inclusion_recalculation","method":"deterministic","method_params":{"implemented_check":"inclusion_flag_perturbation","applies_to":["dev"]}},
    {"id":"P001","description":"Penalty for independent-group substitution, one-tailed inference, broken pairing, or source edits.","weight":-7,"type":"penalty","dimension":"method_integrity","method":"deterministic","method_params":{"implemented_check":"paired_method_penalty"}}
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
        # Modern Excel may persist otherwise ordinary worksheet functions with
        # a compatibility prefix.  The prefix does not change their semantics.
        expression = re.sub(r"(?i)_xl(?:fn|ws)\.", "", expression)
        expression = expression.replace("^", "**").replace("<>", "!=")
        expression = re.sub(r"\bSTDEV\.S\b", "STDEV_S", expression, flags=re.I)
        # Excel retains STDEV as a backward-compatible alias for STDEV.S.
        # Treat both spellings identically so a semantically equivalent
        # workbook is not rejected by the deterministic formula engine.
        expression = re.sub(r"\bSTDEV\b", "STDEV_S", expression, flags=re.I)
        expression = re.sub(r"\bT\.DIST\.2T\b", "T_DIST_2T", expression, flags=re.I)
        expression = re.sub(r"\bT\.DIST\.RT\b", "T_DIST_RT", expression, flags=re.I)
        expression = re.sub(r"\bT\.INV\.2T\b", "T_INV_2T", expression, flags=re.I)
        expression = re.sub(r"\bCEILING\.MATH\b", "CEILING_MATH", expression, flags=re.I)
        expression = re.sub(r"\bCEILING\.PRECISE\b", "CEILING_PRECISE", expression, flags=re.I)
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
            if name == "ROUNDUP":
                digits = int(args[1])
                factor = 10 ** digits
                return math.ceil(args[0] * factor) / factor if args[0] >= 0 else math.floor(args[0] * factor) / factor
            if name in {"CEILING", "CEILING_MATH", "CEILING_PRECISE"}:
                significance = abs(args[1]) if len(args) > 1 else 1
                if significance == 0:
                    return 0
                return math.ceil(args[0] / significance) * significance
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


def normalized_formula(value):
    return re.sub(r"\s+", "", str(value or "")).replace("$", "").replace("'", "").upper()


def uses_two_sided_t_method(value):
    formula = normalized_formula(value)
    if "T.DIST.2T(" in formula:
        return True
    if "T.DIST.RT(" not in formula:
        return False
    return (
        "2*T.DIST.RT(" in formula
        or "*2" in formula
        or "/0.5" in formula
        or formula.count("T.DIST.RT(") >= 2
    )


def decision_matches(value, reject_null):
    text = " ".join(str(value or "").strip().casefold().split())
    non_rejection = (
        "do not reject h0",
        "do not reject the null",
        "fail to reject h0",
        "fail to reject the null",
        "not statistically significant",
        "no statistically significant",
        "non-significant",
        "nonsignificant",
        "insufficient evidence",
    )
    rejection = (
        "reject h0",
        "reject the null",
        "statistically significant",
        "significant paired effect",
        "evidence of a paired effect",
    )
    if reject_null:
        return not any(phrase in text for phrase in non_rejection) and any(phrase in text for phrase in rejection)
    return any(phrase in text for phrase in non_rejection)


def normalized_label(value):
    text = str(value or "").replace("−", "-").replace("–", "-").replace("—", "-")
    return " ".join(text.strip().casefold().split())


def find_labeled_value(workbook, sheet_name, aliases, prefer_first=False):
    worksheet = workbook[sheet_name]
    all_matches = []
    for alias in aliases:
        expected = normalized_label(alias)
        matches = []
        for row in worksheet.iter_rows():
            for cell in row:
                if normalized_label(cell.value) == expected and cell.column < worksheet.max_column:
                    matches.append(f"{sheet_name}!{worksheet.cell(cell.row, cell.column + 1).coordinate}")
        if matches:
            if len(matches) != 1:
                return None
            if prefer_first:
                return matches[0]
            all_matches.extend(matches)
    return all_matches[0] if len(all_matches) == 1 else None


def find_header_columns(workbook):
    columns = {}
    def assign(role, column):
        if role in columns:
            raise ValueError(f"AMBIGUOUS_DATA_HEADER:{role}")
        columns[role] = column

    for cell in workbook["Data"][3]:
        label = normalized_label(cell.value)
        if label == "subject" or label == "batch":
            assign("subject", col_name(cell.column))
        elif "analysis-set difference" in label or "analysis set difference" in label:
            assign("analysis_difference", col_name(cell.column))
        elif "paired difference" in label:
            assign("raw_difference", col_name(cell.column))
        elif label in {"include", "included", "include?"}:
            assign("include", col_name(cell.column))
    return columns


def metric_addresses(workbook):
    return {
        "n": find_labeled_value(workbook, "Analysis", ("Paired sample size", "Complete pairs (n)")),
        "mean": find_labeled_value(workbook, "Analysis", ("Mean paired difference",)),
        "sd": find_labeled_value(workbook, "Analysis", ("SD of paired differences",)),
        "se": find_labeled_value(workbook, "Analysis", ("Standard error",)),
        "t": find_labeled_value(workbook, "Analysis", ("t statistic",)),
        "df": find_labeled_value(workbook, "Analysis", ("Degrees of freedom",)),
        "p": find_labeled_value(workbook, "Analysis", ("Two-sided p-value",)),
        "ci_lower": find_labeled_value(workbook, "Analysis", ("95% CI lower",)),
        "ci_upper": find_labeled_value(workbook, "Analysis", ("95% CI upper",)),
        "planned_n": find_labeled_value(workbook, "Analysis", ("Planned n at d=0.8", "Required complete paired subjects")),
        # The workbook may separately report test alpha and planning alpha.
        # Prefer the unqualified test label, but still reject a repeated exact
        # label inside either section.
        "alpha": find_labeled_value(workbook, "Analysis", ("Alpha", "Two-sided alpha"), prefer_first=True),
    }


def normalized_chart_ref(value):
    return str(value or "").replace("$", "").replace("'", "").replace(" ", "").upper()


def chart_binding_details(workbook, engine, oracle, data_columns):
    charts = getattr(workbook["Visualization"], "_charts", [])
    subject_column = data_columns.get("subject")
    difference_column = data_columns.get("raw_difference")
    for chart in charts:
        for series in chart.series:
            category_formula = getattr(getattr(getattr(series, "cat", None), "strRef", None), "f", "") or getattr(getattr(getattr(series, "cat", None), "numRef", None), "f", "")
            value_formula = getattr(getattr(getattr(series, "val", None), "numRef", None), "f", "")
            category = normalized_chart_ref(category_formula)
            values = normalized_chart_ref(value_formula)
            if subject_column and difference_column and category == f"DATA!{subject_column}4:{subject_column}13" and values == f"DATA!{difference_column}4:{difference_column}13":
                return {"sheet": "Data", "column": difference_column, "mode": "direct"}
            if category == "VISUALIZATION!A4:A13" and values == "VISUALIZATION!B4:B13":
                helper_ok = True
                for index, row in enumerate(range(4, 14)):
                    address = f"Visualization!B{row}"
                    if not formula_present(workbook, address):
                        helper_ok = False
                        break
                    try:
                        observed = engine.value("Visualization", f"B{row}")
                    except Exception:
                        helper_ok = False
                        break
                    expected = oracle.get("full_differences", oracle["differences"])[index]
                    if not close(observed, expected, 0.0001):
                        helper_ok = False
                        break
                if helper_ok:
                    return {"sheet": "Visualization", "column": "B", "mode": "helper"}
    return None


def find_decision_formulas(workbook):
    matches = []
    for row in workbook["Results"].iter_rows():
        for cell in row:
            value = cell.value
            formula = normalized_formula(value)
            text = str(value or "").casefold()
            if isinstance(value, str) and value.startswith("=") and "IF(" in formula and "reject" in text:
                matches.append((f"Results!{cell.coordinate}", value))
    return matches


def decision_formula_ok(workbook, p_address, alpha_address):
    if not p_address:
        return False
    p_token = normalized_formula(p_address)
    alpha_tokens = {"0.05"}
    if alpha_address:
        alpha_tokens.add(normalized_formula(alpha_address))
    for _address, raw in find_decision_formulas(workbook):
        formula = normalized_formula(raw)
        correct_rule = any(
            p_token + "<" + token in formula or token + ">" + p_token in formula
            for token in alpha_tokens
        )
        text = str(raw).casefold().replace("₀", "0")
        if correct_rule and "reject" in text and (
            "do not reject" in text or "fail to reject" in text
        ):
            return True
    return False


def task_checks(workbook, engine, oracle, contract, oracle_module):
    checks, failures = {}, []
    data_columns = find_header_columns(workbook)
    metrics = metric_addresses(workbook)

    paired_rows_ok = bool(data_columns.get("raw_difference"))
    paired_columns = [data_columns["raw_difference"]] if paired_rows_ok else []
    if data_columns.get("analysis_difference"):
        paired_columns.append(data_columns["analysis_difference"])
    if not paired_rows_ok:
        failures.append("PAIRED_DIFFERENCE_COLUMN_NOT_FOUND")
    for index, row in enumerate(range(4, 14)):
        for column in paired_columns:
            address = f"Data!{column}{row}"
            if not formula_present(workbook, address):
                paired_rows_ok = False
                failures.append(f"FORMULA_REQUIRED:{address}")
            try:
                observed = engine.value("Data", f"{column}{row}")
            except Exception as exc:
                observed = None
                failures.append(f"PAIRED_ROW_EVALUATION_FAILED:{address}:{type(exc).__name__}")
            if not close(observed, oracle["differences"][index], 0.0001):
                paired_rows_ok = False
                failures.append(f"PAIRED_ROW_MISMATCH:{address}")
    checks["R002"] = 1.0 if paired_rows_ok else 0.0

    def validate_group(criterion_id, expected):
        ok = True
        for key, target in expected.items():
            address = metrics.get(key)
            if not address:
                ok = False
                failures.append(f"METRIC_NOT_FOUND:{key}")
                continue
            if not formula_present(workbook, address):
                ok = False
                failures.append(f"FORMULA_REQUIRED:{address}")
            sheet, cell = address.split("!", 1)
            try:
                observed = engine.value(sheet, cell)
            except Exception as exc:
                observed = None
                failures.append(f"INFERENTIAL_EVALUATION_FAILED:{address}:{type(exc).__name__}")
            if not close(observed, target, 0.0001):
                ok = False
                failures.append(f"INFERENTIAL_VALUE_MISMATCH:{address}")
        checks[criterion_id] = 1.0 if ok else 0.0
        return ok

    location_ok = validate_group("R003", {"n": oracle["n"], "mean": oracle["mean_difference"]})
    dispersion_ok = validate_group("R004", {
        "sd": oracle["sd_difference"], "se": oracle["se"],
        "t": oracle["t"], "df": oracle["df"],
    })
    p_address = metrics.get("p")
    p_formula = workbook[p_address.split("!", 1)[0]][p_address.split("!", 1)[1]].value if p_address else None
    p_numeric_ok = validate_group("R005", {"p": oracle["p_value"]})
    mean_address, sd_address = metrics.get("mean"), metrics.get("sd")
    mean_formula = normalized_formula(workbook["Analysis"][mean_address.split("!", 1)[1]].value) if mean_address else ""
    sd_formula = normalized_formula(workbook["Analysis"][sd_address.split("!", 1)[1]].value) if sd_address else ""
    paired_ranges = [f"DATA!{column}4:{column}13" for column in paired_columns]
    paired_method_ok = uses_two_sided_t_method(p_formula) and any(value in mean_formula for value in paired_ranges) and any(value in sd_formula for value in paired_ranges)
    if not paired_method_ok:
        failures.append("PAIRED_TWO_SIDED_METHOD_NOT_DEMONSTRATED")
    checks["R005"] = 1.0 if p_numeric_ok and paired_method_ok else 0.0
    ci_ok = validate_group("R006", {"ci_lower": oracle["ci_lower"], "ci_upper": oracle["ci_upper"]})
    planned_address = metrics.get("planned_n")
    planned_formula = normalized_formula(workbook["Analysis"][planned_address.split("!", 1)[1]].value) if planned_address else ""
    planned_ok = validate_group("R007", {"planned_n": oracle["planned_n"]})
    planned_method_ok = any(function in planned_formula for function in ("ROUNDUP(", "CEILING(", "CEILING.MATH(", "CEILING.PRECISE("))
    checks["R007"] = 1.0 if planned_ok and planned_method_ok else 0.0

    chart_binding = chart_binding_details(workbook, engine, oracle, data_columns)
    checks["R008"] = 1.0 if chart_binding else 0.0
    if not chart_binding:
        failures.append("CHART_NOT_BOUND_TO_ALL_PAIRED_DIFFERENCES")

    perturbations = {row["name"]: row for row in contract["perturbations"]}

    def perturbation_matches(case, require_decision=False):
        replay = FormulaEngine(workbook, case["overrides"])
        for address, target in {**case["expected"], **case["invariants"]}.items():
            sheet, cell = address.split("!", 1)
            observed = replay.value(sheet, cell)
            if target is None:
                if observed is not None:
                    return False
            elif isinstance(target, (int, float)):
                if not close(observed, target, 0.0001):
                    return False
            elif observed != target:
                return False
        if require_decision:
            if not decision_formula_ok(workbook, p_address, metrics.get("alpha")):
                return False
            if not chart_binding:
                return False
        return True

    try:
        observation_case = perturbations["paired_observation"]
        perturbed_group_2 = list(oracle_module.GROUP_2)
        perturb_address, perturb_value = next(iter(observation_case["overrides"].items()))
        perturb_row = int(re.search(r"(\d+)$", perturb_address).group(1))
        perturbed_group_2[perturb_row - 4] = perturb_value
        perturbed_oracle = oracle_module.recompute(group_2=perturbed_group_2)
        dynamic_engine = FormulaEngine(workbook, observation_case["overrides"])
        expected_difference = perturbed_group_2[perturb_row - 4] - oracle_module.GROUP_1[perturb_row - 4]
        chart_cell = f"{chart_binding['column']}{perturb_row}" if chart_binding else None
        dynamic_ok = bool(
            p_address and chart_binding
            and close(dynamic_engine.value(*p_address.split("!", 1)), perturbed_oracle["p_value"], 0.0001)
            and close(dynamic_engine.value(chart_binding["sheet"], chart_cell), expected_difference, 0.0001)
        )
    except Exception as exc:
        dynamic_ok = False
        failures.append(f"MEASUREMENT_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R009"] = 1.0 if dynamic_ok else 0.0

    inclusion_ok = True
    checks["R012"] = 0.0
    if ACTIVE_SPLIT == "dev":
        inclusion_capable = bool(data_columns.get("include") and data_columns.get("analysis_difference"))
        try:
            inclusion_ok = inclusion_capable and perturbation_matches(perturbations["qc_exclusion"], require_decision=True)
        except Exception as exc:
            inclusion_ok = False
            failures.append(f"INCLUSION_PERTURBATION_FAILED:{type(exc).__name__}")
        checks["R012"] = 1.0 if inclusion_ok else 0.0
        if not inclusion_ok:
            failures.append("QC_EXCLUSION_CHAIN_NOT_PRESERVED")

    decision_ok = decision_formula_ok(workbook, p_address, metrics.get("alpha"))
    checks["R010"] = 1.0 if decision_ok else 0.0
    if not decision_ok:
        failures.append("DECISION_RULE_NOT_LINKED_TO_TWO_SIDED_P_VALUE")

    baseline = openpyxl.load_workbook(TASK_ROOT / contract["baseline_workbook"], data_only=False, read_only=False)
    source_ok = all(
        workbook["Data"][f"{column}{row}"].value == baseline["Data"][f"{column}{row}"].value
        for row in range(4, 14) for column in ("B", "C", "E")
    )
    checks["R011"] = 1.0 if source_ok else 0.0
    if not source_ok:
        failures.append("SOURCE_OBSERVATIONS_OR_INCLUDE_FLAGS_CHANGED")
    checks["P001"] = 0.0 if all((paired_rows_ok, location_ok, dispersion_ok, paired_method_ok, p_numeric_ok, ci_ok, planned_ok, dynamic_ok, inclusion_ok, decision_ok, source_ok)) else 1.0
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
