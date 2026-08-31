#!/usr/bin/env python3
"""Deterministic task-specific judge for P15-B-HEALTH-REPORT-001."""
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
from pathlib import Path

import openpyxl

CRITERIA = [
    {"id": "R001", "description": "A readable workbook contains every requested analytical and report sheet.", "weight": 1, "type": "positive", "dimension": "file_usability", "method": "deterministic", "method_params": {}},
    {"id": "R002", "description": "Analytical_Data contains exactly the Vermont and New Hampshire 2012-2017 records with correct state/year alignment.", "weight": 4, "type": "positive", "dimension": "data_alignment", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R003", "description": "Both states' baseline mean, current mean, absolute change, and percent change equal the independent replay.", "weight": 4, "type": "positive", "dimension": "metric_correctness", "method": "deterministic", "method_params": {"oracle": "metadata/oracle_recompute.py"}},
    {"id": "R004", "description": "Changing Raw_Vermont 2012 AADR from 150.6 to 151.6 flows through Analytical_Data and the Vermont baseline metric.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Raw_Vermont!D4"}},
    {"id": "R005", "description": "Changing Raw_New_Hampshire 2015 AADR from 149.0 to 150.0 flows through current metrics, report, and chart source cells.", "weight": 3, "type": "positive", "dimension": "native_workbook_semantics", "method": "deterministic", "method_params": {"source_cell": "Raw_New_Hampshire!D7"}},
    {"id": "R006", "description": "All numerical briefing claims and the highest-current-state claim are formulas linked to Period_Metrics.", "weight": 3, "type": "positive", "dimension": "report_traceability", "method": "deterministic", "method_params": {}},
    {"id": "R007", "description": "A column chart on Visualization has categories bound exactly to A4:A5 and values bound exactly to formula-linked B4:B5.", "weight": 4, "type": "positive", "dimension": "chart_binding", "method": "ooxml", "method_params": {}},
    {"id": "R008", "description": "Every supplied Vermont, New Hampshire, and geography-map source field is preserved exactly.", "weight": 3, "type": "positive", "dimension": "change_locality", "method": "deterministic", "method_params": {}},
    {"id": "R009", "description": "Checks formula-link the 12 analytical records, two metric rows, and zero report-to-metric closure residual.", "weight": 2, "type": "positive", "dimension": "auditability", "method": "deterministic", "method_params": {}},
    {"id": "P001", "description": "Penalty for time/geography mixing, broken raw-data lineage, wrong metrics, hardcoded report claims, stale chart binding, or source alteration.", "weight": -7, "type": "penalty", "dimension": "integrity", "method": "deterministic", "method_params": {}},
]

TASK = {
    "task_id": "P15-B-HEALTH-REPORT-001",
    "pass_threshold": 0.7,
    "required_sheets": ["Raw_Vermont", "Raw_New_Hampshire", "Geo_Map", "Analytical_Data", "Period_Metrics", "Report", "Visualization", "Checks"],
    "formula_cells": [
        *[f"Analytical_Data!D{row}" for row in range(4, 16)],
        *[f"Period_Metrics!{column}{row}" for row in (4, 5) for column in ("B", "C", "D", "E")],
        "Report!B4", "Report!B5", "Report!B6", "Visualization!B4", "Visualization!B5",
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
        "vermont_source_cell": "Raw_Vermont!D4",
        "vermont_new_value": 151.6,
        "new_hampshire_source_cell": "Raw_New_Hampshire!D7",
        "new_hampshire_new_value": 150.0,
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
        finally:
            self.stack.discard(key)
        self.memo[key] = result
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
        # Excel's ampersand is string concatenation.  Converting it outside
        # quoted literals lets the safe AST evaluator replay formula-generated
        # briefing sentences without treating them as static text.
        expression = outside(re.compile(r"&"), "+", expression)
        expression = re.sub(r"(?<![<>=!])=(?!=)", "==", expression)
        return self.safe_eval(ast.parse(expression, mode="eval").body)

    def safe_eval(self, node):
        if isinstance(node, ast.Constant): return node.value
        if isinstance(node, ast.List): return [self.safe_eval(item) for item in node.elts]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            value = self.safe_eval(node.operand)
            if isinstance(value, list):
                return [(-item if isinstance(node.op, ast.USub) else item) for item in value]
            return -value if isinstance(node.op, ast.USub) else value
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
            name = node.func.id.upper()
            if name == "IF":
                if len(node.args) != 3:
                    raise ValueError("IF_ARGUMENTS")
                return self.safe_eval(node.args[1] if self.safe_eval(node.args[0]) else node.args[2])
            if name == "IFERROR":
                if len(node.args) != 2:
                    raise ValueError("IFERROR_ARGUMENTS")
                try:
                    return self.safe_eval(node.args[0])
                except (ArithmeticError, TypeError, ValueError):
                    return self.safe_eval(node.args[1])
            args = [self.safe_eval(item) for item in node.args]
            values = [value for group in args for value in (group if isinstance(group, list) else [group]) if value is not None]
            if name == "SUM": return sum(values)
            if name == "AVERAGE": return statistics.mean(values)
            if name == "AVERAGEIFS":
                if len(args) < 3 or len(args) % 2 == 0:
                    raise ValueError("AVERAGEIFS_ARGUMENTS")
                average_range = as_range(args[0], name)
                criteria_pairs = [
                    (as_range(args[index], name), args[index + 1])
                    for index in range(1, len(args), 2)
                ]
                if any(len(criteria_range) != len(average_range) for criteria_range, _ in criteria_pairs):
                    raise ValueError("AVERAGEIFS_RANGE_LENGTH")
                selected = [
                    value
                    for row, value in enumerate(average_range)
                    if isinstance(value, (int, float))
                    and all(excel_criteria_match(criteria_range[row], criterion) for criteria_range, criterion in criteria_pairs)
                ]
                if not selected:
                    raise ValueError("AVERAGEIFS_NO_MATCH")
                return statistics.mean(selected)
            if name == "COUNTIFS":
                if len(args) < 2 or len(args) % 2 != 0:
                    raise ValueError("COUNTIFS_ARGUMENTS")
                criteria_pairs = [
                    (as_range(args[index], name), args[index + 1])
                    for index in range(0, len(args), 2)
                ]
                lengths = {len(criteria_range) for criteria_range, _ in criteria_pairs}
                if len(lengths) != 1:
                    raise ValueError("COUNTIFS_RANGE_LENGTH")
                return sum(
                    all(excel_criteria_match(criteria_range[row], criterion) for criteria_range, criterion in criteria_pairs)
                    for row in range(next(iter(lengths)))
                )
            if name == "MAX":
                numeric_values = [value for value in values if isinstance(value, (int, float))]
                if not numeric_values:
                    raise ValueError("MAX_NO_NUMERIC_VALUES")
                return max(numeric_values)
            if name == "MIN":
                numeric_values = [value for value in values if isinstance(value, (int, float))]
                if not numeric_values:
                    raise ValueError("MIN_NO_NUMERIC_VALUES")
                return min(numeric_values)
            if name == "COUNT": return sum(isinstance(value, (int, float)) for value in values)
            if name == "COUNTA": return sum(value not in (None, "") for value in values)
            if name == "ISERROR":
                if len(args) != 1:
                    raise ValueError("ISERROR_ARGUMENTS")
                return [False for _ in args[0]] if isinstance(args[0], list) else False
            if name == "SUMPRODUCT":
                arrays = [argument if isinstance(argument, list) else [argument] for argument in args]
                lengths = {len(array) for array in arrays}
                if len(lengths) != 1:
                    raise ValueError("SUMPRODUCT_RANGE_LENGTH")
                return sum(math.prod(array[index] for array in arrays) for index in range(next(iter(lengths))))
            if name == "MATCH":
                if len(args) not in (2, 3):
                    raise ValueError("MATCH_ARGUMENTS")
                lookup_range = as_range(args[1], name)
                match_type = 1 if len(args) == 2 else args[2]
                if match_type != 0:
                    raise ValueError("MATCH_ONLY_EXACT_SUPPORTED")
                for index, value in enumerate(lookup_range, start=1):
                    if excel_equal(value, args[0]):
                        return index
                raise ValueError("MATCH_NOT_FOUND")
            if name == "INDEX":
                if len(args) not in (2, 3):
                    raise ValueError("INDEX_ARGUMENTS")
                index_range = as_range(args[0], name)
                row_number = args[1]
                if not isinstance(row_number, (int, float)) or int(row_number) != row_number:
                    raise ValueError("INDEX_ROW_NUMBER")
                row_number = int(row_number)
                if len(args) == 3 and args[2] not in (0, 1):
                    raise ValueError("INDEX_ONLY_SINGLE_COLUMN_SUPPORTED")
                if row_number < 1 or row_number > len(index_range):
                    raise ValueError("INDEX_OUT_OF_RANGE")
                return index_range[row_number - 1]
            if name == "TEXT":
                if len(args) != 2 or not isinstance(args[0], (int, float)):
                    raise ValueError("TEXT_ARGUMENTS")
                value, number_format = args[0], str(args[1])
                sections = number_format.split(";")
                section = sections[1] if value < 0 and len(sections) > 1 else sections[0]
                percent = "%" in section
                decimal_match = re.search(r"\.([0#]+)", section)
                decimals = len(decimal_match.group(1)) if decimal_match else 0
                magnitude = abs(value) * (100 if percent else 1)
                rendered = f"{magnitude:.{decimals}f}"
                if value < 0:
                    rendered = "-" + rendered
                elif section.strip().startswith("+"):
                    rendered = "+" + rendered
                if percent:
                    rendered += "%"
                return rendered
            if name == "ABS": return abs(args[0])
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


def normalized_words(value):
    text = unicodedata.normalize("NFKC", str(value)).casefold()
    text = re.sub(r"[‐‑‒–—−]", "-", text)
    return re.sub(r"[^a-z0-9-]+", " ", text).strip()


def normalize_state_label(value):
    text = normalized_words(value).replace(".", "")
    compact = re.sub(r"[^a-z]", "", text)
    if compact in {"vt", "vermont"}:
        return "Vermont"
    if compact in {"nh", "newhampshire"}:
        return "New Hampshire"
    return value


def normalize_period_label(value):
    text = normalized_words(value)
    role = "baseline" if "baseline" in text else "current" if "current" in text else None
    years = re.findall(r"\b(?:19|20)\d{2}\b", text)
    if role and len(years) == 2:
        return role, int(years[0]), int(years[1])
    return text


def header_matches(role, value):
    text = normalized_words(value)
    words = set(text.split())
    compact = re.sub(r"[^a-z0-9]", "", text)
    if role == "state": return compact in {"state", "statename"}
    if role == "year": return compact in {"year", "reportingyear"}
    if role == "cause": return compact in {"cause", "causename", "causeofdeath"}
    if role == "deaths": return compact in {"death", "deaths", "deathcount", "deathcounts"}
    if role == "aadr": return "aadr" in words or compact == "aadr" or ("ageadjusted" in compact and "rate" in compact)
    if role == "period": return "period" in words and "metric" not in words
    if role == "reporting_region": return "reporting" in words and "region" in words
    if role == "source_key": return "source" in words and "key" in words
    if role == "baseline": return "baseline" in words and ("aadr" in words or "rate" in words or "mean" in words)
    if role == "current": return "current" in words and ("aadr" in words or "rate" in words or "mean" in words)
    if role == "absolute_change": return "absolute" in words and "change" in words
    if role == "percent_change": return ("percent" in words or "percentage" in words) and "change" in words
    if role == "claim": return compact in {"claim", "briefingclaim", "metricclaim"}
    if role == "value": return compact in {"value", "result", "metricvalue"}
    return False


def header_columns(sheet, row, roles):
    columns = {}
    duplicates = set()
    for column in range(1, min(sheet.max_column, 40) + 1):
        value = sheet.cell(row=row, column=column).value
        for role in roles:
            if not header_matches(role, value):
                continue
            if role in columns:
                duplicates.add(role)
            else:
                columns[role] = column
    return columns, duplicates


def find_semantic_table(workbook, sheet_name, roles, allow_missing=False):
    sheet = workbook[sheet_name]
    candidates = []
    for row in range(1, min(sheet.max_row, 30) + 1):
        columns, duplicates = header_columns(sheet, row, roles)
        if not duplicates and all(role in columns for role in roles):
            candidates.append((row, columns))
    if not candidates and allow_missing:
        return None
    if len(candidates) != 1:
        raise ValueError(f"{sheet_name}_SEMANTIC_HEADER_COUNT:{len(candidates)}")
    return candidates[0]


def optional_header_column(sheet, header_row, role):
    columns, duplicates = header_columns(sheet, header_row, (role,))
    if duplicates:
        raise ValueError(f"{sheet.title}_{role}_HEADER_AMBIGUOUS")
    return columns.get(role)


def unique_record_row(workbook, engine, sheet_name, table, state, year):
    header_row, columns = table
    sheet = workbook[sheet_name]
    matches = []
    data_started = False
    for row in range(header_row + 1, sheet.max_row + 1):
        row_state = engine.value(sheet_name, f"{col_name(columns['state'])}{row}")
        row_year = engine.value(sheet_name, f"{col_name(columns['year'])}{row}")
        if row_state in (None, "") and row_year in (None, ""):
            if data_started:
                break
            continue
        data_started = True
        if normalize_state_label(row_state) == normalize_state_label(state) and row_year == year:
            matches.append(row)
    if len(matches) != 1:
        raise ValueError(f"{sheet_name}_RECORD_COUNT:{state}:{year}:{len(matches)}")
    return matches[0]


def unique_state_row(workbook, engine, sheet_name, table, state):
    header_row, columns = table
    sheet = workbook[sheet_name]
    matches = []
    data_started = False
    for row in range(header_row + 1, sheet.max_row + 1):
        row_state = engine.value(sheet_name, f"{col_name(columns['state'])}{row}")
        if row_state in (None, ""):
            if data_started:
                break
            continue
        data_started = True
        if normalize_state_label(row_state) == normalize_state_label(state):
            matches.append(row)
    if len(matches) != 1:
        raise ValueError(f"{sheet_name}_STATE_ROW_COUNT:{state}:{len(matches)}")
    return matches[0]


def semantic_address(table, role, row):
    return f"{col_name(table[1][role])}{row}"


def formula_mentions_sheet(workbook, sheet_name, cell, dependency):
    formula = workbook[sheet_name][cell].value
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    compact = formula.casefold().replace("'", "").replace(" ", "")
    return f"{dependency.casefold()}!" in compact


def formula_cells_in_row(workbook, sheet_name, row):
    sheet = workbook[sheet_name]
    return [
        f"{col_name(column)}{row}"
        for column in range(1, min(sheet.max_column, 40) + 1)
        if isinstance(sheet.cell(row=row, column=column).value, str)
        and sheet.cell(row=row, column=column).value.startswith("=")
    ]


def load_public_sources(root):
    health = {}
    for filename in ("vermont_heart_disease.csv", "new_hampshire_heart_disease.csv"):
        with (root / filename).open(newline="") as handle:
            for record in csv.DictReader(handle):
                key = (normalize_state_label(record["state"]), int(record["year"]))
                health[key] = {
                    "state": record["state"],
                    "year": int(record["year"]),
                    "cause": record.get("cause_name"),
                    "deaths": int(record["deaths"]),
                    "aadr": float(record["aadr"]),
                }
    geography = {}
    with (root / "geography_map.csv").open(newline="") as handle:
        for record in csv.DictReader(handle):
            geography[normalize_state_label(record["state"])] = {
                "state": record["state"],
                "reporting_region": record["reporting_region"],
                "source_key": record["source_key"],
            }
    return health, geography


def source_value_equal(actual, expected, role):
    if role in {"year", "deaths"}:
        return isinstance(actual, (int, float)) and int(actual) == actual and int(actual) == expected
    if role == "aadr":
        return close(actual, expected, 0.000001)
    return actual == expected


def rows(workbook, sheet, start, end, columns):
    values = []
    for row in range(start, end + 1):
        record = tuple(norm(workbook[sheet].cell(row=row, column=column).value) for column in columns)
        if any(value not in (None, "") for value in record): values.append(record)
    return values


def source_tables_match(workbook, root):
    expected_health, expected_geo = load_public_sources(root)
    actual_health = {}
    for sheet in ("Raw_Vermont", "Raw_New_Hampshire"):
        table = find_semantic_table(workbook, sheet, ("state", "year", "deaths", "aadr"))
        header_row, columns = table
        cause_column = optional_header_column(workbook[sheet], header_row, "cause")
        for row in range(header_row + 1, workbook[sheet].max_row + 1):
            state = workbook[sheet].cell(row=row, column=columns["state"]).value
            year = workbook[sheet].cell(row=row, column=columns["year"]).value
            if state in (None, "") and year in (None, ""):
                continue
            if state in (None, "") or not isinstance(year, (int, float)) or int(year) != year:
                return False
            key = (normalize_state_label(state), int(year))
            if key in actual_health:
                return False
            record = {
                "state": state,
                "year": int(year),
                "deaths": workbook[sheet].cell(row=row, column=columns["deaths"]).value,
                "aadr": workbook[sheet].cell(row=row, column=columns["aadr"]).value,
            }
            if cause_column is not None:
                record["cause"] = workbook[sheet].cell(row=row, column=cause_column).value
            actual_health[key] = record
    if set(actual_health) != set(expected_health):
        return False
    for key, expected in expected_health.items():
        actual = actual_health[key]
        for role in ("state", "year", "deaths", "aadr"):
            if not source_value_equal(actual[role], expected[role], role):
                return False
        if "cause" in actual and not source_value_equal(actual["cause"], expected["cause"], "cause"):
            return False

    geo_table = find_semantic_table(workbook, "Geo_Map", ("state", "reporting_region", "source_key"))
    geo_header, geo_columns = geo_table
    actual_geo = {}
    for row in range(geo_header + 1, workbook["Geo_Map"].max_row + 1):
        state = workbook["Geo_Map"].cell(row=row, column=geo_columns["state"]).value
        if state in (None, ""):
            continue
        key = normalize_state_label(state)
        if key in actual_geo:
            return False
        actual_geo[key] = {
            "state": state,
            "reporting_region": workbook["Geo_Map"].cell(row=row, column=geo_columns["reporting_region"]).value,
            "source_key": workbook["Geo_Map"].cell(row=row, column=geo_columns["source_key"]).value,
        }
    if set(actual_geo) != set(expected_geo):
        return False
    return all(
        all(source_value_equal(actual_geo[state][role], expected_geo[state][role], role) for role in ("state", "reporting_region", "source_key"))
        for state in expected_geo
    )


def normalize_chart_ref(value):
    return (value or "").replace("'", "").replace("$", "").lower()


def chart_binding_ok(workbook):
    expected_categories = "visualization!a4:a5"
    expected_values = "visualization!b4:b5"
    for chart in workbook["Visualization"]._charts:
        if getattr(chart, "type", None) != "col":
            continue
        for series in getattr(chart, "ser", []):
            category_ref = None
            if getattr(series, "cat", None) is not None:
                str_ref = getattr(series.cat, "strRef", None)
                num_ref = getattr(series.cat, "numRef", None)
                category_ref = getattr(str_ref, "f", None) or getattr(num_ref, "f", None)
            value_ref = getattr(getattr(getattr(series, "val", None), "numRef", None), "f", None)
            if normalize_chart_ref(category_ref) == expected_categories and normalize_chart_ref(value_ref) == expected_values:
                return True
    return False


def analytical_alignment_ok(workbook, engine, oracle, source_root):
    table = find_semantic_table(workbook, "Analytical_Data", ("state", "year", "deaths", "aadr", "period"))
    header_row, columns = table
    cause_column = optional_header_column(workbook["Analytical_Data"], header_row, "cause")
    region_column = optional_header_column(workbook["Analytical_Data"], header_row, "reporting_region")
    source_key_column = optional_header_column(workbook["Analytical_Data"], header_row, "source_key")
    expected_health, expected_geo = load_public_sources(source_root)
    expected_rows = {
        (normalize_state_label(row[0]), row[1]): {"deaths": row[2], "aadr": row[3]}
        for row in oracle["rows"]
    }
    years = sorted({row[1] for row in oracle["rows"]})
    baseline_label = oracle.get("baseline_label", f"Baseline {years[0]}-{years[2]}")
    current_label = oracle.get("current_label", f"Current {years[3]}-{years[5]}")
    actual = {}
    for row in range(header_row + 1, workbook["Analytical_Data"].max_row + 1):
        state = engine.value("Analytical_Data", f"{col_name(columns['state'])}{row}")
        year = engine.value("Analytical_Data", f"{col_name(columns['year'])}{row}")
        if state in (None, "") and year in (None, ""):
            continue
        if state in (None, "") or not isinstance(year, (int, float)) or int(year) != year:
            return False
        key = (normalize_state_label(state), int(year))
        if key in actual:
            return False
        actual[key] = {
            "deaths": engine.value("Analytical_Data", f"{col_name(columns['deaths'])}{row}"),
            "aadr": engine.value("Analytical_Data", f"{col_name(columns['aadr'])}{row}"),
            "period": engine.value("Analytical_Data", f"{col_name(columns['period'])}{row}"),
        }
        if cause_column is not None:
            actual[key]["cause"] = engine.value("Analytical_Data", f"{col_name(cause_column)}{row}")
        if region_column is not None:
            actual[key]["reporting_region"] = engine.value("Analytical_Data", f"{col_name(region_column)}{row}")
        if source_key_column is not None:
            actual[key]["source_key"] = engine.value("Analytical_Data", f"{col_name(source_key_column)}{row}")
    if set(actual) != set(expected_rows):
        return False
    for key, expected in expected_rows.items():
        record = actual[key]
        if not source_value_equal(record["deaths"], expected["deaths"], "deaths") or not close(record["aadr"], expected["aadr"], 0.000001):
            return False
        expected_period = baseline_label if key[1] <= years[2] else current_label
        if normalize_period_label(record["period"]) != normalize_period_label(expected_period):
            return False
        if "cause" in record and record["cause"] != expected_health[key]["cause"]:
            return False
        if "reporting_region" in record and record["reporting_region"] != expected_geo[key[0]]["reporting_region"]:
            return False
        if "source_key" in record and record["source_key"] != expected_geo[key[0]]["source_key"]:
            return False
    return True


def metric_table(workbook):
    return find_semantic_table(workbook, "Period_Metrics", ("state", "baseline", "current", "absolute_change", "percent_change"))


def metric_address(workbook, engine, table, state, role):
    row = unique_state_row(workbook, engine, "Period_Metrics", table, state)
    return semantic_address(table, role, row)


def semantic_driver(context, name, defaults):
    semantic = context["contract"].get("semantic_drivers", {}).get(name, {})
    driver = dict(defaults)
    driver.update(semantic)
    legacy_prefix = "vermont" if name == "vermont" else "new_hampshire"
    legacy_cell = context["drivers"].get(f"{legacy_prefix}_source_cell")
    if legacy_cell and "raw_sheet" not in semantic:
        driver["raw_sheet"] = legacy_cell.split("!", 1)[0]
    if f"{legacy_prefix}_new_value" in context["drivers"] and "new_value" not in semantic:
        driver["new_value"] = context["drivers"][f"{legacy_prefix}_new_value"]
    return driver


def raw_driver_address(workbook, engine, driver):
    table = find_semantic_table(workbook, driver["raw_sheet"], ("state", "year", "deaths", "aadr"))
    row = unique_record_row(workbook, engine, driver["raw_sheet"], table, driver["state"], driver["year"])
    return f"{driver['raw_sheet']}!{semantic_address(table, 'aadr', row)}"


def report_state_table(workbook):
    return find_semantic_table(
        workbook,
        "Report",
        ("state", "baseline", "current", "absolute_change", "percent_change"),
        allow_missing=True,
    )


def state_name_in_text(text, state):
    return re.sub(r"[^a-z]", "", normalized_words(state)) in re.sub(r"[^a-z]", "", normalized_words(text))


def report_change_formula(workbook, engine, state, expected_value):
    matches = []
    sheet = workbook["Report"]
    for row in range(1, sheet.max_row + 1):
        label = " ".join(str(sheet.cell(row=row, column=column).value or "") for column in range(1, min(sheet.max_column, 12) + 1))
        label_words = normalized_words(label)
        if not state_name_in_text(label, state) or "change" not in label_words:
            continue
        for cell in formula_cells_in_row(workbook, "Report", row):
            try:
                observed = engine.value("Report", cell)
                numeric_claims = []
                if isinstance(observed, str):
                    numeric_claims = [float(value) for value in re.findall(r"[-+]?\d+(?:\.\d+)?", observed)]
                value_matches = close(observed, expected_value, 0.001) or any(
                    close(value, expected_value, 0.051) for value in numeric_claims
                )
                if formula_mentions_sheet(workbook, "Report", cell, "Period_Metrics") and value_matches:
                    matches.append(cell)
            except Exception:
                continue
    if len(matches) != 1:
        raise ValueError(f"REPORT_CHANGE_FORMULA_COUNT:{state}:{len(matches)}")
    return matches[0]


def highest_state_formula(workbook, engine, expected_state):
    matches = []
    sheet = workbook["Report"]
    for row in range(1, sheet.max_row + 1):
        context = " ".join(
            str(sheet.cell(row=context_row, column=column).value or "")
            for context_row in range(max(1, row - 2), row + 1)
            for column in range(1, min(sheet.max_column, 12) + 1)
        )
        context_words = normalized_words(context)
        if not ({"higher", "highest"} & set(context_words.split())) or "current" not in context_words:
            continue
        for cell in formula_cells_in_row(workbook, "Report", row):
            try:
                if (
                    formula_mentions_sheet(workbook, "Report", cell, "Period_Metrics")
                    and state_name_in_text(engine.value("Report", cell), expected_state)
                ):
                    matches.append(cell)
            except Exception:
                continue
    if len(matches) != 1:
        raise ValueError(f"HIGHEST_STATE_FORMULA_COUNT:{len(matches)}")
    return matches[0]


def report_claims_ok(workbook, engine, expected):
    state_table = report_state_table(workbook)
    if state_table is not None:
        for state in ("Vermont", "New Hampshire"):
            row = unique_state_row(workbook, engine, "Report", state_table, state)
            for role, expected_value in zip(("baseline", "current", "absolute_change", "percent_change"), expected["metrics"][state]):
                cell = semantic_address(state_table, role, row)
                if not formula_mentions_sheet(workbook, "Report", cell, "Period_Metrics") or not close(engine.value("Report", cell), expected_value, 0.001):
                    return False
    else:
        for state in ("Vermont", "New Hampshire"):
            report_change_formula(workbook, engine, state, expected["metrics"][state][2])
    highest_state_formula(workbook, engine, expected["highest_current"])
    return True


def report_dynamic_ok(workbook, engine, state, expected_current, expected_absolute):
    state_table = report_state_table(workbook)
    if state_table is not None:
        row = unique_state_row(workbook, engine, "Report", state_table, state)
        cell = semantic_address(state_table, "current", row)
        return formula_mentions_sheet(workbook, "Report", cell, "Period_Metrics") and close(engine.value("Report", cell), expected_current, 0.001)
    report_change_formula(workbook, engine, state, expected_absolute)
    return True


def visualization_current_address(workbook, engine, state):
    table = find_semantic_table(workbook, "Visualization", ("state", "current"))
    row = unique_state_row(workbook, engine, "Visualization", table, state)
    return semantic_address(table, "current", row)


def check_linkage_ok(workbook, engine):
    canonical_formulas = all(formula_present(workbook, f"Checks!B{row}") for row in range(4, 7))
    try:
        canonical_values = close(engine.value("Checks", "B4"), 12) and close(engine.value("Checks", "B5"), 2) and close(engine.value("Checks", "B6"), 0, 0.001)
    except Exception:
        canonical_values = False
    if canonical_formulas and canonical_values:
        return True

    analytical_ok = metric_ok = closure_ok = False
    sheet = workbook["Checks"]
    for row in range(1, sheet.max_row + 1):
        label = normalized_words(sheet.cell(row=row, column=1).value)
        for cell in formula_cells_in_row(workbook, "Checks", row):
            try:
                value = engine.value("Checks", cell)
            except Exception:
                continue
            if "analytical" in label and ({"row", "rows", "record", "records"} & set(label.split())) and close(value, 12):
                analytical_ok = True
            if "period" in label and "metric" in label and formula_mentions_sheet(workbook, "Checks", cell, "Period_Metrics") and isinstance(value, (int, float)) and value >= 2:
                metric_ok = True
            if (
                ({"error", "errors", "closure", "residual", "residuals"} & set(label.split()))
                and close(value, 0, 0.001)
                and formula_mentions_sheet(workbook, "Checks", cell, "Period_Metrics")
                and formula_mentions_sheet(workbook, "Checks", cell, "Report")
            ):
                closure_ok = True
    return analytical_ok and metric_ok and closure_ok


def task_checks(workbook, engine, oracle, context):
    checks, failures = {}, []
    expected = oracle
    try:
        alignment_ok = analytical_alignment_ok(workbook, engine, expected, context["source_root"])
    except Exception as exc:
        alignment_ok = False; failures.append(f"ALIGNMENT_EVALUATION_FAILED:{type(exc).__name__}")
    checks["R002"] = 1.0 if alignment_ok else 0.0

    try:
        metrics_table = metric_table(workbook)
        metrics_ok = True
        for state in ("Vermont", "New Hampshire"):
            expected_metric = expected["metrics"][state]
            for position, role in enumerate(("baseline", "current", "absolute_change", "percent_change")):
                address = metric_address(workbook, engine, metrics_table, state, role)
                metrics_ok = metrics_ok and formula_present(workbook, f"Period_Metrics!{address}") and close(engine.value("Period_Metrics", address), expected_metric[position], 0.001)
    except Exception as exc:
        metrics_ok = False; failures.append(f"METRIC_REPLAY_FAILED:{type(exc).__name__}")
    checks["R003"] = 1.0 if metrics_ok else 0.0

    years = sorted({row[1] for row in expected["rows"]})
    analytical_table = find_semantic_table(workbook, "Analytical_Data", ("state", "year", "deaths", "aadr", "period"))
    vermont_driver = semantic_driver(context, "vermont", {"raw_sheet": "Raw_Vermont", "state": "Vermont", "year": years[0], "field": "aadr", "new_value": 151.6})
    try:
        source = raw_driver_address(workbook, engine, vermont_driver)
        source_sheet, source_cell = source.split("!", 1)
        baseline = engine.value(source_sheet, source_cell)
        new_value = vermont_driver["new_value"]
        perturb = FormulaEngine(workbook, {source: new_value})
        new_baseline = expected["metrics"]["Vermont"][0] + (new_value - baseline) / 3
        analytical_row = unique_record_row(workbook, perturb, "Analytical_Data", analytical_table, vermont_driver["state"], vermont_driver["year"])
        analytical_address = semantic_address(analytical_table, "aadr", analytical_row)
        baseline_address = metric_address(workbook, perturb, metrics_table, "Vermont", "baseline")
        vt_dynamic = (
            formula_present(workbook, f"Analytical_Data!{analytical_address}")
            and formula_present(workbook, f"Period_Metrics!{baseline_address}")
            and close(perturb.value("Analytical_Data", analytical_address), new_value)
            and close(perturb.value("Period_Metrics", baseline_address), new_baseline, 0.001)
        )
    except Exception as exc:
        vt_dynamic = False; failures.append(f"VERMONT_SOURCE_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R004"] = 1.0 if vt_dynamic else 0.0

    new_hampshire_driver = semantic_driver(context, "new_hampshire", {"raw_sheet": "Raw_New_Hampshire", "state": "New Hampshire", "year": years[3], "field": "aadr", "new_value": 150.0})
    try:
        source = raw_driver_address(workbook, engine, new_hampshire_driver)
        source_sheet, source_cell = source.split("!", 1)
        baseline = engine.value(source_sheet, source_cell)
        new_value = new_hampshire_driver["new_value"]
        perturb = FormulaEngine(workbook, {source: new_value})
        new_current = expected["metrics"]["New Hampshire"][1] + (new_value - baseline) / 3
        analytical_row = unique_record_row(workbook, perturb, "Analytical_Data", analytical_table, new_hampshire_driver["state"], new_hampshire_driver["year"])
        analytical_address = semantic_address(analytical_table, "aadr", analytical_row)
        current_address = metric_address(workbook, perturb, metrics_table, "New Hampshire", "current")
        visualization_address = visualization_current_address(workbook, perturb, "New Hampshire")
        expected_absolute = new_current - expected["metrics"]["New Hampshire"][0]
        nh_dynamic = (
            formula_present(workbook, f"Analytical_Data!{analytical_address}")
            and formula_present(workbook, f"Period_Metrics!{current_address}")
            and formula_present(workbook, f"Visualization!{visualization_address}")
            and close(perturb.value("Analytical_Data", analytical_address), new_value)
            and close(perturb.value("Period_Metrics", current_address), new_current, 0.001)
            and close(perturb.value("Visualization", visualization_address), new_current, 0.001)
            and report_dynamic_ok(workbook, perturb, "New Hampshire", new_current, expected_absolute)
        )
    except Exception as exc:
        nh_dynamic = False; failures.append(f"NEW_HAMPSHIRE_SOURCE_PERTURBATION_FAILED:{type(exc).__name__}")
    checks["R005"] = 1.0 if nh_dynamic else 0.0

    try:
        report_ok = report_claims_ok(workbook, engine, expected)
    except Exception as exc:
        report_ok = False; failures.append(f"REPORT_EVALUATION_FAILED:{type(exc).__name__}")
    checks["R006"] = 1.0 if report_ok else 0.0

    visualization_formulas = all(formula_present(workbook, address) for address in ("Visualization!B4", "Visualization!B5"))
    chart_ok = chart_binding_ok(workbook)
    checks["R007"] = 1.0 if visualization_formulas and chart_ok else 0.0
    if not chart_ok: failures.append("CHART_SERIES_OR_CATEGORY_BINDING_MISMATCH")

    protected_ok = source_tables_match(workbook, context["source_root"])
    checks["R008"] = 1.0 if protected_ok else 0.0
    if not protected_ok: failures.append("PROTECTED_HEALTH_SOURCE_CHANGED")

    try:
        check_values_ok = check_linkage_ok(workbook, engine)
    except Exception as exc:
        check_values_ok = False; failures.append(f"CHECK_LINKAGE_FAILED:{type(exc).__name__}")
    checks["R009"] = 1.0 if check_values_ok else 0.0

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
