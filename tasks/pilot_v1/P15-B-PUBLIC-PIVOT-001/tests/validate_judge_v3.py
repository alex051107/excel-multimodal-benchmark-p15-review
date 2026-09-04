#!/usr/bin/env python3
"""Uniform entry point for the focused Judge V3 regression."""
from pathlib import Path
import runpy

runpy.run_path(str(Path(__file__).with_name("judge_v3_regression.py")), run_name="__main__")
