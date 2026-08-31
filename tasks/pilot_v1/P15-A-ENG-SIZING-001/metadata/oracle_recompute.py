#!/usr/bin/env python3
"""Independent SI pump-sizing calculation."""
import json
import math

INPUT = {"flow_lps": 42.0, "static_head_m": 18.0, "pipe_length_m": 140.0, "diameter_mm": 150.0, "friction": 0.02, "efficiency": 0.72, "safety": 0.15, "density": 998.0, "gravity": 9.81}
CATALOG = [("P-120", 45, 25, 11), ("P-180", 58, 36, 18.5), ("P-250", 80, 48, 30)]

def recompute(flow_lps=None, diameter_mm=None):
    x = dict(INPUT)
    if flow_lps is not None: x["flow_lps"] = flow_lps
    if diameter_mm is not None: x["diameter_mm"] = diameter_mm
    q = x["flow_lps"] / 1000
    diameter = x["diameter_mm"] / 1000
    area = math.pi * diameter ** 2 / 4
    velocity = q / area
    velocity_head = velocity ** 2 / (2 * x["gravity"])
    friction_head = x["friction"] * (x["pipe_length_m"] / diameter) * velocity_head
    tdh = (x["static_head_m"] + friction_head) * (1 + x["safety"])
    hydraulic = x["density"] * x["gravity"] * q * tdh / 1000
    shaft = hydraulic / x["efficiency"]
    min_flow, min_motor = x["flow_lps"] * (1 + x["safety"]), shaft * (1 + x["safety"])
    eligibility = ["Eligible" if cap >= min_flow and head >= tdh and motor >= min_motor else "Ineligible" for _, cap, head, motor in CATALOG]
    selected = next((name for (name, _, _, _), status in zip(CATALOG, eligibility) if status == "Eligible"), "No eligible pump")
    return {
        "flow_m3s": q, "diameter_m": diameter, "area_m2": area,
        "velocity": velocity, "velocity_head": velocity_head, "friction_head": friction_head,
        "tdh": tdh, "hydraulic_power": hydraulic, "shaft_power": shaft,
        "minimum_flow": min_flow, "minimum_motor": min_motor,
        "eligibility": eligibility, "selected_pump": selected,
    }

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
