#!/usr/bin/env python3
"""Independent CONFIRM hydraulic sizing replay for a secondary-loop pump."""
import json
import math

INPUT = {"flow_lps": 36.0, "static_head_m": 16.0, "pipe_length_m": 210.0, "diameter_mm": 125.0, "friction": 0.021, "efficiency": 0.76, "safety": 0.12, "density": 997.0, "gravity": 9.81}
CATALOG = [("H-080", 38, 28, 15), ("H-140", 52, 40, 22), ("H-220", 74, 52, 30)]


def recompute(flow_lps=None):
    values = dict(INPUT)
    if flow_lps is not None:
        values["flow_lps"] = flow_lps
    flow_m3s = values["flow_lps"] / 1000
    diameter_m = values["diameter_mm"] / 1000
    area_m2 = math.pi * diameter_m ** 2 / 4
    velocity = flow_m3s / area_m2
    velocity_head = velocity ** 2 / (2 * values["gravity"])
    friction_head = values["friction"] * (values["pipe_length_m"] / diameter_m) * velocity_head
    tdh = (values["static_head_m"] + friction_head) * (1 + values["safety"])
    hydraulic_power = values["density"] * values["gravity"] * flow_m3s * tdh / 1000
    shaft_power = hydraulic_power / values["efficiency"]
    minimum_flow = values["flow_lps"] * (1 + values["safety"])
    minimum_motor = shaft_power * (1 + values["safety"])
    eligibility = [
        "Eligible" if capacity >= minimum_flow and head >= tdh and motor >= minimum_motor else "Ineligible"
        for _, capacity, head, motor in CATALOG
    ]
    selected_pump = next((name for (name, _, _, _), status in zip(CATALOG, eligibility) if status == "Eligible"), "No eligible pump")
    return {
        "flow_m3s": flow_m3s, "diameter_m": diameter_m, "area_m2": area_m2,
        "velocity": velocity, "velocity_head": velocity_head, "friction_head": friction_head,
        "tdh": tdh, "hydraulic_power": hydraulic_power, "shaft_power": shaft_power,
        "minimum_flow": minimum_flow, "minimum_motor": minimum_motor,
        "eligibility": eligibility, "selected_pump": selected_pump,
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
