#!/usr/bin/env python3
"""Independent CONFIRM electricity-scenario replay for a regional generation case."""
import json

DATA = {"coal": 510000.0, "gas": 1420000.0, "wind": 515000.0, "solar": 305000.0, "demand": 3300000.0, "coal_factor": 0.96, "gas_factor": 0.38}
BASE = {"coal_displacement": 0.0, "wind_uplift": 0.0, "solar_uplift": 0.0, "demand_growth": 0.012, "gas_factor": 1.0}
POLICY = {"coal_displacement": 0.11, "wind_uplift": 0.07, "solar_uplift": 0.15, "demand_growth": 0.012, "gas_factor": 1.0}


def case(assumptions):
    demand = DATA["demand"] * (1 + assumptions["demand_growth"])
    coal = DATA["coal"] * (1 - assumptions["coal_displacement"])
    wind = DATA["wind"] * (1 + assumptions["wind_uplift"])
    solar = DATA["solar"] * (1 + assumptions["solar_uplift"])
    gas = (demand - coal - wind - solar) * assumptions["gas_factor"]
    emissions = coal * DATA["coal_factor"] + gas * DATA["gas_factor"]
    return {"demand": demand, "coal": coal, "wind": wind, "solar": solar, "gas": gas, "emissions": emissions, "intensity": emissions / demand}


def recompute(policy_case=None):
    base = case(BASE)
    policy = case(POLICY if policy_case is None else policy_case)
    return {
        "base": base,
        "policy": policy,
        "policy_emissions": policy["emissions"],
        "emissions_reduction": base["emissions"] - policy["emissions"],
        "intensity_reduction": base["intensity"] - policy["intensity"],
        "gas_change": policy["gas"] - base["gas"],
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
