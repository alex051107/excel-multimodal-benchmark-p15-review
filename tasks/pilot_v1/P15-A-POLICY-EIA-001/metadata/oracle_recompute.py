#!/usr/bin/env python3
"""Independent public-electricity policy scenario replay."""
import json

DATA = {"coal": 675000.0, "gas": 1802000.0, "wind": 425000.0, "solar": 238000.0, "demand": 4000000.0, "coal_factor": 1.0, "gas_factor": 0.4}
BASE = {"coal_displacement": 0.0, "wind_uplift": 0.0, "solar_uplift": 0.0, "demand_growth": 0.01, "gas_factor": 1.0}
POLICY = {"coal_displacement": 0.08, "wind_uplift": 0.05, "solar_uplift": 0.12, "demand_growth": 0.01, "gas_factor": 1.0}

def case(x):
    demand = DATA["demand"] * (1 + x["demand_growth"])
    coal = DATA["coal"] * (1 - x["coal_displacement"])
    wind = DATA["wind"] * (1 + x["wind_uplift"])
    solar = DATA["solar"] * (1 + x["solar_uplift"])
    gas = (demand - coal - wind - solar) * x["gas_factor"]
    emissions = coal * DATA["coal_factor"] + gas * DATA["gas_factor"]
    return {"demand": demand, "coal": coal, "wind": wind, "solar": solar, "gas": gas, "emissions": emissions, "intensity": emissions / demand}

def recompute(policy_case=None):
    base, policy = case(BASE), case(POLICY if policy_case is None else policy_case)
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
