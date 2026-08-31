#!/usr/bin/env python3
"""Independent DCF replay; it never reads the reference workbook."""
import json

HISTORICAL_REVENUE = 391035.0
ASSUMPTIONS = {
    "growth": [0.06, 0.055, 0.05, 0.045, 0.04],
    "margin": [0.315, 0.317, 0.32, 0.322, 0.324],
    "tax": [0.21] * 5,
    "da": [0.029] * 5,
    "capex": [0.027] * 5,
    "nwc": [0.012] * 5,
    "wacc": 0.085,
    "terminal_growth": 0.025,
    "cash": 29943.0,
    "debt": 97800.0,
}

def recompute(growth=None, wacc=None, terminal_growth=None):
    a = ASSUMPTIONS
    growth = growth or a["growth"]
    wacc = a["wacc"] if wacc is None else wacc
    terminal_growth = a["terminal_growth"] if terminal_growth is None else terminal_growth
    revenue = []
    prior = HISTORICAL_REVENUE
    for rate in growth:
        prior *= 1 + rate
        revenue.append(prior)
    ebit = [revenue[i] * a["margin"][i] for i in range(5)]
    tax = [ebit[i] * a["tax"][i] for i in range(5)]
    nopat = [ebit[i] - tax[i] for i in range(5)]
    da = [revenue[i] * a["da"][i] for i in range(5)]
    capex = [revenue[i] * a["capex"][i] for i in range(5)]
    nwc = [revenue[i] * a["nwc"][i] for i in range(5)]
    fcf = [nopat[i] + da[i] - capex[i] - nwc[i] for i in range(5)]
    discount = [1 / (1 + wacc) ** (i + 1) for i in range(5)]
    pv = [fcf[i] * discount[i] for i in range(5)]
    terminal_fcf = fcf[-1] * (1 + terminal_growth)
    terminal_value = terminal_fcf / (wacc - terminal_growth)
    pv_terminal = terminal_value * discount[-1]
    enterprise_value = sum(pv) + pv_terminal
    net_debt = a["debt"] - a["cash"]
    sensitivity = {}
    for sensitivity_wacc in (0.075, 0.085, 0.095):
        sensitivity_discount = [1 / (1 + sensitivity_wacc) ** (i + 1) for i in range(5)]
        sensitivity_pv = sum(fcf[i] * sensitivity_discount[i] for i in range(5))
        for sensitivity_growth in (0.02, 0.025, 0.03):
            sensitivity[(sensitivity_wacc, sensitivity_growth)] = (
                sensitivity_pv
                + fcf[-1] * (1 + sensitivity_growth)
                / (sensitivity_wacc - sensitivity_growth)
                * sensitivity_discount[-1]
                - net_debt
            )
    return {
        "revenue": revenue,
        "ebit": ebit,
        "tax": tax,
        "nopat": nopat,
        "da": da,
        "capex": capex,
        "nwc": nwc,
        "fcf": fcf,
        "discount": discount,
        "pv": pv,
        "pv_forecast": sum(pv),
        "terminal_fcf": terminal_fcf,
        "terminal_value": terminal_value,
        "pv_terminal": pv_terminal,
        "cash": a["cash"],
        "debt": a["debt"],
        "net_debt": net_debt,
        "enterprise_value": enterprise_value,
        "equity_value": enterprise_value - net_debt,
        "sensitivity": sensitivity,
    }

def json_safe(value):
    """Serialize tuple-indexed sensitivity coordinates for CLI receipts only."""
    if isinstance(value, dict):
        return {
            ("|".join(str(part) for part in key) if isinstance(key, tuple) else str(key)): json_safe(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    return value


if __name__ == "__main__":
    print(json.dumps(json_safe(recompute()), sort_keys=True))
