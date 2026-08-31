#!/usr/bin/env python3
"""Independent CONFIRM replay for the Seaboard Industrial Systems DCF instance."""
import json

HISTORICAL_REVENUE = 215000.0
ASSUMPTIONS = {
    "growth": [0.065, 0.055, 0.05, 0.045, 0.04],
    "margin": [0.16, 0.162, 0.165, 0.167, 0.17],
    "tax": [0.24] * 5,
    "da": [0.043] * 5,
    "capex": [0.048] * 5,
    "nwc": [0.075] * 5,
    "wacc": 0.092,
    "terminal_growth": 0.026,
    "cash": 18000.0,
    "debt": 46500.0,
}
SENSITIVITY_WACC = [0.082, 0.092, 0.102]
SENSITIVITY_GROWTH = [0.021, 0.026, 0.031]


def recompute(growth=None, wacc=None, terminal_growth=None):
    a = ASSUMPTIONS
    growth = a["growth"] if growth is None else growth
    wacc = a["wacc"] if wacc is None else wacc
    terminal_growth = a["terminal_growth"] if terminal_growth is None else terminal_growth
    revenue, prior = [], HISTORICAL_REVENUE
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
    for sensitivity_wacc in SENSITIVITY_WACC:
        sensitivity_discount = [1 / (1 + sensitivity_wacc) ** (i + 1) for i in range(5)]
        sensitivity_pv = sum(fcf[i] * sensitivity_discount[i] for i in range(5))
        for sensitivity_growth in SENSITIVITY_GROWTH:
            sensitivity[(sensitivity_wacc, sensitivity_growth)] = (
                sensitivity_pv
                + fcf[-1] * (1 + sensitivity_growth)
                / (sensitivity_wacc - sensitivity_growth)
                * sensitivity_discount[-1]
                - net_debt
            )
    return {
        "revenue": revenue, "ebit": ebit, "tax": tax, "nopat": nopat,
        "da": da, "capex": capex, "nwc": nwc, "fcf": fcf,
        "discount": discount, "pv": pv, "pv_forecast": sum(pv),
        "terminal_fcf": terminal_fcf, "terminal_value": terminal_value,
        "pv_terminal": pv_terminal, "cash": a["cash"], "debt": a["debt"],
        "net_debt": net_debt, "enterprise_value": enterprise_value,
        "equity_value": enterprise_value - net_debt, "sensitivity": sensitivity,
    }


if __name__ == "__main__":
    result = recompute()
    result["sensitivity"] = {f"{k[0]}|{k[1]}": v for k, v in result["sensitivity"].items()}
    print(json.dumps(result, sort_keys=True))
