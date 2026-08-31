#!/usr/bin/env python3
"""Independent FP&A operating-model replay."""
import json

ASSUMPTIONS = {
    "volume": [12000, 12900, 13700], "price": [48, 49, 50],
    "margin": [0.46, 0.465, 0.47], "opex": [130000, 134000, 138000],
    "dso": [44, 44, 43], "interest": [0.065, 0.065, 0.065],
    "opening_debt": 220000, "repayment": [15000, 18000, 20000],
    "opening_ar": 58000,
}

def recompute(price_2027=None, dso_2027=None):
    a = {key: list(value) if isinstance(value, list) else value for key, value in ASSUMPTIONS.items()}
    if price_2027 is not None:
        a["price"][-1] = price_2027
    if dso_2027 is not None:
        a["dso"][-1] = dso_2027
    revenue = [a["volume"][i] * a["price"][i] for i in range(3)]
    contribution = [revenue[i] * a["margin"][i] - a["opex"][i] for i in range(3)]
    ar = [revenue[i] / 365 * a["dso"][i] for i in range(3)]
    changes = [ar[0] - a["opening_ar"], ar[1] - ar[0], ar[2] - ar[1]]
    wc_cash = [-x for x in changes]
    opening = [a["opening_debt"]]
    for i in range(2):
        opening.append(opening[-1] - a["repayment"][i])
    interest = [opening[i] * a["interest"][i] for i in range(3)]
    fcf = [contribution[i] + wc_cash[i] - interest[i] for i in range(3)]
    return {
        "revenue_2027": revenue[-1],
        "ar_2027": ar[-1],
        "wc_cash_2027": wc_cash[-1],
        "contribution_2027": contribution[-1],
        "interest_2027": interest[-1],
        "fcf_2027": fcf[-1],
        "dscr_2027": fcf[-1] / a["repayment"][-1],
    }

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
