#!/usr/bin/env python3
"""Independent CONFIRM replay for the Redwood FP&A debugging instance."""
import json

ASSUMPTIONS = {
    "volume": [8300, 8900, 9500], "price": [62, 64, 67],
    "margin": [0.42, 0.425, 0.43], "opex": [95000, 99000, 103000],
    "dso": [41, 40, 39], "interest": [0.07, 0.07, 0.068],
    "opening_debt": 175000, "repayment": [12000, 14000, 16000],
    "opening_ar": 47000,
}


def recompute():
    a = ASSUMPTIONS
    revenue = [a["volume"][i] * a["price"][i] for i in range(3)]
    contribution = [revenue[i] * a["margin"][i] - a["opex"][i] for i in range(3)]
    ar = [revenue[i] / 365 * a["dso"][i] for i in range(3)]
    changes = [ar[0] - a["opening_ar"], ar[1] - ar[0], ar[2] - ar[1]]
    wc_cash = [-value for value in changes]
    opening = [a["opening_debt"]]
    for index in range(2):
        opening.append(opening[-1] - a["repayment"][index])
    interest = [opening[i] * a["interest"][i] for i in range(3)]
    fcf = [contribution[i] + wc_cash[i] - interest[i] for i in range(3)]
    return {
        "revenue_2027": revenue[-1],
        "ar_2027": ar[-1],
        "interest_2027": interest[-1],
        "fcf_2027": fcf[-1],
        "dscr_2027": fcf[-1] / a["repayment"][-1],
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
