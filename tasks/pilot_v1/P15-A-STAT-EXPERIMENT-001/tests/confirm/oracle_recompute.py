#!/usr/bin/env python3
"""Independent CONFIRM paired analysis for batch-release cycle time."""
import json
import math

GROUP_1 = [14.2, 13.8, 15.1, 12.9, 14.7, 13.5, 16.0, 15.4, 14.0, 13.2]
GROUP_2 = [12.8, 12.9, 13.7, 12.1, 13.0, 12.4, 14.2, 13.6, 12.9, 12.0]


def _betacf(a, b, x):
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < 3e-14:
        d = 3e-14
    d = 1.0 / d
    h = d
    for m in range(1, 201):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        d = 3e-14 if abs(d) < 3e-14 else d
        c = 1.0 + aa / c
        c = 3e-14 if abs(c) < 3e-14 else c
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        d = 3e-14 if abs(d) < 3e-14 else d
        c = 1.0 + aa / c
        c = 3e-14 if abs(c) < 3e-14 else c
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 3e-12:
            break
    return h


def _betai(a, b, x):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    factor = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x))
    return factor * _betacf(a, b, x) / a if x < (a + 1) / (a + b + 2) else 1 - factor * _betacf(b, a, 1 - x) / b


def _two_tail(t_value, df):
    return _betai(df / 2, 0.5, df / (df + t_value * t_value))


def _critical_two_tail(alpha, df):
    low, high = 0.0, 20.0
    for _ in range(90):
        middle = (low + high) / 2
        if _two_tail(middle, df) > alpha:
            low = middle
        else:
            high = middle
    return (low + high) / 2


def recompute(group_1=None, group_2=None):
    group_1 = GROUP_1 if group_1 is None else group_1
    group_2 = GROUP_2 if group_2 is None else group_2
    differences = [after - before for before, after in zip(group_1, group_2)]
    n = len(differences)
    mean = sum(differences) / n
    sd = math.sqrt(sum((value - mean) ** 2 for value in differences) / (n - 1))
    se = sd / math.sqrt(n)
    t_stat = mean / se
    p_value = _two_tail(abs(t_stat), n - 1)
    critical = _critical_two_tail(0.05, n - 1)
    planned_n = math.ceil(((1.96 + 0.84) / 0.8) ** 2)
    return {
        "differences": differences, "n": n, "mean_difference": mean,
        "sd_difference": sd, "se": se, "t": t_stat, "df": n - 1,
        "p_value": p_value, "ci_lower": mean - critical * se,
        "ci_upper": mean + critical * se, "planned_n": planned_n,
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
