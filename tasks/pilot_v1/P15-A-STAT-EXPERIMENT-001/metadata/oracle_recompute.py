#!/usr/bin/env python3
"""Independent paired-t analysis of the documented sleep crossover data."""
import json
import math

GROUP_1 = [0.7, -1.6, -0.2, -1.2, -0.1, 3.4, 3.7, 0.8, 0.0, 2.0]
GROUP_2 = [1.9, 0.8, 1.1, 0.1, -0.1, 4.4, 5.5, 1.6, 4.6, 3.4]

def recompute(group_1=None, group_2=None, include=None):
    group_1 = GROUP_1 if group_1 is None else group_1
    group_2 = GROUP_2 if group_2 is None else group_2
    include = [True] * len(group_1) if include is None else include
    full_differences = [b - a for a, b in zip(group_1, group_2)]
    differences = [value for value, keep in zip(full_differences, include) if keep]
    n = len(differences)
    mean = sum(differences) / n
    mean_sq = sum((x - mean) ** 2 for x in differences) / (n - 1)
    sd = math.sqrt(mean_sq)
    se = sd / math.sqrt(n)
    t_stat = mean / se
    # Numerical Recipes continued fraction for the regularized beta function.
    def betacf(a, b, x):
        qab, qap, qam = a + b, a + 1.0, a - 1.0
        c, d = 1.0, 1.0 - qab * x / qap
        if abs(d) < 3e-14: d = 3e-14
        d = 1.0 / d; h = d
        for m in range(1, 201):
            m2 = 2 * m
            aa = m * (b - m) * x / ((qam + m2) * (a + m2))
            d = 1.0 + aa * d; d = 3e-14 if abs(d) < 3e-14 else d
            c = 1.0 + aa / c; c = 3e-14 if abs(c) < 3e-14 else c
            d = 1.0 / d; h *= d * c
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
            d = 1.0 + aa * d; d = 3e-14 if abs(d) < 3e-14 else d
            c = 1.0 + aa / c; c = 3e-14 if abs(c) < 3e-14 else c
            d = 1.0 / d; delta = d * c; h *= delta
            if abs(delta - 1.0) < 3e-12: break
        return h
    def betai(a, b, x):
        if x <= 0: return 0.0
        if x >= 1: return 1.0
        factor = math.exp(math.lgamma(a+b)-math.lgamma(a)-math.lgamma(b)+a*math.log(x)+b*math.log(1-x))
        return factor*betacf(a,b,x)/a if x < (a+1)/(a+b+2) else 1-factor*betacf(b,a,1-x)/b
    def two_tail(t_value, df):
        return betai(df/2, 0.5, df/(df+t_value*t_value))
    def critical_two_tail(alpha, df):
        lo, hi = 0.0, 20.0
        for _ in range(90):
            mid = (lo+hi)/2
            if two_tail(mid, df) > alpha: lo = mid
            else: hi = mid
        return (lo+hi)/2
    p_value = two_tail(abs(t_stat), n - 1)
    critical = critical_two_tail(0.05, n - 1)
    planned_n = math.ceil(((1.96 + 0.84) / 0.8) ** 2)
    return {"differences": differences, "full_differences": full_differences, "n": n, "mean_difference": mean, "sd_difference": sd, "se": se, "t": t_stat, "df": n - 1, "p_value": p_value, "ci_lower": mean - critical * se, "ci_upper": mean + critical * se, "planned_n": planned_n}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
