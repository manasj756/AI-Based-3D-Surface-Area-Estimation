"""Collect the numbers shown on the slides from the result CSVs into slides/data.json."""
import json
import os

import pandas as pd

R = "results"
ms = pd.read_csv(os.path.join(R, "metrics_split.csv"))
mn = pd.read_csv(os.path.join(R, "metrics_noise.csv"))
md = pd.read_csv(os.path.join(R, "metrics_diam.csv"))
sc = pd.read_csv(os.path.join(R, "sphere_curvature.csv"))


def get(df, **kw):
    q = df
    for k, v in kw.items():
        q = q[q[k] == v] if not isinstance(v, float) else q[(q[k] - v).abs() < 1e-9]
    return round(float(q.MAPE.iloc[0]), 2) if len(q) else None


methods = [m for m in ["face", "mc_bin", "mc_grey", "mc_smooth", "lookup", "rf", "pointnet", "cnn"]
           if m in set(ms.method)]
data = {
    "methods": methods,
    "test": {m: {str(N): get(ms, split="test", N=N, method=m) for N in (32, 64)} for m in methods},
    "bias": {m: {str(N): round(float(ms[(ms.split == "test") & (ms.N == N) & (ms.method == m)].bias.iloc[0]), 2)
                 for N in (32, 64)} for m in methods},
    "super": {m: {str(N): get(ms, split="held_super", N=N, method=m) for N in (32, 64)} for m in methods},
    "rough": {m: {str(N): get(ms, split="held_rough", N=N, method=m) for N in (32, 64)} for m in methods},
    "noise": {m: {f"{N}_{s}": get(mn, N=N, noise=s, method=m) for N in (32, 64) for s in (0.0, 0.05, 0.15)}
              for m in methods},
    "small": {m: get(md, diam_bin="<16", method=m) for m in methods},
    "sphere": sc.round(2).to_dict(orient="records"),
}
with open(os.path.join("slides", "data.json"), "w") as f:
    json.dump(data, f, indent=1)
print(json.dumps(data["test"], indent=0))
