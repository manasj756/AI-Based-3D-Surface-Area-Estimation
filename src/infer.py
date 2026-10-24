"""Single-volume inference with every estimator (used by sweep.py and demo.py)."""
from __future__ import annotations

import os

os.environ.setdefault("PYTORCH_ENABLE_MPS_FALLBACK", "1")

import numpy as np
import torch

from classical import all_estimates, mc_mesh, sample_surface_points
from data import points_to_tensor, volumes_to_tensor
from baselines import FEATURE_NAMES, features
from models import CNN3D, PointNet
from train import Y_SCALE
from voxelize import encode

_MODELS: dict = {}


def load_model(kind: str, N: int, root: str = ".", device="cpu"):
    key = (kind, N, root)
    if key not in _MODELS:
        path = os.path.join(root, "weights", f"{kind}_N{N}.pt")
        if not os.path.exists(path):
            _MODELS[key] = None
        else:
            if kind == "pointnet":
                m = PointNet()
            else:
                m = CNN3D(n_extra=len(FEATURE_NAMES) if kind == "cnnfeat" else 0)
            m.load_state_dict(torch.load(path, map_location=device, weights_only=True))
            m = m.to(device).eval()
            if kind == "cnnfeat":
                st = np.load(os.path.join(root, "weights", f"{kind}_N{N}_featstats.npz"))
                m.feat_mu, m.feat_sd = st["mu"], st["sd"]
            _MODELS[key] = m
    return _MODELS[key]


@torch.no_grad()
def estimate_all(img: np.ndarray, root: str = ".", device="cpu", seed: int = 0) -> dict:
    """All classical estimates plus CNN / PointNet corrections when weights exist
    for this grid size. `img` is a float volume with background ~0, object ~1."""
    N = img.shape[0]
    est = all_estimates(img)
    out = {"face": est["A_face"], "mc_bin": est["A_mc_bin"], "mc_grey": est["A_mc_grey"],
           "mc_smooth": est["A_mc_smooth"]}
    cnn = load_model("cnn", N, root, device)
    if cnn is not None:
        y = cnn(volumes_to_tensor(encode(img)[None], device)).item() * Y_SCALE
        out["cnn"] = est["A_mc_grey"] * float(np.exp(y))
    hyb = load_model("cnnfeat", N, root, device)
    if hyb is not None:
        import pandas as pd
        f = (features(pd.DataFrame([{**est, "N": N}])) - hyb.feat_mu) / hyb.feat_sd
        f = torch.from_numpy(f.astype(np.float32)).to(device)
        y = hyb(volumes_to_tensor(encode(img)[None], device), f).item() * Y_SCALE
        out["cnnfeat"] = est["A_mc_grey"] * float(np.exp(y))
    pn = load_model("pointnet", N, root, device)
    if pn is not None:
        v, f = mc_mesh(img)
        p = sample_surface_points(v, f, 2048, np.random.default_rng(seed))
        y = pn(points_to_tensor(p[None], N, device)).item() * Y_SCALE
        out["pointnet"] = est["A_mc_grey"] * float(np.exp(y))
    return out
