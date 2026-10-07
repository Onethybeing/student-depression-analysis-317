"""Collect the numbers the slide deck needs into results/slides_data.json (read by build_slides.js)."""
import json
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RES = ROOT / "results"
T = {p.stem: pd.read_csv(p) for p in (RES / "tables").glob("*.csv")}
R = json.loads((RES / "results.json").read_text())

rates = T["depression_rates_by_group"]


def group_rates(char, order, labels=None):
    sub = rates[rates["Characteristic"] == char].set_index("Group").loc[order]
    return {"labels": labels or order, "values": sub["Depression rate (%)"].round(1).tolist(),
            "n": sub["Students"].astype(int).tolist()}


health = rates[rates["Characteristic"] == "Health condition"].set_index("Group").loc[["Good", "Normal", "Abnormal", "Critical"]]

ame = T["logit_marginal_effects_ame"]
comp = T["marginal_effects_comparison"].set_index("Variable")
key_vars = ["Single", "Mid living expense", "High living expense", "Male", "Rich class"]

roc = T["roc_curve_points"]
grid = np.round(np.linspace(0, 1, 41), 3)


def roc_on_grid(model):
    pts = roc[roc["Model"] == model]
    return [float(pts.loc[pts["False positive rate"] <= g + 1e-9, "True positive rate"].max()) for g in grid]


chi = T["chi_square_tests"]
tt = T["t_tests"].set_index("Variable")
hyp = T["hypothesis_tests"]

data = {
    "R": R,
    "rates": {
        "relationship": group_rates("Relationship status", ["Single", "In a Relationship", "Complicated"],
                                    ["Single", "In a relationship", "Complicated"]),
        "family": group_rates("Family background", ["Lower Middle class", "Middle class", "Rich class"],
                              ["Lower-middle", "Middle", "Rich"]),
        "living": group_rates("Living style", ["Bachelor", "Family"], ["Bachelor", "With family"]),
        "gender": group_rates("Gender", ["Female", "Male"]),
        "expense": group_rates("Living expense", ["Low", "Mid", "High"]),
    },
    "health": {"labels": health.index.tolist(), "depressed": health["Depression rate (%)"].round(1).tolist(),
               "n": health["Students"].astype(int).tolist()},
    "ame": [{"var": r.Variable, "pp": round(100 * r._2, 2), "p": float(r._5)} for r in ame.itertuples()],
    "compare": {"vars": key_vars,
                "logit": [round(100 * comp.loc[v, "Logit AME"], 2) for v in key_vars],
                "probit": [round(100 * comp.loc[v, "Probit AME"], 2) for v in key_vars],
                "lpm": [round(100 * comp.loc[v, "LPM coefficient"], 2) for v in key_vars]},
    "roc": {"grid": grid.tolist(), "logit": roc_on_grid("Logit"), "probit": roc_on_grid("Probit")},
    "chi": [{"var": r.Variable, "chi2": round(r._2, 2), "p": float(r._4), "v": round(r._5, 2)} for r in chi.itertuples()],
    "ttests": {v: {"no": round(tt.loc[v, "Mean (not depressed)"], 2), "yes": round(tt.loc[v, "Mean (depressed)"], 2),
                   "p": float(tt.loc[v, "t-test p-value"])} for v in ["CGPA", "Current semester"]},
    "classification": T["classification_tables"].query("Model == 'Logit'").round(4).to_dict("records"),
    "hypotheses": [{"h": r.Hypothesis, "var": r.Variable, "result": r.Result} for r in hyp.itertuples()],
}
assert data["ame"][0]["var"] == "Male" and ame.columns[1] == "dy/dx" and ame.columns[4] == "P>|z|"
(RES / "slides_data.json").write_text(json.dumps(data, indent=1))
print("Saved results/slides_data.json")
