"""Generate the fictional ED_Visits_2025.csv used by the app.

Run once (python make_dataset.py). The seed is fixed, so every run
produces the same 300 rows and every student sees identical data.
"""
import numpy as np
import pandas as pd

rng = np.random.default_rng(1009)

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
SITES = ["Barrie", "Orillia", "Midland", "Collingwood"]
COMPLAINTS = ["Abdominal pain", "Chest pain", "Shortness of breath", "Injury/Fracture",
              "Fever", "Headache", "Laceration", "Mental health"]


def pick(options, weights):
    w = np.array(weights, dtype=float)
    return options[rng.choice(len(options), p=w / w.sum())]


rows = []
for _ in range(300):
    m = pick(list(range(12)), [13, 12, 9, 8, 7, 7, 8, 8, 7, 8, 9, 12])
    d = int(rng.integers(1, DAYS[m] + 1))
    site = pick(SITES, [40, 25, 18, 17])
    band = pick([0, 1, 2, 3], [16, 30, 30, 24])
    age = [0, 18, 40, 65][band] + int(rng.integers(0, [18, 22, 25, 30][band]))
    sex = "F" if rng.random() < 0.51 else "M"
    ctas = pick([1, 2, 3, 4, 5], [2, 12, 38, 33, 15])
    complaint = (pick(COMPLAINTS, [12, 30, 26, 10, 6, 6, 2, 8]) if ctas <= 2
                 else pick(COMPLAINTS, [18, 8, 10, 20, 14, 10, 12, 8]))
    base = [4, 18, 55, 85, 105][ctas - 1]
    base *= 1.25 if site == "Barrie" else 0.9 if site == "Collingwood" else 1
    base *= 1.15 if m in (0, 1, 11) else 1  # winter surge
    wait = max(0, round(base * np.exp(0.35 * rng.standard_normal())))
    if ctas == 1:
        disp = pick(["Admitted", "Transferred", "Discharged"], [70, 25, 5])
    elif ctas == 2:
        disp = pick(["Admitted", "Discharged", "Transferred", "LWBS"], [40, 52, 6, 2])
    else:
        disp = pick(["Discharged", "Admitted", "LWBS", "Transferred"], [78, 10, 9 + (ctas - 3) * 3, 3])
    r = rng.random()
    los = {"LWBS": 0.5 + r * 2, "Admitted": 8 + r * 16, "Transferred": 5 + r * 8}.get(
        disp, 1.5 + (6 - ctas) * 0.8 + r * 3)
    r = rng.random()
    cost = {"LWBS": 60 + r * 40, "Admitted": 2800 + r * 4200, "Transferred": 1500 + r * 1500}.get(
        disp, 250 + (6 - ctas) * 120 + r * 350)
    rows.append({
        "VisitDate": f"2025-{m + 1:02d}-{d:02d}", "Month": MONTHS[m], "Site": site,
        "Age": age, "AgeGroup": ["0-17", "18-39", "40-64", "65+"][band], "Sex": sex,
        "TriageLevel": ctas, "ChiefComplaint": complaint, "WaitTimeMin": wait,
        "LengthOfStayHrs": round(los, 1), "Disposition": disp, "CostCAD": round(cost),
    })

df = pd.DataFrame(rows).sort_values("VisitDate", kind="stable").reset_index(drop=True)
df.insert(0, "VisitID", [f"ED{i + 1:04d}" for i in range(len(df))])
df.to_csv("data/ED_Visits_2025.csv", index=False)
print(df.head(), "\n", len(df), "rows written")
