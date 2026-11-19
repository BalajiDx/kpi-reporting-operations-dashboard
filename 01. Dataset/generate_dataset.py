"""
Generate the SYNTHETIC raw dataset for the KPI Reporting & Operations Dashboard project.

A content-services operation (data conversion, academic editing, typesetting) handles jobs from publisher clients.
Job-level data, monthly staffing (FTE) and client escalations, July 2023 - December 2024.
All data is randomly generated (seed 61); client names are invented. No real company data is used.

Run:  python generate_dataset.py
Output: ./Raw Dataset/{Jobs,Staffing,Escalations}.csv
(The raw files deliberately contain data-quality problems that clean_dataset.py fixes.)
"""
import os
import numpy as np
import pandas as pd

rng = np.random.default_rng(61)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Raw Dataset")
os.makedirs(OUT, exist_ok=True)

clients = ["Northbridge Press", "Larkfield Academic", "Bluewater Journals", "Meridian Books", "Oakline Publishing",
           "Harbourview Media", "Silverpine Education", "Redstone Scientific", "Westgate Law Reports", "Kestrel Medical"]
cw = np.array([0.18, 0.15, 0.13, 0.11, 0.1, 0.09, 0.08, 0.07, 0.05, 0.04])
teams = {  # team: (job types, pages range, standard turnaround days, base FTE)
    "Data Conversion": (["XML conversion", "ePub conversion", "Legacy digitisation"], (40, 600), 6, 22),
    "Academic Editing": (["Copyediting", "Proofreading", "Reference checking"], (10, 120), 5, 18),
    "Typesetting": (["Journal typesetting", "Book typesetting", "Corrections round"], (8, 400), 4, 14),
}
months = pd.period_range("2023-07", "2024-12", freq="M")
season = {9: 1.25, 10: 1.35, 11: 1.3, 3: 1.15}          # academic publishing peaks

jobs, jid = [], 1
staff = []
for m in months:
    for t, (jt, pr, sla, fte0) in teams.items():
        fte = fte0 * (1 + 0.012 * (m - months[0]).n) * rng.uniform(0.97, 1.03)
        if t == "Typesetting" and m.strftime("%Y-%m") in ("2024-09", "2024-10", "2024-11"):
            fte *= 0.78                                     # attrition: three typesetters left
        staff.append({"Month": m.strftime("%Y-%m"), "Team": t, "FTE": round(fte, 1), "PlannedFTE": round(fte0 * (1 + 0.012 * (m - months[0]).n), 1)})
        load = season.get(m.month, 1.0)
        n = int(rng.poisson({"Data Conversion": 230, "Academic Editing": 330, "Typesetting": 260}[t] * load * (1 + 0.01 * (m - months[0]).n)))
        capacity_ratio = fte / (fte0 * (1 + 0.012 * (m - months[0]).n)) / load
        for _ in range(n):
            rec = m.start_time + pd.Timedelta(days=int(rng.integers(0, m.days_in_month)), hours=int(rng.integers(8, 19)))
            due = rec + pd.Timedelta(days=sla + int(rng.integers(-1, 3)))
            strain = max(0.0, 1.05 - capacity_ratio)
            tat = max(1, rng.gamma(5, (sla * 0.6) / 5) * (1 + strain * 2.0))
            dl = rec + pd.Timedelta(days=float(tat))
            pages = int(rng.integers(*pr))
            err_rate = 0.0012 + strain * 0.004 + (0.0008 if t == "Data Conversion" else 0)
            errs = int(rng.poisson(pages * err_rate))
            rework = int(errs > 1 and rng.random() < 0.35)
            jobs.append({"JobID": f"JB{jid:06d}", "Client": rng.choice(clients, p=cw), "Team": t, "JobType": rng.choice(jt),
                         "ReceivedDate": rec.strftime("%Y-%m-%d"), "DueDate": due.strftime("%Y-%m-%d"),
                         "DeliveredDate": dl.strftime("%Y-%m-%d") if dl <= pd.Timestamp("2024-12-31") else "",
                         "Pages": pages, "QAErrors": errs, "Rework": rework,
                         "Status": "Delivered" if dl <= pd.Timestamp("2024-12-31") else "In progress"})
            jid += 1
jobs, staff = pd.DataFrame(jobs), pd.DataFrame(staff)

esc = []
late = jobs[(jobs.DeliveredDate != "") & (jobs.DeliveredDate > jobs.DueDate)]
for r in late.sample(frac=0.06, random_state=1).itertuples():
    esc.append({"EscalationID": None, "Date": r.DeliveredDate, "Client": r.Client, "Team": r.Team, "JobID": r.JobID, "Reason": "Late delivery"})
for r in jobs[jobs.Rework == 1].sample(frac=0.05, random_state=2).itertuples():
    if r.DeliveredDate:
        esc.append({"EscalationID": None, "Date": r.DeliveredDate, "Client": r.Client, "Team": r.Team, "JobID": r.JobID, "Reason": "Quality issue"})
esc = pd.DataFrame(esc).sort_values("Date").reset_index(drop=True)
esc["EscalationID"] = [f"ESC{i:04d}" for i in range(1, len(esc) + 1)]

# ----------------------------------------------- deliberate data-quality issues
jobs = pd.concat([jobs, jobs.sample(60, random_state=3)], ignore_index=True)            # duplicate job rows
i = rng.choice(len(jobs), 35, replace=False)
jobs.loc[i, "Team"] = ""                                                                 # team missing
j = jobs[jobs.DeliveredDate != ""].sample(25, random_state=4).index
jobs.loc[j, "DeliveredDate"] = (pd.to_datetime(jobs.loc[j, "ReceivedDate"]) - pd.Timedelta(days=3)).dt.strftime("%Y-%m-%d")  # delivered before received
k = rng.choice(len(jobs), 20, replace=False)
jobs.loc[k, "Pages"] = -jobs.loc[k, "Pages"]                                             # negative pages
jobs.loc[jobs.sample(30, random_state=5).index, "Client"] = jobs.Client.str.upper()      # client case

jobs.sample(frac=1, random_state=6).to_csv(os.path.join(OUT, "Jobs.csv"), index=False)
staff.to_csv(os.path.join(OUT, "Staffing.csv"), index=False)
esc[["EscalationID", "Date", "Client", "Team", "JobID", "Reason"]].to_csv(os.path.join(OUT, "Escalations.csv"), index=False)
print(f"jobs {len(jobs):,}  staffing rows {len(staff)}  escalations {len(esc)}")
