# Databricks notebook source
# MAGIC %md
# MAGIC # Travel & Expense Demo - Spend Unified
# MAGIC Builds `<schema>.spend_unified` (one row per travel agency booking / expense claim line, USD) from the files in volume `raw_files`.
# MAGIC
# MAGIC **Files used**
# MAGIC * Travel agency – the newest file whose name contains "Travel agency bookings". Visible tabs only (hidden tabs are old copies); Summary tabs are skipped.
# MAGIC * Expense claims – every file whose name contains "Approved expense claims". Per Report ID the file with the most lines wins (newest on a tie), so an itemised export always beats a one-line-per-report export.
# MAGIC * Staff mapping – the file whose name contains "Staff mapping".
# MAGIC
# MAGIC **Fixes applied (every changed or removed row is written to `agency_adjustments` for review)**
# MAGIC 1. Travel agency hotel rows in a country listed in `LOCAL_CURRENCY_HOTELS` whose nightly rate is above the threshold are in local currency, not USD → converted.
# MAGIC 2. Travel agency duplicates removed – hotels on traveler + invoice date + check-in + check-out + nights + amount (the extract repeats a stay under the parent chain and the brand name); air and rail on identical rows.
# MAGIC 3. Expense claims – amount is the report total in USD (ties to the expense system). Lines shown in the file keep their expense type at the report's FX rate; when the file shows only part of a report the rest becomes "Not itemised in export".
# MAGIC
# MAGIC Dates: travel agency = invoice date, expense claims = report date (same basis as each file's Summary tab). Only rows dated inside PERIOD_START–PERIOD_END (calendar 2026) are kept; the rest are counted in `load_checks`.

# COMMAND ----------

# MAGIC %pip install openpyxl

# COMMAND ----------

dbutils.library.restartPython()

# COMMAND ----------

import os, re, json
import pandas as pd
import numpy as np
from openpyxl import load_workbook

dbutils.widgets.text("schema", "main.travel_expense_demo")
CATALOG, SCHEMA = dbutils.widgets.get("schema").split(".")
FOLDER = f"/Volumes/{CATALOG}/{SCHEMA}/raw_files/"
T = f"{CATALOG}.{SCHEMA}."

# Hotel rows in these countries above the nightly threshold are local currency (agency extract issue: Tokyo stays booked in yen).
# Rate = USD per 1 unit of local currency.
LOCAL_CURRENCY_HOTELS = {"Japan": {"currency": "JPY", "usd_per_unit": 1 / 150, "min_nightly": 5000}}

# Reporting window: rows dated outside it (agency invoice date, expense report date) are left out and counted in load_checks.
PERIOD_START, PERIOD_END = pd.Timestamp("2026-01-01"), pd.Timestamp("2026-12-31")

def files(pattern):
    fs = [f for f in os.listdir(FOLDER) if pattern.lower() in f.lower() and f.lower().endswith(".xlsx")]
    return sorted(fs, key=lambda f: os.path.getmtime(FOLDER + f), reverse=True)  # newest first

def visible_sheets(path):
    wb = load_workbook(path, read_only=True)
    out = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    wb.close()
    return out

agency_FILE = files("Travel agency bookings")[0]
EXP_FILES = files("Approved expense claims")
STAFF_FILE = files("Staff mapping")[0]
print("Agency:", agency_FILE); print("Expenses:", EXP_FILES); print("Staff:", STAFF_FILE)

COUNTRY = {"USA": "United States", "US": "United States", "UK": "United Kingdom"}
TITLES = {"MR", "MRS", "MS", "MISS", "DR", "MSTR", "PROF"}

def name_std(s):
    s = re.sub(r"\(.*?\)", "", str(s)).strip()
    last, _, first = s.partition("/") if "/" in s else s.partition(",")
    first = [w for w in first.split() if w.upper().strip(".") not in TITLES]
    return f"{last.strip()}, {' '.join(first)}".title().strip(", ")

# COMMAND ----------

# ---------- agency ----------
agency_path = FOLDER + agency_FILE
sheets = visible_sheets(agency_path)
adj = []  # adjustments log

def load_agency(kind):
    tabs = [s for s in sheets if s.startswith(kind) and "summary" not in s.lower()]
    d = pd.concat([pd.read_excel(agency_path, sheet_name=s).assign(_tab=s) for s in tabs], ignore_index=True)
    d.columns = [re.sub(r"\s+", " ", c.strip()) for c in d.columns]
    return d, tabs

SPEC = {  # kind: amount col, country col, detail col
    "Air":   ("Total Ticket Amount (USD)", "Destination Country", "Carrier Name"),
    "Hotel": ("Total Amount (USD)",        "Hotel Country",       "Hotel Name"),
    "Rail":  ("Ticket Amount (USD)",       "Destination Country", "Carrier Name"),
}
agency_raw_total, agency_parts, agency_tabs = 0.0, [], {}
for kind, (amt, ctry, det) in SPEC.items():
    d, tabs = load_agency(kind)
    agency_tabs[kind] = tabs
    d[amt] = d[amt].astype(float).round(2)  # file has up to 4 decimals; work in cents so the tie-out is exact
    agency_raw_total += d[amt].sum()
    d["_original_amount"] = d[amt]
    d["_currency_fix"] = None

    if kind == "Hotel":  # fix 1: local-currency hotel rows
        nightly = d[amt] / d["Number Of Nights"].where(d["Number Of Nights"] > 0)
        for c, rule in LOCAL_CURRENCY_HOTELS.items():
            bad = (d[ctry] == c) & (nightly > rule["min_nightly"])
            d.loc[bad, amt] = (d.loc[bad, amt] * rule["usd_per_unit"]).round(2)
            d.loc[bad, "_currency_fix"] = f"{rule['currency']}→USD @ {rule['usd_per_unit']:.6f}"
        key = ["Traveler", "Invoice Date", "Check-In Date", "Check-Out Date", "Number Of Nights", amt]
    else:
        key = [c for c in d.columns if not c.startswith("_")]

    dup = d.duplicated(key, keep="first")  # fix 2: duplicates
    for i, r in d[d["_currency_fix"].notna() | dup].iterrows():
        adj.append({
            "category": kind, "traveler": r["Traveler"], "invoice_date": pd.to_datetime(r["Invoice Date"]).date(),
            "detail": str(r[det]).strip(), "country": r[ctry], "tab": r["_tab"],
            "original_amount": round(float(r["_original_amount"]), 2),
            "corrected_amount": round(float(r[amt]), 2),
            "amount_kept_usd": 0.0 if dup[i] else round(float(r[amt]), 2),
            "action": ("Removed – duplicate" + (" (after currency fix)" if r["_currency_fix"] else "")) if dup[i] else "Converted – " + r["_currency_fix"],
        })
    d = d[~dup]
    agency_parts.append(pd.DataFrame({
        "source": "Travel agency", "category_group": "Travel", "category": kind,
        "detail": d[det].astype(str).str.strip(),
        "employee_name": d["Traveler"].map(name_std),
        "employee_id": None,
        "country": d[ctry].replace(COUNTRY), "country_basis": "destination",
        "entity": d["Agency Name"],
        "spend_date": pd.to_datetime(d["Invoice Date"]).dt.normalize(),
        "amount_usd": d[amt].round(2),
        "purpose_of_trip": d["PURPOSE OF TRIP"].astype("string").str.strip().str.title(),
        "description": None, "payment_method": d["Booking Source"],
        "report_id": None, "cost_center": None,
        "line_basis": np.where(d["_currency_fix"].notna(), "Currency corrected", "As booked"),
        "source_file": agency_FILE,
    }))
agency = pd.concat(agency_parts, ignore_index=True)
adj = pd.DataFrame(adj)
adj_delta = (adj["amount_kept_usd"] - adj["original_amount"]).sum() if len(adj) else 0.0
assert abs(agency["amount_usd"].sum() - (agency_raw_total + adj_delta)) < 0.05, "Travel agency total does not tie to file total less adjustments"
print(f"Travel agency file total {agency_raw_total:,.2f} | adjustments {adj_delta:,.2f} | kept {agency['amount_usd'].sum():,.2f} ({len(agency)} rows)")
if len(adj): print(adj.groupby(["category", "action"])[["original_amount", "amount_kept_usd"]].agg(["count", "sum"]).round(0))

# COMMAND ----------

# ---------- expense claims ----------
TOTAL, LINE = "Total Report Amount (rpt)", "Expense amount"
frames = []
for rank, f in enumerate(EXP_FILES):  # rank 0 = newest
    p = FOLDER + f
    for s in [s for s in visible_sheets(p) if s.lower().startswith("approved")]:
        x = pd.read_excel(p, sheet_name=s, dtype={"Employee ID": str})
        x.columns = [c.strip() for c in x.columns]
        x["_n"] = x.groupby(list(x.columns), dropna=False).cumcount()  # keeps genuine repeats inside one tab
        frames.append(x.assign(_file=f, _rank=rank))
allc = pd.concat(frames, ignore_index=True)
# monthly tabs overlap (a report can sit on the Jan and the Feb tab): drop rows repeated across tabs of the same file
data_cols = [c for c in allc.columns if not c.startswith("_")]
allc = allc.drop_duplicates(subset=data_cols + ["_file", "_n"]).drop(columns="_n")
allc = allc[allc["Approval Status"].astype(str).str.strip().eq("Approved")]
lines_per = allc.groupby(["Report ID", "_file", "_rank"]).size().rename("n").reset_index()
best = lines_per.sort_values(["Report ID", "n", "_rank"], ascending=[True, False, True]).drop_duplicates("Report ID")
exp = allc.merge(best[["Report ID", "_file"]], on=["Report ID", "_file"])
exp_report_total = exp.groupby("Report ID")[TOTAL].first().sum()
print(f"expense claims: {exp['Report ID'].nunique()} reports, {len(exp)} lines; report totals {exp_report_total:,.2f}")
print(best.groupby("_file").size().rename("reports taken from file"))

exp["Report Date"] = pd.to_datetime(exp["Report Date"])
exp["cur"] = exp["Reimbursement Currency"].astype(str).str.strip()
rep = exp.groupby("Report ID").agg(total=(TOTAL, "first"), line_sum=(LINE, "sum"), n=(LINE, "size"),
                                   cur=("cur", "first"), rdate=("Report Date", "first"))
rep["implied_fx"] = rep.total / rep.line_sum.where(rep.line_sum.abs() > 0.005)
rep["ym"] = rep.rdate.dt.to_period("M")
# The export shows ONE expense line per report while the total covers the whole report.
# Reference rate per currency & month = median implied rate, leaving out reports whose line is already in USD (implied rate 1.0,
# e.g. Canadian phone allowances).
cand = rep[(rep.implied_fx > 0) & ((rep.implied_fx - 1).abs() > 0.01)]
ref = cand.groupby(["cur", "ym"]).implied_fx.median().rename("ref_m")
ref_all = cand.groupby("cur").implied_fx.median().rename("ref_all")
rep = rep.join(ref, on=["cur", "ym"]).join(ref_all, on="cur")
rep["ref_fx"] = np.where(rep.cur == "USD", 1.0, rep.ref_m.fillna(rep.ref_all).fillna(1.0))
# complete = the line shown accounts for the report total (in local currency or already in USD); otherwise the rest is not itemised
tol = np.maximum(1.0, 0.03 * rep.total.abs())
rep["complete"] = rep.implied_fx.notna() & (rep.total <= np.maximum(rep.line_sum * rep.ref_fx, rep.line_sum) + tol)
rep["fx"] = np.where(rep.complete, rep.implied_fx, rep.ref_fx)
rep["rest_usd"] = np.where(rep.complete, 0.0, rep.total - rep.line_sum * rep.fx)
rep["basis"] = np.where(rep.n > 1, "Itemised", "Single line")
rep.loc[~rep.complete, "basis"] = "Partial – rest not itemised"

exp = exp.join(rep[["fx", "basis"]], on="Report ID")
exp["amount_usd"] = (exp[LINE] * exp["fx"]).round(2)
# put rounding pennies on the first line so every report ties exactly
first = ~exp.duplicated("Report ID")
diff = (rep.total - rep.rest_usd.round(2)) - exp.groupby("Report ID")["amount_usd"].sum()
exp.loc[first, "amount_usd"] += exp.loc[first, "Report ID"].map(diff).round(2)
exp["expense_type"] = exp["Expense Type"].astype(str).str.strip().str.replace(r"^\d+\.?\s*", "", regex=True).str.strip()

rest = rep[~rep.complete].join(exp[first].set_index("Report ID")[[c for c in exp.columns if c not in ("Report ID", "amount_usd", "expense_type", LINE)]], rsuffix="_c")
rest = rest.reset_index().assign(amount_usd=lambda r: r.rest_usd.round(2), expense_type="Not itemised in export")
exp = pd.concat([exp, rest[exp.columns.intersection(rest.columns)]], ignore_index=True)

CAT = {
 "Hotel": ("Travel", "Hotel"), "Airfare": ("Travel", "Air"), "Airline Fees": ("Travel", "Air"), "Train": ("Travel", "Rail"),
 "Taxi": ("Ground Transport", "Taxi/Car Service"), "Car Services": ("Ground Transport", "Taxi/Car Service"),
 "Car Rental": ("Ground Transport", "Car Rental"), "Fuel": ("Ground Transport", "Fuel/Parking/Tolls"),
 "Parking": ("Ground Transport", "Fuel/Parking/Tolls"), "Tolls/Road Charges": ("Ground Transport", "Fuel/Parking/Tolls"),
 "Personal Car Mileage": ("Ground Transport", "Mileage"), "Subway/Underground": ("Ground Transport", "Public Transit"),
 "Public Transport - Other": ("Ground Transport", "Public Transit"),
 "Breakfast": ("Meals", "Meals"), "Lunch": ("Meals", "Meals"), "Dinner": ("Meals", "Meals"),
 "Working Meals": ("Meals", "Meals"), "Refreshments - non alcohol": ("Meals", "Meals"),
 "Business Entertainment": ("Entertainment & Gifts", "Entertainment"), "Staff Entertainment": ("Entertainment & Gifts", "Entertainment"),
 "Event Catering": ("Entertainment & Gifts", "Events"), "Venue Cost": ("Entertainment & Gifts", "Events"),
 "Event Site Cost": ("Entertainment & Gifts", "Events"),
 "Gifts - Staff": ("Entertainment & Gifts", "Gifts"), "Gifts - Clients": ("Entertainment & Gifts", "Gifts"),
 "Staff Awards/Incentives": ("Entertainment & Gifts", "Gifts"),
 "Mobile/Cellular Phone": ("Telecom & Subscriptions", "Phone"), "Telephone/Fax": ("Telecom & Subscriptions", "Phone"),
 "Internet/Online Fees": ("Telecom & Subscriptions", "Internet/Online"), "Online Services": ("Telecom & Subscriptions", "Internet/Online"),
 "Professional Subscriptions/Fees": ("Telecom & Subscriptions", "Subscriptions"), "Newspapers/Magazines/Books": ("Telecom & Subscriptions", "Subscriptions"),
 "Computer Hardware": ("Office & Other", "IT"), "Computer Software": ("Office & Other", "IT"),
 "Office Supplies": ("Office & Other", "Office"), "Office Equipment": ("Office & Other", "Office"),
 "Home Office Equipment": ("Office & Other", "Office"), "Printing/Photocopying": ("Office & Other", "Office"),
 "Printing": ("Office & Other", "Office"), "Stationery": ("Office & Other", "Office"), "Postage": ("Office & Other", "Office"),
 "Courier/Shipping/Freight": ("Office & Other", "Office"),
 "Conference Fee/Seminar Fee": ("Office & Other", "Training & Conferences"), "Tuition/Training": ("Office & Other", "Training & Conferences"),
 "Visa Fees": ("Office & Other", "Fees"), "Currency Exchange Fees": ("Office & Other", "Fees"),
 "Regulatory filing / statutory fees": ("Office & Other", "Fees"),
 "Medical Other": ("Office & Other", "Other"),
 "Not itemised in export": ("Not itemised", "Not itemised in export"),
}
mapped = exp["expense_type"].map(CAT)
unmapped = sorted(exp.loc[mapped.isna(), "expense_type"].unique())
exp["category_group"] = mapped.map(lambda x: x[0] if isinstance(x, tuple) else "Office & Other")
exp["category"] = mapped.map(lambda x: x[1] if isinstance(x, tuple) else "Other")

expenses = pd.DataFrame({
    "source": "Expense claims", "category_group": exp["category_group"], "category": exp["category"],
    "detail": exp["expense_type"],
    "employee_name": exp["Employee"].map(name_std),
    "employee_id": exp["Employee ID"].astype(str).str.strip(),
    "country": exp["Country"].astype(str).str.strip().replace(COUNTRY), "country_basis": "employee_entity",
    "entity": exp["Company Code"],
    "spend_date": pd.to_datetime(exp["Report Date"]).dt.normalize(),
    "amount_usd": exp["amount_usd"],
    "purpose_of_trip": None, "description": exp["Purpose"], "payment_method": exp["Payment Type"],
    "report_id": exp["Report ID"], "cost_center": exp["Cost Center"].astype(str).str.strip(),
    "line_basis": exp["basis"], "source_file": exp["_file"],
})
assert abs(expenses["amount_usd"].sum() - exp_report_total) < 0.05, "expense claims does not tie to the report totals"
print("unmapped expense claims types (→ Office & Other / Other):", unmapped)
print(rep.groupby("basis").agg(reports=("total", "size"), usd=("total", "sum"), not_itemised=("rest_usd", "sum")).round(0))

# COMMAND ----------

# ---------- people: staff mapping ----------
m = pd.read_excel(FOLDER + STAFF_FILE)
m = m.rename(columns={m.columns[-1]: "headcount_flag"}).drop_duplicates()
m["Cost Center"] = m["Cost Center"].astype(str).str.strip()
m["emp_id"] = m["Employee #"].astype(str).str.strip().str.replace(r"\.0\b", "", regex=True)

def k(s):  # compare names on letters only: O'Donnell = Odonnell, Drummond Smith = Drummond-Smith
    return re.sub(r"[^a-z]", "", str(s).lower())
m["last_key"] = m["Last Name"].map(k)
m["first_key"] = m["First Name"].astype(str).str.split().str[0].map(k)

m["team"] = m["Team"].astype(str).str.strip()
COLS = {"Department category": "department", "CC Group 2": "cc_group_2", "CC Group 3": "cc_group_3"}

dept_by_cc = m.drop_duplicates("Cost Center").set_index("Cost Center")[list(COLS)].rename(columns=COLS)
by_id = m.drop_duplicates("emp_id").set_index("emp_id")[list(COLS) + ["team"]].rename(columns=COLS)

def split_key(name):
    last, _, first = str(name).partition(", ")
    return k(last), (k(first.split()[0]) if first.split() else "")

def candidates(table, last, first, lcol, fcol):
    c = table[table[lcol] == last]
    return c[c[fcol].map(lambda f: bool(f) and bool(first) and (first.startswith(f) or f.startswith(first)))]

unified = pd.concat([agency, expenses], ignore_index=True)

# agency travellers have no employee ID or cost center: borrow them from the same person's expense claims claims (one person only)
cp = expenses.assign(lk=expenses.employee_name.map(lambda n: split_key(n)[0]), fk=expenses.employee_name.map(lambda n: split_key(n)[1]))
cp = cp.groupby(["employee_id", "lk", "fk"]).cost_center.agg(lambda s: s.value_counts().index[0]).rename("cc").reset_index()
via = {}
for n in unified.loc[unified.source == "Travel agency", "employee_name"].unique():
    c = candidates(cp, *split_key(n), "lk", "fk").drop_duplicates("employee_id")
    if len(c) == 1: via[n] = (c.iloc[0].employee_id, c.iloc[0].cc)
b = unified.source.eq("Travel agency") & unified.employee_name.isin(via)
unified.loc[b, "employee_id"] = unified.loc[b, "employee_name"].map(lambda n: via[n][0])
cc_used = unified["cost_center"].where(unified.source.eq("Expense claims"), unified["employee_name"].map(lambda n: via.get(n, (None, None))[1]))

names = {}
for n in unified["employee_name"].unique():
    c = candidates(m, *split_key(n), "last_key", "first_key").drop_duplicates("emp_id")
    names[n] = {} if len(c) != 1 else {"department": c.iloc[0]["Department category"], "cc_group_2": c.iloc[0]["CC Group 2"],
                                       "cc_group_3": c.iloc[0]["CC Group 3"], "team": c.iloc[0]["team"]}
by_name = pd.DataFrame.from_dict(names, orient="index").reindex(index=list(names), columns=["department", "cc_group_2", "cc_group_3", "team"])

idm = unified["employee_id"].isin(by_id.index)
nm = unified["employee_name"].map(lambda n: bool(names.get(n)))
ccm = cc_used.isin(dept_by_cc.index)
for col in ["department", "cc_group_2", "cc_group_3", "team"]:
    v = unified["employee_name"].map(by_name[col])
    v = v.where(~idm, unified["employee_id"].map(by_id[col]))
    if col != "team":  # department comes from the cost center the expense was charged to, when known
        v = v.where(~ccm, cc_used.map(dept_by_cc[col]))
    unified[col] = v.fillna("Unmapped").astype(str)
unified["match_method"] = np.select([ccm & unified.source.eq("Expense claims"), ccm, idm, nm],
                                    ["cost_center", "cost_center_via_expenses", "employee_id", "name"], "none")
unified["person_match"] = np.select([idm, nm, unified.employee_id.notna()], ["employee_id", "name", "expenses_only"], "none")

in_period = unified["spend_date"].between(PERIOD_START, PERIOD_END)
out_of_period = unified[~in_period]
print(f"left out – dated outside {PERIOD_START.date()} to {PERIOD_END.date()}:", len(out_of_period), "rows", round(out_of_period["amount_usd"].sum(), 2))
unified = unified[in_period].copy()
unified["month"] = unified["spend_date"].dt.to_period("M").dt.to_timestamp()
unified["is_credit"] = unified["amount_usd"] < 0
print(unified.groupby(["source", "match_method"])["amount_usd"].agg(["count", "sum"]).round(0))
print("TOTAL:", len(unified), "rows,", round(unified["amount_usd"].sum(), 2))

# COMMAND ----------

# ---------- save ----------
from pyspark.sql import functions as F

def save(df, name, comment):
    df = df.copy()
    for c in df.columns:  # text columns as plain strings so Spark infers one type
        if df[c].dtype == object and c not in ("invoice_date",):
            df[c] = df[c].map(lambda v: None if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA else str(v))
    sdf = spark.createDataFrame(df.astype(object).where(df.notna(), None))
    for c, t in df.dtypes.items():
        if c in ("spend_date", "month", "invoice_date"): sdf = sdf.withColumn(c, F.to_date(c))
        elif c in ("amount_usd", "original_amount", "amount_kept_usd", "corrected_amount", "value"): sdf = sdf.withColumn(c, F.col(c).cast("decimal(14,2)"))
    sdf.write.mode("overwrite").option("overwriteSchema", "true").saveAsTable(T + name)
    spark.sql(f"COMMENT ON TABLE {T}{name} IS '{comment}'")
    return sdf.count()

cols = ["source", "category_group", "category", "detail", "employee_name", "employee_id", "country", "country_basis", "entity",
        "spend_date", "month", "amount_usd", "is_credit", "purpose_of_trip", "description", "payment_method", "report_id",
        "cost_center", "department", "cc_group_2", "cc_group_3", "team", "match_method", "person_match", "line_basis", "source_file"]
n = save(unified[cols], "spend_unified",
         "Travel agency bookings + approved expense claims in USD. Built by notebook Travel & Expense Demo - Spend Unified. Agency rows dated by invoice date, expense claims by report date. line_basis explains how each amount was derived.")
print("spend_unified:", n)
if len(adj):
    save(adj, "agency_adjustments", "Travel agency rows the notebook converted (local currency booked as USD) or removed (duplicates). amount_kept_usd = what is left in spend_unified.")

checks = pd.DataFrame([
    {"check": "Travel agency: file total (USD as delivered)", "value": round(agency_raw_total, 2)},
    {"check": "Travel agency: currency corrections (local currency booked as USD)", "value": round(float((adj.corrected_amount - adj.original_amount).sum()) if len(adj) else 0.0, 2)},
    {"check": "Travel agency: duplicates removed", "value": round(float(-adj.loc[adj.action.str.startswith("Removed"), "corrected_amount"].sum()) if len(adj) else 0.0, 2)},
    {"check": "Travel agency: dated outside the reporting window (left out)", "value": round(float(-out_of_period.loc[out_of_period.source.eq("Travel agency"), "amount_usd"].sum()), 2)},
    {"check": "Travel agency: in spend_unified", "value": round(float(unified.loc[unified.source.eq("Travel agency"), "amount_usd"].sum()), 2)},
    {"check": "Expense claims: report totals (approved)", "value": round(float(exp_report_total), 2)},
    {"check": "Expense claims: not itemised in export", "value": round(float(rep.rest_usd.sum()), 2)},
    {"check": "Expense claims: dated outside the reporting window (left out)", "value": round(float(-out_of_period.loc[out_of_period.source.eq("Expense claims"), "amount_usd"].sum()), 2)},
    {"check": "Expense claims: in spend_unified", "value": round(float(unified.loc[unified.source.eq("Expense claims"), "amount_usd"].sum()), 2)},
])
checks["files"] = json.dumps({"agency": agency_FILE, "agency_tabs": agency_tabs, "expenses": EXP_FILES, "staff": STAFF_FILE})
checks["run_at"] = pd.Timestamp.now(tz="UTC").strftime("%Y-%m-%d %H:%M UTC")
save(checks, "load_checks", "Tie-out of the last spend_unified build: file totals, adjustments, totals loaded.")
display(checks)
