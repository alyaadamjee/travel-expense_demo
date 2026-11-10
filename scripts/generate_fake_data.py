"""Generate the demo's fake input files (no real people, companies, hotels, airlines or amounts).

usage: python scripts/generate_fake_data.py data/      (needs pandas + openpyxl)

Writes three Excel files shaped like the real exports the pipeline was built for:
  Travel agency bookings YTD 2026.xlsx   Air / Hotel / Rail Details tabs (+ a hidden old tab and a Summary tab)
  Approved expense claims 2026.xlsx      expense system export, one line per report (report total covers the whole report)
  Staff mapping.xlsx                     employee -> cost center, department, team (deliberately a few months stale)

Problems are planted on purpose so the notebook's fixes and checks have something to find:
  * September hotel stays repeated under the hotel group name and the brand name (duplicates), plus a few repeated air rows
  * hotel stays in Japan booked in yen in the USD column (one week in October)
  * about 1 in 5 expense reports shows only part of the report in the export ("not itemised")
  * 4 expense reports dated in December 2025 (outside the reporting window)
  * employees hired after the staff mapping was taken (no department / team)
"""
import datetime as dt
import os
import random
import sys

import pandas as pd

OUT = sys.argv[1] if len(sys.argv) > 1 else "data"
rng = random.Random(20261109)
os.makedirs(OUT, exist_ok=True)

FIRST = ["Avery", "Jordan", "Riley", "Casey", "Morgan", "Taylor", "Quinn", "Rowan", "Skyler", "Emerson", "Harper", "Reese",
         "Finley", "Sawyer", "Logan", "Parker", "Hayden", "Blake", "Dakota", "Elliot", "Kendall", "Marlow", "Sasha", "Remy",
         "Noel", "Arden", "Lennon", "Tatum", "Shiloh", "Wren", "Kai", "Milan", "Ari", "Jules", "Devon", "Robin", "Sage", "Peyton"]
LAST = ["Ashdown", "Bellweather", "Carrow", "Dunmore", "Ellery", "Fairbairn", "Garrow", "Hollins", "Ivers", "Jessop", "Kettering",
        "Lowell", "Marchetti", "Northam", "Okonkwo", "Pemberly", "Quarles", "Radcliffe", "Stroud", "Thackery", "Underhill",
        "Vasquez-Lane", "Whitlock", "Yardley", "Zeller", "Abernett", "Brisco", "Calloway", "Darrow", "Esterly", "Fenwick", "Galloway"]
TEAMS = ["NA-North", "NA-South", "NA-East", "NA-West", "NA-Central", "EU-North", "EU-South"]
DEPTS = {  # department -> (cost center group, sample cost centers)
    "Account Management": "Client Services", "Media Strategy": "Client Services", "Analytics": "Client Services",
    "Creative": "Client Services", "Technology": "Operations", "Data Engineering": "Operations", "Finance": "G&A",
    "People": "G&A", "Legal": "G&A", "Sales": "Growth", "Marketing": "Growth", "Leadership": "G&A",
}
ENTITY = {  # home country -> (company code, agency account name, currency, USD per unit)
    "United States": ("US01", "Example Co US Inc", "USD", 1.0), "United Kingdom": ("UK01", "Example Co UK Ltd", "GBP", 1.27),
    "Canada": ("CA01", "Example Co Canada Ltd", "CAD", 0.73), "Germany": ("DE01", "Example Co GmbH", "EUR", 1.08),
    "India": ("IN01", "Example Co India Pvt Ltd", "INR", 0.012),
}
HOME_WEIGHTS = [("United States", 60), ("United Kingdom", 15), ("Canada", 10), ("Germany", 8), ("India", 7)]

# --- employees + staff mapping ---------------------------------------------------------------------
people, used = [], set()
dept_list = list(DEPTS)
cc_by_dept = {d: [f"CC{1000 + 10 * i + k}" for k in range(2)] for i, d in enumerate(dept_list)}
while len(people) < 160:
    f, l = rng.choice(FIRST), rng.choice(LAST)
    if (f, l) in used:
        continue
    used.add((f, l))
    home = rng.choices([h for h, _ in HOME_WEIGHTS], [w for _, w in HOME_WEIGHTS])[0]
    dept = rng.choices(dept_list, [14, 12, 8, 6, 6, 4, 3, 3, 2, 5, 3, 2])[0]  # mostly client-facing
    team = (rng.choice([t for t in TEAMS if t.startswith("EU")]) if home in ("United Kingdom", "Germany") else rng.choice(TEAMS[:5])) \
        if DEPTS[dept] == "Client Services" else "No team"
    people.append({"first": f, "last": l, "id": str(200000 + len(people) * 7 + rng.randint(0, 6)), "home": home, "dept": dept,
                   "cc": rng.choice(cc_by_dept[dept]), "team": team, "travel": rng.choice([0, 0.1, 0.2, 0.3, 0.5, 0.8, 1.2]),
                   "new_hire": len(people) >= 150})  # last 10: hired after the mapping was taken

staff = pd.DataFrame([{"Employee #": p["id"], "Last Name": p["last"], "First Name": p["first"], "Cost Center": p["cc"],
                       "Department category": p["dept"], "CC Group 2": DEPTS[p["dept"]], "CC Group 3": p["dept"],
                       "Team": p["team"], "Headcount": 1} for p in people if not p["new_hire"]])
staff.to_excel(f"{OUT}/Staff mapping.xlsx", index=False, sheet_name="Staff")

# --- travel agency bookings -----------------------------------------------------------------------------
CITIES = {"United States": ["New York", "Chicago", "Seattle", "Austin", "Atlanta"], "United Kingdom": ["London", "Manchester"],
          "Canada": ["Toronto", "Vancouver"], "Germany": ["Berlin", "Munich"], "India": ["Bengaluru", "Mumbai"],
          "France": ["Paris"], "Japan": ["Tokyo"], "Singapore": ["Singapore"]}
HOTEL_GROUPS = {"Harborview Hotel Group": ["Harborview Suites", "Harborview Inn"], "Lumen Hospitality": ["Lumen Grand", "Lumen Stay"],
                "Northgate Resorts": ["Northgate House"], "Aldermoor Hotels": ["Aldermoor City", "Aldermoor Express"]}
AIRLINES = ["Skyward Air", "Bluecrest Airways", "Meridian Air", "Polaris Airlines", "Coastline Air"]
RAIL = ["Metroline Rail", "Crosscountry Express", "Eurolink Rail"]
PURPOSES = ["Client Meeting", "Client Work", "Internal Meeting", "Conference", "Training", "New Business", "Site Visit", None]
SOURCE = ["Online booking tool", "Agent assisted"]


def traveler(p):
    return f"{p['last'].upper()}/{p['first'].upper()} {rng.choice(['MR', 'MS', ''])}".strip()


def rand_date(month):
    return dt.date(2026, month, rng.randint(1, 28))


air, hotel, rail = [], [], []
travelers = [p for p in people if p["travel"] > 0]
for month in range(1, 11):
    for p in travelers:
        for _ in range(int(p["travel"]) + (1 if rng.random() < p["travel"] % 1 else 0)):
            dest = rng.choices(list(CITIES), [50, 14, 8, 8, 6, 5, 5, 4])[0]
            if dest == p["home"] and rng.random() < 0.5:
                dest = "United States" if p["home"] != "United States" else "United Kingdom"
            inv = rand_date(month)
            purpose = rng.choice(PURPOSES)
            agency = ENTITY[p["home"]][1]
            long_haul = dest != p["home"]
            air.append({"Traveler": traveler(p), "Invoice Date": inv, "Agency Name": agency, "Carrier Name": rng.choice(AIRLINES),
                        "Origin City": rng.choice(CITIES[p["home"]]), "Destination City": rng.choice(CITIES[dest]),
                        "Destination Country": dest, "Class Of Service": rng.choice(["Economy"] * 6 + ["Premium Economy", "Business"]),
                        "Total Ticket Amount (USD)": round(rng.uniform(900, 4200) if long_haul else rng.uniform(180, 750), 4),
                        "PURPOSE OF TRIP": purpose, "Booking Source": rng.choice(SOURCE)})
            nights = rng.choice([1, 2, 2, 3, 3, 4, 5])
            if rng.random() < 0.85:
                ci = inv + dt.timedelta(days=rng.randint(3, 20))
                group = rng.choice(list(HOTEL_GROUPS))
                hotel.append({"Traveler": traveler(p), "Invoice Date": inv, "Agency Name": agency,
                              "Hotel Name": rng.choice(HOTEL_GROUPS[group]), "Hotel Chain": group, "Hotel City": rng.choice(CITIES[dest]),
                              "Hotel Country": dest, "Check-In Date": ci, "Check-Out Date": ci + dt.timedelta(days=nights),
                              "Number Of Nights": nights, "Total Amount (USD)": round(nights * rng.uniform(140, 420), 4),
                              "PURPOSE OF TRIP": purpose, "Booking Source": rng.choice(SOURCE)})
            if dest in ("United Kingdom", "Germany", "France") and rng.random() < 0.5:
                rail.append({"Traveler": traveler(p), "Invoice Date": inv, "Agency Name": agency, "Carrier Name": rng.choice(RAIL),
                             "Destination Country": dest, "Ticket Amount (USD)": round(rng.uniform(40, 320), 4),
                             "PURPOSE OF TRIP": purpose, "Booking Source": rng.choice(SOURCE)})
air, hotel, rail = pd.DataFrame(air), pd.DataFrame(hotel), pd.DataFrame(rail)

# planted: September extract repeats hotel stays under the group name, and repeats some air rows
sep = hotel[pd.to_datetime(hotel["Invoice Date"]).dt.month == 9]
dup = sep.sample(frac=0.6, random_state=1).assign(**{"Hotel Name": lambda d: d["Hotel Chain"]})
air_dup = air[pd.to_datetime(air["Invoice Date"]).dt.month == 9].sample(12, random_state=2)
hotel = pd.concat([hotel, dup], ignore_index=True)
air = pd.concat([air, air_dup], ignore_index=True)
# planted: Tokyo stays one October week booked in yen in the USD column (~150 yen per USD)
tokyo = (hotel["Hotel Country"] == "Japan") & pd.to_datetime(hotel["Invoice Date"]).dt.month.eq(10)
hotel.loc[tokyo, "Total Amount (USD)"] = (hotel.loc[tokyo, "Total Amount (USD)"] * 150).round(0)
if not tokyo.any():  # make sure there is at least one
    hotel.loc[hotel.index[-1], ["Hotel Country", "Hotel City", "Total Amount (USD)"]] = ["Japan", "Tokyo", 3 * 260 * 150]

summary = pd.DataFrame({"Category": ["Air", "Hotel", "Rail"],
                        "Total (USD)": [air["Total Ticket Amount (USD)"].sum(), hotel["Total Amount (USD)"].sum(), rail["Ticket Amount (USD)"].sum()]})
with pd.ExcelWriter(f"{OUT}/Travel agency bookings YTD 2026.xlsx", engine="openpyxl") as w:
    summary.to_excel(w, sheet_name="Summary", index=False)
    air.to_excel(w, sheet_name="Air Details", index=False)
    hotel.to_excel(w, sheet_name="Hotel Details", index=False)
    rail.to_excel(w, sheet_name="Rail Details", index=False)
    air.head(50).to_excel(w, sheet_name="Air Details (old)", index=False)
    w.book["Air Details (old)"].sheet_state = "hidden"  # an old copy left in the workbook: must be ignored

# --- expense claims --------------------------------------------------------------------------------------
TYPES = {  # expense type -> (min, max) local amount
    "01. Hotel": (120, 900), "02. Airfare": (150, 1800), "03. Taxi": (12, 90), "04. Car Rental": (80, 600), "05. Fuel": (30, 120),
    "06. Parking": (10, 60), "07. Train": (20, 250), "08. Breakfast": (8, 30), "09. Lunch": (12, 45), "10. Dinner": (25, 160),
    "11. Working Meals": (40, 400), "12. Business Entertainment": (80, 900), "13. Staff Entertainment": (100, 1500),
    "14. Mobile/Cellular Phone": (40, 95), "15. Internet/Online Fees": (20, 80), "16. Professional Subscriptions/Fees": (15, 400),
    "17. Computer Hardware": (40, 1500), "18. Office Supplies": (10, 150), "19. Conference Fee/Seminar Fee": (200, 2500),
    "20. Tuition/Training": (150, 2000), "21. Gifts - Clients": (25, 200), "22. Visa Fees": (60, 250), "23. Personal Car Mileage": (20, 220),
}
TYPE_W = [6, 2, 12, 2, 3, 3, 2, 3, 5, 6, 4, 4, 2, 10, 4, 3, 2, 4, 2, 1, 1, 1, 3]
lines, rid = [], 0
for month in range(1, 11):
    for p in people:
        for _ in range(rng.choice([1, 1, 2, 2, 3, 4])):
            rid += 1
            code, _, cur, usd = ENTITY[p["home"]]
            rdate = rand_date(month)
            n_items = rng.choice([1, 1, 1, 1, 1, 2, 3, 4])
            items = [(t, round(rng.uniform(*TYPES[t]) / usd, 2)) for t in rng.choices(list(TYPES), TYPE_W, k=n_items)]  # local currency
            fx = usd * rng.uniform(0.98, 1.02)
            total = round(sum(a for _, a in items) * fx, 2)
            shown = items if rng.random() < 0.12 else items[:1]  # mostly one line per report; ~1 in 8 exported fully itemised
            for t, a in shown:
                lines.append({"Report ID": f"RPT{100000 + rid}", "Report Name": f"{rdate:%b} expenses", "Report Date": rdate,
                              "Approval Status": "Approved", "Employee": f"{p['last']}, {p['first']}", "Employee ID": p["id"],
                              "Company Code": code, "Country": p["home"], "Cost Center": p["cc"], "Expense Type": t,
                              "Expense amount": a, "Reimbursement Currency": cur, "Total Report Amount (rpt)": total,
                              "Payment Type": rng.choice(["Corporate Card", "Corporate Card", "Out of Pocket"]),
                              "Purpose": rng.choice(["Client visit", "Team offsite", "Monthly phone", "Conference", "Home office", ""])})
# planted: 4 reports dated December 2025 (outside the 2026 window)
for i in range(4):
    p = people[i]
    rid += 1
    lines.append({"Report ID": f"RPT{100000 + rid}", "Report Name": "Dec expenses", "Report Date": dt.date(2025, 12, 15 + i),
                  "Approval Status": "Approved", "Employee": f"{p['last']}, {p['first']}", "Employee ID": p["id"],
                  "Company Code": ENTITY[p["home"]][0], "Country": p["home"], "Cost Center": p["cc"], "Expense Type": "14. Mobile/Cellular Phone",
                  "Expense amount": 46.25, "Reimbursement Currency": "USD", "Total Report Amount (rpt)": 46.25,
                  "Payment Type": "Out of Pocket", "Purpose": "Monthly phone"})
claims = pd.DataFrame(lines)
with pd.ExcelWriter(f"{OUT}/Approved expense claims 2026.xlsx", engine="openpyxl") as w:
    claims.to_excel(w, sheet_name="Approved expenses", index=False)

print(f"{len(people)} employees ({len(staff)} in staff mapping); agency air {len(air)}, hotel {len(hotel)} (incl. {len(dup)} repeats), "
      f"rail {len(rail)}; expense lines {len(claims)} in {claims['Report ID'].nunique()} reports")
