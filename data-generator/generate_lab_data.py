"""Synthetic data generator for the lab analytics Power BI reports.

Creates a fictional contract testing laboratory ("Contoso Analytical Labs")
with ~1,500 client facilities and ~770,000 test records (2023 to Oct 2026),
plus quality-review timings, OOS events, product orders, and consulting work.

Every name, ID, price and date is randomly generated. File names, sheet names,
table names and column names match what the Power BI reports expect, so the
reports can be re-pointed to this folder without changing any DAX.

Usage:
    pip install numpy faker xlsxwriter
    python generate_lab_data.py --out ../data/lab --scale 1.0

--scale 0.1 builds a 10% sample (fast, good for a GitHub "sample" folder).
"""
import argparse
import datetime as dt
import os
import random
from collections import defaultdict

import numpy as np
from faker import Faker

from xlsx_tables import StreamingWorkbook

SEED = 20261117
COMPANY = "Contoso"
DATA_END = dt.date(2026, 10, 30)          # last submission date in the data
QA_END = dt.date(2026, 9, 30)             # quality/OOS extracts run Jan-Sep 2026
YEAR_TESTS = {2023: 85_000, 2024: 125_000, 2025: 232_000, 2026: 330_000}
FACILITY_MAIN = f"{COMPANY} - US Central"
FACILITY_CA = f"{COMPANY} - Canada"

# --------------------------------------------------------------------------- #
# Reference data (generic, public industry terms)
# --------------------------------------------------------------------------- #
# analyte: (category, umbrella, default dosage forms, controlled class, hazardous)
ANALYTES = {
    "Semaglutide": ("GLP-1 / Incretin (Weight Management/Diabetes)", "GLP / Weight", ["Injectable", "Solution"], "N/A", "N"),
    "Tirzepatide": ("GLP-1 / Incretin (Weight Management/Diabetes)", "GLP / Weight", ["Injectable", "Solution"], "N/A", "N"),
    "Liraglutide": ("GLP-1 / Incretin (Weight Management/Diabetes)", "GLP / Weight", ["Injectable"], "N/A", "N"),
    "Testosterone Cypionate": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable", "Injectable"], "Class III", "N"),
    "Testosterone Enanthate": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable"], "Class III", "N"),
    "Estradiol Valerate": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable"], "N/A", "Y"),
    "Estradiol Cypionate": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable"], "N/A", "Y"),
    "Progesterone": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable", "Cream (gm)", "Capsule"], "N/A", "Y"),
    "Nandrolone Decanoate": ("Hormone Replacement / Endocrine", "Endocrine", ["Oil Injectable"], "Class III", "N"),
    "Gonadorelin": ("Other Peptides", "Peptides", ["Lyophilized", "Injectable"], "N/A", "N"),
    "Sermorelin Acetate": ("Other Peptides", "Peptides", ["Lyophilized", "Injectable"], "N/A", "N"),
    "Ipamorelin": ("Other Peptides", "Peptides", ["Injectable", "Lyophilized"], "N/A", "N"),
    "Oxytocin": ("Other Peptides", "Peptides", ["Injectable", "Solution"], "N/A", "N"),
    "Methylcobalamin": ("Vitamins & Cofactors", "Vitamins", ["Injectable", "Solution"], "N/A", "N"),
    "Hydroxocobalamin": ("Vitamins & Cofactors", "Vitamins", ["Injectable"], "N/A", "N"),
    "Cyanocobalamin": ("Vitamins & Cofactors", "Vitamins", ["Injectable"], "N/A", "N"),
    "Ascorbic Acid": ("Vitamins & Cofactors", "Vitamins", ["Injectable", "Aqueous Solution"], "N/A", "N"),
    "Biotin": ("Vitamins & Cofactors", "Vitamins", ["Injectable", "Capsule"], "N/A", "N"),
    "Pyridoxine Hydrochloride": ("Vitamins & Cofactors", "Vitamins", ["Injectable"], "N/A", "N"),
    "Thiamine Hydrochloride": ("Vitamins & Cofactors", "Vitamins", ["Injectable"], "N/A", "N"),
    "Glutathione": ("Nutraceuticals / Metabolic Support", "Nutraceuticals", ["Injectable", "Aqueous Solution"], "N/A", "N"),
    "NAD+": ("Nutraceuticals / Metabolic Support", "Nutraceuticals", ["Injectable", "Lyophilized"], "N/A", "N"),
    "L-Carnitine": ("Nutraceuticals / Metabolic Support", "Nutraceuticals", ["Injectable"], "N/A", "N"),
    "Lipo-C Blend": ("Nutraceuticals / Metabolic Support", "Nutraceuticals", ["Injectable"], "N/A", "N"),
    "Papaverine / Phentolamine / Alprostadil": ("Trimix / ED", "Urology", ["Injectable"], "N/A", "N"),
    "Alprostadil": ("Trimix / ED", "Urology", ["Injectable"], "N/A", "N"),
    "Bevacizumab": ("Ophthalmic", "Ophthalmic", ["Injectable"], "N/A", "Y"),
    "Moxifloxacin": ("Ophthalmic", "Ophthalmic", ["Solution", "Injectable"], "N/A", "N"),
    "Atropine Sulfate": ("Ophthalmic", "Ophthalmic", ["Solution"], "N/A", "N"),
    "Tri-Moxi (Triamcinolone / Moxifloxacin)": ("Ophthalmic", "Ophthalmic", ["Suspension"], "N/A", "N"),
    "Ketamine Hydrochloride": ("Local Anesthetics / Procedural", "Anesthetics", ["Injectable", "Solution"], "Class III", "N"),
    "Lidocaine Hydrochloride": ("Local Anesthetics / Procedural", "Anesthetics", ["Injectable", "Cream (gm)"], "N/A", "N"),
    "Bupivacaine Hydrochloride": ("Local Anesthetics / Procedural", "Anesthetics", ["Injectable"], "N/A", "N"),
    "Ropivacaine Hydrochloride": ("Local Anesthetics / Procedural", "Anesthetics", ["Injectable"], "N/A", "N"),
    "Hydromorphone Hydrochloride": ("Opioids / Analgesic Adjuncts", "Pain / Opioid Adjunct", ["Injectable"], "Class II", "N"),
    "Fentanyl Citrate": ("Opioids / Analgesic Adjuncts", "Pain / Opioid Adjunct", ["Injectable"], "Class II", "N"),
    "Morphine Sulfate": ("Opioids / Analgesic Adjuncts", "Pain / Opioid Adjunct", ["Injectable"], "Class II", "N"),
    "Ketorolac Tromethamine": ("Opioids / Analgesic Adjuncts", "Pain / Opioid Adjunct", ["Injectable"], "N/A", "N"),
    "Gabapentin": ("Neuropathic Pain / Analgesia", "Neuropathic", ["Cream (gm)", "Capsule"], "N/A", "N"),
    "Baclofen": ("Neuropathic Pain / Analgesia", "Neuropathic", ["Injectable", "Cream (gm)"], "N/A", "N"),
    "Methotrexate": ("Chemo / Onc", "Chemo / Onc", ["Injectable"], "N/A", "Y"),
    "Fluorouracil": ("Chemo / Onc", "Chemo / Onc", ["Injectable", "Cream (gm)"], "N/A", "Y"),
    "Mitomycin": ("Chemo / Onc", "Chemo / Onc", ["Lyophilized", "Solution"], "N/A", "Y"),
    "Phenylephrine Hydrochloride": ("Cardiovascular / Hemodynamics", "Cardiovascular", ["Injectable"], "N/A", "N"),
    "Norepinephrine Bitartrate": ("Cardiovascular / Hemodynamics", "Cardiovascular", ["Injectable"], "N/A", "N"),
    "Ephedrine Sulfate": ("Cardiovascular / Hemodynamics", "Cardiovascular", ["Injectable"], "List I", "N"),
    "Vancomycin": ("Antimicrobial", "Antimicrobial", ["Injectable", "Lyophilized"], "N/A", "N"),
    "Cefuroxime": ("Antimicrobial", "Antimicrobial", ["Injectable"], "N/A", "N"),
    "Dexamethasone Sodium Phosphate": ("All Others Therapeutic Areas", "All Others", ["Injectable"], "N/A", "N"),
    "Ondansetron": ("All Others Therapeutic Areas", "All Others", ["Injectable", "Solution"], "N/A", "N"),
    "Hyaluronidase": ("All Others Therapeutic Areas", "All Others", ["Injectable", "Lyophilized"], "N/A", "N"),
    "Phentermine Hydrochloride": ("GLP-1 / Incretin (Weight Management/Diabetes)", "GLP / Weight", ["Capsule"], "Class IV", "N"),
    "Benzyl Alcohol": ("All Others Therapeutic Areas", "All Others", ["Aqueous Solution", "Liquid"], "N/A", "N"),
    "Sodium Bicarbonate": ("All Others Therapeutic Areas", "All Others", ["Injectable", "Aqueous Solution"], "N/A", "N"),
    "Magnesium Sulfate": ("All Others Therapeutic Areas", "All Others", ["Injectable"], "N/A", "N"),
    "Zinc Sulfate": ("Vitamins & Cofactors", "Vitamins", ["Injectable"], "N/A", "N"),
    "Minoxidil": ("All Others Therapeutic Areas", "All Others", ["Solution", "Capsule"], "N/A", "N"),
    "Tranexamic Acid": ("All Others Therapeutic Areas", "All Others", ["Injectable", "Cream (gm)"], "N/A", "N"),
    "Lisinopril": ("Hypertension", "Hypertension", ["Suspension", "Capsule"], "N/A", "N"),
}
# popularity weight per analyte (heavy tail so a few drive most volume)
ANALYTE_WEIGHT = {"Semaglutide": 30, "Tirzepatide": 22, "Testosterone Cypionate": 14, "Methylcobalamin": 10,
                  "Glutathione": 8, "NAD+": 8, "Progesterone": 6, "Sermorelin Acetate": 6, "Ascorbic Acid": 6,
                  "Lipo-C Blend": 6, "Papaverine / Phentolamine / Alprostadil": 5, "Estradiol Valerate": 4,
                  "Bevacizumab": 4, "Lidocaine Hydrochloride": 4, "Ketamine Hydrochloride": 3, "Benzyl Alcohol": 3}

# Test menu: name -> (department, sub-department, base price, TAT target days, popularity, needs analyte)
CORE_TESTS = {
    "USP <71> Sterility Test": ("Microbiology", "Sterile - Clean Room Ops", 160, 14, 120, False),
    "USP <85> Bacterial Endotoxins Test": ("Microbiology", "Sterile - Non Clean Room Ops", 110, 3, 115, False),
    "Potency (Solution)": ("Chemistry", "Chemistry Production", 240, 3, 95, True),
    "Potency (Peptides)": ("Chemistry", "Chemistry Production", 320, 4, 70, True),
    "Potency (Non-Solution)": ("Chemistry", "Chemistry Production", 260, 4, 35, True),
    "Rapid Sterility (BacT/Alert)": ("Microbiology", "Sterile - Clean Room Ops", 210, 7, 30, False),
    "USP <788> Particulate Matter": ("Microbiology", "Sterile - Non Clean Room Ops", 250, 4, 22, False),
    "pH Determination": ("Chemistry", "Chemistry Production", 60, 2, 20, False),
    "Environmental Monitoring Plate Read": ("Microbiology", "Non - Sterile - Bacteria", 35, 5, 25, False),
    "Glove Fingertip Plate Read": ("Microbiology", "Non - Sterile - Bacteria", 35, 5, 10, False),
    "Bioburden Testing": ("Microbiology", "Non - Sterile - Bacteria", 140, 5, 9, False),
    "USP <51> Antimicrobial Effectiveness Test": ("Microbiology", "Sterile - Non Clean Room Ops", 650, 30, 6, False),
    "Container Closure Integrity Test": ("Microbiology", "Sterile - Non Clean Room Ops", 420, 10, 5, False),
    "USP <61> Microbial Enumeration": ("Microbiology", "Non - Sterile - Bacteria", 180, 5, 6, False),
    "USP <62> Specified Organisms": ("Microbiology", "Non - Sterile - Bacteria", 220, 7, 5, False),
    "Growth Promotion": ("Microbiology", "Non - Sterile - Special projects", 95, 5, 4, False),
    "Media Fill Incubation": ("Microbiology", "Sterile - Clean Room Ops", 120, 14, 4, False),
    "Identification by HPLC": ("Chemistry", "Chemistry Production", 180, 3, 6, True),
    "Assay & Impurities": ("Chemistry", "Chemistry Production", 450, 7, 4, True),
    "USP <467> Residual Solvents": ("Chemistry", "Chemistry Production", 380, 7, 3, False),
    "USP <232>/<233> Elemental Impurities": ("Chemistry", "Chemistry Production", 520, 10, 2, False),
    "USP <905> Uniformity of Dosage Units": ("Chemistry", "Chemistry Production", 300, 5, 3, True),
    "Osmolality": ("Chemistry", "Chemistry Production", 90, 3, 2, False),
    "Water Activity": ("Chemistry", "Chemistry R&D", 85, 3, 1.5, False),
    "Beyond-Use Date Stability Study": ("Chemistry", "Chemistry R&D", 2400, 120, 2, True),
    "Method Validation": ("Chemistry", "Chemistry R&D", 9500, 45, 0.4, True),
    "USP <1225> Method Verification": ("Chemistry", "Chemistry R&D", 3800, 30, 0.6, True),
    "USP <1116> Cleanroom Microbial Evaluation": ("Microbiology", "Non - Sterile - Special projects", 750, 10, 0.8, False),
    "Disinfectant Efficacy Study": ("Microbiology", "Non - Sterile - Special projects", 4200, 45, 0.3, False),
    "USP <790> Visible Particulates": ("Microbiology", "Sterile - Non Clean Room Ops", 120, 3, 3, False),
    "USP <789> Particulates in Ophthalmics": ("Microbiology", "Sterile - Non Clean Room Ops", 260, 4, 1, False),
    "USP <831> Refractive Index": ("Chemistry", "Chemistry R&D", 75, 3, 0.5, False),
    "Fungal Identification": ("Microbiology", "Non - Sterile - Bacteria", 160, 10, 1.2, False),
    "Bacterial Identification": ("Microbiology", "Non - Sterile - Bacteria", 150, 7, 1.5, False),
}
MONOGRAPH_DRUGS = ["Biotin", "Lysine Hydrochloride", "Ascorbic Acid", "Methylcobalamin", "Lidocaine", "Bupivacaine",
                   "Ketorolac", "Ondansetron", "Dexamethasone", "Glycine", "Arginine", "Taurine", "Zinc Sulfate",
                   "Magnesium Chloride", "Sodium Chloride", "Dextrose", "Folic Acid", "Niacinamide", "Riboflavin",
                   "Thiamine", "Pyridoxine", "Leucovorin", "Heparin Sodium", "Oxytocin", "Phenylephrine",
                   "Ephedrine", "Glycopyrrolate", "Neostigmine", "Midazolam", "Atropine", "Succinylcholine",
                   "Rocuronium", "Vasopressin", "Calcium Gluconate", "Potassium Chloride", "Tromethamine",
                   "Edetate Disodium", "Polysorbate 80", "Benzyl Benzoate", "Cottonseed Oil", "Grapeseed Oil",
                   "Sesame Oil", "Mannitol", "Sucrose", "Trehalose", "Histidine", "Methionine", "Cysteine",
                   "Carnitine", "Inositol", "Choline Chloride", "Dexpanthenol", "Selenious Acid", "Cupric Chloride",
                   "Manganese Sulfate", "Chromium Chloride", "Ethanol", "Glycerin", "Propylene Glycol", "Hyaluronic Acid"]
ELEMENTS = ["Sodium", "Potassium", "Calcium", "Magnesium", "Zinc", "Copper", "Selenium", "Chromium", "Manganese", "Iron"]

DOSAGE_EXTRAS = ["Plate", "Raw Powder", "Powder Aliquot", "Component", "Liquid", "Swab"]
FACILITY_TYPES = ["503A", "503B", "GMP", "Research", "Non-USA"]
FACILITY_W = [0.885, 0.07, 0.04, 0.004, 0.001]

# real US cities so map visuals geocode correctly; weights skew to large states
CITIES = [
    ("Houston", "TX"), ("Dallas", "TX"), ("Austin", "TX"), ("San Antonio", "TX"), ("Fort Worth", "TX"), ("Plano", "TX"),
    ("The Woodlands", "TX"), ("Sugar Land", "TX"), ("El Paso", "TX"), ("Lubbock", "TX"),
    ("Los Angeles", "CA"), ("San Diego", "CA"), ("Irvine", "CA"), ("San Jose", "CA"), ("Sacramento", "CA"),
    ("Fresno", "CA"), ("Newport Beach", "CA"), ("Beverly Hills", "CA"), ("Torrance", "CA"), ("Oakland", "CA"),
    ("Miami", "FL"), ("Orlando", "FL"), ("Tampa", "FL"), ("Jacksonville", "FL"), ("Boca Raton", "FL"),
    ("New York", "NY"), ("Brooklyn", "NY"), ("Rochester", "NY"), ("Buffalo", "NY"),
    ("Phoenix", "AZ"), ("Scottsdale", "AZ"), ("Chandler", "AZ"), ("Tucson", "AZ"), ("Mesa", "AZ"),
    ("Philadelphia", "PA"), ("Pittsburgh", "PA"), ("Columbus", "OH"), ("Cleveland", "OH"), ("Cincinnati", "OH"),
    ("Chicago", "IL"), ("Naperville", "IL"), ("Nashville", "TN"), ("Memphis", "TN"), ("Knoxville", "TN"),
    ("Atlanta", "GA"), ("Savannah", "GA"), ("Charlotte", "NC"), ("Raleigh", "NC"), ("Denver", "CO"),
    ("Colorado Springs", "CO"), ("Seattle", "WA"), ("Spokane", "WA"), ("Portland", "OR"), ("Las Vegas", "NV"),
    ("Salt Lake City", "UT"), ("Boise", "ID"), ("Kansas City", "MO"), ("St. Louis", "MO"), ("Minneapolis", "MN"),
    ("Detroit", "MI"), ("Grand Rapids", "MI"), ("Indianapolis", "IN"), ("Louisville", "KY"), ("New Orleans", "LA"),
    ("Baton Rouge", "LA"), ("Oklahoma City", "OK"), ("Tulsa", "OK"), ("Birmingham", "AL"), ("Jackson", "MS"),
    ("Little Rock", "AR"), ("Albuquerque", "NM"), ("Omaha", "NE"), ("Des Moines", "IA"), ("Milwaukee", "WI"),
    ("Boston", "MA"), ("Hartford", "CT"), ("Newark", "NJ"), ("Baltimore", "MD"), ("Richmond", "VA"),
    ("Charleston", "SC"), ("Columbia", "SC"), ("Wilmington", "DE"), ("Honolulu", "HI"), ("Anchorage", "AK"),
]
CITY_W = np.array([18 if s == "TX" else 16 if s == "CA" else 5 if s in ("FL", "NY", "AZ") else 2.5 for _, s in CITIES], float)
REGION = {**{s: "South" for s in "TX FL GA NC SC TN AL MS AR LA OK KY VA MD DE".split()},
          **{s: "West" for s in "CA AZ NV UT CO WA OR ID NM HI AK".split()},
          **{s: "Midwest" for s in "OH IL MI IN MO MN WI IA NE".split()},
          **{s: "Northeast" for s in "NY PA MA CT NJ".split()}}
CA_CITIES = [("Toronto", "ON"), ("Ottawa", "ON"), ("Mississauga", "ON"), ("Vancouver", "BC"), ("Calgary", "AB"),
             ("Edmonton", "AB"), ("Montreal", "QC"), ("Winnipeg", "MB"), ("Halifax", "NS")]

NAME_SUFFIX = {"503A": ["Compounding Pharmacy", "Pharmacy", "Apothecary", "Specialty Pharmacy", "Rx",
                        "Wellness Pharmacy", "Compounding Center", "Family Pharmacy"],
               "503B": ["Outsourcing Facility", "Sterile Outsourcing", "Pharma Services"],
               "GMP": ["Pharmaceuticals", "Biologics", "Laboratories", "Life Sciences"],
               "Research": ["Research Institute", "University Research Lab"],
               "Non-USA": ["International Pharmacy", "Global Pharma"]}
NAME_PREFIX = ["Summit", "Cedar", "Harbor", "Pinnacle", "Lone Star", "Bluebonnet", "Riverbend", "Northgate", "Evergreen",
               "Silver Creek", "Magnolia", "Canyon", "Prairie", "Coastal", "Granite", "Redwood", "Lakeside", "Meadow",
               "Falcon", "Heritage", "Liberty", "Keystone", "Sierra", "Mesa", "Aspen", "Willow", "Oakridge", "Crescent",
               "Beacon", "Sterling", "Horizon", "Vista", "Gulf", "Desert", "Bayou", "Highland", "Union", "Frontier",
               "Pioneer", "Golden", "Copper", "Orchard", "Maple", "Juniper", "Sage", "Trinity", "Unity", "Apex",
               "Clearwater", "Stonebridge", "Westfield", "Eastside", "Southpoint", "Northstar", "Brightwater"]

QA_REVIEWERS_N = 8
ITEMS = [("TSA Contact Plate with Neutralizers", "Sleeve of 10.00 Plate", 35),
         ("SDA Contact Plate", "Sleeve of 10.00 Plate", 37),
         ("Glove Fingertip Sampling Plate", "Sleeve of 10.00 Plate", 20),
         ("Environmental Swab Kit", "Kit of 1.00 Kit", 60)]
CONSULT_CODES = [("C100 (Consulting)", 0.86), ("C200 (Quote Charge)", 0.06),
                 ("C300 (Certification, Calibration & Qualification)", 0.04), ("C400 (Finance Charge)", 0.03),
                 ("C500 (Equipment Rental)", 0.01)]
CONSULT_PROJECTS = ["USP <797> Gap Assessment", "USP <800> Hazardous Drug Program Review", "Cleanroom Certification Support",
                    "Smoke Study Review", "Environmental Monitoring Program Design", "SOP Writing Package",
                    "Media Fill Program Setup", "Method Validation Protocol", "Beyond-Use Date Strategy",
                    "FDA 483 Response Support", "Mock Inspection", "Training Workshop", "Facility Design Review"]


def month_starts(start, end):
    d = dt.date(start.year, start.month, 1)
    while d <= end:
        yield d
        d = dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)


def add_business_days(d, n):
    # simple: calendar days * 7/5 rounded
    return d + dt.timedelta(days=int(round(n * 7 / 5)))


def main(out_dir, scale):
    rng = np.random.default_rng(SEED)
    random.seed(SEED)
    fake = Faker("en_US")
    Faker.seed(SEED)
    os.makedirs(out_dir, exist_ok=True)

    # ---------------- test menu ----------------
    tests = dict(CORE_TESTS)
    for d in MONOGRAPH_DRUGS:
        tests[f"USP Monograph - {d}"] = ("Chemistry", "Chemistry Production", int(rng.integers(280, 520)), 7, 0.12, False)
    for e in ELEMENTS:
        tests[f"Content of {e}"] = ("Chemistry", "Chemistry R&D", int(rng.integers(150, 260)), 5, 0.15, False)
    test_names = list(tests)
    test_w = np.array([tests[t][4] for t in test_names], float)
    test_w /= test_w.sum()
    dept_rows = [(tests[t][0], tests[t][1], t) for t in test_names]
    # departments table also lists retired / rarely used methods (as the real catalog does)
    for i in range(1, 41):
        dept_rows.append(("Chemistry", "Chemistry R&D", f"Legacy Method CH-{i:03d}"))
    for i in range(1, 16):
        dept_rows.append(("Microbiology", "Non - Sterile - Special projects", f"Legacy Method MB-{i:03d}"))

    # ---------------- analysts ----------------
    def person():
        f, l = fake.first_name(), fake.last_name()
        return f"{f} {l}", f"{f[0]}{l}"
    chem_analysts = [person() for _ in range(80)]
    micro_analysts = [person() for _ in range(42)]
    qa_reviewers = [person()[0] for _ in range(QA_REVIEWERS_N)]
    analyst_user = {n: u for n, u in chem_analysts + micro_analysts}
    chem_w = rng.lognormal(0, 0.6, len(chem_analysts)); chem_w /= chem_w.sum()
    micro_w = rng.lognormal(0, 0.6, len(micro_analysts)); micro_w /= micro_w.sum()
    # some analysts join late / leave early so month-over-month trends move
    analyst_window = {}
    for n, _ in chem_analysts + micro_analysts:
        s = dt.date(2022, 1, 1) if rng.random() < 0.7 else dt.date(2023, 1, 1) + dt.timedelta(days=int(rng.integers(0, 1200)))
        e = dt.date(2027, 1, 1) if rng.random() < 0.85 else s + dt.timedelta(days=int(rng.integers(200, 900)))
        analyst_window[n] = (s, e)

    # ---------------- samples / therapeutics ----------------
    analyte_names = list(ANALYTES)
    a_w = np.array([ANALYTE_WEIGHT.get(a, 1.0) for a in analyte_names]); a_w /= a_w.sum()
    units = ["mg/mL", "mcg/mL", "mg/g", "mg", "units/mL"]
    sample_universe = []      # (sample_name, analyte, dosage_form)
    for a in analyte_names:
        forms = ANALYTES[a][2]
        n_var = int(rng.integers(60, 260) * (ANALYTE_WEIGHT.get(a, 1) ** 0.4))
        seen = set()
        for _ in range(n_var):
            form = forms[int(rng.integers(len(forms)))]
            strength = round(float(rng.choice([0.25, 0.5, 1, 1.25, 2, 2.5, 3, 5, 10, 12.5, 20, 25, 50, 100, 200])) *
                             float(rng.choice([1, 1, 1, 1.5, 2])), 2)
            unit = units[int(rng.integers(len(units)))] if form not in ("Capsule",) else "mg"
            extra = rng.choice(["", "", "", " with B6", " PF", " in Grapeseed Oil", " MDV", " SDV", " with Glycine"])
            name = f"{a} {strength:g} {unit}{extra} {form}"
            if name in seen:
                continue
            seen.add(name)
            sample_universe.append((name, a, form))
    nonther = [("EM Plate - ISO 5 Hood", "", "Plate"), ("EM Plate - ISO 7 Buffer Room", "", "Plate"),
               ("Glove Fingertip - Operator", "", "Plate"), ("Water for Injection", "", "Liquid"),
               ("Sterile Water Rinse", "", "Liquid"), ("Media Fill Unit", "", "Liquid"),
               ("API Raw Material", "", "Raw Powder"), ("Bulk Powder Aliquot", "", "Powder Aliquot"),
               ("Stopper / Vial Component", "", "Component"), ("Surface Swab", "", "Swab")]
    for base in nonther:
        for k in range(1, 40):
            sample_universe.append((f"{base[0]} #{k:02d}", "", base[2]))
    sample_by_analyte = defaultdict(list)
    for i, s in enumerate(sample_universe):
        sample_by_analyte[s[1]].append(i)

    # ---------------- clients ----------------
    n_us = int(1500 * max(scale, 0.2))
    used = set()
    clients = []
    city_p = CITY_W / CITY_W.sum()
    for i in range(n_us):
        ftype = rng.choice(FACILITY_TYPES, p=FACILITY_W)
        while True:
            nm = f"{rng.choice(NAME_PREFIX)} {fake.last_name() if rng.random() < 0.55 else rng.choice(NAME_PREFIX)} {rng.choice(NAME_SUFFIX[ftype])}"
            if nm not in used:
                used.add(nm)
                break
        city, st = CITIES[rng.choice(len(CITIES), p=city_p)]
        country = "USA"
        if ftype == "Non-USA":
            city, st, country = rng.choice([("Sydney", "NSW", "AU"), ("Dubai", "DU", "UAE"), ("Singapore", "SG", "SG")])
        start = dt.date(2021, 6, 1) + dt.timedelta(days=int(rng.integers(0, 1950)))
        if rng.random() < 0.45:
            start = dt.date(2021, 1, 1) + dt.timedelta(days=int(rng.integers(0, 700)))
        churn = None
        if rng.random() < 0.22:
            churn = start + dt.timedelta(days=int(rng.integers(120, 1300)))
        base = float(rng.lognormal(0.0, 1.35))
        if ftype in ("503B", "GMP"):
            base *= 1.8
        trend = float(rng.normal(1.0, 0.035))     # monthly growth factor
        clients.append(dict(name=nm, city=city, state=st, country=country, ftype=ftype, start=start, churn=churn,
                            base=base, trend=trend, acct=f"AC{100000 + i}", region=REGION.get(st, "Non American"),
                            oos_mult=float(rng.lognormal(0, 0.7)), late_mult=float(rng.lognormal(0, 0.35))))
    # favourite products per client
    for c in clients:
        fav_a = rng.choice(len(analyte_names), size=int(rng.integers(2, 9)), replace=False, p=a_w)
        favs = []
        for ai in fav_a:
            pool = sample_by_analyte[analyte_names[ai]]
            favs.extend(rng.choice(pool, size=min(len(pool), int(rng.integers(2, 12))), replace=False).tolist())
        favs.extend(rng.choice(sample_by_analyte[""], size=int(rng.integers(1, 5)), replace=False).tolist())
        c["favs"] = favs

    # ---------------- submissions ----------------
    months = list(month_starts(dt.date(2023, 1, 1), DATA_END))
    services = []      # list of dicts (one per test)
    sub_seq = defaultdict(int)
    test_seq = 700000
    lot_seq = 150000
    tests_per_sub_mean = 2.5
    for m_idx, m in enumerate(months):
        yr_months = 10 if m.year == 2026 else 12
        season = 1 + 0.08 * np.sin((m.month - 3) / 12 * 2 * np.pi)
        target_tests = YEAR_TESTS[m.year] / yr_months * season * scale
        n_subs = int(target_tests / tests_per_sub_mean)
        w = np.zeros(len(clients))
        for ci, c in enumerate(clients):
            if c["start"] > m or (c["churn"] and c["churn"] < m):
                continue
            age = (m.year - c["start"].year) * 12 + m.month - c["start"].month
            w[ci] = c["base"] * (c["trend"] ** min(age, 40))
        w /= w.sum()
        counts = rng.multinomial(n_subs, w)
        month_end = min(dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1) - dt.timedelta(days=1), DATA_END)
        days = [m + dt.timedelta(days=k) for k in range((month_end - m).days + 1)]
        wdays = [d for d in days if d.weekday() < 5] or days
        for ci in np.nonzero(counts)[0]:
            c = clients[ci]
            for _ in range(counts[ci]):
                sdate = wdays[int(rng.integers(len(wdays)))]
                key = sdate.strftime("%y%m%d")
                sub_seq[key] += 1
                sub_no = f"SUB-{key}-{sub_seq[key]:04d}"
                si = c["favs"][int(rng.integers(len(c["favs"])))]
                sname, analyte, form = sample_universe[si]
                rush = "Y" if rng.random() < (0.30 if c["ftype"] == "503A" else 0.12) else "N"
                lot_seq += 1
                lot = str(lot_seq)
                select = "Y" if rng.random() < 0.63 else ""
                # tests on this submission
                if analyte == "":
                    if form == "Plate":
                        chosen = ["Environmental Monitoring Plate Read"] if "Glove" not in sname else ["Glove Fingertip Plate Read"]
                        chosen *= int(rng.integers(1, 8))
                    else:
                        chosen = list(rng.choice(test_names, size=int(rng.integers(1, 3)), replace=False, p=test_w))
                else:
                    chosen = []
                    pot = "Potency (Peptides)" if ANALYTES[analyte][1] in ("GLP / Weight", "Peptides") and form != "Capsule" \
                        else ("Potency (Non-Solution)" if form in ("Capsule", "Cream (gm)", "Suspension", "Lyophilized") else "Potency (Solution)")
                    if rng.random() < 0.85:
                        chosen.append(pot)
                    if form in ("Injectable", "Oil Injectable", "Lyophilized", "Solution", "Aqueous Solution") and rng.random() < 0.8:
                        chosen.append("USP <71> Sterility Test")
                        if rng.random() < 0.85:
                            chosen.append("USP <85> Bacterial Endotoxins Test")
                    extra_n = int(rng.poisson(0.6))
                    if extra_n:
                        chosen += list(rng.choice(test_names, size=extra_n, p=test_w))
                    if not chosen:
                        chosen = [pot]
                bud = ""
                if rng.random() < 0.06:
                    bud = rng.choice(["0, 30, 60, 90", "0, 45, 90", "0, 90", "0, 30, 60, 90, 120",
                                      "0, 30, 60, 90, 120, 150, 180", "7, 14, 21, 30", "0, 60, 90"])
                rel_date = sdate + dt.timedelta(days=int(rng.choice([0, 0, 1, 1, 2, 3])))
                status_roll = rng.random()
                for tname in chosen:
                    dept, sub_dept, price, tat, _, needs_an = tests[tname]
                    test_seq += 1
                    tgt = tat * (0.5 if rush == "Y" and tat <= 14 else 1)
                    est = add_business_days(rel_date, max(1, tgt))
                    # completion: early / on time / late, late more likely for some clients & big tests
                    p_late = min(0.6, 0.11 * c["late_mult"] * (1.4 if rush == "Y" else 1) * (1.5 if tat >= 30 else 1))
                    r = rng.random()
                    if r < p_late:
                        comp = est + dt.timedelta(days=int(rng.geometric(0.35)))
                    elif r < p_late + 0.36:
                        comp = est - dt.timedelta(days=int(rng.integers(1, max(2, int(tgt * 0.6) + 2))))
                        comp = max(comp, rel_date)
                    else:
                        comp = est
                    pool = chem_analysts if dept == "Chemistry" else micro_analysts
                    pw = chem_w if dept == "Chemistry" else micro_w
                    for _ in range(6):
                        an = pool[rng.choice(len(pool), p=pw)][0]
                        s, e = analyst_window[an]
                        if s <= comp <= e:
                            break
                    pr = 0 if rng.random() < 0.01 else int(round(price * float(rng.choice([1, 1, 1, 0.9, 0.85, 1.1]))))
                    test_point = ""
                    if bud:
                        tp = int(rng.choice([int(x) for x in bud.split(", ")]))
                        test_point = f"T{tp} - {(rel_date + dt.timedelta(days=tp)).strftime('%m/%d/%y')}"
                    services.append(dict(
                        Submission_Number=sub_no, Submission_Date=sdate, Client_Facility=c["name"], Select=select,
                        Facility_Type=c["ftype"], Sample_Name=sname, Dosage_Form=form, Lot_Number=lot,
                        BUD_Test_Points=bud, Test_Point=test_point, Test_Name=tname,
                        Analyte=analyte, Test_Id=f"T{test_seq:07d}",
                        Submission_Release_Date=rel_date, Estimated_Completion_Date=est, Completion_Date=comp,
                        Rush=rush, Analyst=an, Controlled_Substance_Class=ANALYTES[analyte][3] if analyte else "N/A",
                        Hazardous_Drug=ANALYTES[analyte][4] if analyte else "N", Price=pr, Department=dept,
                        _status_roll=status_roll, _client=c))
    print(f"services rows: {len(services):,}")

    # status fields
    for r in services:
        done = r["Completion_Date"] <= DATA_END
        r["Status"] = "Completed" if done else "In Test"
        if r["Submission_Date"].year < 2026:
            r["Submission_Status"] = "Complete"
        else:
            r["Submission_Status"] = "Complete" if r["_status_roll"] < 0.918 else ("In Test" if r["_status_roll"] < 0.998 else "Partial Release")
        if not done:
            r["Submission_Status"] = "In Test"

    raw_cols = ["Submission_Number", "Submission_Date", "Submission_Status", "Client_Facility", "Select", "Facility_Type",
                "Sample_Name", "Dosage_Form", "Lot_Number", "Test_Point", "Test_Name", "Analyte", "Test_Id",
                "Submission_Release_Date", "Estimated_Completion_Date", "Completion_Date", "Rush", "Status", "Analyst",
                "Controlled_Substance_Class", "Hazardous_Drug", "Price", "BUD_Test_Points", "Test_Facility"]
    raw26_cols = ["Submission_Number", "Submission_Date", "Submission_Status", "Client_Facility", "Select", "Sample_Name",
                  "Dosage_Form", "Lot_Number", "Test_Point", "Test_Name", "Analyte", "Test_Id", "Submission_Release_Date",
                  "Estimated_Completion_Date", "Completion_Date", "Rush", "Status", "Analyst",
                  "Controlled_Substance_Class", "Hazardous_Drug", "Price", "Facility_Type", "BUD_Test_Points", "Test_Facility"]
    tests_cols = ["Submission_Number", "Submission_Date", "Submission_Status", "Client_Facility", "Select", "Facility_Type",
                  "Sample_Name", "Dosage_Form", "Lot_Number", "BUD_Test_Points", "Test_Point", "Test_Name", "Analyte",
                  "Test_Id", "Submission_Release_Date", "Estimated_Completion_Date", "Completion_Date", "Rush", "Status",
                  "Analyst", "Controlled_Substance_Class", "Hazardous_Drug", "Price"]

    def rowgen(rows, cols, facility=FACILITY_MAIN, blank_incomplete=False):
        for r in rows:
            out = []
            for col in cols:
                if col == "Test_Facility":
                    out.append(facility)
                elif col == "Completion_Date" and blank_incomplete and r["Status"] != "Completed":
                    out.append("")
                else:
                    out.append(r[col])
            yield out

    by_year = defaultdict(list)
    for r in services:
        by_year[r["Submission_Date"].year].append(r)
    completed26 = [r for r in by_year[2026] if r["Status"] == "Completed"]
    # earlier years only contain completed work
    for y in (2023, 2024, 2025):
        by_year[y] = [r for r in by_year[y] if r["Status"] == "Completed"]

    # ---------------- Therapeutics table ----------------
    used_samples = sorted({r["Sample_Name"] for r in services})
    s_lookup = {s[0]: s for s in sample_universe}
    ther_rows = []
    for s in used_samples:
        a = s_lookup[s][1]
        if not a:
            ther_rows.append((s, "Non-Therapy", "Non-Therapy"))
        elif rng.random() < 0.12:
            ther_rows.append((s, "Noncategorised", "Noncategorised"))
        else:
            ther_rows.append((s, ANALYTES[a][0], ANALYTES[a][1]))

    # ---------------- Client facility table ----------------
    cf_cols = ["ClientFacilityName", "BillTo_City", "BillTo_StateProvince", "BillTo_Country", "Date", "Active Client",
               "Region", "Test Facility", "AccountNo", "Facility_Type"]
    cf_rows = [(c["name"], c["city"], c["state"], c["country"], c["start"].strftime("%Y-%m-%d"),
                "Yes", c["region"] if c["country"] == "USA" else ("Australia" if c["country"] == "AU" else "Non American"),
                "US Central", c["acct"], c["ftype"]) for c in clients]

    # ---------------- write Raw DATA.xlsx ----------------
    print("writing Raw DATA.xlsx ...")
    wb = StreamingWorkbook(os.path.join(out_dir, "Raw DATA.xlsx"))
    for y in (2023, 2024, 2025):
        wb.write_sheet(f"RAW_{y}", raw_cols, rowgen(by_year[y], raw_cols), f"RAW_{y}")
    wb.write_sheet("RAW_2026", raw26_cols, rowgen(completed26, raw26_cols), "RAW_2026")
    wb.write_sheet("Departments", ["Department", "Sub - Department", "Test_Name"], dept_rows, "Table2")
    wb.write_sheet("Client Facility", cf_cols, cf_rows, "Table13")
    wb.write_sheet("Therapeutics", ["Sample name", "Therapy Type", "Umbrella"], ther_rows, "Therapeutics_Table")
    wb.close()

    # ---------------- Tests Completed / Submission Release ----------------
    print("writing Tests Completed.xlsx / Submission Release.xlsx ...")
    wb = StreamingWorkbook(os.path.join(out_dir, "Tests Completed.xlsx"))
    wb.write_sheet("Tests", tests_cols, rowgen(completed26, tests_cols))
    wb.close()
    released = [r for r in services if r["Submission_Release_Date"].year == 2026 and r["Submission_Release_Date"] <= DATA_END]
    wb = StreamingWorkbook(os.path.join(out_dir, "Submission Release.xlsx"))
    wb.write_sheet("Tests", tests_cols, rowgen(released, tests_cols, blank_incomplete=True))
    wb.close()

    # ---------------- Quality Data (QA review timings) ----------------
    print("writing Quality Data.xlsx ...")
    qa_src = [r for r in completed26 if r["Completion_Date"] <= QA_END and rng.random() < 0.82]
    qa_w = rng.lognormal(0, 0.5, QA_REVIEWERS_N); qa_w /= qa_w.sum()

    def qa_rows():
        for r in qa_src:
            prep = 0.0 if rng.random() < 0.55 else float(rng.lognormal(2.8, 0.8))
            ana = float(rng.lognormal(1.0, 1.4))
            proc = 0.0 if rng.random() < 0.97 else float(rng.lognormal(1, 1))
            rev = float(rng.lognormal(1.3, 1.2))
            hold = 0.0 if rng.random() < 0.93 else float(rng.lognormal(3.5, 1.2))
            qarev = 0.0 if rng.random() < 0.8 else float(rng.lognormal(2.0, 1.0))
            qahold = 0.0 if rng.random() < 0.97 else float(rng.lognormal(3.0, 1.3))
            qaapp = 0 if rng.random() < 0.995 else int(rng.integers(1, 20))
            total = prep + ana + proc + rev + hold + qarev + qahold
            ts = dt.datetime.combine(r["Completion_Date"], dt.time(int(rng.integers(7, 22)), int(rng.integers(0, 60)), int(rng.integers(0, 60))))
            yield (r["Submission_Number"], r["Test_Name"], r["Department"], r["Dosage_Form"], r["Analyst"],
                   r["Client_Facility"], r["Facility_Type"], round(prep, 2), round(ana, 2), round(proc, 2), round(rev, 2),
                   round(hold, 2), 0, round(qarev, 2), round(qahold, 2), qaapp, round(total, 2), ts,
                   qa_reviewers[rng.choice(QA_REVIEWERS_N, p=qa_w)])
    wb = StreamingWorkbook(os.path.join(out_dir, "Quality Data.xlsx"))
    wb.write_sheet("Quality Data", ["Submission", "Test_Name", "Department", "Dosage_Form", "Analyst", "Client_Name",
                                    "Facility_Type", "Time_In_Sample_Prep", "Time_In_Sample_Analysis",
                                    "Time_In_Data_Processing", "Time_In_Data_Review", "Time_In_On_Hold",
                                    "Time_In_Approved", "Time_In_Qa_Review", "Time_In_Qa_Hold", "Time_In_Qa_Approved",
                                    "Time_To_Complete", "Completed", "Approved_By"], qa_rows())
    wb.close()

    # ---------------- OOS ----------------
    print("writing OOS files ...")
    analyte_oos = {a: float(rng.lognormal(0, 0.6)) for a in analyte_names}
    analyst_oos = {n: float(rng.lognormal(0, 0.5)) for n in analyst_user}
    oos_pool = [r for r in completed26 if r["Completion_Date"] <= QA_END and r["Completion_Date"] >= dt.date(2026, 1, 2)]
    oos_rows, tech = [], defaultdict(lambda: [0, 0])
    inv_seq = 260000
    inv_by_sub = {}
    for r in oos_pool:
        key = (r["Completion_Date"], r["Analyst"], r["Department"], r["Test_Name"])
        tech[key][0] += 1
        p = 0.0026 * r["_client"]["oos_mult"] * analyte_oos.get(r["Analyte"], 1.0) * analyst_oos[r["Analyst"]]
        if r["Test_Name"].startswith("Potency"):
            p *= 1.6
        if rng.random() < min(p, 0.25):
            tech[key][1] += 1
            if r["Submission_Number"] not in inv_by_sub:
                inv_seq += int(rng.integers(1, 4))
                inv_by_sub[r["Submission_Number"]] = str(inv_seq)
            if not r["Analyte"]:
                ttype = r["Test_Name"]
            elif r["Test_Name"].endswith(")"):
                ttype = f"{r['Test_Name']} ({r['Analyte']})"
            else:
                ttype = f"{r['Test_Name']} Samples ({r['Analyte']})"
            oos_rows.append((r["Department"], ttype, r["Completion_Date"], r["Submission_Number"], r["Client_Facility"],
                             inv_by_sub[r["Submission_Number"]], r["Dosage_Form"]))
    wb = StreamingWorkbook(os.path.join(out_dir, "OosByClient.xlsx"))
    wb.write_sheet("OosByClient", ["Department", "TestType", "TestDate", "SubmissionId", "ClientName", "OOSNumber",
                                   "DosageForm"], oos_rows)
    wb.close()
    tech_rows = [(k[0], k[1], k[2], k[3], v[0], v[1], analyst_user[k[1]]) for k, v in sorted(tech.items()) if v[1] > 0]
    wb = StreamingWorkbook(os.path.join(out_dir, "OosByTechnician.xlsx"))
    wb.write_sheet("OosByTechnician", ["Date", "Analyst", "Department", "TestType", "TestsPerformed", "NumberOfOOS",
                                       "Username"], tech_rows)
    wb.close()

    # ---------------- Canada ----------------
    print("writing Raw Data Canada.xlsx ...")
    ca_clients = []
    for i in range(int(20 * max(scale, 0.5))):
        city, prov = CA_CITIES[int(rng.integers(len(CA_CITIES)))]
        ca_clients.append(dict(name=f"{rng.choice(NAME_PREFIX)} {fake.last_name()} Pharmacy ({prov})", city=city, prov=prov))
    ca_rows = []
    for i in range(int(640 * max(scale, 0.2))):
        c = ca_clients[int(rng.integers(len(ca_clients)))]
        sdate = dt.date(2026, 3, 1) + dt.timedelta(days=int(rng.integers(0, (DATA_END - dt.date(2026, 3, 1)).days)))
        tname = rng.choice(["USP <71> Sterility Test", "USP <85> Bacterial Endotoxins Test", "Potency (Solution)",
                            "Potency (Peptides)", "Environmental Monitoring Plate Read"], p=[0.3, 0.25, 0.2, 0.15, 0.1])
        sname, a, form = sample_universe[int(rng.integers(len(sample_universe)))]
        est = add_business_days(sdate, tests[tname][3])
        comp = est + dt.timedelta(days=int(rng.integers(-2, 3)))
        if comp > DATA_END:
            continue
        test_seq += 1
        ca_rows.append((f"CAN-{sdate.strftime('%y%m%d')}-{i:04d}", sdate, "Complete", c["name"], "", "503A", sname, form,
                        str(900000 + i), "", "", tname, a, f"T{test_seq:07d}", sdate, est, comp,
                        "N", "Completed", micro_analysts[int(rng.integers(len(micro_analysts)))][0], "N/A", "N",
                        int(tests[tname][2] * 1.35), FACILITY_CA))
    ca_cols = ["Submission_Number", "Submission_Date", "Submission_Status", "Client_Facility", "Select", "Facility_Type",
               "Sample_Name", "Dosage_Form", "Lot_Number", "BUD_Test_Points", "Test_Point", "Test_Name", "Analyte",
               "Test_Id", "Submission_Release_Date", "Estimated_Completion_Date", "Completion_Date", "Rush", "Status",
               "Analyst", "Controlled_Substance_Class", "Hazardous_Drug", "Price", "Test_Facility"]
    wb = StreamingWorkbook(os.path.join(out_dir, "Raw Data Canada.xlsx"))
    wb.write_sheet("Canada Submissions", ca_cols, ca_rows, "Table1")
    wb.write_sheet("Canada Client Facility", ["ClientFacilityName", "BillTo_City", "BillTo_StateProvince", "BillTo_Country",
                                               "Date", "Active Client", "Region", "Test Facility"],
                   [(c["name"], c["city"], c["prov"], "CA", "2026-03-01", "Yes", "Canada", "Canada") for c in ca_clients],
                   "Table2")
    wb.close()

    # ---------------- Products (EM plates & swab kits) ----------------
    print("writing Product Test.xlsx ...")
    prod_clients = [c for c in clients if rng.random() < 0.19]
    orders, lines = [], []
    o_seq = 50000
    for i in range(int(1870 * max(scale, 0.2))):
        c = prod_clients[int(rng.integers(len(prod_clients)))]
        odate = dt.date(2023, 1, 4) + dt.timedelta(days=int(rng.integers(0, (DATA_END - dt.date(2023, 1, 4)).days)))
        o_seq += 1
        ono = f"SO-{o_seq}"
        n_lines = 1 if rng.random() < 0.7 else 2
        picked = rng.choice(len(ITEMS), size=n_lines, replace=False, p=[0.6, 0.25, 0.1, 0.05])
        line_total = 0
        for it in picked:
            item, pack, up = ITEMS[it]
            qty = int(max(1, rng.lognormal(2.0, 0.8)))
            ext = qty * up
            line_total += ext
            lines.append((ono, item, pack, qty, qty, up, ext))
        sh = round(float(max(0, line_total * rng.uniform(0.08, 0.25))), 2)
        tax = 0 if rng.random() < 0.9 else int(line_total * 0.0825)
        disc = 0 if rng.random() < 0.93 else int(line_total * 0.1)
        user = fake.name()
        addr = fake.street_address() + f", {c['city']}, {c['state']}"
        orders.append((ono, c["name"], c["acct"], int(800000 + i), f"PO{int(rng.integers(10000, 99999))}", odate,
                       "Completed", user, fake.numerify("###-555-####"), "", f"orders@{c['name'].split()[0].lower()}.example.com",
                       "" if rng.random() < 0.95 else "SPRING10", addr, addr, n_lines, line_total, sh, tax, disc,
                       round(line_total + sh + tax - disc, 2)))
    wb = StreamingWorkbook(os.path.join(out_dir, "Product Test.xlsx"))
    wb.write_sheet("Client Orders", ["Order_Number", "Client_Facility", "Account_Number", "Web_Order_Number", "PO_Number",
                                     "Order_Date", "Status", "Order_Placed_By", "Phone", "Fax", "Email", "Promo_Code",
                                     "Ship_To_Address", "Bill_To_Address", "Line_Count", "Line_Total",
                                     "Shipping_And_Handling", "Tax", "Discount", "Order_Total"], orders)
    wb.write_sheet("Order Details", ["Order_Number", "Item", "Packaging", "Ordered", "Shipped", "Unit_Price", "Ext_Price"], lines)
    wb.close()

    # ---------------- Consulting ----------------
    print("writing Consulting Activities.xlsx ...")
    cons_clients = [c for c in clients if rng.random() < 0.1]
    staff = [fake.name() for _ in range(6)]
    cons = []
    codes, code_p = zip(*CONSULT_CODES)
    for i in range(int(645 * max(scale, 0.2))):
        c = cons_clients[int(rng.integers(len(cons_clients)))]
        code = rng.choice(codes, p=code_p)
        est = dt.date(2023, 1, 15) + dt.timedelta(days=int(rng.integers(0, (DATA_END - dt.date(2023, 1, 15)).days)))
        comp = est + dt.timedelta(days=int(rng.integers(-10, 20)))
        if comp > DATA_END:
            comp = DATA_END
        price = float(round(rng.lognormal(7.4, 1.1), -1)) if code.startswith("C100") else float(rng.choice([250, 500, 750, 1200]))
        exp = 0.0 if rng.random() < 0.8 else float(round(rng.uniform(150, 2500), 2))
        proj = str(rng.choice(CONSULT_PROJECTS))
        cons.append((f"CA-{10000 + i}", c["name"], proj, f"PO{int(rng.integers(10000, 99999))}", f"Q-{int(rng.integers(2000, 9999))}",
                     int(30000 + i), code, FACILITY_MAIN, est, "Y", comp, "Y", comp + dt.timedelta(days=int(rng.integers(0, 10))),
                     price, f"{proj} for {c['name']}", exp, "Travel" if exp else "", staff[int(rng.integers(len(staff)))]))
    wb = StreamingWorkbook(os.path.join(out_dir, "Consulting Activities.xlsx"))
    wb.write_sheet("Consulting Activities", ["Activity_Id", "Client_Facility_Name", "Project_Name", "PO_Number",
                                             "Quote_Number", "Invoice_Number", "Billing_Codes", "Billing_Domain",
                                             "Estimated_Completion_Date", "Completed", "Completion_Date", "Invoiced",
                                             "Invoice_Submitted_Date", "Price", "Description", "Expenses_Price",
                                             "Expenses_Description", "Entered_By"], cons)
    wb.close()

    print("summary:", {y: len(v) for y, v in by_year.items()}, "released 2026:", len(released),
          "QA:", len(qa_src), "OOS:", len(oos_rows), "OOS tech:", len(tech_rows),
          "orders:", len(orders), "consulting:", len(cons), "clients:", len(clients))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="../data/lab")
    ap.add_argument("--scale", type=float, default=1.0)
    a = ap.parse_args()
    main(a.out, a.scale)
