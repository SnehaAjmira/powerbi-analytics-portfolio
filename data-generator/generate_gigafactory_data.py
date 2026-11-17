"""Synthetic data generator for a battery cell gigafactory construction project.

Fictional project: "Project Saguaro - GF1", a greenfield lithium-ion battery
cell gigafactory (cell manufacturing building with electrode, dry-room
assembly and formation areas, central utility building, NMP solvent recovery
plant, electrolyte storage, wastewater treatment and substation) delivered
for a fictional owner, Contoso Energy. All companies,
people, numbers and dates are randomly generated.

The data covers what a project information and analytics team reports on
across the project lifecycle:

  Engineering     document register (IFR / IFA / IFC / as-built), RFIs, submittals
  Procurement     PO lines vs required-on-site dates, long-lead equipment
  Construction    weekly earned value by construction work package (CWP)
  Completions     tag register, inspection test records (ITRs, A and B
                  check sheets), punch list, system / subsystem turnover
  HSE & Quality   weekly safety statistics, nonconformance reports

Like a real project, the data does not live in one place. It is written out
as six source systems (see export_sources.py):

  sql_server/          engineering database (reference data, documents, RFIs, submittals)
  erp_procurement/     ERP purchase order export (Excel, ERP vendor numbers)
  p6_schedule/         schedule / progress export (Excel)
  completions_system/  daily extract files (CSV folder) + tag register + punch list
  hse_api/             safety statistics from a REST API (JSON, OData style)
  quality_sharepoint/  NCR log exported from a SharePoint list (Excel)

Usage:
    pip install -r requirements.txt
    python generate_gigafactory_data.py --out ../data/gigafactory --scale 1.0
"""
import argparse
import os

import numpy as np
import pandas as pd

from export_sources import export_sources

SEED = 4242
DATA_DATE = pd.Timestamp("2026-11-13")       # status date (Friday week-ending)
NTP = pd.Timestamp("2024-04-01")             # notice to proceed
rng = np.random.default_rng(SEED)

DISCIPLINES = [
    ("CIV", "Civil", "Construction"), ("STR", "Structural", "Construction"), ("ARC", "Architectural", "Construction"),
    ("MEC", "Mechanical / HVAC", "Mechanical"), ("PPG", "Process Piping", "Mechanical"),
    ("PLB", "Plumbing & Process Drains", "Mechanical"), ("ELE", "Electrical", "Electrical"),
    ("INC", "Instrumentation & Controls", "Electrical"), ("FPR", "Fire Protection & Life Safety", "Mechanical"),
    ("PEQ", "Process Equipment Install & Hookup", "Process"),
]
DISC_CODES = [d[0] for d in DISCIPLINES]

BUILDINGS = {
    "CELL": ("Cell Manufacturing Building", ["L1 Electrode", "L2 Assembly Dry Room", "L3 Formation & Aging", "Roof"],
             ["N1", "N2", "N3", "S1", "S2", "S3"]),
    "CUB": ("Central Utility Building", ["L1", "L2", "Roof"], ["E", "C", "W"]),
    "OSB": ("Office Support Building", ["L1", "L2"], ["A"]),
    "NRU": ("NMP Solvent Recovery Plant", ["Grade"], ["N", "S"]),
    "ELB": ("Electrolyte & Chemical Storage", ["L1", "L2"], ["A"]),
    "WWT": ("Wastewater Treatment", ["Grade", "L1"], ["A", "B"]),
    "SUB": ("Main Substation", ["Grade"], ["A", "B"]),
    "UGS": ("Site Underground & Civil", ["Below Grade"], ["N", "S", "E", "W"]),
}
MAIN_BUILDING = "CELL"

SYSTEMS = [  # code, name, discipline, building, n subsystems, tags per subsystem
    ("RODI", "RO/DI Process Water", "PPG", "CUB", 20, 200), ("PCW", "Process Cooling Water", "PPG", "CUB", 26, 220),
    ("CDA", "Clean Dry Air", "PPG", "CUB", 14, 160), ("GN2", "Nitrogen Distribution", "PPG", "CUB", 12, 150),
    ("VAC", "Process Vacuum", "PPG", "CELL", 14, 150), ("DRY", "Dry Room Dehumidification", "MEC", "CELL", 30, 260),
    ("NMP", "NMP Solvent Recovery", "MEC", "NRU", 16, 200), ("PEX", "Process Exhaust", "MEC", "CELL", 18, 150),
    ("GEX", "General Exhaust", "MEC", "CELL", 10, 110), ("MAU", "Make-up Air Units", "MEC", "CELL", 12, 140),
    ("CHW", "Chilled Water", "MEC", "CUB", 20, 190), ("HHW", "Heating Hot Water", "MEC", "CUB", 10, 120),
    ("ELY", "Electrolyte Distribution", "PPG", "ELB", 14, 170), ("SLR", "Electrode Slurry Transfer", "PPG", "CELL", 12, 150),
    ("CHS", "Chemical Storage & Transfer", "PPG", "ELB", 8, 140), ("IWW", "Industrial Wastewater Treatment", "PLB", "WWT", 12, 150),
    ("PDR", "Process Drains", "PLB", "CELL", 14, 110), ("MVD", "Medium Voltage Distribution", "ELE", "SUB", 14, 120),
    ("LVD", "Low Voltage Distribution", "ELE", "CELL", 34, 240), ("FCP", "Formation Charger Power", "ELE", "CELL", 20, 180),
    ("UPS", "UPS & Critical Power", "ELE", "CUB", 10, 90), ("EMG", "Emergency Generators", "ELE", "CUB", 6, 110),
    ("LTG", "Lighting & Small Power", "ELE", "OSB", 16, 160), ("FPS", "Fire Sprinkler", "FPR", "CELL", 22, 120),
    ("FAD", "Fire Alarm & Detection", "FPR", "CELL", 18, 200), ("TRP", "Thermal Runaway Detection & Suppression", "FPR", "CELL", 14, 160),
    ("FMC", "Facility Monitoring & Control (FMCS)", "INC", "CELL", 20, 260), ("BMS", "Building Management System", "INC", "OSB", 10, 180),
    ("GDS", "Gas & Vapor Detection", "INC", "CELL", 10, 150), ("EQH", "Process Equipment Hookup", "PEQ", "CELL", 60, 140),
]
PRIORITY_1 = {"MVD", "UPS", "EMG", "CHW", "PCW", "CDA", "GN2", "FAD"}
PRIORITY_2 = {"RODI", "MAU", "DRY", "PEX", "FMC", "LVD"}

TAG_TYPES = {
    "PPG": [("V", "Valve", 0.45), ("PI", "Pressure Indicator", 0.12), ("FT", "Flow Transmitter", 0.08), ("P", "Pump", 0.05),
            ("F", "Filter Housing", 0.08), ("SP", "Pipe Spool", 0.22)],
    "MEC": [("D", "Damper", 0.3), ("FAN", "Fan", 0.12), ("AHU", "Air Handling Unit", 0.06), ("DH", "Desiccant Dehumidifier", 0.06),
            ("VAV", "Terminal Unit", 0.26), ("TT", "Temperature / Dew Point Transmitter", 0.12), ("HX", "Heat Exchanger", 0.08)],
    "PLB": [("V", "Valve", 0.4), ("P", "Pump", 0.15), ("TK", "Tank", 0.1), ("LT", "Level Transmitter", 0.15), ("SP", "Pipe Spool", 0.2)],
    "ELE": [("PNL", "Panelboard", 0.2), ("XFMR", "Transformer", 0.06), ("SWGR", "Switchgear Section", 0.06),
            ("MCC", "Motor Control Center Bucket", 0.18), ("CBL", "Cable", 0.5)],
    "INC": [("PLC", "Controller", 0.04), ("AI", "Analog Input Point", 0.4), ("DI", "Digital Input Point", 0.36),
            ("GD", "Gas / Vapor Detector", 0.12), ("JB", "Junction Box", 0.08)],
    "FPR": [("SPK", "Sprinkler Zone", 0.3), ("SD", "Smoke Detector", 0.35), ("FACP", "Fire Alarm Panel", 0.05),
            ("VESDA", "Aspirating Detector", 0.1), ("TRD", "Thermal Runaway Detector", 0.05), ("FV", "Flow / Tamper Switch", 0.15)],
    "PEQ": [("POC", "Point of Connection", 0.6), ("UD", "Utility Drop", 0.2), ("SKD", "Equipment Skid", 0.08),
            ("LEX", "Local Exhaust Connection", 0.12)],
}
ITR_FORMS = {  # discipline -> list of (form code, description, stage)
    "PPG": [("M-A01", "Piping Installation Check", "A"), ("M-A04", "Pressure Test", "A"), ("M-B02", "Flushing & Purge", "B"), ("M-B05", "Cleanliness / Dew Point Certification", "B")],
    "MEC": [("H-A01", "Equipment Installation Check", "A"), ("H-A03", "Duct Leakage Test", "A"), ("H-B01", "Pre-functional Check", "B"), ("H-B04", "Test, Adjust & Balance", "B")],
    "PLB": [("P-A01", "Installation Check", "A"), ("P-A02", "Hydrostatic Test", "A"), ("P-B01", "Pre-functional Check", "B")],
    "ELE": [("E-A01", "Installation Check", "A"), ("E-A03", "Insulation Resistance (Megger)", "A"), ("E-B02", "Energization Check", "B"), ("E-B04", "Protective Relay Test", "B")],
    "INC": [("I-A01", "Instrument Installation Check", "A"), ("I-B01", "Calibration", "B"), ("I-B02", "Loop Check", "B")],
    "FPR": [("F-A01", "Installation Check", "A"), ("F-A02", "Hydrostatic Test", "A"), ("F-B01", "Device Functional Test", "B")],
    "PEQ": [("Q-A01", "Hookup Installation Check", "A"), ("Q-A02", "Leak Test", "A"), ("Q-B01", "Utility Verification", "B")],
}

SUBCONTRACTORS = [
    ("SUB01", "Sonoran Earthworks", "CIV"), ("SUB02", "Red Rock Concrete", "CIV"), ("SUB03", "Ironwood Steel Erectors", "STR"),
    ("SUB04", "Mesa Architectural Systems", "ARC"), ("SUB05", "Dry Room Builders West", "ARC"), ("SUB06", "Saguaro Mechanical", "MEC"),
    ("SUB07", "Desert Air Sheet Metal", "MEC"), ("SUB08", "Pinnacle Process Piping", "PPG"), ("SUB09", "Precision Process Systems", "PPG"),
    ("SUB10", "Gila Plumbing Contractors", "PLB"), ("SUB11", "Copper State Electric", "ELE"), ("SUB12", "Sunbelt Power Systems", "ELE"),
    ("SUB13", "Precision Controls Integration", "INC"), ("SUB14", "Valley Fire Protection", "FPR"), ("SUB15", "EquipLink Hookup Services", "PEQ"),
    ("SUB16", "Arroyo Insulation", "MEC"), ("SUB17", "Phoenix Rigging & Setting", "PEQ"), ("SUB18", "Catalina Scaffold", "CIV"),
]
SUPPLIERS = ["Apex Dry Room Products", "Northfield Chillers", "Summit Switchgear", "Delta Process Components", "Granite Transformers",
             "Polaris Air Handling", "Westline Valves", "Crestview Pumps", "Meridian Filtration", "Keystone Cable & Wire",
             "Pinnacle Dehumidification", "Vantage Solvent Recovery", "Orion Instruments", "Cobalt Controls", "Frontier Fire Equipment",
             "Atlas Generators", "Beacon UPS Systems", "Harbor Steel Works", "Evergreen Tanks", "Ridgeway Duct Products",
             "Silverline Pipe Supply", "Lumen Lighting Supply", "Clearwater Treatment Systems", "Titan Cranes & Rigging"]
COMMODITIES = {
    "MEC": [("Desiccant Dehumidifier", "EA", 850000, True), ("Make-up Air Unit", "EA", 650000, True), ("NMP Recovery Unit", "EA", 3200000, True),
            ("Centrifugal Chiller", "EA", 1200000, True), ("Ductwork - Stainless", "LF", 220, False), ("Dampers", "EA", 2600, False)],
    "PPG": [("Process Valves", "EA", 1400, False), ("Stainless Steel Pipe", "LF", 85, False), ("PVDF Piping", "LF", 140, False),
            ("RO/DI Water Skid", "EA", 1600000, True), ("Electrolyte Storage Tank", "EA", 700000, True), ("Slurry Transfer Skid", "EA", 260000, True)],
    "PLB": [("Double-Contained Drain Pipe", "LF", 190, False), ("Neutralization Tank", "EA", 280000, True), ("Transfer Pumps", "EA", 38000, False)],
    "ELE": [("MV Switchgear Lineup", "EA", 1900000, True), ("Power Transformer", "EA", 1400000, True), ("LV Switchboards", "EA", 210000, True),
            ("Busway", "LF", 950, False), ("Power Cable", "LF", 22, False), ("Cable Tray", "LF", 48, False)],
    "INC": [("PLC Panels", "EA", 120000, False), ("Gas / Vapor Detectors", "EA", 6800, False), ("Field Instruments", "EA", 3200, False)],
    "FPR": [("Sprinkler Pipe & Fittings", "LF", 36, False), ("VESDA Units", "EA", 9500, False), ("Fire Alarm Panels", "EA", 64000, False)],
    "STR": [("Structural Steel", "TON", 5200, False), ("Metal Deck", "SF", 9, False), ("Embed Plates", "EA", 260, False)],
    "CIV": [("Precast Duct Bank", "LF", 410, False), ("Rebar", "TON", 1700, False), ("Ready-Mix Concrete", "CY", 210, False)],
    "ARC": [("Dry Room Insulated Panels", "SF", 95, False), ("Epoxy Floor System", "SF", 18, False), ("Ceiling Grid", "SF", 44, False)],
    "PEQ": [("Equipment Utility Drops", "EA", 14000, False), ("Process Equipment Skid Install Kits", "EA", 180000, True)],
}
SPEC_SECTIONS = {
    "CIV": ["31 23 00 Excavation and Fill", "33 05 00 Common Work for Utilities", "03 30 00 Cast-in-Place Concrete"],
    "STR": ["05 12 00 Structural Steel Framing", "03 41 00 Precast Structural Concrete", "05 31 00 Steel Decking"],
    "ARC": ["13 21 26 Dry Room Assemblies", "09 67 23 Resinous Flooring", "08 71 00 Door Hardware"],
    "MEC": ["23 31 13 Metal Ducts", "23 73 00 Air Handling Units", "23 84 16 Dehumidification Units", "23 05 93 Testing, Adjusting and Balancing"],
    "PPG": ["22 61 13 Process Water Piping (RO/DI)", "40 05 13 Common Work for Process Piping", "40 05 63 Process Valves"],
    "PLB": ["22 66 00 Chemical-Waste Systems", "22 13 00 Facility Sanitary Sewerage"],
    "ELE": ["26 13 00 Medium-Voltage Switchgear", "26 24 13 Switchboards", "26 05 19 Low-Voltage Conductors"],
    "INC": ["25 50 00 Integrated Automation Facility Controls", "28 46 00 Fire Detection and Alarm", "40 91 00 Instrumentation"],
    "FPR": ["21 13 13 Wet-Pipe Sprinkler Systems", "28 31 00 Fire Detection and Alarm"],
    "PEQ": ["40 06 00 Process Equipment Hookup Schedules", "41 22 00 Cranes and Hoists"],
}
RFI_SUBJECTS = ["Clash / Interference", "Missing Information", "Design Clarification", "Field Condition Differs",
                "Substitution Request", "Tolerance / Dimension Conflict", "Embed / Penetration Location", "Spec Conflict"]
PUNCH_TYPES = ["Missing Label / Tag", "Incomplete Insulation", "Support Missing / Incorrect", "Damaged Finish",
               "Torque Not Verified", "Documentation Incomplete", "Cleanliness", "Leak", "Wiring Termination",
               "Access / Clearance", "Seismic Bracing", "Firestopping"]
NCR_CATS = ["Material Nonconformance", "Workmanship", "Weld Quality", "Moisture / Dew Point Excursion", "Wrong Revision Used",
            "Calibration Expired", "Document Control", "Installation Tolerance"]
ROOT_CAUSE = ["Procedure Not Followed", "Training", "Supplier Defect", "Design Error", "Communication", "Inadequate Inspection"]


def s_curve(t):
    """Logistic S-curve on 0..1 time scale returning cumulative fraction."""
    t = np.clip(t, 0, 1)
    y = 1 / (1 + np.exp(-10 * (t - 0.5)))
    y0, y1 = 1 / (1 + np.exp(5)), 1 / (1 + np.exp(-5))
    return (y - y0) / (y1 - y0)


def rand_dates(start, end, n):
    span = (pd.Timestamp(end) - pd.Timestamp(start)).days
    return pd.Timestamp(start) + pd.to_timedelta(rng.integers(0, max(span, 1), n), unit="D")


def clip_future(s):
    """Blank out actual dates that fall after the data date."""
    s = pd.to_datetime(s)
    return s.where(s <= DATA_DATE)


def main(out, scale):
    os.makedirs(out, exist_ok=True)

    # ------------------------------------------------------------------ dims
    dim_disc = pd.DataFrame(DISCIPLINES, columns=["Discipline_Code", "Discipline", "Discipline_Group"])
    areas = []
    for b, (bname, levels, zones) in BUILDINGS.items():
        for lv in levels:
            for z in zones:
                code = f"{b}-{lv.split()[0]}-{z}"
                areas.append((code, b, bname, lv, z))
    dim_area = pd.DataFrame(areas, columns=["Area_Code", "Building_Code", "Building", "Level", "Zone"])
    dim_area["Area_Sort"] = range(1, len(dim_area) + 1)

    comp = [("GC01", "Contoso Builders (General Contractor)", "General Contractor", "ALL"), ("OWN01", "Contoso Energy (Owner)", "Owner", "ALL"),
            ("DOR01", "Ridgeline Design Partners", "Design Partner", "ALL")]
    comp += [(c, n, "Subcontractor", d) for c, n, d in SUBCONTRACTORS]
    comp += [(f"SUP{i + 1:02d}", s, "Supplier", "") for i, s in enumerate(SUPPLIERS)]
    dim_company = pd.DataFrame(comp, columns=["Company_ID", "Company_Name", "Company_Type", "Primary_Discipline"])
    subs_by_disc = {d: [c for c, _, dd in SUBCONTRACTORS if dd == d] for d in DISC_CODES}

    # systems & subsystems with planned mechanical completion (MC) and turnover dates
    sysrows, subrows = [], []
    mc_start, mc_end = pd.Timestamp("2026-05-01"), pd.Timestamp("2027-09-30")
    for i, (code, name, disc, bld, nsub, tps) in enumerate(SYSTEMS):
        prio = 1 if code in PRIORITY_1 else 2 if code in PRIORITY_2 else 3
        sysrows.append((code, name, disc, bld, prio))
        nsub = max(2, int(round(nsub * min(1, scale ** 0.5))))
        base = mc_start + pd.Timedelta(days=int((prio - 1) * 120 + rng.integers(0, 90)))
        for k in range(1, nsub + 1):
            pmc = base + pd.Timedelta(days=int(k / nsub * 260 + rng.integers(-20, 20)))
            pmc = min(max(pmc, mc_start), mc_end)
            bld_areas = dim_area[dim_area.Building_Code == bld].Area_Code.tolist()
            subrows.append((f"{code}-{k:03d}", code, f"{name} - Subsystem {k:03d}", rng.choice(bld_areas), pmc,
                            pmc + pd.Timedelta(days=int(rng.integers(21, 60))), max(2, int(tps * rng.uniform(0.9, 2.1) * scale))))
    dim_system = pd.DataFrame(sysrows, columns=["System_Code", "System_Name", "Discipline_Code", "Building_Code", "Priority_Group"])
    sub = pd.DataFrame(subrows, columns=["Subsystem_Code", "System_Code", "Subsystem_Name", "Area_Code", "Planned_MC_Date",
                                         "Planned_Turnover_Date", "NTags"])

    # ------------------------------------------------------------------ CWPs + weekly earned value
    cwps = []
    disc_window = {"CIV": ("2024-09-01", 0.55), "STR": ("2024-12-01", 0.5), "ARC": ("2025-05-01", 0.6), "MEC": ("2025-06-01", 0.75),
                   "PPG": ("2025-08-01", 0.8), "PLB": ("2025-06-01", 0.7), "ELE": ("2025-05-01", 0.8), "INC": ("2025-11-01", 0.7),
                   "FPR": ("2025-07-01", 0.65), "PEQ": ("2026-03-01", 0.6)}
    n = 0
    for _, a in dim_area.iterrows():
        for d in DISC_CODES:
            if a.Building_Code == "UGS" and d not in ("CIV", "ELE", "PLB"):
                continue
            if a.Building_Code in ("NRU", "SUB", "WWT") and d in ("ARC", "PEQ"):
                continue
            if rng.random() < 0.15:
                continue
            n += 1
            start = pd.Timestamp(disc_window[d][0]) + pd.Timedelta(days=int(rng.integers(0, 200)))
            dur_wk = int(rng.integers(45, 105) * disc_window[d][1] + 35)
            budget = float(rng.lognormal(10.2, 0.7)) * (2.5 if a.Building_Code == MAIN_BUILDING else 1)
            perf = float(rng.normal(0.97, 0.08))        # productivity (earned / actual)
            slip = float(rng.normal(1.06, 0.08))        # schedule stretch
            cwps.append((f"CWP-{a.Building_Code}-{d}-{n:04d}", a.Area_Code, d, rng.choice(subs_by_disc[d] or ["GC01"]),
                         start, start + pd.Timedelta(weeks=dur_wk), round(budget, 0), perf, slip))
    cwp = pd.DataFrame(cwps, columns=["CWP_ID", "Area_Code", "Discipline_Code", "Company_ID", "Planned_Start", "Planned_Finish",
                                      "Budget_Hours", "PerfF", "SlipF"])
    weeks = pd.date_range("2024-09-06", "2027-12-31", freq="W-FRI")
    ev = []
    for r in cwp.itertuples():
        dur = (r.Planned_Finish - r.Planned_Start).days
        wk = weeks[(weeks >= r.Planned_Start - pd.Timedelta(days=7)) & (weeks <= r.Planned_Finish + pd.Timedelta(days=int(dur * (r.SlipF - 1)) + 120))]
        if len(wk) == 0:
            continue
        t_plan = (wk - r.Planned_Start).days / dur
        cum_plan = s_curve(t_plan) * r.Budget_Hours
        t_act = (wk - r.Planned_Start).days / (dur * r.SlipF)
        cum_earn = s_curve(t_act) * r.Budget_Hours * np.clip(rng.normal(1, 0.01, len(wk)), 0.97, 1.03)
        cum_earn = np.maximum.accumulate(np.minimum(cum_earn, r.Budget_Hours))
        p_plan = np.diff(np.r_[0, cum_plan])
        p_earn = np.diff(np.r_[0, cum_earn])
        p_act = p_earn / r.PerfF * rng.lognormal(0, 0.12, len(wk))
        past = wk <= DATA_DATE
        df = pd.DataFrame({"Week_Ending": wk, "CWP_ID": r.CWP_ID, "Planned_Hours": p_plan.round(1),
                           "Earned_Hours": np.where(past, p_earn.round(1), np.nan),
                           "Actual_Hours": np.where(past, p_act.round(1), np.nan)})
        df = df[(df.Planned_Hours > 0.05) | (df.Earned_Hours > 0.05)]
        ev.append(df)
    fact_progress = pd.concat(ev, ignore_index=True)
    cwp_out = cwp.drop(columns=["PerfF", "SlipF"])

    # ------------------------------------------------------------------ document register
    n_docs = int(30000 * scale)
    doc_types = np.array(["Drawing", "P&ID", "Isometric", "Specification", "Datasheet", "Calculation", "Vendor Drawing", "Procedure", "Report"])
    dt_p = np.array([0.34, 0.06, 0.2, 0.05, 0.1, 0.06, 0.13, 0.03, 0.03])
    d_disc = rng.choice(DISC_CODES, n_docs, p=[0.07, 0.1, 0.08, 0.15, 0.2, 0.05, 0.16, 0.09, 0.05, 0.05])
    d_type = rng.choice(doc_types, n_docs, p=dt_p)
    d_area = rng.choice(dim_area.Area_Code.values, n_docs)
    d_ifc_plan = NTP + pd.to_timedelta(rng.integers(60, 1060, n_docs), unit="D")
    vendor = d_type == "Vendor Drawing"
    d_orig = np.where(vendor, rng.choice([f"SUP{i + 1:02d}" for i in range(len(SUPPLIERS))], n_docs),
                      np.where(rng.random(n_docs) < 0.3, "DOR01", "GC01"))
    abbrev = {"Drawing": "DWG", "P&ID": "PID", "Isometric": "ISO", "Specification": "SPC", "Datasheet": "DS",
              "Calculation": "CAL", "Vendor Drawing": "VDR", "Procedure": "PRC", "Report": "RPT"}
    doc_no = [f"SAG-{a.split('-')[0]}-{d}-{abbrev[t]}-{i:05d}" for i, (a, d, t) in enumerate(zip(d_area, d_disc, d_type), start=10001)]
    docSlipF = rng.normal(12, 25, n_docs)          # days late vs plan (engineering)
    rows = []
    rev_seq = ["A", "B", "0", "1", "2", "3"]
    purpose = {"A": "IFR", "B": "IFA", "0": "IFC", "1": "IFC", "2": "IFC", "3": "IFC"}
    for i in range(n_docs):
        n_ifc_revs = int(rng.choice([1, 1, 1, 2, 2, 3, 4], p=[0.3, 0.2, 0.15, 0.15, 0.1, 0.07, 0.03]))
        seq = (["A", "B"] if rng.random() < 0.7 else ["A"]) + rev_seq[2:2 + n_ifc_revs]
        ifc_plan = d_ifc_plan[i]
        offs = {"A": -110, "B": -50, "0": 0, "1": 75, "2": 150, "3": 230}
        for j, rv in enumerate(seq):
            plan = ifc_plan + pd.Timedelta(days=offs[rv] + int(rng.integers(-10, 10)))
            act = plan + pd.Timedelta(days=int(docSlipF[i] + rng.normal(0, 10)))
            due = act + pd.Timedelta(days=14 if purpose[rv] != "IFC" else 10)
            ret = act + pd.Timedelta(days=int(max(2, rng.gamma(2.5, 5))))
            code = rng.choice(["1", "2", "3", "4"], p=[0.42, 0.38, 0.15, 0.05]) if purpose[rv] in ("IFR", "IFA") or vendor[i] else ""
            rows.append((doc_no[i], rv, purpose[rv], plan, act, due, ret, code, "Y" if j == len(seq) - 1 else "N"))
        if d_type[i] in ("Drawing", "P&ID", "Isometric") and rng.random() < 0.5:
            plan = ifc_plan + pd.Timedelta(days=int(rng.integers(500, 800)))
            rows.append((doc_no[i], "AB", "As-Built", plan, plan + pd.Timedelta(days=int(rng.normal(20, 30))), pd.NaT, pd.NaT, "", "N"))
    docs = pd.DataFrame(rows, columns=["Doc_No", "Revision", "Purpose_Code", "Planned_Issue_Date", "Actual_Issue_Date",
                                       "Review_Due_Date", "Review_Return_Date", "Review_Code", "Is_Latest_Planned"])
    meta = pd.DataFrame({"Doc_No": doc_no, "Discipline_Code": d_disc, "Doc_Type": d_type, "Area_Code": d_area,
                         "Originator_ID": d_orig, "Is_Vendor_Doc": np.where(vendor, "Y", "N")})
    docs = docs.merge(meta, on="Doc_No")
    docs["Actual_Issue_Date"] = clip_future(docs["Actual_Issue_Date"])
    issued = docs["Actual_Issue_Date"].notna()
    docs.loc[~issued, ["Review_Due_Date", "Review_Return_Date", "Review_Code"]] = [pd.NaT, pd.NaT, ""]
    docs["Review_Return_Date"] = clip_future(docs["Review_Return_Date"])
    docs.loc[docs.Review_Return_Date.isna(), "Review_Code"] = ""
    docs.loc[docs.Purpose_Code.isin(["IFC", "As-Built"]) & (docs.Is_Vendor_Doc == "N"), ["Review_Due_Date", "Review_Return_Date"]] = [pd.NaT, pd.NaT]
    docs.insert(0, "Doc_Rev_ID", np.arange(1, len(docs) + 1))
    docs = docs[["Doc_Rev_ID", "Doc_No", "Revision", "Purpose_Code", "Discipline_Code", "Doc_Type", "Area_Code", "Originator_ID",
                 "Is_Vendor_Doc", "Planned_Issue_Date", "Actual_Issue_Date", "Review_Due_Date", "Review_Return_Date",
                 "Review_Code", "Is_Latest_Planned"]]

    # ------------------------------------------------------------------ RFIs
    n_rfi = int(28000 * scale)
    raised = rand_dates("2024-10-01", DATA_DATE, n_rfi)
    # more RFIs as construction ramps
    keep = rng.random(n_rfi) < np.clip((raised - pd.Timestamp("2024-10-01")).days / 500, 0.15, 1)
    raised = raised[keep]
    n_rfi = len(raised)
    rdisc = rng.choice(DISC_CODES, n_rfi, p=[0.08, 0.12, 0.1, 0.15, 0.18, 0.06, 0.14, 0.07, 0.05, 0.05])
    prio = rng.choice(["Critical", "High", "Medium", "Low"], n_rfi, p=[0.06, 0.24, 0.5, 0.2])
    req_days = pd.Series(prio).map({"Critical": 3, "High": 7, "Medium": 10, "Low": 14}).values
    resp = rng.gamma(2.0, pd.Series(prio).map({"Critical": 2.5, "High": 4.5, "Medium": 6.5, "Low": 9}).values)
    responded = clip_future(raised + pd.to_timedelta(resp.round(), unit="D"))
    closed = clip_future(responded + pd.to_timedelta(rng.gamma(1.5, 3, n_rfi).round(), unit="D"))
    status = np.where(closed.notna(), "Closed", np.where(responded.notna(), "Responded", "Open"))
    status = np.where(rng.random(n_rfi) < 0.02, "Void", status)
    rfi = pd.DataFrame({
        "RFI_No": [f"RFI-{i:05d}" for i in range(1, n_rfi + 1)], "Discipline_Code": rdisc,
        "Area_Code": rng.choice(dim_area.Area_Code.values, n_rfi),
        "Raised_By_ID": [rng.choice(subs_by_disc[d]) for d in rdisc],
        "Subject_Category": rng.choice(RFI_SUBJECTS, n_rfi), "Priority": prio, "Date_Raised": raised,
        "Date_Required": raised + pd.to_timedelta(req_days, unit="D"), "Date_Responded": responded, "Date_Closed": closed,
        "Status": status, "Cost_Impact": np.where(rng.random(n_rfi) < 0.11, "Y", "N"),
        "Schedule_Impact_Days": np.where(rng.random(n_rfi) < 0.08, rng.integers(1, 30, n_rfi), 0),
        "Linked_Doc_No": rng.choice(doc_no, n_rfi)}).sort_values("Date_Raised")

    # ------------------------------------------------------------------ submittals
    n_sub = int(42000 * scale)
    sdisc = rng.choice(DISC_CODES, n_sub, p=[0.07, 0.1, 0.1, 0.16, 0.18, 0.06, 0.16, 0.07, 0.05, 0.05])
    sdate = rand_dates("2024-08-01", DATA_DATE, n_sub)
    stype = rng.choice(["Shop Drawing", "Product Data", "Test Report", "Sample", "O&M Manual", "Certificate", "Method Statement"],
                       n_sub, p=[0.3, 0.27, 0.15, 0.04, 0.08, 0.1, 0.06])
    due = sdate + pd.Timedelta(days=14)
    ret = clip_future(sdate + pd.to_timedelta(rng.gamma(3, 4.5, n_sub).round(), unit="D"))
    action = np.where(ret.notna(), rng.choice(["A - Approved", "B - Approved as Noted", "C - Revise & Resubmit", "D - Rejected"],
                                                n_sub, p=[0.45, 0.33, 0.17, 0.05]), "")
    subm = pd.DataFrame({
        "Submittal_No": [f"SUBM-{i:05d}" for i in range(1, n_sub + 1)],
        "Spec_Section": [rng.choice(SPEC_SECTIONS[d]) for d in sdisc], "Discipline_Code": sdisc,
        "Company_ID": [rng.choice(subs_by_disc[d]) for d in sdisc], "Submittal_Type": stype,
        "Revision": rng.choice([0, 0, 0, 1, 1, 2], n_sub), "Date_Submitted": sdate, "Date_Due": due, "Date_Returned": ret,
        "Action_Code": action, "Status": np.where(ret.notna(), "Returned", "Under Review"),
        "Critical_Path": np.where(rng.random(n_sub) < 0.09, "Y", "N")}).sort_values("Date_Submitted")

    # ------------------------------------------------------------------ procurement
    n_po = int(5200 * scale)
    po_rows = []
    sup_ids = [f"SUP{i + 1:02d}" for i in range(len(SUPPLIERS))]
    for p in range(1, n_po + 1):
        d = rng.choice(DISC_CODES, p=[0.06, 0.08, 0.07, 0.16, 0.2, 0.06, 0.18, 0.08, 0.05, 0.06])
        sup = rng.choice(sup_ids)
        odate = rand_dates("2024-06-01", "2026-10-31", 1)[0]
        nlines = int(rng.integers(3, 18))
        for ln in range(1, nlines + 1):
            opts = COMMODITIES[d]
            wts = np.array([0.012 if o[3] else 1.0 for o in opts]); wts /= wts.sum()
            com, unit, price, long_lead = opts[int(rng.choice(len(opts), p=wts))]
            qty = int(rng.integers(1, 3)) if price > 100000 else int(rng.integers(1, 12)) if price > 5000 else int(rng.lognormal(4.2, 1.0)) + 1
            lead = int(rng.normal(52, 10) * 7) if long_lead else int(rng.normal(12, 4) * 7)
            ros = odate + pd.Timedelta(days=lead + int(rng.integers(14, 90)))
            promised = odate + pd.Timedelta(days=lead)
            forecast = promised + pd.Timedelta(days=int(rng.normal(8, 25) if long_lead else rng.normal(3, 12)))
            delivered = forecast + pd.Timedelta(days=int(rng.normal(0, 5)))
            po_rows.append((f"PO-{p:05d}", ln, d, sup, com, unit, qty, round(qty * price * rng.uniform(0.9, 1.1), 2),
                            "Y" if long_lead else "N", odate, ros, promised, forecast, delivered))
    proc = pd.DataFrame(po_rows, columns=["PO_No", "PO_Line", "Discipline_Code", "Supplier_ID", "Commodity", "Unit", "Quantity",
                                          "Line_Value_USD", "Long_Lead", "Order_Date", "ROS_Date", "Promised_Date",
                                          "Forecast_Delivery_Date", "Actual_Delivery_Date"])
    proc["Actual_Delivery_Date"] = clip_future(proc["Actual_Delivery_Date"])
    proc["Area_Code"] = rng.choice(dim_area.Area_Code.values, len(proc))
    proc["Expediting_Status"] = np.select(
        [proc.Actual_Delivery_Date.notna(), proc.Forecast_Delivery_Date > proc.ROS_Date,
         (proc.ROS_Date - proc.Forecast_Delivery_Date).dt.days < 14], ["Delivered", "Late vs ROS", "At Risk"], "On Track")

    # ------------------------------------------------------------------ tags, ITRs, punch, subsystem actuals
    tag_rows = []
    sysdisc = dict(zip(dim_system.System_Code, dim_system.Discipline_Code))
    for s in sub.itertuples():
        d = sysdisc[s.System_Code]
        types = TAG_TYPES[d]
        codes = rng.choice(len(types), s.NTags, p=[t[2] for t in types])
        # install progress depends on how close planned MC is
        days_to_mc = (s.Planned_MC_Date - DATA_DATE).days
        p_inst = float(np.clip(1 - days_to_mc / 300, 0.02, 0.995))
        if days_to_mc < -15 and rng.random() < 0.8:
            p_inst = 1.0   # most subsystems past planned MC are fully installed
        inst = rng.random(s.NTags) < p_inst
        for k in range(s.NTags):
            ttype = types[codes[k]]
            idate = s.Planned_MC_Date - pd.Timedelta(days=int(rng.integers(20, 260)))
            if inst[k] and idate > DATA_DATE:
                idate = DATA_DATE - pd.Timedelta(days=int(rng.integers(0, 60)))
            tag_rows.append((f"{s.System_Code}-{ttype[0]}-{s.Subsystem_Code[-3:]}{k + 1:04d}", ttype[1], d, s.System_Code,
                             s.Subsystem_Code, s.Area_Code, "Installed" if inst[k] else "Not Installed",
                             idate if inst[k] else pd.NaT))
    tags = pd.DataFrame(tag_rows, columns=["Tag_No", "Tag_Type", "Discipline_Code", "System_Code", "Subsystem_Code", "Area_Code",
                                           "Install_Status", "Install_Date"])

    # ITRs: each tag gets the A forms and B forms for its discipline (not every form applies to every tag)
    itr_parts = []
    sub_mc = dict(zip(sub.Subsystem_Code, sub.Planned_MC_Date))
    for d, forms in ITR_FORMS.items():
        t = tags[tags.Discipline_Code == d]
        if t.empty:
            continue
        for code, desc, stage in forms:
            take = rng.random(len(t)) < (0.95 if stage == "A" else 0.6)
            tt = t[take]
            pmc = tt.Subsystem_Code.map(sub_mc)
            planned = pmc - pd.to_timedelta(rng.integers(20, 90, len(tt)), unit="D") if stage == "A" else \
                pmc + pd.to_timedelta(rng.integers(5, 50, len(tt)), unit="D")
            base = tt.Install_Date.fillna(pd.Timestamp("2100-01-01"))
            done_dt = np.maximum(base + pd.to_timedelta(rng.integers(3, 45, len(tt)), unit="D"),
                                 planned + pd.to_timedelta(rng.normal(4, 12, len(tt)).round(), unit="D"))
            done_dt = clip_future(pd.Series(done_dt, index=tt.index))
            if stage == "B":   # B sheets only after A done; keep fewer complete
                done_dt = done_dt.where(rng.random(len(tt)) < 0.8)
            itr_parts.append(pd.DataFrame({"Tag_No": tt.Tag_No.values, "Subsystem_Code": tt.Subsystem_Code.values,
                                           "Discipline_Code": d, "ITR_Form": code, "ITR_Description": desc, "Stage": stage,
                                           "Planned_Date": planned.values, "Completed_Date": done_dt.values}))
    itr = pd.concat(itr_parts, ignore_index=True)
    itr.insert(0, "ITR_ID", [f"ITR-{i:07d}" for i in range(1, len(itr) + 1)])
    itr["Status"] = np.where(itr.Completed_Date.notna(), "Completed", "Not Started")
    accepted = itr.Completed_Date.notna() & (rng.random(len(itr)) < 0.85)
    itr.loc[accepted, "Status"] = "Signed Off"
    itr["Completed_By_ID"] = itr.Discipline_Code.map(lambda d: rng.choice(subs_by_disc[d]))

    # punch: raised during walkdowns from ~60 days before MC
    pr = []
    for s in sub.itertuples():
        walk = s.Planned_MC_Date - pd.Timedelta(days=60)
        if walk > DATA_DATE:
            continue
        n_p = int(rng.poisson(max(3, s.NTags * 0.75)))
        rd = walk + pd.to_timedelta(rng.integers(0, 75, n_p), unit="D")
        rd = rd[rd <= DATA_DATE]
        if len(rd) == 0:
            continue
        cat = rng.choice(["A", "B", "C"], len(rd), p=[0.18, 0.52, 0.30])
        clear = rd + pd.to_timedelta(rng.gamma(2, np.where(cat == "A", 7, 14)).round(), unit="D")
        st_tags = tags.loc[tags.Subsystem_Code == s.Subsystem_Code, "Tag_No"].values
        d = sysdisc[s.System_Code]
        for k in range(len(rd)):
            pr.append((s.Subsystem_Code, rng.choice(st_tags), s.Area_Code, d, cat[k], rng.choice(PUNCH_TYPES), rd[k],
                       rd[k] + pd.Timedelta(days=14 if cat[k] == "A" else 30), clear[k],
                       rng.choice(["GC01", "OWN01", "GC01", "GC01"]), rng.choice(subs_by_disc[d])))
    punch = pd.DataFrame(pr, columns=["Subsystem_Code", "Tag_No", "Area_Code", "Discipline_Code", "Category", "Punch_Type",
                                      "Date_Raised", "Target_Date", "Date_Cleared", "Raised_By_ID", "Responsible_ID"])
    punch["Date_Cleared"] = clip_future(punch["Date_Cleared"])
    punch["Status"] = np.where(punch.Date_Cleared.notna(), "Cleared", "Open")
    punch.insert(0, "Punch_No", [f"PL-{i:06d}" for i in range(1, len(punch) + 1)])

    # subsystem actual MC / turnover: all A ITRs complete and no open Cat A punch
    a_open = itr[(itr.Stage == "A") & itr.Completed_Date.isna()].groupby("Subsystem_Code").size()
    a_last = itr[itr.Stage == "A"].groupby("Subsystem_Code").Completed_Date.max()
    pa_open = punch[(punch.Category == "A") & (punch.Status == "Open")].groupby("Subsystem_Code").size()
    pa_last = punch[punch.Category == "A"].groupby("Subsystem_Code").Date_Cleared.max()
    sub["Actual_MC_Date"] = pd.NaT
    ok = ~sub.Subsystem_Code.isin(a_open.index) & ~sub.Subsystem_Code.isin(pa_open.index) & sub.Subsystem_Code.isin(a_last.index)
    sub.loc[ok, "Actual_MC_Date"] = pd.concat([sub.loc[ok, "Subsystem_Code"].map(a_last),
                                               sub.loc[ok, "Subsystem_Code"].map(pa_last)], axis=1).max(axis=1)
    sub["Actual_Turnover_Date"] = clip_future(sub["Actual_MC_Date"] + pd.to_timedelta(rng.integers(14, 50, len(sub)), unit="D"))
    sub["Actual_Turnover_Date"] = sub["Actual_Turnover_Date"].where(sub.Actual_MC_Date.notna())
    dim_subsystem = sub.drop(columns="NTags")

    # ------------------------------------------------------------------ safety & quality
    wk_hist = weeks[weeks <= DATA_DATE]
    saf = []
    for c, _, d in SUBCONTRACTORS + [("GC01", "", "ALL")]:
        start = pd.Timestamp(disc_window.get(d, ("2024-09-01", 0))[0]) - pd.Timedelta(days=30)
        for wkend in wk_hist[wk_hist >= start]:
            ramp = min(1.0, (wkend - start).days / 300)
            head = int(max(5, rng.normal(220 if c != "GC01" else 400, 50) * (0.2 + 0.8 * ramp) * scale ** 0.3))
            mh = head * rng.uniform(48, 58)
            rec = rng.poisson(mh * 0.9 / 200000)
            saf.append((wkend, c, head, round(mh), rec, rng.poisson(mh * 0.12 / 200000) if rec else 0,
                        rng.poisson(mh * 4 / 200000), rng.poisson(mh * 18 / 200000), rng.poisson(mh * 300 / 200000)))
    safety = pd.DataFrame(saf, columns=["Week_Ending", "Company_ID", "Headcount", "Manhours", "Recordables", "Lost_Time_Injuries",
                                        "First_Aids", "Near_Misses", "Safety_Observations"])
    n_ncr = int(2600 * scale)
    ndisc = rng.choice(DISC_CODES, n_ncr)
    opened = rand_dates("2024-11-01", DATA_DATE, n_ncr)
    sev = rng.choice(["Minor", "Major", "Critical"], n_ncr, p=[0.7, 0.26, 0.04])
    ncr = pd.DataFrame({"NCR_No": [f"NCR-{i:05d}" for i in range(1, n_ncr + 1)], "Discipline_Code": ndisc,
                        "Company_ID": [rng.choice(subs_by_disc[d]) for d in ndisc],
                        "Area_Code": rng.choice(dim_area.Area_Code.values, n_ncr), "Category": rng.choice(NCR_CATS, n_ncr),
                        "Severity": sev, "Root_Cause": rng.choice(ROOT_CAUSE, n_ncr), "Date_Opened": opened,
                        "Date_Closed": clip_future(opened + pd.to_timedelta(rng.gamma(2, np.where(sev == "Minor", 8, 18)).round(), unit="D"))})
    ncr["Status"] = np.where(ncr.Date_Closed.notna(), "Closed", "Open")

    # ------------------------------------------------------------------ write as separate source systems
    project = pd.DataFrame({"Data_Date": [DATA_DATE], "Project": ["Project Saguaro - GF1 Cell Gigafactory"],
                            "Owner": ["Contoso Energy (fictional)"]})
    tables = {
        "dim_discipline": dim_disc, "dim_area": dim_area, "dim_company": dim_company, "dim_system": dim_system,
        "dim_subsystem": dim_subsystem, "dim_cwp": cwp_out, "project_info": project,
        "fact_document_revisions": docs, "fact_rfi": rfi, "fact_submittals": subm, "fact_procurement": proc,
        "fact_progress_weekly": fact_progress, "fact_tag_register": tags, "fact_itr": itr, "fact_punch": punch,
        "fact_safety_weekly": safety, "fact_ncr": ncr,
    }
    for name, df in tables.items():
        print(f"{name:28s}{len(df):>9,} rows")
    export_sources(tables, out, rng)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="../data/gigafactory")
    ap.add_argument("--scale", type=float, default=1.0)
    a = ap.parse_args()
    main(a.out, a.scale)
