"""Integrated project KPI report across all six Project Saguaro source systems.

Pulls from SQL Server, the ERP Excel export, the P6 export, the completions
system CSV folder, the HSE JSON API file and the SharePoint NCR export; checks
that the keys line up between systems; and writes one KPI table per discipline.

Usage:
    pip install pandas openpyxl            (add pyodbc to read from SQL Server)
    python integrated_kpi_report.py --root C:\\PowerBI\\gigafactory
    python integrated_kpi_report.py --root C:\\PowerBI\\gigafactory --sql-server "localhost\\SQLEXPRESS"

Without --sql-server the SQL tables are read from the CSV files the database is loaded from.
"""
import argparse
import glob
import json
import os

import pandas as pd

SQL_TABLES = {"discipline": "ref.discipline", "area": "ref.area", "company": "ref.company", "xref": "ref.company_erp_xref",
              "subsystem": "ref.subsystem", "project": "ref.project_info", "docs": "eng.document_revision",
              "rfi": "eng.rfi", "submittals": "eng.submittal"}


def read_sql_sources(root, server):
    """Read the engineering database, from SQL Server if a server is given, else from its load files."""
    if server:
        import pyodbc
        cn = pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};"
                            "DATABASE=SaguaroPIIM;Trusted_Connection=yes;")
        out = {k: pd.read_sql(f"SELECT * FROM {t}", cn) for k, t in SQL_TABLES.items()}
        cn.close()
    else:
        csv = os.path.join(root, "sql_server", "csv")
        out = {k: pd.read_csv(os.path.join(csv, t.replace(".", "_") + ".csv"), low_memory=False)
               for k, t in SQL_TABLES.items()}
    for df in out.values():
        for c in df.columns:
            if "Date" in c:
                df[c] = pd.to_datetime(df[c], errors="coerce")
    return out


def read_other_sources(root):
    erp = pd.read_excel(os.path.join(root, "erp_procurement", "PO_Lines_Export.xlsx"), sheet_name="PO_Lines")
    p6_cwp = pd.read_excel(os.path.join(root, "p6_schedule", "Construction_Progress_Export.xlsx"), sheet_name="WBS_CWP")
    p6_wk = pd.read_excel(os.path.join(root, "p6_schedule", "Construction_Progress_Export.xlsx"), sheet_name="Weekly_Progress")
    files = sorted(glob.glob(os.path.join(root, "completions_system", "itr_extracts", "itr_extract_*.csv")))
    itr = pd.concat((pd.read_csv(f, parse_dates=["Planned_Date", "Completed_Date"]).assign(Source_File=os.path.basename(f))
                     for f in files), ignore_index=True)
    tags = pd.read_csv(os.path.join(root, "completions_system", "tag_register.csv"), parse_dates=["Install_Date"])
    punch = pd.read_csv(os.path.join(root, "completions_system", "punch_list.csv"),
                        parse_dates=["Date_Raised", "Target_Date", "Date_Cleared"])
    hse = pd.DataFrame(json.load(open(os.path.join(root, "hse_api", "safety_weekly.json")))["value"])
    hse["weekEnding"] = pd.to_datetime(hse["weekEnding"])
    ncr = pd.read_excel(os.path.join(root, "quality_sharepoint", "NCR_Log.xlsx"), sheet_name="NCR Log")
    return dict(erp=erp, p6_cwp=p6_cwp, p6_wk=p6_wk, itr=itr, itr_files=len(files), tags=tags, punch=punch, hse=hse, ncr=ncr)


def data_quality_checks(s, o):
    """Key integrity between systems. On a live project these catch broken extracts and unmapped vendors."""
    checks = []
    unmapped = ~o["erp"]["Vendor No"].isin(s["xref"]["ERP_Vendor_No"])
    checks.append(("ERP vendor numbers without a company cross-reference", int(unmapped.sum())))
    checks.append(("ITR rows duplicated across extract files", int(o["itr"].duplicated("ITR_ID").sum())))
    checks.append(("ITRs on subsystems missing from the SQL subsystem table",
                   int((~o["itr"]["Subsystem_Code"].isin(s["subsystem"]["Subsystem_Code"])).sum())))
    checks.append(("ITRs on tags missing from the tag register", int((~o["itr"]["Tag_No"].isin(o["tags"]["Tag_No"])).sum())))
    checks.append(("Punch items on unknown tags", int((~o["punch"]["Tag_No"].isin(o["tags"]["Tag_No"])).sum())))
    checks.append(("P6 work packages with an unknown area code",
                   int((~o["p6_cwp"]["Area"].isin(s["area"]["Area_Code"])).sum())))
    checks.append(("HSE rows for companies not in the company table",
                   int((~o["hse"]["companyId"].isin(s["company"]["Company_ID"])).sum())))
    checks.append(("NCRs raised against companies not in the company table",
                   int((~o["ncr"]["Contractor"].isin(s["company"]["Company_ID"])).sum())))
    return pd.DataFrame(checks, columns=["Check", "Rows failing"])


def kpis_by_discipline(s, o):
    dd = s["project"]["Data_Date"].iloc[0]
    docs, rfi, sub = s["docs"], s["rfi"], s["submittals"]
    ifc = docs[docs.Revision.astype(str) == "0"]
    eng = ifc.groupby("Discipline_Code").agg(
        IFC_Docs=("Doc_No", "count"),
        IFC_Planned_Pct=("Planned_Issue_Date", lambda x: (x <= dd).mean() * 100),
        IFC_Issued_Pct=("Actual_Issue_Date", lambda x: x.notna().mean() * 100))
    live = rfi[rfi.Status != "Void"]
    rf = live.groupby("Discipline_Code").agg(
        RFIs_Open=("Date_Closed", lambda x: x.isna().sum()),
        RFIs_Overdue=("RFI_No", lambda x: ((live.loc[x.index, "Date_Responded"].isna())
                                           & (live.loc[x.index, "Date_Required"] < dd)).sum()),
        Avg_Days_To_Respond=("Date_Responded", lambda x: (x - live.loc[x.index, "Date_Raised"]).dt.days.mean()))
    sb = sub.groupby("Discipline_Code").agg(
        Submittals_Under_Review=("Status", lambda x: (x == "Under Review").sum()))

    erp = o["erp"].rename(columns={"Discipline": "Discipline_Code"})
    erp["late"] = erp["Expediting Status"].eq("Late vs ROS")
    pr = erp.groupby("Discipline_Code").agg(
        PO_Value_USD_M=("Net Value (USD)", lambda x: x.sum() / 1e6),
        PO_Lines_Late_vs_ROS=("late", "sum"))

    wk = o["p6_wk"].merge(o["p6_cwp"][["CWP ID", "Discipline", "Budgeted Units (hrs)"]], on="CWP ID")
    wk = wk[pd.to_datetime(wk["Week Ending"]) <= dd]
    budget = o["p6_cwp"].groupby("Discipline")["Budgeted Units (hrs)"].sum()
    ev = wk.groupby("Discipline")[["Planned Units", "Earned Units", "Actual Units"]].sum()
    ev = pd.DataFrame({"Earned_Pct": ev["Earned Units"] / budget * 100,
                       "Planned_Pct": ev["Planned Units"] / budget * 100,
                       "SPI": ev["Earned Units"] / ev["Planned Units"],
                       "CPI": ev["Earned Units"] / ev["Actual Units"]})
    ev.index.name = "Discipline_Code"

    itr = o["itr"].drop_duplicates("ITR_ID")
    a = itr[itr.Stage == "A"].groupby("Discipline_Code")["Completed_Date"].apply(lambda x: x.notna().mean() * 100)
    p = o["punch"]
    pa = p[(p.Category == "A") & p.Date_Cleared.isna()].groupby("Discipline_Code").size()
    comp = pd.DataFrame({"ITR_A_Complete_Pct": a, "Punch_A_Open": pa})

    ncr = o["ncr"].groupby("Discipline")["NCR Status"].apply(lambda x: (x == "Open").sum()).rename("NCRs_Open")
    ncr.index.name = "Discipline_Code"

    out = (s["discipline"].set_index("Discipline_Code")[["Discipline"]]
           .join([eng, rf, sb, pr, ev, comp, ncr]).fillna({"Punch_A_Open": 0, "NCRs_Open": 0}))
    return out.round(2).reset_index()


def project_summary(s, o):
    dd = s["project"]["Data_Date"].iloc[0]
    k = kpis_by_discipline(s, o)
    hse = o["hse"]
    wk = o["p6_wk"][pd.to_datetime(o["p6_wk"]["Week Ending"]) <= dd]
    return pd.DataFrame([
        ("Data date", dd.date()),
        ("IFC documents issued %", round(k.IFC_Issued_Pct.mul(k.IFC_Docs).sum() / k.IFC_Docs.sum(), 1)),
        ("Construction earned % (hours)", round(wk["Earned Units"].sum() / o["p6_cwp"]["Budgeted Units (hrs)"].sum() * 100, 1)),
        ("SPI", round(wk["Earned Units"].sum() / wk["Planned Units"].sum(), 2)),
        ("CPI", round(wk["Earned Units"].sum() / wk["Actual Units"].sum(), 2)),
        ("Open RFIs", int(k.RFIs_Open.sum())),
        ("PO lines late vs required-on-site", int(k.PO_Lines_Late_vs_ROS.sum())),
        ("Subsystems mechanically complete", int(s["subsystem"]["Actual_MC_Date"].notna().sum())),
        ("Open category A punch", int(k.Punch_A_Open.sum())),
        ("TRIR (project to date)", round(hse.recordables.sum() * 200000 / hse.manhours.sum(), 2)),
        ("Open NCRs", int(k.NCRs_Open.sum())),
    ], columns=["KPI", "Value"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=os.path.join("..", "..", "data", "gigafactory"))
    ap.add_argument("--sql-server", default=None)
    ap.add_argument("--out", default="kpi_report.xlsx")
    a = ap.parse_args()
    s = read_sql_sources(a.root, a.sql_server)
    o = read_other_sources(a.root)
    print(f"Loaded 6 sources: {sum(len(v) for v in s.values()):,} SQL rows, {len(o['erp']):,} ERP lines, "
          f"{len(o['p6_wk']):,} P6 rows, {len(o['itr']):,} ITRs from {o['itr_files']} files, "
          f"{len(o['hse']):,} HSE rows, {len(o['ncr']):,} NCRs")
    dq, kpi, summ = data_quality_checks(s, o), kpis_by_discipline(s, o), project_summary(s, o)
    print("\nData quality checks\n", dq.to_string(index=False))
    print("\nProject summary\n", summ.to_string(index=False))
    with pd.ExcelWriter(a.out) as xw:
        summ.to_excel(xw, sheet_name="Summary", index=False)
        kpi.to_excel(xw, sheet_name="KPIs by Discipline", index=False)
        dq.to_excel(xw, sheet_name="Data Quality", index=False)
    print(f"\nWrote {a.out}")


if __name__ == "__main__":
    main()
