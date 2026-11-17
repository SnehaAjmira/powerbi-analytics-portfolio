"""Write the generated gigafactory tables out as six separate "source systems".

A real project keeps its data in different places, with different column
names, ID formats and file types. This module recreates that on purpose so
the Power BI model has to connect to, clean and integrate several sources:

  sql_server/          SQL Server engineering database: DDL, load script, CSVs, views, analysis queries
  erp_procurement/     ERP purchase order export (Excel). Uses ERP vendor numbers, not company IDs
  p6_schedule/         Schedule / progress export (Excel) with P6-style column names
  completions_system/  ITR extracts split into monthly CSV files + tag register + punch list
  hse_api/             Weekly safety statistics as an OData-style JSON API response
  quality_sharepoint/  NCR log exported from a SharePoint list (Excel), with SharePoint metadata columns
"""
import datetime as dt
import json
import os

import numpy as np
import pandas as pd

from xlsx_tables import StreamingWorkbook

# --------------------------------------------------------------------------- SQL Server
SQL_TABLES = [  # (schema.table, source frame, primary key, foreign keys {col: ref table(col)})
    ("ref.discipline", "dim_discipline", "Discipline_Code", {}),
    ("ref.area", "dim_area", "Area_Code", {}),
    ("ref.company", "dim_company", "Company_ID", {}),
    ("ref.company_erp_xref", "company_erp_xref", "Company_ID", {"Company_ID": "ref.company(Company_ID)"}),
    ("ref.facility_system", "dim_system", "System_Code", {}),
    ("ref.subsystem", "dim_subsystem", "Subsystem_Code",
     {"System_Code": "ref.facility_system(System_Code)", "Area_Code": "ref.area(Area_Code)"}),
    ("ref.project_info", "project_info", None, {}),
    ("eng.document_revision", "fact_document_revisions", "Doc_Rev_ID",
     {"Discipline_Code": "ref.discipline(Discipline_Code)", "Area_Code": "ref.area(Area_Code)",
      "Originator_ID": "ref.company(Company_ID)"}),
    ("eng.rfi", "fact_rfi", "RFI_No",
     {"Discipline_Code": "ref.discipline(Discipline_Code)", "Area_Code": "ref.area(Area_Code)",
      "Raised_By_ID": "ref.company(Company_ID)"}),
    ("eng.submittal", "fact_submittals", "Submittal_No",
     {"Discipline_Code": "ref.discipline(Discipline_Code)", "Company_ID": "ref.company(Company_ID)"}),
]


def _sql_type(s: pd.Series) -> str:
    if pd.api.types.is_datetime64_any_dtype(s):
        return "date"
    if pd.api.types.is_integer_dtype(s):
        return "int"
    if pd.api.types.is_float_dtype(s):
        return "decimal(18,2)"
    n = int(s.astype(str).str.len().max() or 1)
    size = max(10, int(np.ceil(n * 1.5 / 10) * 10))
    return f"nvarchar({size})"


def _write_sql_server(tables, root):
    csv_dir = os.path.join(root, "csv")
    os.makedirs(csv_dir, exist_ok=True)
    ddl = ["-- Engineering database for Project Saguaro (fictional). SQL Server 2017 or later.",
           "-- Run in SSMS or Azure Data Studio against your SQL Server / SQL Server Express instance.", "",
           "IF DB_ID('SaguaroPIIM') IS NULL CREATE DATABASE SaguaroPIIM;", "GO", "USE SaguaroPIIM;", "GO",
           "IF SCHEMA_ID('ref') IS NULL EXEC('CREATE SCHEMA ref');",
           "IF SCHEMA_ID('eng') IS NULL EXEC('CREATE SCHEMA eng');",
           "IF SCHEMA_ID('rpt') IS NULL EXEC('CREATE SCHEMA rpt');", "GO", ""]
    for full, *_ in reversed(SQL_TABLES):
        ddl.append(f"IF OBJECT_ID('{full}') IS NOT NULL DROP TABLE {full};")
    ddl += ["GO", ""]
    load = ["-- Loads the CSV files into SaguaroPIIM.",
            "-- 1. Copy the sql_server\\csv folder to a path the SQL Server service can read, e.g. C:\\PowerBI\\gigafactory\\sql_server\\csv\\",
            "-- 2. Set @path below, then run the whole script.", "",
            "USE SaguaroPIIM;", "GO", "SET NOCOUNT ON;",
            "DECLARE @path nvarchar(400) = N'C:\\PowerBI\\gigafactory\\sql_server\\csv\\';",
            "DECLARE @sql nvarchar(max);", "DECLARE @n int;", ""]
    for full, src, pk, fks in SQL_TABLES:
        df = tables[src]
        cols = []
        for c in df.columns:
            null = "NOT NULL" if c == pk else "NULL"
            cols.append(f"    [{c}] {_sql_type(df[c])} {null}")
        if pk:
            cols.append(f"    CONSTRAINT PK_{full.replace('.', '_')} PRIMARY KEY ([{pk}])")
        for c, ref in fks.items():
            cols.append(f"    CONSTRAINT FK_{full.replace('.', '_')}_{c} FOREIGN KEY ([{c}]) REFERENCES {ref}")
        ddl.append(f"CREATE TABLE {full} (\n" + ",\n".join(cols) + "\n);\nGO\n")
        fname = full.replace(".", "_") + ".csv"
        df.to_csv(os.path.join(csv_dir, fname), index=False, date_format="%Y-%m-%d")
        load.append(f"SET @sql = N'BULK INSERT {full} FROM ''' + @path + N'{fname}'' "
                    f"WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';")
        load.append("EXEC sp_executesql @sql;")
        load.append(f"SELECT @n = COUNT(*) FROM {full};")
        load.append(f"PRINT '{full}: ' + CAST(@n AS varchar(20)) + ' rows';")
        load.append("")
    open(os.path.join(root, "01_create_database.sql"), "w").write("\n".join(ddl))
    open(os.path.join(root, "02_load_data.sql"), "w").write("\n".join(load))


# --------------------------------------------------------------------------- helpers
def _py(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if v is pd.NaT:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    return v


def _rows(df):
    for rec in df.itertuples(index=False, name=None):
        yield [_py(v) for v in rec]


def _iso(v):
    return None if pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d")


# --------------------------------------------------------------------------- main export
def export_sources(tables, out, rng):
    os.makedirs(out, exist_ok=True)

    # ERP vendor numbers: the ERP does not know the project company IDs
    comp = tables["dim_company"]
    ext = comp[comp.Company_Type.isin(["Supplier", "Subcontractor"])].copy()
    nums = rng.choice(np.arange(100000, 999999), len(ext), replace=False)
    xref = pd.DataFrame({"Company_ID": ext.Company_ID.values, "ERP_Vendor_No": [f"V{n}" for n in nums]})
    tables = dict(tables, company_erp_xref=xref)

    # 1. SQL Server engineering database
    _write_sql_server(tables, os.path.join(out, "sql_server"))

    # 2. ERP purchase order export
    vmap = dict(zip(xref.Company_ID, xref.ERP_Vendor_No))
    p = tables["fact_procurement"].copy()
    erp = pd.DataFrame({
        "PO Number": p.PO_No, "Item": p.PO_Line * 10, "Vendor No": p.Supplier_ID.map(vmap),
        "Material Description": p.Commodity, "UoM": p.Unit, "Order Qty": p.Quantity,
        "Net Value (USD)": p.Line_Value_USD, "Long Lead Flag": np.where(p.Long_Lead == "Y", "X", ""),
        "PO Date": p.Order_Date, "Required On Site": p.ROS_Date, "Vendor Promised Date": p.Promised_Date,
        "Expedited Forecast Date": p.Forecast_Delivery_Date, "Goods Receipt Date": p.Actual_Delivery_Date,
        "Discipline": p.Discipline_Code, "Delivery Area": p.Area_Code, "Expediting Status": p.Expediting_Status})
    d = os.path.join(out, "erp_procurement"); os.makedirs(d, exist_ok=True)
    wb = StreamingWorkbook(os.path.join(d, "PO_Lines_Export.xlsx"))
    wb.write_sheet("PO_Lines", list(erp.columns), _rows(erp))
    wb.close()

    # 3. P6 schedule / progress export
    c = tables["dim_cwp"]
    cwp = pd.DataFrame({"CWP ID": c.CWP_ID, "Area": c.Area_Code, "Discipline": c.Discipline_Code,
                        "Contractor": c.Company_ID, "Planned Start": c.Planned_Start,
                        "Planned Finish": c.Planned_Finish, "Budgeted Units (hrs)": c.Budget_Hours})
    w_ = tables["fact_progress_weekly"]
    wk = pd.DataFrame({"Week Ending": w_.Week_Ending, "CWP ID": w_.CWP_ID, "Planned Units": w_.Planned_Hours,
                       "Earned Units": w_.Earned_Hours, "Actual Units": w_.Actual_Hours})
    d = os.path.join(out, "p6_schedule"); os.makedirs(d, exist_ok=True)
    wb = StreamingWorkbook(os.path.join(d, "Construction_Progress_Export.xlsx"))
    wb.write_sheet("WBS_CWP", list(cwp.columns), _rows(cwp))
    wb.write_sheet("Weekly_Progress", list(wk.columns), _rows(wk))
    wb.close()

    # 4. Completions system: monthly ITR extract files + registers
    d = os.path.join(out, "completions_system"); os.makedirs(os.path.join(d, "itr_extracts"), exist_ok=True)
    itr = tables["fact_itr"]
    month = pd.to_datetime(itr.Planned_Date).dt.strftime("%Y_%m")
    for m, part in itr.groupby(month):
        part.to_csv(os.path.join(d, "itr_extracts", f"itr_extract_{m}.csv"), index=False, date_format="%Y-%m-%d")
    tables["fact_tag_register"].to_csv(os.path.join(d, "tag_register.csv"), index=False, date_format="%Y-%m-%d")
    tables["fact_punch"].to_csv(os.path.join(d, "punch_list.csv"), index=False, date_format="%Y-%m-%d")

    # 5. HSE REST API response (OData style, camelCase fields)
    s = tables["fact_safety_weekly"]
    payload = {"@odata.context": "https://hse.example.com/api/v1/$metadata#WeeklyStatistics",
               "value": [{"weekEnding": _iso(r.Week_Ending), "companyId": r.Company_ID, "headcount": int(r.Headcount),
                          "manhours": int(r.Manhours), "recordables": int(r.Recordables),
                          "lostTimeInjuries": int(r.Lost_Time_Injuries), "firstAids": int(r.First_Aids),
                          "nearMisses": int(r.Near_Misses), "observations": int(r.Safety_Observations)}
                         for r in s.itertuples()]}
    d = os.path.join(out, "hse_api"); os.makedirs(d, exist_ok=True)
    json.dump(payload, open(os.path.join(d, "safety_weekly.json"), "w"), indent=1)

    # 6. SharePoint list export (NCR log)
    n = tables["fact_ncr"]
    created = pd.to_datetime(n.Date_Opened)
    sp = pd.DataFrame({"ID": np.arange(1, len(n) + 1), "Title": n.NCR_No, "Discipline": n.Discipline_Code,
                       "Contractor": n.Company_ID, "Area": n.Area_Code, "Category": n.Category, "Severity": n.Severity,
                       "Root Cause": n.Root_Cause, "Date Opened": n.Date_Opened, "Date Closed": n.Date_Closed,
                       "NCR Status": n.Status, "Created": created + pd.to_timedelta(rng.integers(8, 17, len(n)), unit="h"),
                       "Modified": created + pd.to_timedelta(rng.integers(1, 40, len(n)), unit="D"),
                       "Item Type": "Item", "Path": "sites/ProjectSaguaro/Lists/NCR Log"})
    sp["Modified"] = sp["Modified"].where(sp["Modified"] <= pd.Timestamp("2026-11-13"), pd.Timestamp("2026-11-13 16:00"))
    d = os.path.join(out, "quality_sharepoint"); os.makedirs(d, exist_ok=True)
    wb = StreamingWorkbook(os.path.join(d, "NCR_Log.xlsx"))
    wb.write_sheet("NCR Log", list(sp.columns), _rows(sp))
    wb.close()
    print("sources written to", out)
