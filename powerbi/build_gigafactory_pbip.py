"""Build the Power BI Project (.pbip) for the gigafactory project analytics report.

Semantic model (model.bim) is assembled from the files in 04-gigafactory-project-analytics:
  powerquery/all_sources.m        queries (SQL Server, ERP, P6, completions, HSE API, SharePoint)
  dax/01_tables_and_columns.dax   calculated tables and columns
  dax/02_measures.dax             measures
  model/relationships.md          relationships
Column lists and types are read from the generated source files so the model matches the queries.
The report (PBIR) has four pages: Overview, Engineering & RFIs, Procurement & Construction, Completions & Safety.

Usage: python build_gigafactory_pbip.py --root .. --theme <path to a Power BI base theme json>
"""
import argparse
import glob
import json
import os
import re
import shutil
import uuid

import pandas as pd

NAME = "Gigafactory Project Analytics"
VIS_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json"
GREEN_BORDER, GREEN_FILL = "#D6ECD6", "#F6FBF6"


def lt():
    return str(uuid.uuid4())


def split(text):
    return text.replace("\r\n", "\n").strip("\n").split("\n")


# ------------------------------------------------------------------ queries
def parse_queries(m_text):
    blocks = re.split(r"^// ----- ", m_text, flags=re.M)[1:]
    out = {}
    for b in blocks:
        name = b.split()[0]
        body = b[b.index("\nlet") + 1:]
        body = body[:re.search(r"^in\n\s+\w+\s*$", body, re.M).end()]
        out[name] = body.strip("\n")
    return out


def dtype_of(name, series):
    if "Date" in name or name in ("Week_Ending", "Planned_Start", "Planned_Finish"):
        return "dateTime"
    if pd.api.types.is_integer_dtype(series):
        return "int64"
    if pd.api.types.is_float_dtype(series):
        return "double"
    return "string"


def table_columns(root):
    G = os.path.join(root, "data", "gigafactory")
    csv = lambda *p: pd.read_csv(os.path.join(G, *p), nrows=2000, low_memory=False)
    cols = {}
    sql = {"dim_discipline": "ref_discipline", "dim_area": "ref_area", "dim_company": "ref_company",
           "dim_system": "ref_facility_system", "dim_subsystem": "ref_subsystem", "project_info": "ref_project_info",
           "fact_document_revisions": "eng_document_revision", "fact_rfi": "eng_rfi", "fact_submittals": "eng_submittal"}
    for t, f in sql.items():
        df = csv("sql_server", "csv", f + ".csv")
        cols[t] = [(c, dtype_of(c, df[c])) for c in df.columns]
    for t in ("fact_document_revisions",):
        cols[t] = [(c, "string" if c in ("Revision", "Review_Code") else d) for c, d in cols[t]]
    cols["fact_procurement"] = [("PO_No", "string"), ("Vendor_No", "string"), ("Commodity", "string"), ("Unit", "string"),
                                ("Quantity", "int64"), ("Line_Value_USD", "double"), ("Order_Date", "dateTime"),
                                ("ROS_Date", "dateTime"), ("Promised_Date", "dateTime"), ("Forecast_Delivery_Date", "dateTime"),
                                ("Actual_Delivery_Date", "dateTime"), ("Discipline_Code", "string"), ("Area_Code", "string"),
                                ("Expediting_Status", "string"), ("PO_Line", "int64"), ("Long_Lead", "string"), ("Supplier_ID", "string")]
    cols["dim_cwp"] = [("CWP_ID", "string"), ("Area_Code", "string"), ("Discipline_Code", "string"), ("Company_ID", "string"),
                       ("Planned_Start", "dateTime"), ("Planned_Finish", "dateTime"), ("Budget_Hours", "double")]
    cols["fact_progress_weekly"] = [("Week_Ending", "dateTime"), ("CWP_ID", "string"), ("Planned_Hours", "double"),
                                    ("Earned_Hours", "double"), ("Actual_Hours", "double")]
    cols["fact_itr"] = [("Source_File", "string")] + [(c, "dateTime" if "Date" in c else "string") for c in
                       ["ITR_ID", "Tag_No", "Subsystem_Code", "Discipline_Code", "ITR_Form", "ITR_Description", "Stage",
                        "Planned_Date", "Completed_Date", "Status", "Completed_By_ID"]]
    for t, f in (("fact_tag_register", "tag_register.csv"), ("fact_punch", "punch_list.csv")):
        df = csv("completions_system", f)
        cols[t] = [(c, "dateTime" if "Date" in c else "string") for c in df.columns]
    cols["fact_safety_weekly"] = [("Week_Ending", "dateTime"), ("Company_ID", "string")] + \
        [(c, "int64") for c in ["Headcount", "Manhours", "Recordables", "Lost_Time_Injuries", "First_Aids", "Near_Misses", "Safety_Observations"]]
    cols["fact_ncr"] = [("NCR_No", "string"), ("Discipline_Code", "string"), ("Company_ID", "string"), ("Area_Code", "string"),
                        ("Category", "string"), ("Severity", "string"), ("Root_Cause", "string"), ("Date_Opened", "dateTime"),
                        ("Date_Closed", "dateTime"), ("Status", "string")]
    return cols


def parse_measures(text):
    body = text.split("DEFINE", 1)[1].split("\nEVALUATE", 1)[0]
    parts = re.split(r"^\s*MEASURE '_Measures'\[([^\]]+)\] =", body, flags=re.M)
    out = []
    for i in range(1, len(parts), 2):
        expr = "\n".join(l[8:] if l.startswith("        ") else l.strip() for l in parts[i + 1].strip("\n").split("\n"))
        expr = re.sub(r"\n\s*// -{5,}.*$", "", expr.rstrip(), flags=re.S).rstrip()
        out.append((parts[i], expr.strip()))
    return out


def parse_calc_columns(text):
    out = []
    for m in re.finditer(r"^(\w+)\[([^\]]+)\] =\n(.*?)(?=\n\n|\Z)", text, re.M | re.S):
        out.append((m.group(1), m.group(2), m.group(3).strip()))
    return out


def parse_relationships(md):
    rels = []
    for row in re.findall(r"^\| (\w+)\[(\w+)\] \| (\w+)\[(\w+)\] \| (\S+?) ", md, re.M):
        rels.append({"name": lt(), "fromTable": row[0], "fromColumn": row[1], "toTable": row[2], "toColumn": row[3],
                     **({"isActive": False} if "No" in row[4] else {})})
    return rels


def build_model(root):
    P = os.path.join(root, "04-gigafactory-project-analytics")
    queries = parse_queries(open(os.path.join(P, "powerquery", "all_sources.m")).read())
    cols = table_columns(root)
    tables = []
    for name, m in queries.items():
        if name == "company_erp_xref":
            continue
        tables.append({"name": name, "lineageTag": lt(),
                       "columns": [{"name": c, "dataType": d, "sourceColumn": c, "lineageTag": lt(),
                                    "summarizeBy": "none" if d != "double" else "sum",
                                    **({"formatString": "Short Date"} if d == "dateTime" else {}),
                                    "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}]} for c, d in cols[name]],
                       "partitions": [{"name": name, "mode": "import", "source": {"type": "m", "expression": split(m)}}],
                       "annotations": [{"name": "PBI_ResultType", "value": "Table"}]})
    T = {t["name"]: t for t in tables}
    calc_cols = parse_calc_columns(open(os.path.join(P, "dax", "01_tables_and_columns.dax")).read())
    ctypes = {"Age Days": "int64", "Age Bucket": "string", "Age Bucket Sort": "int64", "Float Days": "int64",
              "Review Cycle Days": "int64", "MC Status": "string"}
    for t, c, expr in calc_cols:
        if t in T:
            T[t]["columns"].append({"type": "calculated", "name": c, "dataType": ctypes.get(c, "string"), "isDataTypeInferred": True,
                                    "expression": split(expr), "lineageTag": lt(), "summarizeBy": "none"})
    for c in T["fact_rfi"]["columns"]:
        if c["name"] == "Age Bucket":
            c["sortByColumn"] = "Age Bucket Sort"
    dd = re.search(r"^dim_date =\n(.*?)\n// Sort", open(os.path.join(P, "dax", "01_tables_and_columns.dax")).read(), re.S | re.M).group(1)
    dcols = [("Date", "dateTime"), ("Year", "int64"), ("Month No", "int64"), ("Month", "string"), ("Year Month", "string"),
             ("Quarter", "string"), ("Year Quarter", "string"), ("Week Ending", "dateTime")]
    date_tbl = {"name": "dim_date", "dataCategory": "Time", "lineageTag": lt(),
                "columns": [{"type": "calculatedTableColumn", "name": c, "dataType": d, "isNameInferred": True,
                             "isDataTypeInferred": True, "sourceColumn": f"[{c}]", "lineageTag": lt(), "summarizeBy": "none",
                             **({"isKey": True} if c == "Date" else {}),
                             **({"formatString": "Short Date"} if d == "dateTime" else {}),
                             **({"sortByColumn": "Month No"} if c == "Month" else {})} for c, d in dcols],
                "partitions": [{"name": "dim_date", "mode": "import", "source": {"type": "calculated", "expression": split(dd)}}]}
    fmt = lambda n: ("0.0%;-0.0%;0.0%" if "%" in n else "0.00" if n in ("SPI", "CPI", "Productivity Factor", "Weekly SPI", "TRIR",
                     "LTIR", "TRIR Rolling 12M", "Near Miss Ratio") else "\\$#,0;(\\$#,0);\\$#,0" if n in ("PO Value", "Value at Risk")
                     else "dd mmm yyyy" if "Date" in n else "#,0.0" if "days" in n.lower() or "Avg" in n else "#,0")
    meas = [{"name": n, "expression": split(e), "lineageTag": lt(), "formatString": fmt(n)}
            for n, e in parse_measures(open(os.path.join(P, "dax", "02_measures.dax")).read())]
    for mm in meas:
        if mm["name"].endswith("Color"):
            mm.pop("formatString")
    mtab = {"name": "_Measures", "lineageTag": lt(),
            "columns": [{"type": "calculatedTableColumn", "name": "x", "dataType": "string", "isHidden": True, "isNameInferred": True,
                         "isDataTypeInferred": True, "sourceColumn": "[x]", "lineageTag": lt(), "summarizeBy": "none"}],
            "partitions": [{"name": "_Measures", "mode": "import", "source": {"type": "calculated", "expression": ['ROW ( "x", BLANK () )']}}],
            "measures": meas}
    tables = [mtab, date_tbl] + tables
    rels = parse_relationships(open(os.path.join(P, "model", "relationships.md")).read())
    for r in rels:
        if r["toTable"] == "dim_date":
            r["joinOnDateBehavior"] = "datePartOnly"
    params = [("SqlServer", "localhost\\\\SQLEXPRESS"), ("SqlDatabase", "SaguaroPIIM"), ("SourceRoot", "C:\\\\PowerBI\\\\gigafactory\\\\")]
    exprs = [{"name": n, "kind": "m", "lineageTag": lt(),
              "expression": [f'"{v}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'.replace("\\\\", "\\")],
              "annotations": [{"name": "PBI_ResultType", "value": "Text"}]} for n, v in params]
    exprs.append({"name": "company_erp_xref", "kind": "m", "lineageTag": lt(), "expression": split(queries["company_erp_xref"]),
                  "annotations": [{"name": "PBI_ResultType", "value": "Table"}]})
    model = {"culture": "en-US", "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
             "defaultPowerBIDataSourceVersion": "powerBI_V3", "sourceQueryCulture": "en-US",
             "tables": tables, "relationships": rels, "expressions": exprs,
             "annotations": [{"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
                             {"name": "PBI_QueryOrder", "value": json.dumps([p[0] for p in params] + list(queries))}]}
    return {"compatibilityLevel": 1567, "model": model}


# ------------------------------------------------------------------ report
def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def field(entity, prop, measure=True):
    k = "Measure" if measure else "Column"
    return {k: {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def proj(entity, prop, measure=True):
    return {"field": field(entity, prop, measure), "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}


def container(title, card=False):
    o = {"title": [{"properties": {"show": lit("true"), "text": lit(f"'{title}'"), "fontSize": lit("11D")}}],
         "border": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit(f"'{GREEN_BORDER}'")}},
                                    "radius": lit("8D")}}],
         "background": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit(f"'{GREEN_FILL if card else '#FFFFFF'}'")}},
                                        "transparency": lit("0D")}}]}
    return o


def visual(vtype, x, y, w, h, title, query, z, card=False, objects=None):
    v = {"$schema": VIS_SCHEMA, "name": uuid.uuid4().hex[:20],
         "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
         "visual": {"visualType": vtype, "query": {"queryState": query}, "visualContainerObjects": container(title, card),
                    "drillFilterOtherVisuals": True}}
    if objects:
        v["visual"]["objects"] = objects
    return v


def card(measure, x, y, w=150, h=92, z=0):
    return visual("card", x, y, w, h, measure, {"Values": {"projections": [proj("_Measures", measure)]}}, z, card=True,
                  objects={"labels": [{"properties": {"fontSize": lit("20D")}}], "categoryLabels": [{"properties": {"show": lit("false")}}]})


def line(title, measures, x, y, w, h, z, axis=("dim_date", "Week Ending")):
    return visual("lineChart", x, y, w, h, title,
                  {"Category": {"projections": [dict(proj(axis[0], axis[1], False), active=True)]},
                   "Y": {"projections": [proj("_Measures", m) for m in measures]}}, z)


def bar(title, cat, measures, x, y, w, h, z):
    return visual("clusteredBarChart", x, y, w, h, title,
                  {"Category": {"projections": [dict(proj(cat[0], cat[1], False), active=True)]},
                   "Y": {"projections": [proj("_Measures", m) for m in measures]}}, z)


def table(title, cols, measures, x, y, w, h, z):
    return visual("tableEx", x, y, w, h, title,
                  {"Values": {"projections": [proj(t, c, False) for t, c in cols] + [proj("_Measures", m) for m in measures]}}, z)


def slicer(title, col, x, y, w, h, z):
    return visual("slicer", x, y, w, h, title, {"Values": {"projections": [dict(proj(col[0], col[1], False), active=True)]}}, z,
                  objects={"data": [{"properties": {"mode": lit("'Dropdown'")}}]})


def pages():
    kpis = ["IFC % Complete", "Earned % to Date", "SPI", "CPI", "RFIs Open", "Lines Late vs ROS", "Punch A Open", "TRIR"]
    p1 = [card(m, 16 + i * 157, 16, z=i) for i, m in enumerate(kpis)]
    p1 += [line("Construction progress: planned vs earned (% of budget hours)", ["Planned % (Cum)", "Earned % (Cum)"], 16, 124, 616, 280, 20),
           line("IFC documents: planned vs actual (cumulative)", ["IFC Planned (Cum)", "IFC Actual (Cum)"], 648, 124, 616, 280, 21),
           line("Mechanical completion: planned vs actual subsystems", ["MC Planned (Cum)", "MC Actual (Cum)"], 16, 420, 616, 284, 22),
           bar("Open RFIs and overdue by discipline", ("dim_discipline", "Discipline"), ["RFIs Open", "RFIs Overdue"], 648, 420, 616, 284, 23)]
    p2 = [slicer("Discipline", ("dim_discipline", "Discipline"), 16, 16, 240, 70, 0),
          card("Reviews Overdue", 272, 16, 180, 70, 1), card("Avg Review Cycle (days)", 468, 16, 180, 70, 2),
          card("First-Pass Approval %", 664, 16, 180, 70, 3), card("Avg Days to Respond", 860, 16, 180, 70, 4),
          card("Submittals Overdue", 1056, 16, 208, 70, 5),
          line("Open RFIs over time", ["RFIs Open", "RFIs Overdue"], 16, 102, 616, 290, 10),
          table("IFC progress by discipline", [("dim_discipline", "Discipline")], ["IFC % Planned", "IFC % Complete", "IFC Behind Plan"], 648, 102, 616, 290, 11),
          bar("Open RFIs by area (building)", ("dim_area", "Building"), ["RFIs Open"], 16, 408, 616, 296, 12),
          table("Submittal performance by company", [("dim_company", "Company_Name")], ["Submittals", "Submittal Approval %", "Avg Submittal Turnaround (days)"], 648, 408, 616, 296, 13)]
    p3 = [card("PO Value", 16, 16, 200, 80, 0), card("Value at Risk", 232, 16, 200, 80, 1), card("Delivered On Time %", 448, 16, 200, 80, 2),
          card("Long-Lead Items Open", 664, 16, 200, 80, 3), card("Productivity Factor", 880, 16, 180, 80, 4), card("Weekly SPI", 1076, 16, 188, 80, 5),
          bar("Value at risk by supplier", ("dim_company", "Company_Name"), ["Value at Risk"], 16, 112, 616, 292, 10),
          bar("PO lines late vs required-on-site by discipline", ("dim_discipline", "Discipline"), ["Lines Late vs ROS"], 648, 112, 616, 292, 11),
          line("Earned value: planned, earned and actual hours (cumulative)", ["Planned Hours (Cum)", "Earned Hours (Cum)", "Actual Hours (Cum)"], 16, 420, 616, 284, 12),
          table("SPI and CPI by building", [("dim_area", "Building")], ["SPI", "CPI", "Earned % to Date"], 648, 420, 616, 284, 13)]
    p4 = [card("ITR A %", 16, 16, 200, 80, 0), card("ITR B %", 232, 16, 200, 80, 1), card("Punch A Open", 448, 16, 200, 80, 2),
          card("Subsystems Behind MC Plan", 664, 16, 200, 80, 3), card("TRIR", 880, 16, 180, 80, 4), card("NCRs Open", 1076, 16, 188, 80, 5),
          line("Turnover: planned vs actual subsystems", ["Turnover Planned (Cum)", "Turnover Actual (Cum)"], 16, 112, 616, 292, 10),
          bar("ITR A completion by system", ("dim_system", "System_Name"), ["ITR A %"], 648, 112, 616, 292, 11),
          line("Open punch (all) and category A", ["Punch Open", "Punch A Open"], 16, 420, 616, 284, 12),
          line("TRIR, rolling 12 months", ["TRIR Rolling 12M"], 648, 420, 616, 284, 13, axis=("dim_date", "Year Month"))]
    return [("Overview", p1), ("Engineering & RFIs", p2), ("Procurement & Construction", p3), ("Completions & Safety", p4)]


def write_report(out, theme):
    rep = os.path.join(out, f"{NAME}.Report")
    d = os.path.join(rep, "definition")
    os.makedirs(os.path.join(d, "pages"), exist_ok=True)
    theme_name = os.path.splitext(os.path.basename(theme))[0]
    os.makedirs(os.path.join(rep, "StaticResources", "SharedResources", "BaseThemes"), exist_ok=True)
    shutil.copy(theme, os.path.join(rep, "StaticResources", "SharedResources", "BaseThemes", os.path.basename(theme)))
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
               "version": "2.0.0"}, open(os.path.join(d, "version.json"), "w"), indent=2)
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
               "themeCollection": {"baseTheme": {"name": theme_name, "reportVersionAtImport": {"visual": "2.12.0", "report": "3.4.0", "page": "2.3.1"},
                                                 "type": "SharedResources"}},
               "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                                     "items": [{"name": theme_name, "path": f"BaseThemes/{os.path.basename(theme)}", "type": "BaseTheme"}]}],
               "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                            "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                            "useDefaultAggregateDisplayName": True}},
              open(os.path.join(d, "report.json"), "w"), indent=2)
    order = []
    for title, visuals in pages():
        pid = uuid.uuid4().hex[:20]
        order.append(pid)
        pdir = os.path.join(d, "pages", pid)
        os.makedirs(pdir, exist_ok=True)
        json.dump({"$schema": PAGE_SCHEMA, "name": pid, "displayName": title, "displayOption": "FitToPage", "height": 720, "width": 1280,
                   "objects": {"background": [{"properties": {"color": {"solid": {"color": lit("'#FFFFFF'")}}, "transparency": lit("0D")}}],
                               "outspace": [{"properties": {"color": {"solid": {"color": lit("'#F3FAF3'")}}}}]}},
                  open(os.path.join(pdir, "page.json"), "w"), indent=2)
        for v in visuals:
            vd = os.path.join(pdir, "visuals", v["name"])
            os.makedirs(vd, exist_ok=True)
            json.dump(v, open(os.path.join(vd, "visual.json"), "w"), indent=2)
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
               "pageOrder": order, "activePageName": order[0]}, open(os.path.join(d, "pages", "pages.json"), "w"), indent=2)
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
               "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}},
              open(os.path.join(rep, "definition.pbir"), "w"), indent=2)
    return rep


def main():
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from pbix_to_pbip import platform, write_project
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="..")
    ap.add_argument("--theme", required=True)
    a = ap.parse_args()
    out = os.path.join(a.root, "04-gigafactory-project-analytics", "pbip")
    shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out)
    write_project(build_model(a.root), NAME, out)
    rep = write_report(out, a.theme)
    platform(rep, "Report", NAME)
    print("wrote", out)


if __name__ == "__main__":
    main()
