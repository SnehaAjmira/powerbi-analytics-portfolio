"""Convert a .pbix into a Power BI Project (.pbip) that points at a new data folder.

What it does
  * Reads the model metadata (tables, columns, measures, calculated tables and
    columns, calculation groups, hierarchies, auto date tables, relationships,
    Power Query) from the .pbix and writes it as a TMSL model.bim.
  * Replaces every hard-coded File.Contents("C:\\...\\file.xlsx") path with a
    DataFolder parameter, so the model refreshes from any folder.
  * Copies the report definition (PBIR pages and visuals) and removes saved
    filter selections and expanded tree nodes that hold values not found in the
    new data (for example names from the original data source).
  * Optionally replaces report background images.

No data rows are copied: the model is empty until it is refreshed in Power BI Desktop.

Usage
  python pbix_to_pbip.py report.pbix OutName out_dir --data-folder "C:\\PowerBI\\lab\\"
         [--allowed-values values.txt] [--backdrop new.png]
"""
import argparse
import json
import os
import re
import shutil
import uuid
import zipfile

from pbixray import PBIXRay

DT = {1: "automatic", 2: "string", 6: "int64", 8: "double", 9: "dateTime", 10: "decimal", 11: "boolean",
      17: "binary", 19: "unknown", 20: "variant"}
SUMM = {1: "default", 2: "none", 3: "sum", 4: "min", 5: "max", 6: "count", 7: "average", 8: "distinctCount"}
CARD = {1: "one", 2: "many"}
XF = {1: "oneDirection", 2: "bothDirections", 3: "automatic"}
SEC = {1: "oneDirection", 2: "bothDirections", 3: "none"}
ANNOT_TYPES = {1: "model", 3: "table", 4: "column", 8: "measure", 9: "hierarchy"}
PATH_RE = re.compile(r'File\.Contents\("[^"]*\\([^"\\]+)"\)')


def q(conn, sql):
    table = re.search(r"from \[?(\w+)\]?", sql).group(1)
    cols = [r[1] for r in conn.execute(f"pragma table_info([{table}])")]
    return [dict(zip(cols, r)) for r in conn.execute(sql)]


def lines(text):
    return (text or "").replace("\r\n", "\n").split("\n")


def build_model(p, data_folder):
    conn = p._metadata._meta._db.conn
    model = q(conn, "select * from Model")[0]
    tables = q(conn, "select * from [Table]")
    keep = {t["ID"]: t for t in tables if t["SystemFlags"] in (0, 2) and not re.match(r"^[HRU]\$", t["Name"])}
    cols = q(conn, "select * from [Column]")
    colname = {c["ID"]: (c["ExplicitName"] or c["InferredName"]) for c in cols}
    col_table = {c["ID"]: c["TableID"] for c in cols}
    ann = q(conn, "select * from Annotation")
    ann_by = {}
    for a in ann:
        ann_by.setdefault((a["ObjectType"], a["ObjectID"]), []).append({"name": a["Name"], "value": a["Value"]})
    fsd = {f["ID"]: f["Expression"] for f in q(conn, "select * from FormatStringDefinition")}
    rels = q(conn, "select * from Relationship")
    rel_name = {r["ID"]: r["Name"] for r in rels}
    hier = q(conn, "select * from Hierarchy")
    hier_by_id = {h["ID"]: h for h in hier}
    levels = q(conn, "select * from Level")
    variations = q(conn, "select * from Variation")
    parts = q(conn, "select * from Partition")
    measures = q(conn, "select * from Measure")
    cgroups = {g["TableID"]: g for g in q(conn, "select * from CalculationGroup")}
    citems = q(conn, "select * from CalculationItem")

    def fix_m(text):
        return PATH_RE.sub(r'File.Contents(DataFolder & "\1")', text or "")

    out_tables = []
    for tid, t in keep.items():
        tj = {"name": t["Name"], "lineageTag": t["LineageTag"]}
        if t["IsHidden"]: tj["isHidden"] = True
        if t["IsPrivate"]: tj["isPrivate"] = True
        if t["ShowAsVariationsOnly"]: tj["showAsVariationsOnly"] = True
        if t["DataCategory"]: tj["dataCategory"] = t["DataCategory"]
        if tid in cgroups:
            g = cgroups[tid]
            items = [{"name": i["Name"], "expression": lines(i["Expression"]), "ordinal": i["Ordinal"]}
                     for i in citems if i["CalculationGroupID"] == g["ID"]]
            for i, src in zip(items, [i for i in citems if i["CalculationGroupID"] == g["ID"]]):
                if src["FormatStringDefinitionID"]:
                    i["formatStringDefinition"] = {"expression": lines(fsd[src["FormatStringDefinitionID"]])}
            tj["calculationGroup"] = {"precedence": g["Precedence"], "calculationItems": items}
        cj = []
        for c in [c for c in cols if c["TableID"] == tid and c["Type"] != 3]:
            name = colname[c["ID"]]
            dt = c["ExplicitDataType"] if c["ExplicitDataType"] not in (None, 1, 19) else c["InferredDataType"]
            j = {"name": name, "dataType": DT.get(dt, "string")}
            if c["Type"] == 2:
                j["type"] = "calculated"
                j["isDataTypeInferred"] = c["ExplicitDataType"] in (None, 1, 19)
                j["expression"] = lines(c["Expression"])
            elif c["Type"] == 4:
                j["type"] = "calculatedTableColumn"
                j["isNameInferred"] = c["ExplicitName"] is None
                j["isDataTypeInferred"] = c["ExplicitDataType"] in (None, 1, 19)
                j["sourceColumn"] = c["SourceColumn"]
            else:
                j["sourceColumn"] = c["SourceColumn"]
            for k_src, k_dst in (("FormatString", "formatString"), ("DisplayFolder", "displayFolder"),
                                 ("DataCategory", "dataCategory"), ("LineageTag", "lineageTag")):
                if c.get(k_src):
                    j[k_dst] = c[k_src]
            if c["IsHidden"]: j["isHidden"] = True
            if c["IsKey"]: j["isKey"] = True
            if c["SummarizeBy"] in SUMM: j["summarizeBy"] = SUMM[c["SummarizeBy"]]
            if c["SortByColumnID"]: j["sortByColumn"] = colname[c["SortByColumnID"]]
            if c["IsAvailableInMDX"] == 0: j["isAvailableInMdx"] = False
            vs = [v for v in variations if v["ColumnID"] == c["ID"]]
            if vs:
                j["variations"] = []
                for v in vs:
                    h = hier_by_id[v["DefaultHierarchyID"]]
                    j["variations"].append({"name": v["Name"], "relationship": rel_name[v["RelationshipID"]],
                                            "defaultHierarchy": {"table": keep[h["TableID"]]["Name"], "hierarchy": h["Name"]},
                                            "isDefault": bool(v["IsDefault"])})
            a = ann_by.get((4, c["ID"]))
            if a: j["annotations"] = a
            cj.append(j)
        tj["columns"] = cj
        mj = []
        for m in [m for m in measures if m["TableID"] == tid]:
            j = {"name": m["Name"], "expression": lines(m["Expression"])}
            for k_src, k_dst in (("FormatString", "formatString"), ("DisplayFolder", "displayFolder"),
                                 ("Description", "description"), ("LineageTag", "lineageTag"), ("DataCategory", "dataCategory")):
                if m.get(k_src):
                    j[k_dst] = m[k_src]
            if m["IsHidden"]: j["isHidden"] = True
            if m["FormatStringDefinitionID"]:
                j["formatStringDefinition"] = {"expression": lines(fsd[m["FormatStringDefinitionID"]])}
            a = ann_by.get((8, m["ID"]))
            if a: j["annotations"] = a
            mj.append(j)
        if mj: tj["measures"] = mj
        hj = []
        for h in [h for h in hier if h["TableID"] == tid]:
            lv = sorted([l for l in levels if l["HierarchyID"] == h["ID"]], key=lambda l: l["Ordinal"])
            j = {"name": h["Name"], "lineageTag": h["LineageTag"],
                 "levels": [{"name": l["Name"], "ordinal": l["Ordinal"], "column": colname[l["ColumnID"]],
                             "lineageTag": l["LineageTag"]} for l in lv]}
            if h["IsHidden"]: j["isHidden"] = True
            a = ann_by.get((9, h["ID"]))
            if a: j["annotations"] = a
            hj.append(j)
        if hj: tj["hierarchies"] = hj
        pj = []
        for pt in [x for x in parts if x["TableID"] == tid]:
            if pt["Type"] == 4:
                src = {"type": "m", "expression": lines(fix_m(pt["QueryDefinition"]))}
            elif pt["Type"] == 2:
                src = {"type": "calculated", "expression": lines(pt["QueryDefinition"])}
            elif pt["Type"] == 7:
                src = {"type": "calculationGroup"}
            else:
                continue
            j = {"name": pt["Name"], "source": src}
            if pt["Type"] != 7:
                j["mode"] = "import"
            pj.append(j)
        tj["partitions"] = pj
        a = ann_by.get((3, tid))
        if a: tj["annotations"] = a
        out_tables.append(tj)

    out_rels = []
    for r in rels:
        if r["FromTableID"] not in keep or r["ToTableID"] not in keep:
            continue
        j = {"name": r["Name"], "fromTable": keep[r["FromTableID"]]["Name"], "fromColumn": colname[r["FromColumnID"]],
             "toTable": keep[r["ToTableID"]]["Name"], "toColumn": colname[r["ToColumnID"]]}
        if not r["IsActive"]: j["isActive"] = False
        if r["CrossFilteringBehavior"] != 1: j["crossFilteringBehavior"] = XF[r["CrossFilteringBehavior"]]
        if r["JoinOnDateBehavior"] == 2: j["joinOnDateBehavior"] = "datePartOnly"
        if r["FromCardinality"] != 2: j["fromCardinality"] = CARD[r["FromCardinality"]]
        if r["ToCardinality"] != 1: j["toCardinality"] = CARD[r["ToCardinality"]]
        if r["SecurityFilteringBehavior"] == 2: j["securityFilteringBehavior"] = "bothDirections"
        if r["RelyOnReferentialIntegrity"]: j["relyOnReferentialIntegrity"] = True
        out_rels.append(j)

    exprs = [{"name": "DataFolder", "kind": "m",
              "expression": [f'"{data_folder}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'],
              "lineageTag": str(uuid.uuid4()),
              "annotations": [{"name": "PBI_ResultType", "value": "Text"}, {"name": "PBI_NavigationStepName", "value": "Navigation"}]}]
    for e in q(conn, "select * from Expression"):
        j = {"name": e["Name"], "kind": "m", "expression": lines(fix_m(e["Expression"]))}
        if e["LineageTag"]: j["lineageTag"] = e["LineageTag"]
        exprs.append(j)

    m = {"culture": model["Culture"] or "en-US",
         "dataAccessOptions": json.loads(model["DataAccessOptions"]) if model["DataAccessOptions"] else {"legacyRedirects": True, "returnErrorValuesAsNull": True},
         "defaultPowerBIDataSourceVersion": "powerBI_V3", "sourceQueryCulture": model["SourceQueryCulture"] or "en-US",
         "tables": out_tables, "relationships": out_rels, "expressions": exprs,
         "annotations": [a for a in ann_by.get((1, model["ID"]), []) if a["name"] != "PBI_ProTooling"]}
    if model["DiscourageImplicitMeasures"]:
        m["discourageImplicitMeasures"] = True
    qo = [a for a in m["annotations"] if a["name"] == "PBI_QueryOrder"]
    if qo:
        order = json.loads(qo[0]["value"])
        qo[0]["value"] = json.dumps(["DataFolder"] + [x for x in order if x != "DataFolder"])
    return {"compatibilityLevel": 1567, "model": m}


# --------------------------------------------------------------------------- report scrubbing
def literals(o, out):
    if isinstance(o, dict):
        if "Literal" in o and isinstance(o["Literal"], dict) and isinstance(o["Literal"].get("Value"), str):
            v = o["Literal"]["Value"]
            if len(v) >= 2 and v[0] == "'" and v[-1] == "'":
                out.append(v[1:-1].replace("''", "'"))
        for x in o.values():
            literals(x, out)
    elif isinstance(o, list):
        for x in o:
            literals(x, out)
    return out


def scrub(o, bad):
    def has_bad(x):
        return any(bad(v) for v in literals(x, []))
    if isinstance(o, dict):
        if "filter" in o and has_bad(o["filter"]):
            del o["filter"]
        for k in list(o.keys()):
            o[k] = scrub(o[k], bad)
        return o
    if isinstance(o, list):
        return [scrub(x, bad) for x in o if not has_bad(x)]
    return o


def write_report(pbix, name, out_dir, bad, backdrop):
    rep_dir = os.path.join(out_dir, f"{name}.Report")
    with zipfile.ZipFile(pbix) as z:
        for n in z.namelist():
            if not n.startswith("Report/") or n.endswith("/"):
                continue
            dest = os.path.join(rep_dir, n[len("Report/"):])
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            data = z.read(n)
            if n.endswith(".json"):
                j = json.loads(data.decode("utf-8-sig"))
                j = scrub(j, bad)
                data = json.dumps(j, ensure_ascii=False, indent=2).encode("utf-8")
            elif backdrop and re.search(r"RegisteredResources/Backdrop[^/]*\.png$", n):
                data = open(backdrop, "rb").read()
            open(dest, "wb").write(data)
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
               "version": "4.0", "datasetReference": {"byPath": {"path": f"../{name}.SemanticModel"}}},
              open(os.path.join(rep_dir, "definition.pbir"), "w"), indent=2)
    platform(rep_dir, "Report", name)
    return rep_dir


def platform(d, kind, name):
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
               "metadata": {"type": kind, "displayName": name}, "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}},
              open(os.path.join(d, ".platform"), "w"), indent=2)


def write_project(model_bim, name, out_dir):
    sm = os.path.join(out_dir, f"{name}.SemanticModel")
    os.makedirs(sm, exist_ok=True)
    json.dump(model_bim, open(os.path.join(sm, "model.bim"), "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    json.dump({"version": "1.0", "settings": {}}, open(os.path.join(sm, "definition.pbism"), "w"), indent=2)
    platform(sm, "SemanticModel", name)
    json.dump({"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
               "version": "1.0", "artifacts": [{"report": {"path": f"{name}.Report"}}],
               "settings": {"enableAutoRecovery": True}}, open(os.path.join(out_dir, f"{name}.pbip"), "w"), indent=2)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pbix"); ap.add_argument("name"); ap.add_argument("out_dir")
    ap.add_argument("--data-folder", required=True)
    ap.add_argument("--allowed-values", help="text file, one value per line, that may stay in saved filters")
    ap.add_argument("--blocked-values", help="text file, one value per line, that must be removed")
    ap.add_argument("--backdrop")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    p = PBIXRay(a.pbix)
    write_project(build_model(p, a.data_folder), a.name, a.out_dir)
    allowed = set(open(a.allowed_values, encoding="utf-8").read().split("\n")) if a.allowed_values else None
    blocked = set(open(a.blocked_values, encoding="utf-8").read().split("\n")) if a.blocked_values else set()

    def bad(v):
        lv = v.lower()
        if "eagle" in lv or v in blocked:
            return True
        return allowed is not None and v not in allowed and not re.fullmatch(r"[\d.,\-\s%:/TZ]*", v) and v not in ("true", "false")
    write_report(a.pbix, a.name, a.out_dir, bad, a.backdrop)
    print("wrote", os.path.join(a.out_dir, a.name + ".pbip"))


if __name__ == "__main__":
    main()
