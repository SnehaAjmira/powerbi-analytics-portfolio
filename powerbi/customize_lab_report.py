"""Finish the Lab Test Performance report after conversion.

* Removes visuals that point at tables that are not in the model (unfinished
  field-parameter slicers from the original file).
* Completes the KPI page: fills the three empty cards and removes the
  "work in progress" text box.
* Adds a "KPI Dashboard" overview page (first page) built from copies of the
  report's own working visuals: KPI cards, turnaround status by month, rush
  late tests, completed trend, top analytes and tests by type.

Usage: python customize_lab_report.py "<path to Lab Test Performance.Report>" "<path to model.bim>"
"""
import copy
import glob
import json
import os
import re
import shutil
import sys
import uuid


def load(p):
    return json.load(open(p, encoding="utf-8"))


def save(p, j):
    json.dump(j, open(p, "w", encoding="utf-8"), indent=2, ensure_ascii=False)


def new_id():
    return uuid.uuid4().hex[:20]


def entities(o, out):
    if isinstance(o, dict):
        if "Entity" in o and isinstance(o["Entity"], str):
            out.add(o["Entity"])
        for v in o.values():
            entities(v, out)
    elif isinstance(o, list):
        for v in o:
            entities(v, out)
    return out


def set_measure(card, entity, measure):
    q = card["visual"]["query"]["queryState"]["Values"]["projections"][0]
    q["field"] = {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": measure}}
    q["queryRef"] = f"{entity}.{measure}"
    q["nativeQueryRef"] = measure
    card["visual"]["query"].pop("sortDefinition", None)
    return card


def main(rep, bim):
    model = load(bim)["model"]
    tables = {t["name"] for t in model["tables"]}
    pages_dir = os.path.join(rep, "definition", "pages")

    # 1. drop visuals that reference missing tables
    removed = 0
    for vf in glob.glob(os.path.join(pages_dir, "*", "visuals", "*", "visual.json")):
        missing = entities(load(vf).get("visual", {}), set()) - tables
        if missing:
            shutil.rmtree(os.path.dirname(vf))
            removed += 1
    print("removed visuals referencing missing tables:", removed)

    pages = {}
    for pf in glob.glob(os.path.join(pages_dir, "*", "page.json")):
        pages[load(pf)["displayName"].strip()] = os.path.dirname(pf)
    visuals = {}
    for vf in glob.glob(os.path.join(pages_dir, "*", "visuals", "*", "visual.json")):
        visuals[os.path.basename(os.path.dirname(vf))] = vf

    # 2. complete the KPI page
    kpi = pages["KPI"]
    template_card = load(visuals["ff4f95e95a835a5eaa4a"])          # Avg TAT Days card
    fills = {"87c6bcc28076f09286da": ("Tests Complete", "T.Tests Completed"),
             "988e9c55c24d7c885bf3": ("Tests Release", "T.Tests Released"),
             "ad686de560dcc3a3af3b": ("_OOS Measures", "OOS Client Rate")}
    for vid, (ent, meas) in fills.items():
        vf = visuals.get(vid)
        if not vf:
            continue
        old = load(vf)
        card = copy.deepcopy(template_card)
        card["name"], card["position"] = old["name"], old["position"]
        save(vf, set_measure(card, ent, meas))
    for vf in glob.glob(os.path.join(kpi, "visuals", "*", "visual.json")):
        v = load(vf)
        if v.get("visual", {}).get("visualType") == "textbox" and "PROGRESS" in json.dumps(v):
            shutil.rmtree(os.path.dirname(vf))

    # 3. new overview page
    pid = new_id()
    pdir = os.path.join(pages_dir, pid)
    os.makedirs(os.path.join(pdir, "visuals"))
    page = load(os.path.join(kpi, "page.json"))
    page["name"], page["displayName"] = pid, "KPI Dashboard"
    save(os.path.join(pdir, "page.json"), page)

    def place(src_id, x, y, w, h, z, measure=None):
        v = copy.deepcopy(load(visuals[src_id]))
        v["name"] = new_id()
        v["position"] = {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}
        v.pop("parentGroupName", None)
        if measure:
            set_measure(v, *measure)
        d = os.path.join(pdir, "visuals", v["name"])
        os.makedirs(d)
        save(os.path.join(d, "visual.json"), v)

    place("278ce1a5544065bed34a", 6, 2, 230, 53, 0)                     # date range slicer
    cards = [("Tests Complete", "T.Tests Completed"), ("Tests Release", "T.Tests Released"), ("_Measures", "Avg TAT Days"),
             ("_Measures", "% Rush"), ("_OOS Measures", "OOS Client Rate"), ("_Measures", "Active Clients")]
    for i, m in enumerate(cards):
        place("ff4f95e95a835a5eaa4a", 8 + i * 212, 62, 200, 118, 100 + i, m)
    place("9c3448420134c3572cd0", 8, 190, 830, 262, 200)                # completed early / on time / late by month
    place("24fb12c4039c2de06359", 846, 190, 426, 262, 201)              # late tests: rush vs standard
    place("e5a539701885b4813175", 8, 460, 416, 256, 202)                # tests completed trend
    place("bad866e8c1c5be0c0db4", 432, 460, 416, 256, 203)              # tests completed by analyte
    place("b54c6996ca8d74904332", 856, 460, 416, 256, 204)              # tests completed by test
    pj = os.path.join(pages_dir, "pages.json")
    meta = load(pj)
    meta["pageOrder"] = [pid] + [p for p in meta["pageOrder"] if os.path.isdir(os.path.join(pages_dir, p))]
    meta["activePageName"] = pid
    save(pj, meta)
    print("added page KPI Dashboard", pid)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
