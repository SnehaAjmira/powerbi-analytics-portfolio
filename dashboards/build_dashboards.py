"""Build the interactive HTML dashboards (GitHub Pages) and PNG screenshots.

Reads the synthetic data in data/lab and data/gigafactory, computes the same
measures as the Power BI reports, and writes:

  docs/index.html                         landing page
  docs/<dashboard>.html                   one page per project (Plotly, interactive)
  docs/assets/plotly.min.js               local copy so pages work offline
  <project>/screenshots/dashboard.png     full-page screenshot for each README

Usage:
    pip install pandas numpy plotly openpyxl
    python build_dashboards.py --root ..            (screenshots need Google Chrome)
"""
import argparse
import glob
import json
import os
import pickle
import shutil
import subprocess

import numpy as np
import pandas as pd
import plotly
import plotly.graph_objects as go

# ---------------------------------------------------------------- look and feel
# Data colors: validated categorical palette (fixed order). Page chrome: white with faint green accents.
C = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
GOOD, WARN, SERIOUS, CRIT, NEUTRAL = "#0ca30c", "#fab219", "#ec835a", "#d03b3b", "#a3a29c"
INK, INK2, MUTED, GRID = "#0b0b0b", "#52514e", "#8a8985", "#ececea"
FONT = "Segoe UI, Inter, Arial, sans-serif"

CSS = """
*{box-sizing:border-box} body{margin:0;background:#ffffff;color:#0b0b0b;font-family:Segoe UI,Inter,Arial,sans-serif}
header{background:#f3faf3;border-bottom:1px solid #d6ecd6;padding:22px 32px 18px}
header h1{margin:0;font-size:24px;font-weight:600} header p{margin:6px 0 0;color:#52514e;font-size:14px}
nav{padding:10px 32px;border-bottom:1px solid #e3f1e3;background:#fff;font-size:13px}
nav a{color:#2e6b34;text-decoration:none;margin-right:18px} nav a.on{font-weight:600;border-bottom:2px solid #7cc47f}
main{padding:20px 32px 40px;max-width:1440px;margin:0 auto}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:18px}
.kpi{background:#f6fbf6;border:1px solid #d6ecd6;border-radius:10px;padding:12px 14px}
.kpi .l{font-size:12px;color:#52514e} .kpi .v{font-size:26px;font-weight:600;margin-top:4px}
.kpi .n{font-size:11.5px;color:#6b6a66;margin-top:2px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.card{background:#fff;border:1px solid #e1efe1;border-radius:10px;padding:12px 14px 6px;min-width:0}
.card.wide{grid-column:1/-1} .card h3{margin:0 0 2px;font-size:15px;font-weight:600}
.card .sub{font-size:12px;color:#6b6a66;margin-bottom:4px}
footer{color:#8a8985;font-size:12px;padding:0 32px 28px;max-width:1440px;margin:0 auto}
.dl{display:inline-block;margin-top:10px;margin-right:8px;background:#ffffff;border:1px solid #7cc47f;color:#2e6b34;border-radius:8px;padding:6px 12px;font-size:13px;text-decoration:none;font-weight:600}
.dl:hover{background:#e8f5e8} .badge{display:inline-block;background:#e8f5e8;color:#2e6b34;border-radius:999px;padding:2px 10px;font-size:12px;margin-left:8px;vertical-align:middle}
.links a{display:block;background:#f6fbf6;border:1px solid #d6ecd6;border-radius:10px;padding:16px 18px;margin-bottom:12px;color:#0b0b0b;text-decoration:none}
.links a b{display:block;font-size:17px;margin-bottom:4px} .links a span{color:#52514e;font-size:13px}
"""

DOWNLOADS = {"lab-test-performance.html": ("01-lab-test-performance", "lab-data-v1.0"),
             "pareto-80-20.html": ("02-pareto-80-20-analysis", "lab-data-v1.0"),
             "client-rfm.html": ("03-client-segmentation-rfm", "lab-data-v1.0"),
             "gigafactory-project-analytics.html": ("04-gigafactory-project-analytics", "gigafactory-data-v1.0")}
REPO = "https://github.com/SnehaAjmira/powerbi-analytics-portfolio"
PAGES = [("lab-test-performance.html", "Lab Test Performance"), ("pareto-80-20.html", "Pareto 80/20"),
         ("client-rfm.html", "Client RFM & Early Warning"), ("gigafactory-project-analytics.html", "Gigafactory Project Analytics")]


def style(fig, height=340, legend=True):
    fig.update_layout(
        height=height, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
        font=dict(family=FONT, size=12, color=INK2), hoverlabel=dict(font_family=FONT, bgcolor="#ffffff"),
        showlegend=legend, legend=dict(orientation="h", y=1.08, x=0, traceorder="normal", font=dict(size=12, color=INK2)),
        bargap=0.25)
    fig.update_xaxes(showgrid=False, linecolor="#cfcfcb", tickfont=dict(color=MUTED), title_font=dict(color=MUTED))
    fig.update_yaxes(gridcolor=GRID, zeroline=False, tickfont=dict(color=MUTED), title_font=dict(color=MUTED))
    return fig


def line(fig, x, y, name, color, dash=None, width=2):
    fig.add_trace(go.Scatter(x=x, y=y, name=name, mode="lines", line=dict(color=color, width=width, dash=dash),
                             hovertemplate="%{x|%d %b %Y}<br>" + name + ": %{y:,.0f}<extra></extra>"))


def hbar(labels, values, color=C[0], fmt=",.0f", height=340, text=None, colors=None):
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker=dict(color=colors or color, cornerradius=4),
                           text=text, textposition="outside", cliponaxis=False,
                           hovertemplate="%{y}: %{x:" + fmt + "}<extra></extra>"))
    fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")
    vmax = max([v for v in values if v == v] or [1])
    fig.update_xaxes(range=[0, vmax * (1.28 if text is not None else 1.05)])
    return style(fig, height, legend=False)


def page(fname, title, subtitle, kpis, cards, out_dir, note):
    nav = "".join(f'<a class="{"on" if f == fname else ""}" href="{f}">{t}</a>' for f, t in PAGES)
    proj, tag = DOWNLOADS[fname]
    dl = (f'<a class="dl" href="downloads/{proj}-powerbi.zip">Download Power BI file (.pbip)</a>'
          f'<a class="dl" href="{REPO}/releases/tag/{tag}">Download the data</a>'
          f'<a class="dl" href="{REPO}/tree/main/{proj}">Source and DAX</a>')
    kp = "".join(f'<div class="kpi"><div class="l">{l}</div><div class="v">{v}</div><div class="n">{n}</div></div>'
                 for l, v, n in kpis)
    body = []
    for c in cards:
        t, sub, fig, wide = c
        html = fig.to_html(full_html=False, include_plotlyjs=False, config={"displaylogo": False, "responsive": True, "displayModeBar": False})
        body.append(f'<div class="card{" wide" if wide else ""}"><h3>{t}</h3><div class="sub">{sub}</div>{html}</div>')
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><script src="assets/plotly.min.js"></script><style>{CSS}</style></head><body>
<header><h1>{title}<span class="badge">Synthetic data</span></h1><p>{subtitle}</p>{dl}</header><nav><a href="index.html">Home</a>{nav}</nav>
<main><div class="kpis">{kp}</div><div class="grid">{''.join(body)}</div></main>
<footer>{note} All companies, people and figures are randomly generated. Built by Sai Sneha Ajmira.</footer></body></html>"""
    open(os.path.join(out_dir, fname), "w", encoding="utf-8").write(doc)
    rows = sum(1 if c[3] else 0.5 for c in cards)
    return int(330 + np.ceil(rows) * 430)


# ---------------------------------------------------------------- lab data
def load_lab(root):
    cache = os.path.join(root, "data", "lab", "_cache.pkl")
    if os.path.exists(cache):
        return pickle.load(open(cache, "rb"))
    L = os.path.join(root, "data", "lab")
    d = {"raw": pd.concat([pd.read_excel(os.path.join(L, "Raw DATA.xlsx"), sheet_name=f"RAW_{y}") for y in (2023, 2024, 2025, 2026)],
                          ignore_index=True),
         "cf": pd.read_excel(os.path.join(L, "Raw DATA.xlsx"), sheet_name="Client Facility"),
         "tc": pd.read_excel(os.path.join(L, "Tests Completed.xlsx")),
         "rel": pd.read_excel(os.path.join(L, "Submission Release.xlsx")),
         "qa": pd.read_excel(os.path.join(L, "Quality Data.xlsx")),
         "oos": pd.read_excel(os.path.join(L, "OosByClient.xlsx"))}
    pickle.dump(d, open(cache, "wb"))
    return d


def lab_page(d, out):
    tc, rel, qa, oos = d["tc"].copy(), d["rel"].copy(), d["qa"], d["oos"]
    for df in (tc, rel):
        for c in ("Submission_Date", "Submission_Release_Date", "Estimated_Completion_Date", "Completion_Date"):
            df[c] = pd.to_datetime(df[c])
    tc["TAT_Status"] = np.select([tc.Completion_Date < tc.Estimated_Completion_Date,
                                  tc.Completion_Date == tc.Estimated_Completion_Date], ["Early", "On Time"], "Late")
    tc["TAT_Days"] = (tc.Completion_Date - tc.Submission_Release_Date).dt.days
    tc26 = tc[tc.Completion_Date.dt.year == 2026]
    n = len(tc26)
    ontime = (tc26.TAT_Status != "Late").mean()
    jan_sep = tc26[tc26.Completion_Date <= "2026-09-30"]
    oos_rate = len(oos) / max(len(jan_sep), 1) * 1000
    kpis = [("Tests completed 2026", f"{n:,}", "Jan to Oct"), ("Early or on time", f"{ontime:.1%}", "vs estimated completion"),
            ("Average TAT", f"{tc26.TAT_Days.mean():.1f} days", "release to completion"),
            ("Rush share", f"{(tc26.Rush == 'Y').mean():.1%}", "of completed tests"),
            ("OOS rate", f"{oos_rate:.1f} / 1,000", "tests, Jan to Sep"),
            ("Active clients", f"{tc26.Client_Facility.nunique():,}", "with completed tests")]

    m = pd.date_range("2026-01-01", "2026-10-01", freq="MS")
    relm = rel.groupby(rel.Submission_Release_Date.dt.to_period("M")).size()
    comm = tc26.groupby(tc26.Completion_Date.dt.to_period("M")).size()
    p = [x.to_period("M") for x in m]
    f1 = go.Figure()
    f1.add_trace(go.Scatter(x=m, y=[relm.get(x, 0) for x in p], name="Released", mode="lines+markers",
                            line=dict(color=C[0], width=2), marker=dict(size=8, line=dict(color="#fff", width=2)),
                            hovertemplate="%{x|%b %Y}<br>Released: %{y:,}<extra></extra>"))
    f1.add_trace(go.Scatter(x=m, y=[comm.get(x, 0) for x in p], name="Completed", mode="lines+markers",
                            line=dict(color=C[1], width=2), marker=dict(size=8, line=dict(color="#fff", width=2)),
                            hovertemplate="%{x|%b %Y}<br>Completed: %{y:,}<extra></extra>"))
    style(f1).update_yaxes(title="Tests")

    st = (tc26.groupby([tc26.Completion_Date.dt.to_period("M"), "TAT_Status"]).size().unstack(fill_value=0))
    st = st.div(st.sum(axis=1), axis=0) * 100
    f2 = go.Figure()
    for s, col in (("Early", GOOD), ("On Time", C[0]), ("Late", CRIT)):
        f2.add_trace(go.Bar(x=[x.to_timestamp() for x in st.index], y=st.get(s, 0), name=s, marker=dict(color=col),
                            hovertemplate="%{x|%b %Y}<br>" + s + ": %{y:.1f}%<extra></extra>"))
    style(f2).update_layout(barmode="stack", bargap=0.35)
    f2.update_traces(marker_line=dict(color="#fff", width=2))
    f2.update_yaxes(title="% of tests", range=[0, 100])

    top = tc26[tc26.Analyte.notna()].groupby("Analyte").size().nlargest(10)
    f3 = hbar(list(top.index), list(top.values), C[0], text=[f"{v:,}" for v in top.values])

    dep = tc26.groupby(["Test_Name"]).agg(n=("Test_Id", "count"), tat=("TAT_Days", "mean")).nlargest(8, "n")
    rush = tc26[tc26.Test_Name.isin(dep.index)].groupby(["Test_Name", "Rush"]).TAT_Days.mean().unstack()
    f4 = go.Figure()
    for r, nm, col in (("N", "Standard", C[0]), ("Y", "Rush", C[1])):
        f4.add_trace(go.Bar(y=list(rush.index), x=rush[r].round(1), name=nm, orientation="h", marker=dict(color=col, cornerradius=4),
                            hovertemplate="%{y}<br>" + nm + ": %{x:.1f} days<extra></extra>"))
    style(f4, 360).update_layout(barmode="group", bargap=0.3, bargroupgap=0.08)
    f4.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")
    f4.update_xaxes(title="Average TAT (days)", showgrid=True, gridcolor=GRID)

    vol = jan_sep.groupby("Client_Facility").size()
    oc = oos.groupby("ClientName").size()
    lab_rate = len(oos) / len(jan_sep)
    idx = (oc.reindex(vol.index).fillna(0) / vol / lab_rate)[vol >= 400].nlargest(12)
    flag = ["Review" if v >= 2 else "Watch" if v >= 1.3 else "Keep" for v in idx.values]
    cols = [CRIT if f == "Review" else SERIOUS if f == "Watch" else GOOD for f in flag]
    f5 = hbar(list(idx.index), list(idx.round(2).values), colors=cols, fmt=".2f",
              text=[f"{v:.2f}x  {f}" for v, f in zip(idx.values, flag)], height=380)
    f5.add_vline(x=1, line=dict(color=MUTED, dash="dot", width=1))
    f5.update_xaxes(title="OOS rate vs lab average (1.0 = average)")

    stages = {"Sample prep": "Time_In_Sample_Prep", "Analysis": "Time_In_Sample_Analysis", "Data review": "Time_In_Data_Review",
              "On hold": "Time_In_On_Hold", "QA review": "Time_In_Qa_Review", "QA hold": "Time_In_Qa_Hold"}
    sv = pd.Series({k: qa[v].mean() for k, v in stages.items()})
    f6 = hbar(list(sv.index), list(sv.round(1).values), C[2], fmt=".1f", text=[f"{v:.1f} h" for v in sv.values], height=300)
    f6.update_xaxes(title="Average hours per test")

    cards = [("Tests released vs completed", "Monthly, 2026. Completed above released means the backlog is shrinking.", f1, False),
             ("Turnaround status", "Share of completed tests finished early, on time or late against the estimated date", f2, False),
             ("Top 10 analytes", "Completed tests in 2026", f3, False),
             ("Turnaround by test: standard vs rush", "Eight highest-volume tests, average days from release to completion", f4, False),
             ("OOS watchlist", "Clients with 400+ tests (Jan to Sep): out-of-spec rate relative to the lab average", f5, False),
             ("Where QA time goes", "Average hours per test spent in each workflow stage", f6, False)]
    return page("lab-test-performance.html", "Lab Test Performance & Quality",
                "Throughput, turnaround, analyst workload and out-of-spec results for a contract testing laboratory (Contoso Analytical Labs).",
                kpis, cards, out, "Source: lab LIMS extracts (Tests Completed, Submission Release, Quality Data, OOS).")


def pareto_page(d, out):
    raw = d["raw"]
    rev = raw.groupby("Client_Facility").Price.sum().sort_values(ascending=False)
    cum = rev.cumsum() / rev.sum() * 100
    n80 = int((cum < 80).sum() + 1)
    trev = raw.groupby("Test_Name").Price.sum().sort_values(ascending=False)
    tcum = trev.cumsum() / trev.sum() * 100
    t80 = int((tcum < 80).sum() + 1)
    vol = raw.groupby("Client_Facility").size().sort_values(ascending=False)
    vcum = vol.cumsum() / vol.sum() * 100
    kpis = [("Total revenue", f"${rev.sum() / 1e6:,.1f}M", "2023 to Oct 2026"), ("Clients", f"{len(rev):,}", "with test revenue"),
            ("Clients for 80% of revenue", f"{n80:,}", f"{n80 / len(rev):.0%} of clients"),
            ("Top 10 clients", f"{rev.head(10).sum() / rev.sum():.1%}", "share of revenue"),
            ("Tests for 80% of revenue", f"{t80}", f"of {len(trev)} tests"),
            ("Clients for 80% of volume", f"{int((vcum < 80).sum() + 1):,}", "by number of tests")]

    def curve(c, n, label):
        x = np.arange(1, len(c) + 1) / len(c) * 100
        f = go.Figure()
        f.add_trace(go.Scatter(x=x, y=c.values, mode="lines", line=dict(color=C[0], width=2), name="Cumulative share",
                               customdata=np.array(c.index), fill="tozeroy", fillcolor="rgba(42,120,214,0.08)",
                               hovertemplate="%{customdata}<br>Top %{x:.1f}% of " + label + "<br>Cumulative: %{y:.1f}%<extra></extra>"))
        f.add_hline(y=80, line=dict(color=MUTED, dash="dot", width=1))
        xn = n / len(c) * 100
        f.add_vline(x=xn, line=dict(color=C[1], width=1.5))
        f.add_annotation(x=xn, y=40, text=f"{n:,} {label} = 80%", showarrow=False, xanchor="left", xshift=6,
                         font=dict(color=INK, size=12))
        style(f, 340, legend=False).update_xaxes(title=f"% of {label} (ranked by revenue)", range=[0, 100])
        f.update_yaxes(title="Cumulative % of revenue", range=[0, 102])
        return f

    f1 = curve(cum, n80, "clients")
    top = rev.head(15)
    f2 = hbar(list(top.index), list((top / 1e3).round(1).values), C[0], fmt=",.1f",
              text=[f"${v / 1e3:,.0f}K" for v in top.values], height=440)
    f2.update_xaxes(title="Revenue ($K)")
    f3 = curve(tcum, t80, "tests")
    tt = trev.head(10)
    cols = [C[0] if i < t80 else NEUTRAL for i in range(len(tt))]
    f4 = hbar(list(tt.index), list((tt / 1e6).round(2).values), colors=cols, fmt=".2f",
              text=[f"${v / 1e6:,.1f}M" for v in tt.values], height=340)
    f4.update_xaxes(title="Revenue ($M)")
    yr = raw.assign(Y=pd.to_datetime(raw.Submission_Date).dt.year, A=raw.Client_Facility.isin(rev.index[:n80]))
    g = yr.groupby(["Y", "A"]).Price.sum().unstack() / 1e6
    f5 = go.Figure()
    for a, nm, col in ((True, "A+ clients (top 80%)", C[0]), (False, "Others", C[1])):
        f5.add_trace(go.Bar(x=g.index.astype(str), y=g[a].round(2), name=nm, marker=dict(color=col, cornerradius=4),
                            hovertemplate="%{x}<br>" + nm + ": $%{y:.2f}M<extra></extra>"))
    style(f5, 340).update_layout(barmode="group", bargroupgap=0.08)
    f5.update_yaxes(title="Revenue ($M)")
    cards = [("Client revenue Pareto", "Cumulative share of revenue as clients are added from largest to smallest", f1, False),
             ("Top 15 clients by revenue", "All years", f2, False),
             ("Test revenue Pareto", "Cumulative share of revenue across the test menu", f3, False),
             ("Top 10 tests by revenue", "Blue = inside the 80% group", f4, False),
             ("A+ clients vs others by year", "Revenue from the clients that make up 80% of all-time revenue (2026 is Jan to Oct)", f5, True)]
    return page("pareto-80-20.html", "Pareto (80/20) Analysis: Clients and Tests",
                "Which clients and which tests drive revenue, and how concentrated the business is.",
                kpis, cards, out, "Source: lab LIMS extracts RAW_2023 to RAW_2026.")


def rfm_page(d, out):
    raw = d["raw"].copy()
    raw["Submission_Date"] = pd.to_datetime(raw.Submission_Date)
    asof = raw.Submission_Date.max()
    last12 = raw[raw.Submission_Date > asof - pd.DateOffset(months=12)]
    g = raw.groupby("Client_Facility").agg(last=("Submission_Date", "max"))
    g["months"] = ((asof - g["last"]).dt.days / 30.4).round(1)
    g["F"] = last12.groupby("Client_Facility").size().reindex(g.index).fillna(0)
    g["M"] = last12.groupby("Client_Facility").Price.sum().reindex(g.index).fillna(0)
    g["R_s"] = pd.cut(g.months, [-1, 1, 3, 6, 12, 1e9], labels=[5, 4, 3, 2, 1]).astype(int)
    def q(s):
        out = pd.Series(1, index=s.index)
        pos = s[s > 0]
        out[pos.index] = pd.qcut(pos.rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
        return out
    g["F_s"], g["M_s"] = q(g.F), q(g.M)
    r6 = raw[raw.Submission_Date > asof - pd.DateOffset(months=6)].groupby("Client_Facility").Price.sum()
    p6 = raw[(raw.Submission_Date <= asof - pd.DateOffset(months=6)) &
             (raw.Submission_Date > asof - pd.DateOffset(months=12))].groupby("Client_Facility").Price.sum()
    chg = (r6.reindex(g.index).fillna(0) - p6.reindex(g.index).fillna(0)) / p6.reindex(g.index).replace(0, np.nan)
    g["chg"] = chg

    def cohort(r):
        if r.R_s >= 4 and r.F_s >= 4 and r.M_s >= 4: return "Champions"
        if r.R_s >= 3 and r.F_s >= 4: return "Loyal"
        if r.R_s >= 3 and r.M_s >= 4: return "Big spenders"
        if r.R_s <= 2 and (r.F_s >= 3 or r.M_s >= 3): return "At risk"
        if r.R_s >= 4 and r.F_s <= 2: return "New / promising"
        if r.R_s == 1: return "Lost"
        return "Needs attention"
    g["Cohort"] = g.apply(cohort, axis=1)

    def band(c):
        if pd.isna(c): return None
        return "Accelerating" if c >= 0.2 else "Improving" if c >= 0.05 else "Stable" if c > -0.05 else \
            "Early decline" if c > -0.25 else "Sharp decline"
    g["Band"] = g.chg.map(band)
    active = g[g.months <= 12]
    rev12 = last12.Price.sum()
    prev12 = raw[(raw.Submission_Date <= asof - pd.DateOffset(months=12)) &
                 (raw.Submission_Date > asof - pd.DateOffset(months=24))].Price.sum()
    kpis = [("Active clients", f"{len(active):,}", "ordered in the last 12 months"),
            ("Revenue, last 12 months", f"${rev12 / 1e6:,.1f}M", f"{rev12 / prev12 - 1:+.1%} vs prior 12 months"),
            ("Champions", f"{(g.Cohort == 'Champions').sum():,}", "recent, frequent, high spend"),
            ("At risk", f"{(g.Cohort == 'At risk').sum():,}", "valuable but going quiet"),
            ("Sharp decline", f"{(g.Band == 'Sharp decline').sum():,}", "6-month revenue down 25%+"),
            ("Revenue in decline", f"${g[g.Band.isin(['Early decline', 'Sharp decline'])].M.sum() / 1e6:,.1f}M",
             "12-month revenue of declining clients")]

    order = ["Champions", "Loyal", "Big spenders", "New / promising", "Needs attention", "At risk", "Lost"]
    cc = g.Cohort.value_counts().reindex(order).fillna(0)
    cm = g.groupby("Cohort").M.sum().reindex(order).fillna(0)
    f1 = hbar(order, list(cc.values), C[0], text=[f"{int(c):,} clients · ${m / 1e6:,.1f}M" for c, m in zip(cc.values, cm.values)],
              height=330)
    f1.update_xaxes(title="Clients")

    bo = ["Accelerating", "Improving", "Stable", "Early decline", "Sharp decline"]
    bc = g.Band.value_counts().reindex(bo).fillna(0)
    bcol = [GOOD, "#7cc47f", NEUTRAL, SERIOUS, CRIT]
    f2 = go.Figure(go.Bar(x=bo, y=bc.values, marker=dict(color=bcol, cornerradius=4), text=[f"{int(v):,}" for v in bc.values],
                          textposition="outside", cliponaxis=False, hovertemplate="%{x}: %{y:,} clients<extra></extra>"))
    style(f2, 330, legend=False).update_yaxes(title="Clients")

    mon = raw.assign(Mo=raw.Submission_Date.dt.to_period("M").dt.to_timestamp())
    mon = mon[mon.Mo < pd.Timestamp(asof.year, asof.month, 1)]   # drop the partial last month
    mt = mon.groupby(["Mo", "Facility_Type"]).Price.sum().unstack().fillna(0) / 1e3
    f3 = go.Figure()
    for i, t in enumerate(["503A", "503B", "GMP"]):
        if t in mt:
            f3.add_trace(go.Scatter(x=mt.index, y=mt[t].round(1), name=t, mode="lines", line=dict(color=C[i], width=2),
                                    hovertemplate="%{x|%b %Y}<br>" + t + ": $%{y:,.0f}K<extra></extra>"))
    style(f3, 330).update_yaxes(title="Revenue ($K)")

    s = active[active.M > 0]
    f4 = go.Figure()
    for i, (nm, mask) in enumerate((("Champions", s.Cohort == "Champions"), ("At risk", s.Cohort == "At risk"),
                                    ("Other active", ~s.Cohort.isin(["Champions", "At risk"])))):
        sub = s[mask]
        f4.add_trace(go.Scatter(x=sub.F, y=sub.M, mode="markers", name=nm, customdata=np.array(sub.index),
                                marker=dict(size=8, color=[C[0], C[1], C[2]][i], opacity=0.75, line=dict(color="#fff", width=1)),
                                hovertemplate="%{customdata}<br>%{x:,} tests · $%{y:,.0f}<extra>" + nm + "</extra>"))
    style(f4, 330).update_xaxes(type="log", title="Tests, last 12 months (log)")
    f4.update_yaxes(type="log", title="Revenue, last 12 months (log)")

    cf = d["cf"]
    st = last12.merge(cf[["ClientFacilityName", "BillTo_StateProvince"]], left_on="Client_Facility",
                      right_on="ClientFacilityName", how="left").groupby("BillTo_StateProvince").Price.sum().nlargest(12)
    f5 = hbar(list(st.index), list((st / 1e3).round(0).values), C[0], text=[f"${v / 1e3:,.0f}K" for v in st.values], height=420)
    f5.update_xaxes(title="Revenue, last 12 months ($K)")

    watch = g[g.Band == "Sharp decline"].nlargest(8, "M")
    f6 = go.Figure(go.Table(
        header=dict(values=["Client", "Cohort", "12M revenue", "6M change", "Months since last order"],
                    fill_color="#f3faf3", line_color="#d6ecd6", font=dict(color=INK, size=12), align="left"),
        cells=dict(values=[watch.index, watch.Cohort, [f"${v:,.0f}" for v in watch.M], [f"{v:+.0%}" for v in watch.chg],
                           watch.months], fill_color="#ffffff", line_color="#ececea", font=dict(color=INK2, size=12), align="left",
                   height=30), columnwidth=[2.6, 1.2, 1.1, 0.9, 1.1]))
    style(f6, 420, legend=False)
    cards = [("Behavior cohorts", "RFM scoring: recency, frequency (tests) and monetary value over the last 12 months", f1, False),
             ("Early warning bands", "Last 6 months of revenue vs the 6 months before", f2, False),
             ("Monthly revenue by facility type", "503A pharmacies, 503B outsourcing facilities, GMP manufacturers", f3, False),
             ("Frequency vs spend", "Active clients, last 12 months. Hover for client names.", f4, False),
             ("Revenue by state", "Top 12 states, last 12 months", f5, False),
             ("Call list: largest clients in sharp decline", "Sorted by 12-month revenue", f6, False)]
    return page("client-rfm.html", "Client Performance, RFM Segmentation & Early Warning",
                "Who the best clients are, who is slipping, and where revenue is at risk.",
                kpis, cards, out, f"As of {asof:%d %b %Y}. Source: lab LIMS extracts and client facility master.")


# ---------------------------------------------------------------- gigafactory
def giga_page(root, out):
    G = os.path.join(root, "data", "gigafactory")
    sq = lambda n: pd.read_csv(os.path.join(G, "sql_server", "csv", n + ".csv"), low_memory=False)
    dd = pd.Timestamp(sq("ref_project_info").Data_Date.iloc[0])
    docs = sq("eng_document_revision")
    for c in ("Planned_Issue_Date", "Actual_Issue_Date"):
        docs[c] = pd.to_datetime(docs[c])
    rfi = sq("eng_rfi")
    for c in ("Date_Raised", "Date_Required", "Date_Responded", "Date_Closed"):
        rfi[c] = pd.to_datetime(rfi[c])
    sub = sq("ref_subsystem")
    for c in ("Planned_MC_Date", "Actual_MC_Date"):
        sub[c] = pd.to_datetime(sub[c])
    disc = sq("ref_discipline").set_index("Discipline_Code").Discipline
    erp = pd.read_excel(os.path.join(G, "erp_procurement", "PO_Lines_Export.xlsx"))
    cwp = pd.read_excel(os.path.join(G, "p6_schedule", "Construction_Progress_Export.xlsx"), sheet_name="WBS_CWP")
    wk = pd.read_excel(os.path.join(G, "p6_schedule", "Construction_Progress_Export.xlsx"), sheet_name="Weekly_Progress")
    wk["Week Ending"] = pd.to_datetime(wk["Week Ending"])
    itr = pd.concat([pd.read_csv(f, parse_dates=["Planned_Date", "Completed_Date"])
                     for f in glob.glob(os.path.join(G, "completions_system", "itr_extracts", "*.csv"))])
    punch = pd.read_csv(os.path.join(G, "completions_system", "punch_list.csv"), parse_dates=["Date_Raised", "Date_Cleared"])
    hse = pd.DataFrame(json.load(open(os.path.join(G, "hse_api", "safety_weekly.json")))["value"])
    hse["weekEnding"] = pd.to_datetime(hse.weekEnding)

    ifc = docs[docs.Revision.astype(str) == "0"]
    budget = cwp["Budgeted Units (hrs)"].sum()
    h = wk[wk["Week Ending"] <= dd]
    spi = h["Earned Units"].sum() / h["Planned Units"].sum()
    cpi = h["Earned Units"].sum() / h["Actual Units"].sum()
    open_rfi = ((rfi.Status != "Void") & rfi.Date_Closed.isna()).sum()
    trir = hse.recordables.sum() * 200000 / hse.manhours.sum()
    kpis = [("IFC drawings issued", f"{ifc.Actual_Issue_Date.notna().mean():.1%}", f"plan {(ifc.Planned_Issue_Date <= dd).mean():.1%}"),
            ("Construction earned", f"{h['Earned Units'].sum() / budget:.1%}", f"plan {h['Planned Units'].sum() / budget:.1%}"),
            ("SPI / CPI", f"{spi:.2f}/{cpi:.2f}", "below 1.00 = behind / over"),
            ("Open RFIs", f"{open_rfi:,}", f"{((rfi.Status != 'Void') & rfi.Date_Responded.isna() & (rfi.Date_Required < dd)).sum():,} overdue"),
            ("PO lines late vs ROS", f"{(erp['Expediting Status'] == 'Late vs ROS').sum():,}", "forecast after required-on-site"),
            ("Subsystems at MC", f"{sub.Actual_MC_Date.notna().sum()} / {(sub.Planned_MC_Date <= dd).sum()}", "actual / planned to date"),
            ("Open punch, Cat A", f"{((punch.Category == 'A') & punch.Date_Cleared.isna()).sum():,}", "blocks mechanical completion"),
            ("TRIR", f"{trir:.2f}", "per 200,000 hours, project to date")]

    weeks = pd.date_range("2024-06-07", "2027-09-24", freq="W-FRI")
    marker = dict(color=MUTED, dash="dot", width=1)

    def cum(series_dates, upto=None):
        s = series_dates.dropna().sort_values().values
        y = np.searchsorted(s, weeks.values, side="right").astype(float)
        if upto is not None:
            y[weeks > upto] = np.nan
        return y

    f1 = go.Figure()
    line(f1, weeks, cum(ifc.Planned_Issue_Date), "Planned", C[0], dash="dash")
    line(f1, weeks, cum(ifc.Actual_Issue_Date, dd), "Actual", C[1])
    style(f1).update_yaxes(title="IFC documents (cumulative)")
    f1.add_vline(x=dd, line=marker)

    ws = wk.groupby("Week Ending")[["Planned Units", "Earned Units", "Actual Units"]].sum().cumsum() / budget * 100
    ws.loc[ws.index > dd, ["Earned Units", "Actual Units"]] = np.nan
    f2 = go.Figure()
    for col, nm, c, dsh in (("Planned Units", "Planned", C[0], "dash"), ("Earned Units", "Earned", C[1], None),
                            ("Actual Units", "Actual hours spent", C[2], None)):
        f2.add_trace(go.Scatter(x=ws.index, y=ws[col].round(1), name=nm, mode="lines", line=dict(color=c, width=2, dash=dsh),
                                hovertemplate="%{x|%d %b %Y}<br>" + nm + ": %{y:.1f}% of budget<extra></extra>"))
    style(f2).update_yaxes(title="% of budget hours")
    f2.add_vline(x=dd, line=marker)

    live = rfi[rfi.Status != "Void"]
    wk_hist = weeks[weeks <= dd]
    rs, cl = np.sort(live.Date_Raised.values), np.sort(live.Date_Closed.dropna().values)
    open_series = np.searchsorted(rs, wk_hist.values, side="right") - np.searchsorted(cl, wk_hist.values, side="right")
    f3 = go.Figure()
    line(f3, wk_hist, open_series, "Open RFIs", C[0])
    style(f3, legend=False).update_yaxes(title="Open RFIs")

    o = live[live.Date_Closed.isna()].copy()
    o["age"] = (dd - o.Date_Raised).dt.days
    o["bucket"] = pd.cut(o.age, [-1, 7, 14, 30, 1e9], labels=["0-7 days", "8-14 days", "15-30 days", "Over 30 days"])
    ab = o.groupby(["Discipline_Code", "bucket"], observed=False).size().unstack(fill_value=0)
    ab = ab.loc[ab.sum(axis=1).sort_values().index]
    f4 = go.Figure()
    for i, b in enumerate(ab.columns):
        f4.add_trace(go.Bar(y=[disc[x] for x in ab.index], x=ab[b], name=str(b), orientation="h",
                            marker=dict(color=["#86b6ef", "#3987e5", "#1c5cab", "#0d366b"][i]),
                            hovertemplate="%{y}<br>" + str(b) + ": %{x}<extra></extra>"))
    style(f4).update_layout(barmode="stack")
    f4.update_traces(marker_line=dict(color="#fff", width=2))
    f4.update_xaxes(title="Open RFIs", showgrid=True, gridcolor=GRID)
    f4.update_yaxes(gridcolor="rgba(0,0,0,0)")

    op = erp[erp["Expediting Status"] != "Delivered"]
    es = op.groupby(["Discipline", "Expediting Status"]).size().unstack(fill_value=0)
    es = es.loc[es.sum(axis=1).sort_values().index]
    f5 = go.Figure()
    for s, c in (("On Track", C[0]), ("At Risk", WARN), ("Late vs ROS", CRIT)):
        if s in es:
            f5.add_trace(go.Bar(y=[disc[x] for x in es.index], x=es[s], name=s, orientation="h", marker=dict(color=c),
                                hovertemplate="%{y}<br>" + s + ": %{x:,} lines<extra></extra>"))
    style(f5).update_layout(barmode="stack")
    f5.update_traces(marker_line=dict(color="#fff", width=2))
    f5.update_xaxes(title="PO lines", showgrid=True, gridcolor=GRID)
    f5.update_yaxes(gridcolor="rgba(0,0,0,0)")

    f6 = go.Figure()
    line(f6, weeks, cum(sub.Planned_MC_Date), "Planned MC", C[0], dash="dash")
    line(f6, weeks, cum(sub.Actual_MC_Date, dd), "Actual MC", C[1])
    style(f6).update_yaxes(title="Subsystems (cumulative)")
    f6.add_vline(x=dd, line=marker)

    wh = weeks[(weeks >= "2026-02-01") & (weeks <= dd)]
    f7 = go.Figure()
    for i, cat in enumerate("ABC"):
        p = punch[punch.Category == cat]
        r_, c_ = np.sort(p.Date_Raised.values), np.sort(p.Date_Cleared.dropna().values)
        y = np.searchsorted(r_, wh.values, side="right") - np.searchsorted(c_, wh.values, side="right")
        line(f7, wh, y, f"Category {cat}", [C[7], C[3], C[0]][i])
    style(f7).update_yaxes(title="Open punch items")

    a = itr[itr.Stage == "A"]
    sysc = a.Subsystem_Code.str.split("-").str[0]
    ia = a.groupby(sysc).Completed_Date.apply(lambda x: x.notna().mean() * 100).sort_values()
    ia = ia[ia > 0].tail(15)
    f8 = hbar(list(ia.index), list(ia.round(1).values), C[2], fmt=".1f", text=[f"{v:.0f}%" for v in ia.values], height=340)
    f8.update_xaxes(title="A check sheets complete (%)", range=[0, 105])

    hm = hse.set_index("weekEnding").resample("MS")[["recordables", "manhours"]].sum()
    roll = hm.rolling(12, min_periods=6).sum()
    tr = (roll.recordables * 200000 / roll.manhours).round(2).dropna()
    f9 = go.Figure()
    f9.add_trace(go.Scatter(x=tr.index, y=tr, mode="lines", line=dict(color=C[0], width=2), name="TRIR",
                            hovertemplate="%{x|%b %Y}<br>TRIR (rolling 12M): %{y:.2f}<extra></extra>"))
    style(f9, legend=False).update_yaxes(title="TRIR, rolling 12 months", rangemode="tozero")

    cards = [("Engineering: IFC documents", "Issued for construction, planned vs actual. Dotted line = data date.", f1, False),
             ("Construction progress (earned value)", "Cumulative planned, earned and actual hours as % of budget", f2, False),
             ("Open RFI trend", "Requests for information open at each week end", f3, False),
             ("Open RFI aging by discipline", "As of the data date", f4, False),
             ("Open PO lines by expediting status", "Lines not yet delivered, by discipline (from the ERP export)", f5, False),
             ("Mechanical completion", "Subsystems mechanically complete, planned vs actual", f6, False),
             ("Punch list burndown", "Open items by category since walkdowns began", f7, False),
             ("Check sheet (ITR A) completion by system", "Systems with work started", f8, False),
             ("Safety: TRIR", "Recordable incidents per 200,000 hours, rolling 12 months (from the HSE API)", f9, True)]
    return page("gigafactory-project-analytics.html", "Project Saguaro: Gigafactory Project Analytics",
                f"Fictional lithium-ion battery cell gigafactory. Data from six systems (SQL Server, ERP, P6, completions, HSE API, SharePoint). Status date {dd:%d %b %Y}.",
                kpis, cards, out, "Sources: SQL Server engineering database, ERP PO export, P6 progress export, completions extracts, HSE API, SharePoint NCR log.")


def index_page(out):
    items = [("lab-test-performance.html", "Lab Test Performance & Quality",
              "Throughput, turnaround, rush work, QA time and an out-of-spec client watchlist."),
             ("pareto-80-20.html", "Pareto 80/20 Analysis", "How concentrated revenue is across clients and tests."),
             ("client-rfm.html", "Client RFM Segmentation & Early Warning", "Behavior cohorts, declining clients and a call list."),
             ("gigafactory-project-analytics.html", "Gigafactory Project Analytics",
              "Engineering, RFIs, procurement, earned value, completions and safety from six data sources.")]
    links = "".join(f'<a href="{f}"><b>{t}</b><span>{s}</span></a>'
                    f'<div style="margin:-6px 0 16px 2px"><a class="dl" style="display:inline-block;padding:4px 10px;margin:0" href="downloads/{DOWNLOADS[f][0]}-powerbi.zip">Power BI file (.pbip zip)</a></div>'
                    for f, t, s in items)
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Power BI & Analytics Portfolio: Sai Sneha Ajmira</title><style>{CSS}</style></head><body>
<header><h1>Power BI & Analytics Portfolio<span class="badge">Synthetic data</span></h1>
<p>Interactive versions of the Power BI reports in <a href="https://github.com/SnehaAjmira/powerbi-analytics-portfolio" style="color:#2e6b34">this repository</a>. Sai Sneha Ajmira.</p></header>
<main style="max-width:860px"><div class="links">{links}</div></main>
<footer>All data is synthetic and generated by the scripts in the repository.</footer></body></html>"""
    open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(doc)


def screenshot(html, png, height):
    chrome = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chrome")
    if not chrome:
        print("Chrome not found; skipping", png)
        return
    subprocess.run([chrome, "--headless=new", "--no-sandbox", "--disable-gpu", "--hide-scrollbars",
                    f"--window-size=1440,{height}", "--virtual-time-budget=15000", f"--screenshot={png}",
                    "file://" + os.path.abspath(html)], check=True, capture_output=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="..")
    a = ap.parse_args()
    out = os.path.join(a.root, "docs")
    os.makedirs(os.path.join(out, "assets"), exist_ok=True)
    shutil.copy(os.path.join(os.path.dirname(plotly.__file__), "package_data", "plotly.min.js"),
                os.path.join(out, "assets", "plotly.min.js"))
    lab = load_lab(a.root)
    builds = [("lab-test-performance.html", lambda: lab_page(lab, out), "01-lab-test-performance"),
              ("pareto-80-20.html", lambda: pareto_page(lab, out), "02-pareto-80-20-analysis"),
              ("client-rfm.html", lambda: rfm_page(lab, out), "03-client-segmentation-rfm"),
              ("gigafactory-project-analytics.html", lambda: giga_page(a.root, out), "04-gigafactory-project-analytics")]
    index_page(out)
    os.makedirs(os.path.join(out, "img"), exist_ok=True)
    for f, fn, proj in builds:
        h = fn()
        png = os.path.join(a.root, proj, "screenshots", "dashboard.png")
        os.makedirs(os.path.dirname(png), exist_ok=True)
        screenshot(os.path.join(out, f), png, h)
        shutil.copy(png, os.path.join(out, "img", f.replace(".html", ".png"))) if os.path.exists(png) else None
        print("built", f)


if __name__ == "__main__":
    main()
