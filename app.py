"""TLP Escalation Dashboard — Access Issue, Stolen, dan TT CME (tema gelap).

Membaca file .xlsb (sheet `List Access Issue`, `List Stolen`, `TT CME`) dengan pyxlsb.
Tata letak: header + tombol aksi, bar filter horizontal, 6 kartu KPI, tab navigasi,
kartu Executive Insight, lalu grafik dalam kartu.

Grafik summary A–E per bulan mengikuti pivot di sheet Summary:
  A = akumulasi TT open bulan sebelumnya      B = TT baru per bulan
  C = A + B (akumulasi TT open per bulan)     D = TT close per bulan (C − E)
  E = sisa TT yang masih open s/d bulan tsb   Progress = D / C

Jalankan:  streamlit run app.py
"""
from __future__ import annotations

import html
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pyxlsb import open_workbook

APP_DIR = Path(__file__).resolve().parent
ALL = "Semua"

# Kategori dari teks bebas (Stolen: Material Stolen, CME: Description) — aturan kata kunci, urutan = prioritas.
STOLEN_RULES = [
    ("No stolen / investigasi", r"no stolen|not stolen|tidak ada (?:yang )?hilang|investigat"),
    ("Battery", r"batt|batre|baterai|bater|vrla|accu|aki\b"),
    ("RRU / Antena", r"rru|antena|antenna|radio"),
    ("Grounding", r"ground|arde|bc ?50|penangkal"),
    ("Kabel / Feeder", r"cable|kabel|feeder|cabel"),
    ("Rectifier / Modul", r"rectif|modul|module|psu"),
    ("Genset / BBM", r"genset|solar|bbm|fuel"),
    ("Pagar / Gembok", r"pagar|fence|gembok|padlock|kunci|gate"),
]
CME_RULES = [
    ("Arrester", r"arrest"),
    ("Grounding", r"ground|\bpit\b|bak (?:kontrol|control)"),
    ("APAR", r"\bapar\b"),
    ("Lampu / OBL", r"yard ?lamp|lampu|\bobl\b|obstruction"),
    ("PLN / Tegangan", r"\bpln\b|voltage|tegangan|kwh"),
    ("ACPDB / Kelistrikan", r"acpdb|mcb|panel|kabel|cable|listrik"),
    ("Pagar / Shelter", r"pagar|fence|shelter|kawat|bordes|ladder|tangga|pintu|door"),
    ("AC / Cooling", r"\bac\b|aircon|cooling|fan\b"),
]

CATEGORIES = {
    "Access Issue": dict(
        sheet="List Access Issue", month="Month Open", status="Case Status (Open/Closed)", tlp="GROUP",
        cause="Problem Type", cause_label="Main Cause", severity="Severity", aging="Aging Time (Days)",
        year="Year Open", site=["Site ID", "Site ID IOH", "Site Name", "Remarks from FLM"], icon="🚧"),
    "Stolen": dict(
        sheet="List Stolen", month="Month Open", status="Status TT ToCo/OpCo", tlp="Toco Group Name",
        cause="Material Stolen", cause_rules=STOLEN_RULES, cause_label="Material Stolen",
        severity="TT Status [Critical/Major]", aging="Aging TT", year="Year Open",
        site=["IOH ID", "Old Site ID", "Site Name", "TT IOH", "Material Stolen"], icon="🔓"),
    "TT CME": dict(
        sheet="TT CME", month="Create Month", status="STATUS TT TLP", tlp="Owner",
        cause="Description", cause_rules=CME_RULES, cause_label="Kategori CME", severity="Severity",
        aging="Aging TT CLOSED", year=None, site=["Site ID", "Site Name", "TT Number", "Description"],
        backlog="Backlog", icon="🛠️"),
}
REGION_COL = "New Region"

# Palet untuk latar gelap.
BLUE, CYAN, PINK, YELLOW, PURPLE, TEAL, GREEN, RED, ORANGE = (
    "#3b9cff", "#22d3ee", "#ff5c8a", "#fbbf24", "#a78bfa", "#2dd4bf", "#10d39a", "#ff4d6d", "#ff9f43")
TEXT, TEXT_2, MUTED = "#e6edf8", "#c3cee6", "#8fa3c8"
GRID = "rgba(148,170,220,0.14)"
FONT = 'Inter, "Segoe UI", system-ui, sans-serif'
SERIES = {
    "A": ("A: Akumulasi TT open bulan sebelumnya", "#7c8db5"),
    "B": ("B: TT baru per bulan", BLUE),
    "C": ("C: Akumulasi TT open per bulan", YELLOW),
    "D": ("D: TT close per bulan", GREEN),
    "E": ("E: Sisa TT masih open [C−D]", RED),
}

st.set_page_config(page_title="TLP Escalation Dashboard", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown(f"""
<style>
  .stApp {{background: radial-gradient(1200px 600px at 85% -10%, #1d3f8f 0%, rgba(29,63,143,0) 60%),
                       linear-gradient(160deg, #081226 0%, #0c1c42 45%, #10285c 100%); background-attachment: fixed;}}
  [data-testid="stHeader"] {{background: transparent;}}
  .block-container {{padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1600px;}}

  /* Header */
  .hdr {{display:flex; align-items:center; gap:14px;}}
  .hdr .logo {{width:46px; height:46px; border-radius:12px; display:flex; align-items:center; justify-content:center;
              background:linear-gradient(135deg,#ff9f43,#ff6b2c); color:#fff; font-weight:800; font-size:.8rem;
              box-shadow:0 4px 14px rgba(255,120,50,.35);}}
  .hdr h1 {{font-size:1.45rem; margin:0; padding:0; color:{TEXT}; font-weight:800; letter-spacing:-.01em;}}
  .hdr .pill {{display:inline-block; margin-left:10px; padding:2px 10px; border-radius:999px; font-size:.68rem;
              font-weight:600; color:#9ff3ff; background:rgba(34,211,238,.12); border:1px solid rgba(34,211,238,.45);
              vertical-align:middle;}}
  .hdr .sub {{color:{MUTED}; font-size:.78rem; margin-top:2px;}}

  /* Bar filter */
  .st-key-filterbar {{background:rgba(14,30,68,.78); border:1px solid rgba(96,140,230,.28) !important;
                      border-radius:16px; padding:10px 16px 4px 16px; box-shadow:0 8px 24px rgba(0,0,0,.25);}}
  [data-testid="stWidgetLabel"] p {{text-transform:uppercase; font-size:.68rem !important; font-weight:700;
                                   letter-spacing:.06em; color:{MUTED} !important;}}
  [data-baseweb="select"] > div, .stTextInput input {{background:rgba(8,20,48,.9) !important;
                                   border-color:rgba(96,140,230,.35) !important;}}
  .st-key-filterbar [data-baseweb="select"] div, .st-key-filterbar input {{font-size:.8rem !important;}}
  .st-key-filterbar [data-baseweb="select"] > div {{padding-left:2px;}}
  .st-key-filterbar [data-testid="stColumn"] {{min-width:0;}}

  /* Kartu KPI */
  .kpis {{display:grid; grid-template-columns:repeat(6, minmax(0,1fr)); gap:14px; margin:14px 0 6px;}}
  .kpi {{background:rgba(14,30,68,.82); border:1px solid rgba(96,140,230,.22); border-top:3px solid var(--c);
         border-radius:14px; padding:14px 16px 12px; min-height:118px; box-shadow:0 8px 22px rgba(0,0,0,.25);}}
  .kpi .lbl {{font-size:.68rem; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:{TEXT_2};}}
  .kpi .val {{font-size:1.9rem; font-weight:800; color:var(--c); line-height:1.25; margin-top:6px;
              white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}}
  .kpi .val.txt {{font-size:1.25rem; margin-top:10px;}}
  .kpi .sub {{font-size:.74rem; color:{MUTED}; margin-top:4px;}}

  /* Tab navigasi berbentuk pill */
  [role="tablist"] {{gap:8px; flex-wrap:wrap; border-bottom:1px solid rgba(96,140,230,.2); padding-bottom:10px;}}
  [data-testid="stTab"] {{background:rgba(14,30,68,.6); border:1px solid rgba(96,140,230,.2); border-radius:10px;
                         padding:6px 14px !important; height:auto; transition:background .15s;}}
  [data-testid="stTab"] p {{font-size:.86rem; font-weight:600; color:{TEXT_2};}}
  [data-testid="stTab"]:hover {{background:rgba(59,156,255,.16);}}
  [data-testid="stTab"][aria-selected="true"] {{background:linear-gradient(135deg,#2f7df6,#3b9cff);
                                               border-color:#3b9cff; box-shadow:0 4px 14px rgba(59,156,255,.35);}}
  [data-testid="stTab"][aria-selected="true"] p {{color:#fff;}}
  [role="tablist"] > :not([role="tab"]), [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {{display:none;}}

  /* Kartu Executive Insight */
  .insights {{display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:14px; margin:6px 0 14px;}}
  .ins {{background:rgba(14,30,68,.82); border:1px solid rgba(96,140,230,.22); border-left:4px solid var(--c);
         border-radius:14px; padding:14px 16px; box-shadow:0 8px 22px rgba(0,0,0,.25);}}
  .ins h4 {{margin:0 0 6px; font-size:.92rem; color:{TEXT}; font-weight:700;}}
  .ins p {{margin:0; font-size:.83rem; line-height:1.5; color:{TEXT_2};}}
  .ins b {{color:#fff;}}

  /* Kartu grafik */
  [class*="st-key-card_"] {{background:rgba(14,30,68,.82); border:1px solid rgba(96,140,230,.22) !important;
                            border-radius:14px; padding:12px 14px 6px; box-shadow:0 8px 22px rgba(0,0,0,.25);}}
  .ctitle {{font-size:.95rem; font-weight:700; color:{TEXT}; margin:2px 0 4px;}}
  .badge {{display:inline-block; font-size:.66rem; font-weight:700; padding:2px 7px; border-radius:6px;
           margin-right:8px; vertical-align:middle; color:var(--c); border:1px solid var(--c);
           background:color-mix(in srgb, var(--c) 14%, transparent);}}
</style>
""", unsafe_allow_html=True)


# ---------- membaca .xlsb dengan pyxlsb ----------

MAX_EMPTY_RUN = 500  # berhenti setelah 500 baris kosong berturut-turut


def _clean_name(v, i: int, seen: dict) -> str:
    name = " ".join(str(v).split()) if v not in (None, "") else f"Kolom_{i + 1}"
    if name in seen:
        seen[name] += 1
        name = f"{name}_{seen[name]}"
    else:
        seen[name] = 0
    return name


def read_sheet(wb, sheet: str) -> pd.DataFrame:
    rows, empty_run = [], 0
    with wb.get_sheet(sheet) as sh:
        for r in sh.rows(sparse=True):
            vals = [c.v for c in r]
            if all(v is None or v == "" for v in vals):
                empty_run += 1
                if empty_run >= MAX_EMPTY_RUN:
                    break
                continue
            empty_run = 0
            rows.append(vals)
    if not rows:
        return pd.DataFrame()
    seen: dict = {}
    header = [_clean_name(v, i, seen) for i, v in enumerate(rows[0])]
    width = len(header)
    body = [(r + [None] * width)[:width] for r in rows[1:]]
    return pd.DataFrame(body, columns=header).dropna(how="all")


@st.cache_data(show_spinner="Membaca file .xlsb (pyxlsb)…", max_entries=4)
def load_all(key: str, _source: str | bytes) -> dict[str, pd.DataFrame]:
    src = io.BytesIO(_source) if isinstance(_source, bytes) else _source
    with open_workbook(src) as wb:
        return {k: read_sheet(wb, cfg["sheet"]) for k, cfg in CATEGORIES.items() if cfg["sheet"] in wb.sheets}


# ---------- persiapan data ----------

def to_period(s: pd.Series) -> pd.Series:
    text = s.astype("string").str.strip()
    out = pd.to_datetime(text, format="%b_%Y", errors="coerce")
    num = pd.to_numeric(s, errors="coerce")
    serial = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    out = out.fillna(serial.where(num > 20000))
    out = out.fillna(pd.to_datetime(text.where(out.isna()), errors="coerce"))
    return out.dt.to_period("M")


def classify(text: pd.Series, rules: list[tuple[str, str]]) -> pd.Series:
    low = text.astype("string").str.lower().fillna("")
    out = pd.Series("Lainnya", index=text.index, dtype="string")
    for label, pattern in reversed(rules):
        out = out.mask(low.str.contains(pattern, regex=True), label)
    return out


def _col(df: pd.DataFrame, name: str | None, default="(Kosong)") -> pd.Series:
    if name and name in df.columns:
        return df[name].astype("string").str.strip().replace("", pd.NA).fillna(default)
    return pd.Series(default, index=df.index, dtype="string")


@st.cache_data(show_spinner=False, max_entries=8)
def prepare(key: str, _df: pd.DataFrame, cat: str) -> pd.DataFrame:
    cfg = CATEGORIES[cat]
    df = _df.copy()
    df["_period"] = to_period(df[cfg["month"]]) if cfg["month"] in df.columns else pd.NaT
    df = df[df["_period"].notna()].copy()
    status = _col(df, cfg["status"], "").str.lower()
    df["_status"] = status.map(lambda s: "Open" if s == "open" else "Cancel" if s.startswith("cancel")
                               else "Closed" if s.startswith("close") else "Lainnya")
    df["_region"] = _col(df, REGION_COL)
    df["_tlp"] = _col(df, cfg["tlp"])
    raw_cause = _col(df, cfg["cause"])
    df["_cause"] = classify(raw_cause, cfg["cause_rules"]) if cfg.get("cause_rules") else raw_cause
    df["_sev"] = _col(df, cfg["severity"]).str.title()
    aging = pd.to_numeric(df[cfg["aging"]], errors="coerce") if cfg["aging"] in df.columns else pd.Series(
        float("nan"), index=df.index)
    df["_aging"] = aging.where(aging >= 0)
    df["_year"] = (_col(df, cfg["year"]).str.replace(r"\.0$", "", regex=True) if cfg["year"]
                   else df["_period"].dt.year.astype("string"))
    df["_backlog"] = _col(df, cfg.get("backlog"), "Normal TT")
    df["_month"] = df["_period"].dt.strftime("%b_%Y")
    df["_search"] = (pd.concat([_col(df, c, "") for c in cfg["site"] if c in df.columns], axis=1)
                     .agg(" ".join, axis=1).str.lower()) if any(c in df.columns for c in cfg["site"]) else ""
    return df


# ---------- perhitungan ----------

def backlog_summary(df: pd.DataFrame, months: list) -> pd.DataFrame:
    idx = pd.PeriodIndex(months, freq="M")
    new = df.groupby("_period").size().reindex(idx, fill_value=0)
    still_open = df[df["_status"].eq("Open")].groupby("_period").size().reindex(idx, fill_value=0)
    e = still_open.cumsum()
    a = e.shift(1, fill_value=0)
    c = a + new
    out = pd.DataFrame({"A": a, "B": new, "C": c, "D": c - e, "E": e})
    out["Progress (%)"] = (out["D"] / out["C"].where(out["C"] > 0) * 100).round(1)
    out["Remaining (%)"] = (out["E"] / out["C"].where(out["C"] > 0) * 100).round(1)
    out.index = [p.strftime("%b_%Y") for p in idx]
    out.index.name = "Bulan"
    return out


def fmt(v) -> str:
    return f"{v:,.0f}".replace(",", ".")


def pct(v, digits: int = 1) -> str:
    return "-" if pd.isna(v) else f"{v:.{digits}f}%".replace(".", ",")


def days(v) -> str:
    return "-" if pd.isna(v) else f"{v:.2f}".replace(".", ",")


def esc(s) -> str:
    return html.escape(str(s))


# ---------- gaya grafik gelap ----------

def dark(fig: go.Figure, height: int, legend_top: bool = True) -> go.Figure:
    fig.update_layout(
        height=height, font=dict(family=FONT, size=12, color=TEXT_2),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", separators=",.",
        margin=dict(l=8, r=8, t=46 if legend_top else 16, b=8),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=1.04, yanchor="bottom",
                    font=dict(size=11, color=TEXT_2), itemsizing="constant", bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="#0f2350", bordercolor="#3b9cff", font=dict(family=FONT, color=TEXT)),
        bargap=0.28, barcornerradius=4,
    )
    fig.update_xaxes(gridcolor=GRID, linecolor="rgba(148,170,220,.3)", zeroline=False, tickfont=dict(color=TEXT_2))
    fig.update_yaxes(gridcolor=GRID, linecolor="rgba(0,0,0,0)", zeroline=False, tickfont=dict(color=MUTED))
    return fig


def wrap(label, width: int = 14) -> str:
    words, lines, cur = str(label).split(), [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > width:
            lines.append(cur)
            cur = w
        else:
            cur = f"{cur} {w}".strip()
    lines.append(cur)
    return "<br>".join(lines[:3]) + ("…" if len(lines) > 3 else "")


def bar_text(values, fmt_fn=fmt):
    return [fmt_fn(v) if pd.notna(v) else "" for v in values]


def summary_chart(t: pd.DataFrame) -> go.Figure:
    x, n = t.index.tolist(), len(t)
    fig = go.Figure()
    for key in ["A", "B", "C", "D"]:
        name, color = SERIES[key]
        fig.add_trace(go.Bar(
            x=x, y=t[key], name=name, marker_color=color, text=bar_text(t[key]),
            textposition="outside", textangle=-90 if n > 6 else 0, textfont=dict(size=10, color=TEXT_2),
            cliponaxis=False, hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>"))
    name, color = SERIES["E"]
    fig.add_trace(go.Scatter(
        x=x, y=t["E"], name=name, yaxis="y2", mode="lines+markers+text", text=bar_text(t["E"]),
        textposition="top center", textfont=dict(size=11, color=color), cliponaxis=False,
        line=dict(color=color, width=2.5), marker=dict(size=8, color=color, line=dict(color="#0c1c42", width=2)),
        hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>"))
    bar_max = float(t[["A", "B", "C", "D"]].to_numpy().max() or 1)
    e_max = float(t["E"].max() or 1)
    step = max(1, round(e_max / 2))
    dark(fig, 520)
    fig.update_layout(
        barmode="group", bargroupgap=0.06, hovermode="x unified", margin=dict(t=90),
        legend=dict(x=0, xanchor="left", y=1.02, entrywidth=270, entrywidthmode="pixels"),
        yaxis=dict(range=[0, bar_max * 1.9], tickformat=",.0f", title=dict(text="Jumlah TT", font=dict(color=MUTED))),
        yaxis2=dict(overlaying="y", side="right", range=[-e_max * 2.1, e_max * 1.15], showgrid=False,
                    tickmode="array", tickvals=[0, step, step * 2] if step * 2 <= e_max * 1.15 else [0, step],
                    tickformat=",.0f", tickfont=dict(color=RED), zeroline=False,
                    title=dict(text="Sisa TT open (E)", font=dict(color=RED))),
        xaxis=dict(tickangle=-30 if n > 8 else 0, showgrid=False))
    return fig


def progress_chart(t: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for col, color, pos in [("Progress (%)", GREEN, "inside"), ("Remaining (%)", RED, "outside")]:
        fig.add_trace(go.Bar(x=t.index, y=t[col], name=col.replace(" (%)", ""), marker_color=color,
                             text=bar_text(t[col], pct), textposition=pos, insidetextanchor="middle",
                             textangle=0, textfont=dict(size=10, color="#fff" if pos == "inside" else TEXT_2),
                             hovertemplate=f"<b>%{{x}}</b><br>{col}: %{{y:.1f}}%<extra></extra>"))
    dark(fig, 300)
    fig.update_layout(barmode="stack", uniformtext=dict(minsize=8, mode="hide"),
                      yaxis=dict(range=[0, 112], ticksuffix="%"), xaxis=dict(tickangle=-30, showgrid=False))
    return fig


def open_close_region(d: pd.DataFrame) -> go.Figure:
    ct = pd.crosstab(d["_region"], d["_status"]).reindex(columns=["Closed", "Open"], fill_value=0)
    ct = ct.loc[ct.sum(axis=1).sort_values(ascending=False).index]
    fig = go.Figure()
    for col, color, label in [("Closed", BLUE, "CLOSE"), ("Open", PINK, "OPEN")]:
        fig.add_trace(go.Bar(x=ct.index, y=ct[col], name=label, marker_color=color, text=bar_text(ct[col]),
                             textposition="outside", textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
                             hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}}<extra></extra>"))
    dark(fig, 360)
    fig.update_layout(barmode="group", yaxis=dict(range=[0, (ct.to_numpy().max() or 1) * 1.18]))
    return fig


def mttr_by(d: pd.DataFrame, col: str, top: int = 12, height: int = 360) -> go.Figure:
    closed = d[d["_status"].eq("Closed")]
    g = closed.groupby(col)["_aging"].agg(["mean", "size"]).sort_values("size", ascending=False).head(top)
    g = g.sort_values("mean", ascending=False)
    avg = closed["_aging"].mean()
    fig = go.Figure(go.Bar(
        x=[wrap(i) for i in g.index], y=g["mean"], marker_color=[RED if v > avg else GREEN for v in g["mean"]],
        text=[f"{days(v)} d" for v in g["mean"]], textposition="outside", textfont=dict(size=11, color=TEXT_2),
        cliponaxis=False, customdata=g["size"],
        hovertemplate="<b>%{x}</b><br>MTTR: %{y:.2f} hari<br>%{customdata:,} tiket close<extra></extra>"))
    if pd.notna(avg):
        fig.add_hline(y=avg, line=dict(color=YELLOW, dash="dash", width=1.5),
                      annotation_text=f"Rata-rata {days(avg)} d", annotation_position="top right",
                      annotation_font=dict(color=YELLOW, size=11))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=False, margin=dict(t=26),
                      yaxis=dict(range=[0, (g["mean"].max() if len(g) else 1) * 1.22], title="hari"),
                      xaxis=dict(tickangle=-30 if len(g) > 8 else 0, tickfont=dict(size=10)))
    return fig


def ranked(d: pd.DataFrame, col: str, top: int = 10, color: str = TEAL, height: int = 360) -> go.Figure:
    vc = d[col].value_counts().head(top)[::-1]
    share = vc / max(len(d), 1) * 100
    fig = go.Figure(go.Bar(
        y=vc.index, x=vc.values, orientation="h", marker_color=color,
        text=[f"{fmt(v)}  ({pct(s)})" for v, s in zip(vc.values, share)], textposition="outside",
        textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:,} TT<extra></extra>"))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=False, margin=dict(t=16, r=70),
                      xaxis=dict(range=[0, (vc.max() if len(vc) else 1) * 1.3], showgrid=True),
                      yaxis=dict(automargin=True))
    return fig


def pareto(d: pd.DataFrame, col: str, top: int = 15) -> go.Figure:
    vc = d[col].value_counts()
    cum = vc.cumsum() / max(vc.sum(), 1) * 100
    vc, cum = vc.head(top), cum.head(top)
    fig = go.Figure()
    labels = [wrap(i) for i in vc.index]
    fig.add_trace(go.Bar(x=labels, y=vc.values, name="Jumlah TT", marker_color=PURPLE, text=bar_text(vc.values),
                         textposition="outside", textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
                         hovertemplate="<b>%{x}</b><br>%{y:,} TT<extra></extra>"))
    fig.add_trace(go.Scatter(x=labels, y=cum.values, name="Kumulatif %", yaxis="y2",
                             mode="lines+markers+text", text=bar_text(cum.values, lambda v: pct(v, 0)),
                             textposition="top center", textfont=dict(size=10, color=YELLOW), cliponaxis=False,
                             line=dict(color=YELLOW, width=2), marker=dict(size=7),
                             hovertemplate="<b>%{x}</b><br>Kumulatif: %{y:.1f}%<extra></extra>"))
    dark(fig, 400)
    fig.update_layout(yaxis=dict(range=[0, (vc.max() if len(vc) else 1) * 1.6]),
                      yaxis2=dict(overlaying="y", side="right", range=[-60, 115], showgrid=False,
                                  tickvals=[0, 50, 100], ticksuffix="%", tickfont=dict(color=YELLOW)),
                      xaxis=dict(tickangle=-30 if len(vc) > 8 else 0, tickfont=dict(size=10)))
    return fig


def monthly_status(d: pd.DataFrame) -> go.Figure:
    ct = pd.crosstab(d["_period"], d["_status"]).reindex(columns=["Closed", "Open"], fill_value=0).sort_index()
    x = [p.strftime("%b_%Y") for p in ct.index]
    fig = go.Figure()
    for col, color, label in [("Closed", BLUE, "CLOSE"), ("Open", PINK, "OPEN")]:
        fig.add_trace(go.Bar(x=x, y=ct[col], name=label, marker_color=color,
                             text=bar_text(ct[col].where(ct[col] > 0)), textposition="inside",
                             insidetextanchor="middle", textangle=0, textfont=dict(size=10, color="#fff"),
                             hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}}<extra></extra>"))
    total = ct.sum(axis=1)
    fig.add_trace(go.Scatter(x=x, y=total, mode="text", text=bar_text(total), textposition="top center",
                             textfont=dict(size=11, color=TEXT), showlegend=False, hoverinfo="skip"))
    dark(fig, 360)
    fig.update_layout(barmode="stack", uniformtext=dict(minsize=8, mode="hide"),
                      yaxis=dict(range=[0, (total.max() if len(total) else 1) * 1.15]),
                      xaxis=dict(tickangle=-30 if len(x) > 8 else 0, showgrid=False))
    return fig


def chart_card(key: str, badge: str, color: str, title: str, fig: go.Figure) -> None:
    with st.container(key=f"card_{key}"):
        st.markdown(f"<div class='ctitle'><span class='badge' style='--c:{color}'>{esc(badge)}</span>"
                    f"{esc(title)}</div>", unsafe_allow_html=True)
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "displayModeBar": False},
                        key=f"fig_{key}")


# ---------- sumber data ----------

files = sorted([p for p in APP_DIR.glob("*.xlsb") if not p.name.startswith("~$")],
               key=lambda p: p.stat().st_mtime, reverse=True)
FILTER_KEYS = ("f_year", "f_region", "f_tlp", "f_cause", "f_status", "f_month", "f_sev", "f_backlog", "f_q")

head_l, head_r = st.columns([1.15, 1], vertical_alignment="center")
with head_r:
    b1, b2, b3, b4 = st.columns([1.55, 1.05, 1, 1])
    with b1.popover("📂 Update dari File Excel", width="stretch"):
        upload = st.file_uploader("Unggah file .xlsb", type=["xlsb"], key="upload")
        file_pick = None
        if files:
            labels = {f"{p.name} · {datetime.fromtimestamp(p.stat().st_mtime):%d-%m-%Y %H:%M}": p for p in files}
            file_pick = labels[st.selectbox("Atau pilih file di folder aplikasi", list(labels))]
        if st.button("🔄 Muat ulang data", width="stretch"):
            st.cache_data.clear()
            st.rerun()
    with b2.popover("📘 Cara Update", width="stretch"):
        st.markdown(
            f"1. Simpan file `.xlsb` terbaru ke folder **{APP_DIR}** (atau unggah lewat *Update dari File Excel*).\n"
            "2. Pastikan sheet **List Access Issue**, **List Stolen**, dan **TT CME** ada, header di baris pertama.\n"
            "3. Klik **Muat ulang data**. File terbaru otomatis dipilih.\n\n"
            "Grafik A–E: TT berstatus *Cancel* tidak dihitung (sama dengan pivot Summary).")
    export_slot = b3.empty()
    if b4.button("↻ Reset Filter", width="stretch"):
        for k in [k for k in st.session_state if str(k).startswith(FILTER_KEYS) or k == "ae_range"]:
            st.session_state.pop(k, None)
        st.rerun()

if upload is not None:
    source, source_name, cache_key = upload.getvalue(), upload.name, f"upload:{upload.name}:{upload.size}"
    mtime = datetime.now()
elif file_pick is not None:
    source, source_name, cache_key = str(file_pick), file_pick.name, f"{file_pick}:{file_pick.stat().st_mtime}"
    mtime = datetime.fromtimestamp(file_pick.stat().st_mtime)
else:
    source = None

with head_l:
    pill = f"<span class='pill'>{esc(source_name)}</span>" if source else ""
    sub = (f"Sumber: <b>{esc(source_name)}</b> • Diperbarui: {mtime:%Y-%m-%d %H:%M:%S} • Dibaca dengan pyxlsb"
           if source else "Belum ada file data")
    st.markdown(f"<div class='hdr'><div class='logo'>TLP</div><div><h1>TLP Escalation Dashboard{pill}</h1>"
                f"<div class='sub'>{sub}</div></div></div>", unsafe_allow_html=True)

if source is None:
    st.info(f"Belum ada file .xlsb. Letakkan file di folder **{APP_DIR}** atau klik **Update dari File Excel**.")
    st.stop()

try:
    sheets = load_all(cache_key, source)
except Exception as exc:
    st.error(f"Gagal membaca **{source_name}**: {exc}")
    st.stop()
if not sheets:
    st.error("Sheet List Access Issue / List Stolen / TT CME tidak ditemukan di file ini.")
    st.stop()


# ---------- bar filter horizontal (Menggunakan Multiselect / Ceklis) ----------

def opts(series: pd.Series, label: str, sort_by_count: bool = True) -> list[str]:
    vals = series.value_counts().index.tolist() if sort_by_count else sorted(series.dropna().unique())
    return [str(v) for v in vals]


with st.container(key="filterbar"):
    r1 = st.columns([1.15, 1.05, 1.15, 1.15, 1.25, 1.3, 1.1, 1.1, 1.35])
    cat = r1[0].selectbox("Kategori", [c for c in CATEGORIES if c in sheets],
                          format_func=lambda c: f"{CATEGORIES[c]['icon']} {c}", key="f_cat")
    cfg = CATEGORIES[cat]
    base = prepare(f"{cache_key}:{cat}", sheets[cat], cat)
    
    if "backlog" in cfg:
        bl_opts = opts(base["_backlog"], "Backlog")
        f_backlog = r1[7].multiselect("Backlog", bl_opts, default=[], key=f"f_backlog_{cat}")
    else:
        f_backlog = []

    f_year = r1[1].multiselect("Years", sorted(opts(base["_year"], "Tahun"), reverse=True), key=f"f_year_{cat}")
    f_region = r1[2].multiselect("Region", opts(base["_region"], "Region"), key=f"f_region_{cat}")
    f_tlp = r1[3].multiselect("TLP (ToCo)", opts(base["_tlp"], "TLP"), key=f"f_tlp_{cat}")
    f_cause = r1[4].multiselect(cfg["cause_label"], opts(base["_cause"], "Kategori"), key=f"f_cause_{cat}")
    
    status_opts = ["Open", "Closed", "Cancel"]
    f_status = r1[5].multiselect("Status", status_opts, key=f"f_status_{cat}",
                               help="Pilih status yang ingin ditampilkan.")
                               
    months_all = sorted(base["_period"].unique())
    f_month = r1[6].multiselect("Bulan (Open)", [p.strftime("%b_%Y") for p in months_all[::-1]],
                              key=f"f_month_{cat}")
                              
    if not cfg.get("backlog"):
        f_sev = r1[7].multiselect("Severity", opts(base["_sev"], "Severity"), key=f"f_sev_{cat}")
    else:
        f_sev = []
        
    f_q = r1[8].text_input("Cari Site / Issue", placeholder="Ketik Site ID, nama, TT, deskripsi…", key=f"f_q_{cat}")

# Logika penyaringan data (jika multiselect kosong, maka semua data ditampilkan)
def apply_multiselect(data, col, selected):
    if selected:
        return data[data[col].isin(selected)]
    return data

d = base
d = apply_multiselect(d, "_year", f_year)
d = apply_multiselect(d, "_region", f_region)
d = apply_multiselect(d, "_tlp", f_tlp)
d = apply_multiselect(d, "_cause", f_cause)
d = apply_multiselect(d, "_sev", f_sev)
d = apply_multiselect(d, "_backlog", f_backlog)
d = apply_multiselect(d, "_month", f_month)

if f_q.strip():
    d = d[d["_search"].str.contains(re.escape(f_q.strip().lower()), regex=True, na=False)]

dim = d[d["_status"] != "Cancel"]  # dasar grafik A–E (Cancel tidak dihitung)

if f_status:
    d = d[d["_status"].isin(f_status)]

export_cols = [c for c in d.columns if not c.startswith("_")]
export_slot.download_button("⬇️ Export CSV", d[export_cols].to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"TLP_{cat.replace(' ', '_')}_{datetime.now():%Y%m%d_%H%M}.csv",
                            mime="text/csv", width="stretch")

# ---------- kartu KPI ----------

total = len(d)
closed = d[d["_status"] == "Closed"]
open_ = d[d["_status"] == "Open"]
close_rate = len(closed) / total * 100 if total else float("nan")
mttr = closed["_aging"].mean()
tlp_vc = d["_tlp"].value_counts()
cause_vc = d["_cause"].value_counts()


def kpi(label, value, sub, color, text_value=False):
    cls = "val txt" if text_value else "val"
    return (f"<div class='kpi' style='--c:{color}'><div class='lbl'>{esc(label)}</div>"
            f"<div class='{cls}' title='{esc(value)}'>{esc(value)}</div><div class='sub'>{sub}</div></div>")


top_tlp = (tlp_vc.index[0], tlp_vc.iloc[0]) if len(tlp_vc) else ("-", 0)
top_cause = (cause_vc.index[0], cause_vc.iloc[0]) if len(cause_vc) else ("-", 0)
st.markdown("<div class='kpis'>" + "".join([
    kpi(f"Total TT {cat}", fmt(total), f"Dari {fmt(len(base))} total record", BLUE),
    kpi("Ticket Closed", fmt(len(closed)), f"Close Rate: {pct(close_rate, 2)}", CYAN),
    kpi("Ticket Masih Open", fmt(len(open_)), f"Avg Aging: {days(open_['_aging'].mean())} hari", PINK),
    kpi("Overall MTTR", f"{days(mttr)} d", f"Dari {fmt(len(closed))} tiket CLOSE", YELLOW),
    kpi("Top TLP Volume", top_tlp[0], f"{fmt(top_tlp[1])} TT ({pct(top_tlp[1] / max(total, 1) * 100)})",
        PURPLE, True),
    kpi(f"Top {cfg['cause_label']}", top_cause[0],
        f"{fmt(top_cause[1])} TT ({pct(top_cause[1] / max(total, 1) * 100)})", TEAL, True),
]) + "</div>", unsafe_allow_html=True)

if d.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini. Klik **Reset Filter** untuk mengembalikan.")
    st.stop()


# ---------- executive insight ----------

def insights(d: pd.DataFrame) -> list[tuple[str, str, str]]:
    n = max(len(d), 1)
    reg = d["_region"].value_counts()
    r_top = reg.head(2)
    open_tlp = d[d["_status"] == "Open"]["_tlp"].value_counts().head(2)
    open_txt = (" (terbanyak di " + " & ".join(f"<b>{esc(k)}</b> {fmt(v)} TT" for k, v in open_tlp.items()) + ")"
                if len(open_tlp) else "")
    p1 = (" & ".join(f"Region <b>{esc(k)} ({fmt(v)} TT)</b>" for k, v in r_top.items())
          + f" menyumbang <b>{pct(r_top.sum() / n * 100)}</b> dari total tiket {esc(cat)}. "
          f"Saat ini tersisa <b>{fmt((d['_status'] == 'Open').sum())} tiket OPEN</b>{open_txt}.")

    cl = d[d["_status"] == "Closed"]
    g = cl.groupby("_tlp")["_aging"].agg(["mean", "size"]).sort_values("size", ascending=False).head(4)
    avg = cl["_aging"].mean()
    if len(g):
        fast = g["mean"].idxmin()
        slow = g[g["mean"] > avg].sort_values("mean", ascending=False)
        slow_txt = (", sementara " + ", ".join(f"<b>{esc(k)} ({days(v)}d)</b>" for k, v in slow["mean"].items())
                    + " berada di atas rata-rata.") if len(slow) else "; semua TLP utama di bawah rata-rata."
        p2 = (f"Grand average MTTR berada di <b>{days(avg)} hari</b>. Di antara {len(g)} TLP utama, "
              f"<b>{esc(fast)}</b> mencatat MTTR tercepat (<b>{days(g.loc[fast, 'mean'])} hari</b>){slow_txt}")
    else:
        p2 = "Belum ada tiket CLOSE pada filter ini."

    tv = d["_tlp"].value_counts().head(4)
    p3 = (f"<b>{len(tv)} TLP teratas</b> (" + ", ".join(f"{esc(k)} {fmt(v)}" for k, v in tv.items())
          + f") mendominasi <b>{pct(tv.sum() / n * 100)}</b> dari seluruh tiket {esc(cat)}.")

    cv = d["_cause"].value_counts().head(3)
    p4 = (f"{len(cv)} isu teratas — " + ", ".join(f"<b>{esc(k)} ({fmt(v)} TT)</b>" for k, v in cv.items())
          + f" — menyumbang <b>{pct(cv.sum() / n * 100)}</b> dari seluruh tiket.")
    return [("📍 Regional & TLP Open/Close", p1, BLUE), ("⏱️ Performa MTTR per TLP", p2, YELLOW),
            ("🏢 Konsentrasi TLP (Pareto)", p3, CYAN), (f"🔎 {cfg['cause_label']} Dominan", p4, PINK)]


# ---------- tab navigasi ----------

tabs = st.tabs(["📊 Executive Summary", "📈 Trend A–E", "⏱️ MTTR per TLP", "🏢 Jumlah TT per TLP",
                f"🔎 {cfg['cause_label']}", f"🗂️ Data Explorer ({fmt(len(d))} TT)"])

with tabs[0]:
    st.markdown("<div class='insights'>" + "".join(
        f"<div class='ins' style='--c:{c}'><h4>{esc(t)}</h4><p>{body}</p></div>" for t, body, c in insights(d))
        + "</div>", unsafe_allow_html=True)
    a, b = st.columns(2)
    with a:
        chart_card("oc_region", "PVT-1", BLUE, "Open vs Close Ticket per Region", open_close_region(d))
    with b:
        chart_card("mttr_tlp", "PVT-2", YELLOW, "Average MTTR (Days) per TLP · Status CLOSE", mttr_by(d, "_tlp"))
    a, b = st.columns(2)
    with a:
        chart_card("monthly", "PVT-3", CYAN, "Jumlah TT per Bulan Open (Close/Open)", monthly_status(d))
    with b:
        chart_card("cause", "PVT-4", PINK, f"Top {cfg['cause_label']}", ranked(d, "_cause", 10, PINK))

with tabs[1]:
    months_sel = sorted(dim["_period"].unique())
    labels_m = [p.strftime("%b_%Y") for p in months_sel]
    if labels_m:
        m_from, m_to = st.select_slider("Periode grafik A–E (bulan open)", options=labels_m,
                                        value=(labels_m[max(0, len(labels_m) - 13)], labels_m[-1]), key="ae_range")
        win = months_sel[labels_m.index(m_from): labels_m.index(m_to) + 1]
        t = backlog_summary(dim[dim["_period"].between(win[0], win[-1])], win)
        chart_card("ae", "A–E", BLUE, f"Trend TT {cat} · {m_from} – {m_to}", summary_chart(t))
        chart_card("progress", "%", GREEN, f"Progress & Remaining TT {cat}", progress_chart(t))
        with st.expander("Tabel summary A–E"):
            st.dataframe(t.rename(columns={k: v[0] for k, v in SERIES.items()}), width="stretch",
                         column_config={"Progress (%)": st.column_config.NumberColumn(format="%.1f%%"),
                                        "Remaining (%)": st.column_config.NumberColumn(format="%.1f%%")})
        st.caption("Grafik A–E mengikuti filter di atas kecuali Status & Bulan; TT Cancel tidak dihitung.")

with tabs[2]:
    chart_card("mttr_tlp2", "PVT-2", YELLOW, "Average MTTR (Days) per TLP", mttr_by(d, "_tlp", 15, 400))
    a, b = st.columns(2)
    with a:
        chart_card("mttr_region", "MTTR", ORANGE, "Average MTTR (Days) per Region", mttr_by(d, "_region"))
    with b:
        chart_card("mttr_cause", "MTTR", PINK, f"Average MTTR (Days) per {cfg['cause_label']}",
                   mttr_by(d, "_cause"))

with tabs[3]:
    chart_card("pareto", "PVT-3", PURPLE, "Jumlah TT per TLP (Pareto)", pareto(d, "_tlp"))
    a, b = st.columns(2)
    with a:
        chart_card("tlp_open", "OPEN", PINK, "TT masih OPEN per TLP",
                   ranked(d[d["_status"] == "Open"], "_tlp", 10, PINK))
    with b:
        chart_card("tlp_sev", "SEV", CYAN, "Jumlah TT per Severity", ranked(d, "_sev", 8, CYAN, 300))

with tabs[4]:
    a, b = st.columns(2)
    with a:
        chart_card("cause2", "PVT-4", TEAL, f"Top {cfg['cause_label']}", ranked(d, "_cause", 12, TEAL, 420))
    with b:
        chart_card("cause_pareto", "PVT-4", PURPLE, f"Pareto {cfg['cause_label']}", pareto(d, "_cause", 12))
    if cfg.get("cause_rules"):
        st.caption(f"{cfg['cause_label']} dikelompokkan otomatis dari kolom **{cfg['cause']}** (teks bebas) "
                   "berdasarkan kata kunci; yang tidak cocok masuk *Lainnya*.")

with tabs[5]:
    default_cols = [c for c in (cfg["site"] + [REGION_COL, cfg["tlp"], cfg["status"], cfg["month"], cfg["aging"],
                                               cfg["severity"]]) if c in d.columns]
    cols = st.multiselect("Kolom ditampilkan", export_cols, default=list(dict.fromkeys(default_cols)))
    shown = d[cols].copy()
    for c in shown.columns:
        if shown[c].dtype == object:
            shown[c] = shown[c].astype("string")
    st.dataframe(shown, hide_index=True, width="stretch", height=560)
