"""Dashboard Summary TLP — Access Issue, Stolen, dan TT CME.

Membaca file .xlsb (sheet `List Access Issue`, `List Stolen`, `TT CME`) dengan pyxlsb, lalu
menampilkan grafik summary A–E per bulan seperti pivot di sheet Summary:
  A = akumulasi TT open bulan sebelumnya      B = TT baru per bulan
  C = A + B (akumulasi TT open per bulan)     D = TT close per bulan (C − E)
  E = sisa TT yang masih open s/d bulan tsb   Progress = D / C

Jalankan:  streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pyxlsb import open_workbook

APP_DIR = Path(__file__).resolve().parent

# Konfigurasi tiap summary: sheet, kolom bulan open, kolom status, kolom owner/grup, filter tetap.
SUMMARIES = {
    "access": dict(title="Access Issue", sheet="List Access Issue", month="Month Open",
                   status="Case Status (Open/Closed)", owner="GROUP", icon="🚧"),
    "stolen": dict(title="Stolen", sheet="List Stolen", month="Month Open",
                   status="Status TT ToCo/OpCo", owner="Toco Group Name", icon="🔓"),
    "cme": dict(title="TT CME", sheet="TT CME", month="Create Month",
                status="STATUS TT TLP", owner="Owner", backlog="Backlog", icon="🛠️"),
}
REGION_COL = "New Region"

SERIES = {
    "A": ("A: Akumulasi TT open bulan sebelumnya", "#898781"),
    "B": ("B: TT baru per bulan", "#2a78d6"),
    "C": ("C: Akumulasi TT open per bulan", "#eda100"),
    "D": ("D: TT close per bulan", "#1baf7a"),
    "E": ("E: Sisa TT masih open [C−D]", "#d03b3b"),
}
INK, INK_2, MUTED, GRID = "#0b0b0b", "#52514e", "#898781", "#e1e0d9"
FONT = 'system-ui, -apple-system, "Segoe UI", sans-serif'

st.set_page_config(page_title="Summary TLP", page_icon="📊", layout="wide")
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem; padding-bottom: 2rem;}
      [data-testid="stMetric"] {background: #f4f3ef; border-radius: 10px; padding: 12px 16px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------- membaca .xlsb dengan pyxlsb ----------

def _clean_name(v, i: int, seen: dict) -> str:
    name = " ".join(str(v).split()) if v not in (None, "") else f"Kolom_{i + 1}"
    if name in seen:  # nama kolom ganda diberi akhiran agar unik
        seen[name] += 1
        name = f"{name}_{seen[name]}"
    else:
        seen[name] = 0
    return name


MAX_EMPTY_RUN = 500  # berhenti setelah 500 baris kosong berturut-turut


def read_sheet(wb, sheet: str) -> pd.DataFrame:
    """Baca satu sheet .xlsb dari workbook pyxlsb yang sudah dibuka; baris pertama = header.

    Sebagian sheet (mis. List Stolen) berformat sampai ~1 juta baris walau datanya hanya
    belasan ribu; pembacaan dihentikan setelah deretan baris kosong agar tetap cepat.
    """
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
    df = pd.DataFrame(body, columns=header)
    return df.dropna(how="all")


@st.cache_data(show_spinner="Membaca file .xlsb (pyxlsb)…", max_entries=4)
def load_all(key: str, _source: str | bytes) -> dict[str, pd.DataFrame]:
    # `key` (path + waktu ubah / nama unggahan) menjadi kunci cache; isi file tidak di-hash.
    # Workbook dibuka sekali saja: membuka ulang = membaca ulang shared strings (lambat).
    src = io.BytesIO(_source) if isinstance(_source, bytes) else _source
    with open_workbook(src) as wb:
        return {k: read_sheet(wb, cfg["sheet"]) for k, cfg in SUMMARIES.items() if cfg["sheet"] in wb.sheets}


# ---------- perhitungan summary A–E ----------

def to_period(s: pd.Series) -> pd.Series:
    """'Oct_2025' / tanggal / serial Excel -> Period bulanan."""
    text = s.astype("string").str.strip()
    out = pd.to_datetime(text, format="%b_%Y", errors="coerce")
    num = pd.to_numeric(s, errors="coerce")
    serial = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    out = out.fillna(serial.where(num > 20000))
    out = out.fillna(pd.to_datetime(text.where(out.isna()), errors="coerce"))
    return out.dt.to_period("M")


def prepare(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    df = df.copy()
    df["_period"] = to_period(df[cfg["month"]]) if cfg["month"] in df.columns else pd.NaT
    df["_status"] = df[cfg["status"]].astype("string").str.strip().str.lower() if cfg["status"] in df.columns else ""
    for col in [REGION_COL, cfg["owner"], cfg.get("backlog")]:
        if col and col in df.columns:
            df[col] = df[col].astype("string").str.strip().fillna("(Kosong)")
    return df[df["_period"].notna()]


def backlog_summary(df: pd.DataFrame, months: pd.PeriodIndex) -> pd.DataFrame:
    """Tabel A–E per bulan. TT sebelum bulan pertama jendela tidak dihitung (A awal = 0)."""
    in_window = df["_period"].between(months[0], months[-1])
    df = df[in_window]
    new = df.groupby("_period").size().reindex(months, fill_value=0)
    still_open = df[df["_status"].eq("open")].groupby("_period").size().reindex(months, fill_value=0)
    e = still_open.cumsum()
    a = e.shift(1, fill_value=0)
    c = a + new
    d = c - e
    out = pd.DataFrame({"A": a, "B": new, "C": c, "D": d, "E": e})
    out["Progress (%)"] = (out["D"] / out["C"].where(out["C"] > 0) * 100).round(1)
    out["Remaining (%)"] = (out["E"] / out["C"].where(out["C"] > 0) * 100).round(1)
    out.index = [p.strftime("%b_%Y") for p in months]
    out.index.name = "Bulan"
    return out


# ---------- grafik ----------

def summary_chart(t: pd.DataFrame, title: str) -> go.Figure:
    x = t.index.tolist()
    n = len(x)
    fig = go.Figure()
    for key in ["A", "B", "C", "D"]:
        name, color = SERIES[key]
        fig.add_trace(go.Bar(
            x=x, y=t[key], name=name, marker_color=color, offsetgroup=key,
            text=[f"{v:,.0f}".replace(",", ".") for v in t[key]],
            # Label vertikal di ujung bar: 4 bar per bulan tidak cukup lebar untuk label mendatar.
            textposition="outside", textangle=-90 if n > 6 else 0,
            textfont=dict(size=10, color=INK_2), cliponaxis=False,
            hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>",
        ))
    name, color = SERIES["E"]
    fig.add_trace(go.Scatter(
        x=x, y=t["E"], name=name, yaxis="y2", mode="lines+markers+text",
        line=dict(color=color, width=2.5), marker=dict(size=8, color=color, line=dict(color="#ffffff", width=2)),
        text=[f"{v:,.0f}".replace(",", ".") for v in t["E"]], textposition="top center",
        textfont=dict(size=11, color=color), cliponaxis=False,
        hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>",
    ))
    bar_max = float(t[["A", "B", "C", "D"]].to_numpy().max() or 1)
    e_max = float(t["E"].max() or 1)
    # Bar menempati ~55% bawah area plot, garis E ~35% atas: label keduanya tidak bertabrakan.
    # Sumbu kanan sengaja dimulai dari negatif; tick hanya ditampilkan untuk nilai ≥ 0.
    step = max(1, round(e_max / 2))
    e_ticks = [0, step, step * 2] if step * 2 <= e_max * 1.15 else [0, step]
    fig.update_layout(
        title=dict(text=title, x=0, xanchor="left", y=0.98, yanchor="top", font=dict(size=16, color=INK)),
        height=540,
        barmode="group", bargap=0.22, bargroupgap=0.06,
        font=dict(family=FONT, size=12, color=INK_2),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        separators=",.",
        # Legend mendatar di bawah judul, dengan jarak antar item, terpisah dari area plot.
        legend=dict(orientation="h", x=0, xanchor="left", y=1.02, yanchor="bottom",
                    font=dict(size=11), itemsizing="constant", entrywidth=260, entrywidthmode="pixels",
                    traceorder="normal", bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=10, r=10, t=120, b=10),
        hovermode="x unified",
        hoverlabel=dict(bgcolor="#ffffff", bordercolor="#c3c2b7", font=dict(family=FONT, color=INK)),
        yaxis=dict(title=dict(text="Jumlah TT", font=dict(size=11, color=MUTED)), range=[0, bar_max * 1.9],
                   gridcolor=GRID, tickformat=",.0f", tickfont=dict(color=MUTED), zeroline=False),
        yaxis2=dict(title=dict(text="Sisa TT open (E)", font=dict(size=11, color=SERIES["E"][1])),
                    overlaying="y", side="right", range=[-e_max * 2.1, e_max * 1.15], showgrid=False,
                    tickmode="array", tickvals=e_ticks, tickformat=",.0f",
                    tickfont=dict(color=SERIES["E"][1]), zeroline=False),
        xaxis=dict(tickangle=-30 if n > 8 else 0, tickfont=dict(color=INK_2), showgrid=False, linecolor="#c3c2b7"),
    )
    return fig


def progress_chart(t: pd.DataFrame, title: str) -> go.Figure:
    x = t.index.tolist()
    fig = go.Figure()
    for col, color in [("Progress (%)", "#1baf7a"), ("Remaining (%)", "#d03b3b")]:
        fig.add_trace(go.Bar(
            x=x, y=t[col], name=col.replace(" (%)", ""), marker_color=color,
            text=[f"{v:.1f}%".replace(".", ",") if pd.notna(v) else "" for v in t[col]],
            textposition="inside" if col.startswith("Progress") else "outside",
            insidetextanchor="middle", textangle=0, textfont=dict(size=10),
            hovertemplate=f"<b>%{{x}}</b><br>{col}: %{{y:.1f}}%<extra></extra>",
        ))
    fig.update_layout(
        title=dict(text=title, x=0, xanchor="left", font=dict(size=14, color=INK)),
        barmode="stack", height=300, bargap=0.3,
        font=dict(family=FONT, size=11, color=INK_2), separators=",.",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", x=1, xanchor="right", y=1.02, yanchor="bottom", entrywidth=110,
                    entrywidthmode="pixels"),
        margin=dict(l=10, r=10, t=60, b=10),
        uniformtext=dict(minsize=8, mode="hide"),
        yaxis=dict(range=[0, 108], ticksuffix="%", gridcolor=GRID, tickfont=dict(color=MUTED)),
        xaxis=dict(tickangle=-30 if len(x) > 8 else 0, showgrid=False),
    )
    return fig


def fmt(v) -> str:
    return f"{v:,.0f}".replace(",", ".")


# ---------- sidebar: sumber data ----------

st.sidebar.header("📁 Sumber data")
files = sorted([p for p in APP_DIR.glob("*.xlsb") if not p.name.startswith("~$")],
               key=lambda p: p.stat().st_mtime, reverse=True)
upload = st.sidebar.file_uploader("Atau unggah file .xlsb", type=["xlsb"])

if upload is not None:
    source, source_name, cache_key = upload.getvalue(), upload.name, f"upload:{upload.name}:{upload.size}"
elif files:
    labels = {f"{p.name}  ·  {datetime.fromtimestamp(p.stat().st_mtime):%d-%m-%Y %H:%M}": p for p in files}
    path = labels[st.sidebar.selectbox("File terdeteksi (terbaru di atas)", list(labels))]
    source, source_name, cache_key = str(path), path.name, f"{path}:{path.stat().st_mtime}"
    if st.sidebar.button("🔄 Muat ulang data", width="stretch"):
        st.cache_data.clear()
        st.rerun()
else:
    st.title("📊 Summary TLP")
    st.info(f"Belum ada file .xlsb. Letakkan file di folder **{APP_DIR}** atau unggah lewat sidebar.")
    st.stop()

try:
    sheets = load_all(cache_key, source)
except Exception as exc:  # file rusak / bukan xlsb
    st.error(f"Gagal membaca **{source_name}**: {exc}")
    st.stop()

missing = [cfg["sheet"] for k, cfg in SUMMARIES.items() if k not in sheets]
if missing:
    st.warning(f"Sheet tidak ditemukan: {', '.join(missing)}")
if not sheets:
    st.stop()

data = {k: prepare(df, SUMMARIES[k]) for k, df in sheets.items()}

# ---------- sidebar: filter ----------

st.sidebar.header("🔎 Filter")
all_months = sorted(set().union(*[set(d["_period"].unique()) for d in data.values()]))
labels_m = [p.strftime("%b_%Y") for p in all_months]
default_from = labels_m[max(0, len(labels_m) - 13)]
m_from, m_to = st.sidebar.select_slider("Periode (bulan open)", options=labels_m,
                                        value=(default_from, labels_m[-1]))
months = pd.period_range(all_months[labels_m.index(m_from)], all_months[labels_m.index(m_to)], freq="M")

regions = sorted(set().union(*[set(d[REGION_COL].dropna()) for d in data.values() if REGION_COL in d.columns]))
sel_regions = st.sidebar.multiselect("New Region", regions, placeholder="Semua")
include_cancel = st.sidebar.toggle("Sertakan TT berstatus Cancel", value=False,
                                   help="Pivot di sheet Summary tidak menghitung TT Cancel.")
cme_backlog = None
if "cme" in data and "Backlog" in data["cme"].columns:
    opts = data["cme"]["Backlog"].value_counts().index.tolist()
    default = [o for o in opts if o.lower() == "normal tt"]
    cme_backlog = st.sidebar.multiselect("TT CME · Backlog", opts, default=default, placeholder="Semua")

st.sidebar.caption("Filter owner/grup tersedia di masing-masing grafik.")

# ---------- konten ----------

st.title("📊 Summary TLP — Access Issue, Stolen & TT CME")
st.caption(f"Sumber: **{source_name}** (dibaca dengan pyxlsb) · Periode {m_from} s/d {m_to}"
           + (f" · Region: {', '.join(sel_regions)}" if sel_regions else ""))

with st.expander("Cara membaca grafik A–E"):
    st.markdown(
        "- **A** — akumulasi TT yang masih open dari bulan sebelumnya (= E bulan lalu)\n"
        "- **B** — TT baru yang dibuka pada bulan tersebut\n"
        "- **C** — akumulasi TT open per bulan (A + B)\n"
        "- **D** — TT yang sudah close (C − E)\n"
        "- **E** — sisa TT yang masih open sampai bulan tersebut (garis merah, sumbu kanan)\n\n"
        "TT yang dibuka sebelum bulan awal periode tidak dihitung, sama seperti pivot di sheet Summary."
    )

tabs = st.tabs([f"{SUMMARIES[k]['icon']} {SUMMARIES[k]['title']}" for k in data])
for tab, (key, df) in zip(tabs, data.items()):
    cfg = SUMMARIES[key]
    with tab:
        d = df
        if sel_regions and REGION_COL in d.columns:
            d = d[d[REGION_COL].isin(sel_regions)]
        if not include_cancel:
            d = d[~d["_status"].fillna("").str.startswith("cancel")]
        if key == "cme" and cme_backlog:
            d = d[d["Backlog"].isin(cme_backlog)]
        owner_col = cfg["owner"]
        if owner_col in d.columns:
            owners = d[owner_col].value_counts().index.tolist()
            sel = st.multiselect(f"Owner / grup ({owner_col})", owners, placeholder="Semua owner",
                                 key=f"owner_{key}")
            if sel:
                d = d[d[owner_col].isin(sel)]

        t = backlog_summary(d, months)
        last = t.iloc[-1]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("TT baru (periode)", fmt(t["B"].sum()))
        c2.metric(f"TT close · {t.index[-1]}", fmt(last["D"]))
        c3.metric("Sisa TT open", fmt(last["E"]),
                  delta=f"{fmt(last['E'] - last['A'])} vs bulan lalu" if len(t) > 1 else None,
                  delta_color="inverse")
        c4.metric(f"Progress · {t.index[-1]}", f"{last['Progress (%)']:.1f}%".replace(".", ",")
                  if pd.notna(last["Progress (%)"]) else "-")

        st.plotly_chart(summary_chart(t, f"Trend TT {cfg['title']} · {m_from} – {m_to}"), width="stretch",
                        config={"displaylogo": False}, key=f"chart_{key}")
        st.plotly_chart(progress_chart(t, f"Progress & remaining TT {cfg['title']} (%)"), width="stretch",
                        config={"displayModeBar": False}, key=f"progress_{key}")
        with st.expander("Tabel summary"):
            shown = t.rename(columns={k: v[0] for k, v in SERIES.items()})
            st.dataframe(shown, width="stretch",
                         column_config={"Progress (%)": st.column_config.NumberColumn(format="%.1f%%"),
                                        "Remaining (%)": st.column_config.NumberColumn(format="%.1f%%")})
            st.download_button("⬇️ Unduh tabel (.csv)", shown.to_csv().encode("utf-8-sig"),
                               file_name=f"summary_{key}.csv", mime="text/csv", key=f"dl_{key}")
