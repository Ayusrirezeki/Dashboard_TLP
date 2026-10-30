"""TLP Dashboard — Access Issue, Stolen, dan TT CME.

Fitur:
  * Membaca file .xlsb (sheet `List Access Issue`, `List Stolen`, `TT CME`) dengan pyxlsb.
    Tabel SoW (A) per ToCo diambil dari sheet `performance access` / `Performance Stolen`
    bila ada, dan bisa disunting dari dashboard.
  * Bar filter horizontal (multiselect), 6 kartu KPI, tab navigasi, Executive Insight.
  * Grafik Impact Service & SoW: donut impact per ToCo, pie SoW vs sites impact,
    pie SoW vs sites stolen, dan pie top 5 root cause (header case).
  * Upload data mingguan bersama: file disimpan ke `shared_data/latest_tlp_data.xlsb`,
    log pengunggah di `shared_data/upload_log.json`; semua pengguna ter-refresh otomatis.
  * Export Excel (tabel angka tiap grafik) dan PowerPoint berisi grafik native yang bisa diedit
    (python-pptx, tanpa perlu Google Chrome).

Aturan Impact Service (sama dengan rumus COUNTIFS di Excel):
  Site Status (kolom ke-2) = "Down" dan Case Status ∈ {Open, Closed}.
  Total TT Impact Service Released = jumlah TT; Total Sites Impact Service (B) = jumlah site unik.

Grafik summary A–E per bulan mengikuti pivot sheet Summary:
  A = akumulasi TT open bulan sebelumnya      B = TT baru per bulan
  C = A + B (akumulasi TT open per bulan)     D = TT close per bulan (C − E)
  E = sisa TT yang masih open s/d bulan tsb   Progress = D / C

Jalankan:  streamlit run app.py
"""
from __future__ import annotations

import streamlit as st

# set_page_config harus menjadi perintah Streamlit pertama dan hanya dipanggil sekali.
st.set_page_config(page_title="TLP Dashboard", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")

import html  # noqa: E402
import io  # noqa: E402
import json  # noqa: E402
import os  # noqa: E402
import re  # noqa: E402
import time  # noqa: E402
from datetime import datetime  # noqa: E402
from pathlib import Path  # noqa: E402

import pandas as pd  # noqa: E402
import plotly.graph_objects as go  # noqa: E402
from pyxlsb import open_workbook  # noqa: E402

# =====================================================================================
# Konstanta
# =====================================================================================

APP_DIR = Path(__file__).resolve().parent
SHARED_DIR = APP_DIR / "shared_data"
LATEST_FILE = SHARED_DIR / "latest_tlp_data.xlsb"
LOG_FILE = SHARED_DIR / "upload_log.json"
LOCK_FILE = SHARED_DIR / "upload.lock"
SOW_FILE = SHARED_DIR / "sow_override.json"
LOCK_TIMEOUT_S = 180       # lock unggahan dianggap basi setelah 3 menit
REFRESH_EVERY = "15s"      # interval cek data baru untuk auto-refresh semua pengguna
PREP_VERSION = 7           # naikkan bila kolom turunan di prepare() berubah, agar cache lama tidak dipakai

# Warna khas Indosat di atas latar gelap.
MAGENTA, YELLOW, RED, TOSCA = "#E20074", "#FFD100", "#ED1C24", "#00A3AD"
PURPLE, ORANGE, TOSCA_2, PINK_2 = "#9B3D97", "#F7941D", "#4FD1D9", "#FF6FB5"
PALETTE = [MAGENTA, YELLOW, TOSCA, RED, PURPLE, ORANGE, TOSCA_2, PINK_2]
BG, CARD = "#ffffff", "#ffffff"
GOLD = "#B88A00"           # kuning gelap untuk teks di atas latar putih
TEXT, TEXT_2, MUTED = "#1a1c2e", "#3d4058", "#6b7090"
GRID = "rgba(30,35,70,0.09)"
BORDER, SHADOW = "rgba(20,25,60,.10)", "0 6px 18px rgba(20,25,60,.07)"
FONT = 'Inter, "Segoe UI", system-ui, sans-serif'
GREY, TEAL = "#7F7F7F", "#26B5A7"
# Seri grafik Trends [General] — label & warna mengikuti grafik Excel.
SERIES = {
    "A": ("A: Accumulative of TT Open previous Month", YELLOW),
    "B": ("B: New TT are released per Month", MAGENTA),
    "C": ("C: Accumulative TT Open Per Month", GREY),
    "D": ("D: Achievement TT close per Month", TEAL),
    "E": ("E: The remaining of TT still OPEN per month [C-D]", RED),
}
# Warna bawaan Excel (pie LTOA & Root Cause di Excel memakai palet ini).
EXCEL_COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"]
REGION_ORDER = ["SUM", "KALISU", "JAYA", "JAVA"]

# Kategori dari teks bebas (Stolen: Material Stolen, CME: Description) — urutan = prioritas.
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

# Nama ToCo di tiap sheet berbeda (GROUP1, Toco Group Name, Owner) -> kode singkat seragam.
TOCO_RULES = [
    ("PTI", r"\bpti\b|protelindo|profesional telekom"),
    ("DMT", r"\bdmt\b|mitratel|dayamitra"),
    ("TBG", r"\btbg\b|tower bersama"),
    ("CMI", r"\bcmi\b|epid|menara assetco"),
    ("DHOST", r"dhost"),
    ("BALI TOWER", r"bali tower"),
    ("GIHON", r"gihon"),
    ("PST", r"\bpst\b|persada sokka"),
    ("EBT", r"\bebt\b|era bangun"),
    ("EDOTCO", r"edotco"),
    ("IBSW", r"ibsw|infrastruktur bisnis sejahtera"),
    ("IBS", r"inti bangun sejahtera|\bibs\b"),
    ("KIRANA", r"kirana"),
    ("TRIVIEW", r"triview"),
    ("GAMETRACO", r"gametraco"),
    ("MSN", r"\bmsn\b|menara selular"),
    ("MASSADA", r"massada"),
    ("WINET", r"winet|wireless network"),
    ("MATRINDO", r"matrindo"),
    ("BALCOM", r"balcom"),
    ("STP", r"\bstp\b|solusi tunas"),
    ("IOH", r"^ioh$|indosat"),
]
# SoW (A) cadangan (dari dokumen referensi) bila sheet performance tidak memuat ToCo tersebut.
# Urutan grup ToCo di grafik (mengikuti tabel Total site K1 / ToCO di Excel).
TOCO_ORDER = ["PTI", "DMT", "TBG", "CMI", "DHOST", "BALI TOWER", "GIHON", "EBT", "EDOTCO", "IBSW", "TRIVIEW",
              "KIRANA", "GAMETRACO", "MSN", "MASSADA", "WINET", "MATRINDO", "BALCOM"]
DEFAULT_SOW = {"PTI": 21449, "DMT": 12004, "TBG": 11358, "CMI": 8009, "DHOST": 847, "BALI TOWER": 638,
               "GIHON": 530, "EBT": 269, "EDOTCO": 97, "IBSW": 92, "KIRANA": 46, "TRIVIEW": 42,
               "GAMETRACO": 33, "MSN": 29, "MASSADA": 24, "WINET": 20}

SITE_STATUS = "Site Status (Up/Down/LTOA/ Plan to Relocate)"
CATEGORIES = {
    "Access Issue": dict(
        sheet="List Access Issue", month="Month Open", status="Case Status (Open/Closed)",
        toco=["GROUP1", "GROUP"], owner="Site Owner", cause="Problem Type", root=["Case"],
        severity="Severity", aging="Aging Time (Days)", year="Year Open", site_id="Site ID",
        impact=[f"{SITE_STATUS}_1", SITE_STATUS],
        search=["Site ID", "Site ID IOH", "Site Name", "Remarks from FLM", "Case"], tt="Huawei TT", icon="🚧"),
    "Stolen": dict(
        sheet="List Stolen", month="Month Open", status="Status TT ToCo/OpCo", toco=["Toco Group Name"],
        owner="Toco Owner",
        cause="Material Stolen", cause_rules=STOLEN_RULES, root=[], severity="TT Status [Critical/Major]",
        aging="Aging TT", year="Year Open", site_id="IOH ID", impact=[], tt="TT IOH", count_tt=True,
        search=["IOH ID", "Old Site ID", "Site Name", "TT IOH", "Material Stolen"], icon="🔓"),
    "TT CME": dict(
        sheet="TT CME", month="Create Month", status="STATUS TT TLP", toco=["Owner"], owner="Owner",
        cause="Description", cause_rules=CME_RULES, root=[], severity="Severity", aging="Aging TT CLOSED",
        year=None, site_id="Site ID", impact=[], backlog="Backlog", tt="TT Number", count_tt=True,
        search=["Site ID", "Site Name", "TT Number", "Description"], icon="🛠️"),
}
REGION_COL = "New Region"

# =====================================================================================
# Gaya halaman (tema gelap Indosat)
# =====================================================================================

st.markdown(f"""
<style>
  .stApp {{background:#ffffff;}}
  [data-testid="stHeader"] {{background: transparent;}}
  .block-container {{padding-top: 1.1rem; padding-bottom: 2rem; max-width: 1640px;}}

  .hdr {{display:flex; align-items:center; gap:14px;}}
  .hdr .logo {{width:48px; height:48px; border-radius:14px; display:flex; align-items:center; justify-content:center;
              background:conic-gradient(from 210deg, {MAGENTA}, {RED}, {YELLOW}, {TOSCA}, {MAGENTA});
              color:#fff; font-weight:800; font-size:.78rem; text-shadow:0 1px 3px rgba(0,0,0,.5);
              box-shadow:0 4px 18px rgba(226,0,116,.45);}}
  .hdr h1 {{font-size:1.45rem; margin:0; padding:0; color:{TEXT}; font-weight:800; letter-spacing:-.01em;}}
  .hdr .pill {{display:inline-block; margin-left:10px; padding:2px 10px; border-radius:999px; font-size:.68rem;
              font-weight:700; color:{MAGENTA}; background:rgba(226,0,116,.07); border:1px solid rgba(226,0,116,.45);
              vertical-align:middle;}}
  .hdr .sub {{color:{MUTED}; font-size:.78rem; margin-top:2px;}}
  .lock {{display:inline-flex; align-items:center; gap:8px; padding:6px 12px; border-radius:10px; font-size:.78rem;
          color:{TEXT_2}; background:rgba(0,163,173,.12); border:1px solid rgba(0,163,173,.45);}}
  .lock b {{color:{TEXT};}}
  .lock .dot {{width:8px; height:8px; border-radius:50%; background:{TOSCA}; box-shadow:0 0 8px {TOSCA};
               animation: pulse 1.6s infinite;}}
  @keyframes pulse {{0%,100% {{opacity:1}} 50% {{opacity:.35}}}}

  .st-key-filterbar {{background:{CARD}; border:1px solid rgba(226,0,116,.28) !important; border-radius:16px;
                      padding:10px 16px 4px 16px; box-shadow:{SHADOW};}}
  [data-testid="stWidgetLabel"] p {{text-transform:uppercase; font-size:.68rem !important; font-weight:700;
                                   letter-spacing:.06em; color:{MUTED} !important;}}
  [data-baseweb="select"] > div, .stTextInput input {{background:#ffffff !important;
                                   border-color:rgba(226,0,116,.32) !important;}}
  .st-key-filterbar [data-baseweb="select"] div, .st-key-filterbar input {{font-size:.8rem !important;}}
  .st-key-filterbar [data-baseweb="tag"] {{background:rgba(226,0,116,.85) !important;}}
  .st-key-filterbar [data-testid="stColumn"] {{min-width:0;}}

  .kpis {{display:grid; grid-template-columns:repeat(6, minmax(0,1fr)); gap:14px; margin:14px 0 6px;}}
  .kpi {{background:{CARD}; border:1px solid {BORDER}; border-top:3px solid var(--c);
         border-radius:14px; padding:14px 16px 12px; min-height:118px; box-shadow:{SHADOW};}}
  .kpi .lbl {{font-size:.68rem; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:{TEXT_2};}}
  .kpi .val {{font-size:1.9rem; font-weight:800; color:var(--c); line-height:1.25; margin-top:6px;
              white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}}
  .kpi .val.txt {{font-size:1.25rem; margin-top:10px;}}
  .kpi .sub {{font-size:.74rem; color:{MUTED}; margin-top:4px;}}

  [role="tablist"] {{gap:8px; flex-wrap:wrap; border-bottom:1px solid rgba(226,0,116,.2); padding-bottom:10px;}}
  [data-testid="stTab"] {{background:#f4f5f9; border:1px solid {BORDER}; border-radius:10px;
                         padding:6px 14px !important; height:auto; transition:background .15s;}}
  [data-testid="stTab"] p {{font-size:.86rem; font-weight:600; color:{TEXT_2};}}
  [data-testid="stTab"]:hover {{background:rgba(226,0,116,.16);}}
  [data-testid="stTab"][aria-selected="true"] {{background:linear-gradient(135deg,{MAGENTA},#ff3d9a);
                                               border-color:{MAGENTA}; box-shadow:0 4px 16px rgba(226,0,116,.45);}}
  [data-testid="stTab"][aria-selected="true"] p {{color:#fff;}}
  [role="tablist"] > :not([role="tab"]), [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {{display:none;}}

  .insights {{display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:14px; margin:6px 0 14px;}}
  .ins {{background:{CARD}; border:1px solid {BORDER}; border-left:4px solid var(--c);
         border-radius:14px; padding:14px 16px; box-shadow:{SHADOW};}}
  .ins h4 {{margin:0 0 6px; font-size:.92rem; color:{TEXT}; font-weight:700;}}
  .ins p {{margin:0; font-size:.83rem; line-height:1.5; color:{TEXT_2};}}
  .ins b {{color:{TEXT};}}

  .section {{font-size:1.02rem; font-weight:800; color:{TEXT}; margin:10px 0 6px; display:flex; gap:8px;
             align-items:center;}}
  .section .bar {{width:5px; height:20px; border-radius:3px;
                 background:linear-gradient({MAGENTA},{YELLOW});}}
  [class*="st-key-card_"] {{background:{CARD}; border:1px solid {BORDER} !important;
                            border-radius:14px; padding:12px 14px 6px; box-shadow:{SHADOW};}}
  /* Kartu berdampingan dibuat sama tinggi (border rata). */
  [data-testid="stHorizontalBlock"] {{align-items:stretch;}}
  [data-testid="stColumn"] > [data-testid="stVerticalBlock"] {{height:100%;}}
  [data-testid="stColumn"] [data-testid="stLayoutWrapper"]:has(> [class*="st-key-card_"]) {{flex:1 1 auto;}}
  [data-testid="stColumn"] [class*="st-key-card_"] {{height:100%;}}
  .ctitle {{font-size:.95rem; font-weight:700; color:{TEXT}; margin:2px 0 4px;}}
  .badge {{display:inline-block; font-size:.66rem; font-weight:700; padding:2px 7px; border-radius:6px;
           margin-right:8px; vertical-align:middle; color:var(--c); border:1px solid var(--c);
           background:color-mix(in srgb, var(--c) 14%, transparent);}}
</style>
""", unsafe_allow_html=True)


# =====================================================================================
# Utilitas format
# =====================================================================================

def fmt(v) -> str:
    return "-" if v is None or pd.isna(v) else f"{v:,.0f}".replace(",", ".")


def pct(v, digits: int = 1) -> str:
    return "-" if v is None or pd.isna(v) else f"{v:.{digits}f}%".replace(".", ",")


def days(v) -> str:
    return "-" if v is None or pd.isna(v) else f"{v:.2f}".replace(".", ",")


def esc(s) -> str:
    return html.escape(str(s))


def toco_code(name) -> str:
    text = " ".join(str(name).split())
    low = text.lower()
    for code, pattern in TOCO_RULES:
        if re.search(pattern, low):
            return code
    return text.upper() if text and text.lower() not in ("nan", "none") else "(Kosong)"


# =====================================================================================
# Data bersama: upload, lock, log
# =====================================================================================

def read_log() -> list[dict]:
    try:
        return json.loads(LOG_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def resolve_source() -> tuple[Path | None, dict]:
    """File aktif: unggahan bersama terbaru, atau file .xlsb terbaru di folder aplikasi."""
    if LATEST_FILE.exists():
        log = read_log()
        meta = log[-1] if log else {"uploader": "-", "timestamp": datetime.fromtimestamp(
            LATEST_FILE.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"), "filename": LATEST_FILE.name}
        return LATEST_FILE, meta
    local = sorted([p for p in APP_DIR.glob("*.xlsb") if not p.name.startswith("~$")],
                   key=lambda p: p.stat().st_mtime, reverse=True)
    if not local:
        return None, {}
    p = local[0]
    return p, {"uploader": "File lokal", "filename": p.name,
               "timestamp": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")}


def source_signature() -> str:
    path, _ = resolve_source()
    if path is None:
        return "none"
    stat = path.stat()
    return f"{path}:{stat.st_mtime_ns}:{stat.st_size}"


def active_lock() -> dict | None:
    try:
        info = json.loads(LOCK_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if time.time() - info.get("since", 0) > LOCK_TIMEOUT_S:
        return None
    return info


def save_upload(data: bytes, uploader: str, filename: str) -> dict:
    """Simpan unggahan ke shared_data dengan lock agar dua unggahan tidak saling menimpa."""
    SHARED_DIR.mkdir(exist_ok=True)
    lock = active_lock()
    if lock:
        raise RuntimeError(f"Data sedang dikunci oleh {lock.get('uploader', '-')} sejak "
                           f"{datetime.fromtimestamp(lock['since']):%H:%M:%S}. Coba lagi sebentar.")
    LOCK_FILE.write_text(json.dumps({"uploader": uploader, "since": time.time()}), encoding="utf-8")
    tmp = SHARED_DIR / "latest_tlp_data.tmp.xlsb"
    try:
        tmp.write_bytes(data)
        with open_workbook(str(tmp)) as wb:
            found = [c["sheet"] for c in CATEGORIES.values() if c["sheet"] in wb.sheets]
        if not found:
            raise ValueError("File tidak memuat sheet List Access Issue / List Stolen / TT CME.")
        os.replace(tmp, LATEST_FILE)  # penggantian atomik: pembaca tidak pernah melihat file setengah jadi
        entry = {"uploader": uploader, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "filename": filename, "size_mb": round(len(data) / 1e6, 2), "sheets": found}
        log = (read_log() + [entry])[-100:]
        LOG_FILE.write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")
        return entry
    finally:
        tmp.unlink(missing_ok=True)
        LOCK_FILE.unlink(missing_ok=True)


def load_sow_override() -> dict[str, float]:
    try:
        return {k: float(v) for k, v in json.loads(SOW_FILE.read_text(encoding="utf-8")).items()}
    except (OSError, ValueError):
        return {}


def save_sow_override(values: dict[str, float]) -> None:
    SHARED_DIR.mkdir(exist_ok=True)
    SOW_FILE.write_text(json.dumps(values, indent=2), encoding="utf-8")


# =====================================================================================
# Membaca .xlsb dengan pyxlsb
# =====================================================================================

MAX_EMPTY_RUN = 500  # berhenti setelah 500 baris kosong berturut-turut


def _clean_name(v, i: int, seen: dict) -> str:
    name = " ".join(str(v).split()) if v not in (None, "") else f"Kolom_{i + 1}"
    if name in seen:  # nama kolom ganda diberi akhiran agar unik (mis. Site Status ..._1)
        seen[name] += 1
        name = f"{name}_{seen[name]}"
    else:
        seen[name] = 0
    return name


def _sheet_rows(wb, sheet: str) -> list[list]:
    """Semua baris berisi dari satu sheet. Sheet yang berformat sampai ~1 juta baris walau
    datanya belasan ribu dihentikan setelah deretan baris kosong agar tetap cepat."""
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
    return rows


def _rows_to_frame(rows: list[list]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    seen: dict = {}
    header = [_clean_name(v, i, seen) for i, v in enumerate(rows[0])]
    width = len(header)
    body = [(r + [None] * width)[:width] for r in rows[1:]]
    return pd.DataFrame(body, columns=header).dropna(how="all").dropna(axis=1, how="all")


def _extract_sow(rows: list[list]) -> dict[str, float]:
    """Cari tabel 'ToCo | SoW (A)' / 'ToCo | Total Sites' di sheet performance dan baca isinya."""
    out: dict[str, float] = {}
    for r, row in enumerate(rows):
        for c in range(1, len(row)):
            head = str(row[c] or "").strip().lower()
            left = str(row[c - 1] or "").strip().lower()
            if left != "toco" or head not in ("sow (a)", "total sites"):
                continue
            for below in rows[r + 1:]:
                name = below[c - 1] if c - 1 < len(below) else None
                val = below[c] if c < len(below) else None
                if name in (None, "") or not isinstance(val, (int, float)):
                    break
                out.setdefault(toco_code(name), float(val))
    return out


COUPA_COLS = ["TLP2", "Site ID", "Status", "Kategori Issue", "SLA", "Problem type", "Weeks", "Region"]


def _frame_from_header(rows: list[list], must_have: tuple[str, ...]) -> pd.DataFrame:
    """Sheet yang header-nya tidak di baris pertama (ada catatan di atas): cari baris header dulu."""
    for i, row in enumerate(rows[:50]):
        cells = {" ".join(str(v).split()) for v in row if v not in (None, "")}
        if all(m in cells for m in must_have):
            return _rows_to_frame(rows[i:])
    return pd.DataFrame()


def load_workbook(path: str) -> tuple[dict[str, pd.DataFrame], dict[str, float], pd.DataFrame, dict]:
    # Workbook dibuka sekali saja: membuka ulang = membaca ulang shared strings (lambat).
    sheets, sow = {}, {}
    with open_workbook(path) as wb:
        for cat, cfg in CATEGORIES.items():
            if cfg["sheet"] in wb.sheets:
                sheets[cat] = _rows_to_frame(_sheet_rows(wb, cfg["sheet"]))
        for name in ["performance access", "Performance Stolen", "sheet tambahan ke TLP"]:
            if name in wb.sheets:
                for k, v in _extract_sow(_sheet_rows(wb, name)).items():
                    sow.setdefault(k, v)
        # Sheet "LIST SITE" (daftar site per TLP) -> Total site K1 = COUNTIFS(Owner, TLP).
        list_site = next((n for n in wb.sheets if n.strip().lower().startswith("list site")), None)
        sites = _rows_to_frame(_sheet_rows(wb, list_site)) if list_site else pd.DataFrame()
        # COUPA Report: sheet data TT bulanan ("Data TT Sept 2026", dst.) + sheet "Site" (total site per TLP).
        coupa: dict = {"tt": pd.DataFrame(), "site": pd.DataFrame(), "tt_sheet": None}
        tt_sheet = next((n for n in wb.sheets if n.strip().lower().startswith("data tt")), None)
        if tt_sheet:
            tt = _frame_from_header(_sheet_rows(wb, tt_sheet), ("TLP2", "Site ID"))
            coupa["tt"] = tt[[c for c in COUPA_COLS if c in tt.columns]]
            coupa["tt_sheet"] = tt_sheet
        site_sheet = next((n for n in wb.sheets if n.strip().lower() == "site"), None)
        if site_sheet:
            coupa["site"] = _frame_from_header(_sheet_rows(wb, site_sheet), ("GROUP", "Total Site"))
    return sheets, sow, sites, coupa


# =====================================================================================
# Persiapan data
# =====================================================================================

def to_period(s: pd.Series) -> pd.Series:
    """'Oct_2025' / tanggal / serial Excel -> Period bulanan."""
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
    for label, pattern in reversed(rules):  # aturan pertama menang
        out = out.mask(low.str.contains(pattern, regex=True), label)
    return out


def _col(df: pd.DataFrame, name: str | None, default="(Kosong)") -> pd.Series:
    if name and name in df.columns:
        return df[name].astype("string").str.strip().replace("", pd.NA).fillna(default)
    return pd.Series(default, index=df.index, dtype="string")


def _first(df: pd.DataFrame, names: list[str]) -> str | None:
    return next((n for n in names if n in df.columns and df[n].notna().any()), None)


GROUP_ORDER = ["PTI", "DMT", "CMI", "TBG", "Others"]


def group_label(name) -> str:
    """Kolom GROUP (PTI - Group, DAYAMITRA..., CMI - Group, TOWER BERSAMA, PT, OTHERS) -> label singkat."""
    code = toco_code(name)
    return code if code in GROUP_ORDER else "Others"


def _unify_case(s: pd.Series) -> pd.Series:
    """Samakan ejaan yang hanya beda huruf besar/kecil (Dhost/DHOST/DHost) ke ejaan terbanyak."""
    key = s.str.lower()
    canon = s.groupby(key).agg(lambda v: v.value_counts().index[0])
    return key.map(canon).astype("string")


def _clean_root(s: pd.Series) -> pd.Series:
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.str.replace(r"(?i)^kompesasi", "Kompensasi", regex=True)
    return s.where(s.isin(["(Kosong)"]), s.str.slice(0, 1).str.upper() + s.str.slice(1))


def prepare(_df: pd.DataFrame, cat: str) -> pd.DataFrame:
    """Tambah kolom turunan (_status, _toco, _impact, ...). `_df` dipakai langsung (tanpa salinan)."""
    cfg = CATEGORIES[cat]
    df = _df
    df["_period"] = to_period(df[cfg["month"]]) if cfg["month"] in df.columns else pd.NaT
    df = df[df["_period"].notna()].copy()
    status = _col(df, cfg["status"], "").str.lower()
    df["_status"] = status.map(lambda s: "Open" if s == "open" else "Cancel" if s.startswith("cancel")
                               else "Closed" if s.startswith("close") else "Lainnya")
    df["_cat"] = cat
    df["_region"] = _col(df, REGION_COL)
    df["_toco_raw"] = _col(df, _first(df, cfg["toco"]))  # nama asli (GROUP1 / Toco Group Name / Owner)
    df["_toco"] = df["_toco_raw"].map(toco_code)
    df["_owner"] = _unify_case(_col(df, cfg.get("owner")))
    raw_cause = _col(df, cfg["cause"])
    df["_cause"] = classify(raw_cause, cfg["cause_rules"]) if cfg.get("cause_rules") else raw_cause
    root_col = _first(df, cfg["root"])
    df["_root"] = _unify_case(_clean_root(_col(df, root_col))) if root_col else df["_cause"]
    df["_sev"] = _col(df, cfg["severity"]).str.title()
    aging = (pd.to_numeric(df[cfg["aging"]], errors="coerce") if cfg["aging"] in df.columns
             else pd.Series(float("nan"), index=df.index))
    df["_aging"] = aging.where(aging >= 0)  # aging negatif = salah input tanggal, diabaikan dari MTTR
    df["_year"] = (_col(df, cfg["year"]).str.replace(r"\.0$", "", regex=True) if cfg["year"]
                   else df["_period"].dt.year.astype("string"))
    df["_backlog"] = _col(df, cfg.get("backlog"), "Normal TT")
    df["_month"] = df["_period"].dt.strftime("%b_%Y")
    df["_site"] = _col(df, cfg["site_id"], "")
    # Baris yang punya nomor TT (mis. "Count of Huawei TT" di pivot); tanpa kolom TT, semua baris dihitung.
    df["_has_tt"] = _col(df, cfg.get("tt"), "").ne("") if cfg.get("tt") in df.columns else True
    # Kolom untuk tab Report TT Access (pivot mingguan/bulanan seperti di Excel).
    df["_cw"] = _col(df, "CW Open", "")
    df["_sla"] = _col(df, "SLA")
    df["_sitecat"] = _col(df, "Site Category").str.upper()
    df["_grp"] = _col(df, "GROUP").map(group_label)
    impact_col = _first(df, cfg["impact"])
    df["_impact"] = _col(df, impact_col, "").str.lower().eq("down") if impact_col else False
    present = [c for c in cfg["search"] if c in df.columns]
    df["_search"] = (pd.concat([_col(df, c, "") for c in present], axis=1).agg(" ".join, axis=1).str.lower()
                     if present else "")
    return df


def _pm_column(df: pd.DataFrame) -> str | None:
    """Kolom status PM di LIST SITE (berisi 'NY PM' / 'DONE PM'; kolom O di Excel), dicari dari isinya."""
    for c in df.columns:
        if df[c].astype("string").str.strip().str.upper().isin(["NY PM", "DONE PM"]).any():
            return c
    return None


def prepare_sites(_df: pd.DataFrame) -> pd.DataFrame:
    """LIST SITE: kolom Owner -> grup ToCo; 'TLP/MO OTHERS' dipisah sebagai OTHERS.
    _pm = status PM (NY PM / DONE PM), _stype = Site Type (MACRO/MICRO/IBC), _site = IOH ID."""
    cols = ["_owner_raw", "_toco", "_region", "_pm", "_stype", "_site"]
    if _df.empty or "Owner" not in _df.columns:
        return pd.DataFrame(columns=cols)
    pm_col = _pm_column(_df)
    df = pd.DataFrame({"_owner_raw": _col(_df, "Owner"), "_region": _col(_df, "Region"),
                       "_pm": _col(_df, pm_col, "").str.upper().str.split().str.join(" "),
                       "_stype": _col(_df, next((c for c in _df.columns if c.strip().lower() == "site type"), None), "")
                       .str.upper(),
                       "_site": _col(_df, "IOH ID", "")})
    others = df["_owner_raw"].str.lower().str.contains("other")
    df["_toco"] = df["_owner_raw"].map(toco_code).where(~others, "OTHERS")
    return df[df["_owner_raw"] != "(Kosong)"]


@st.cache_resource(show_spinner="Membaca file .xlsb (pyxlsb)…", max_entries=1)
def load_data(signature: str, path: str) -> tuple[dict[str, pd.DataFrame], dict[str, float], pd.DataFrame, dict]:
    """Baca + siapkan data sekali per versi file. st.cache_resource membagi objek yang sama ke semua
    pengguna tanpa menyalin (hemat memori — penting di Streamlit Community Cloud yang dibatasi ±1 GB).
    `signature` (path + waktu ubah + ukuran + versi) menjadi kunci: data baru otomatis dibaca ulang dan
    versi lama dibuang (max_entries=1). Jangan mengubah DataFrame hasil fungsi ini di tempat."""
    sheets, sow, sites, coupa = load_workbook(path)
    prepared = {cat: prepare(sheets.pop(cat), cat) for cat in list(sheets)}
    return prepared, sow, prepare_sites(sites), coupa


def apply_filters(df: pd.DataFrame, f: dict, *, status: bool = True, month: bool = True) -> pd.DataFrame:
    for col, key in [("_year", "years"), ("_region", "regions"), ("_toco", "tlp"), ("_cause", "cause"),
                     ("_sev", "sev")]:
        if f[key]:
            df = df[df[col].isin(f[key])]
    if f["backlog"]:  # Backlog hanya berlaku untuk TT CME
        df = df[(df["_cat"] != "TT CME") | df["_backlog"].isin(f["backlog"])]
    if f["q"]:
        df = df[df["_search"].str.contains(re.escape(f["q"].lower()), regex=True, na=False)]
    if month and f["months"]:
        df = df[df["_month"].isin(f["months"])]
    if status:
        df = df[df["_status"].isin(f["status"])] if f["status"] else df[df["_status"] != "Cancel"]
    return df


# =====================================================================================
# Perhitungan
# =====================================================================================

def backlog_summary(df: pd.DataFrame, months: list) -> pd.DataFrame:
    """Tabel A–E per bulan. TT sebelum bulan pertama tidak dihitung (A awal = 0)."""
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


def impact_table(access: pd.DataFrame, stolen: pd.DataFrame, sow: dict[str, float]) -> pd.DataFrame:
    """Tabel per ToCo seperti sheet performance: SoW (A), TT & sites impact service, stolen."""
    imp = access[access["_impact"] & access["_status"].isin(["Open", "Closed"])] if len(access) else access
    st_ = stolen[stolen["_status"].isin(["Open", "Closed"])] if len(stolen) else stolen
    t = pd.DataFrame({
        "SoW (A)": pd.Series(sow, dtype=float),
        "Total TT Impact Service Released": imp.groupby("_toco").size() if len(imp) else pd.Series(dtype=float),
        "Total Sites Impact Service (B)": imp.groupby("_toco")["_site"].nunique() if len(imp) else pd.Series(dtype=float),
        "MTTR Average (Days)": imp.groupby("_toco")["_aging"].mean() if len(imp) else pd.Series(dtype=float),
        "Total TT Stolen": st_.groupby("_toco").size() if len(st_) else pd.Series(dtype=float),
        "Total Sites Stolen": st_.groupby("_toco")["_site"].nunique() if len(st_) else pd.Series(dtype=float),
    })
    for c in ["Total TT Impact Service Released", "Total Sites Impact Service (B)", "Total TT Stolen",
              "Total Sites Stolen"]:
        t[c] = t[c].fillna(0)
    t["% (B/A)"] = t["Total Sites Impact Service (B)"] / t["SoW (A)"].where(t["SoW (A)"] > 0) * 100
    t["% Stolen / SoW"] = t["Total Sites Stolen"] / t["SoW (A)"].where(t["SoW (A)"] > 0) * 100
    t = t[(t["SoW (A)"].fillna(0) > 0) | (t["Total TT Impact Service Released"] > 0) | (t["Total TT Stolen"] > 0)]
    t = t.drop(index=["(Kosong)"], errors="ignore")
    t = t.sort_values(["SoW (A)", "Total TT Impact Service Released"], ascending=False, na_position="last")
    t.index.name = "ToCo"
    return t


def top_n(series: pd.Series, n: int = 5, other: str = "Lainnya") -> pd.Series:
    s = series[series > 0].sort_values(ascending=False)
    if len(s) <= n:
        return s
    return pd.concat([s.head(n), pd.Series({other: s.iloc[n:].sum()})])


# =====================================================================================
# Grafik (tema gelap)
# =====================================================================================

def dark(fig: go.Figure, height: int, legend_top: bool = True) -> go.Figure:
    fig.update_layout(
        height=height, font=dict(family=FONT, size=12, color=TEXT_2),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", separators=",.",
        margin=dict(l=8, r=8, t=46 if legend_top else 16, b=8),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=1.04, yanchor="bottom",
                    font=dict(size=11, color=TEXT_2), itemsizing="constant", bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="#ffffff", bordercolor=MAGENTA, font=dict(family=FONT, color=TEXT)),
        bargap=0.28, barcornerradius=4,
    )
    fig.update_xaxes(gridcolor=GRID, linecolor="rgba(30,35,70,.25)", zeroline=False, tickfont=dict(color=TEXT_2))
    fig.update_yaxes(gridcolor=GRID, linecolor="rgba(0,0,0,0)", zeroline=False, tickfont=dict(color=MUTED))
    return fig


def wrap(label, width: int = 14) -> str:
    """Pecah label panjang jadi beberapa baris agar label sumbu tidak bertumpuk."""
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


def empty_fig(msg: str, height: int = 320) -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(text=msg, showarrow=False, font=dict(color=MUTED, size=13), x=0.5, y=0.5,
                       xref="paper", yref="paper")
    dark(fig, height, legend_top=False)
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return fig


def _light(color: str) -> bool:
    """Warna terang (kuning, abu muda, ...) -> label di atasnya pakai teks gelap."""
    h = str(color).lstrip("#")[:6]
    if len(h) != 6:
        return False
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.299 * r + 0.587 * g + 0.114 * b > 150


def pie_3d(labels, values, colors, *, hole: float = 0.0, texts=None, height: int = 380,
           center: str | None = None) -> go.Figure:
    """Pie bergaya '3D': irisan sedikit ditarik, tepi gelap, dan bayangan elips di bawahnya."""
    values = [float(v) for v in values]
    fig = go.Figure()
    # Lapisan bayangan (pie lebih gelap sedikit di bawah) memberi kesan ketebalan.
    fig.add_trace(go.Pie(labels=labels, values=values, hole=hole, sort=False, direction="clockwise",
                         rotation=90, domain=dict(x=[0.02, 0.98], y=[0.0, 0.92]), showlegend=False,
                         marker=dict(colors=["rgba(20,25,60,.16)"] * len(values), line=dict(width=0)),
                         textinfo="none", hoverinfo="skip"))
    fig.add_trace(go.Pie(
        labels=labels, values=values, hole=hole, sort=False, direction="clockwise", rotation=90,
        domain=dict(x=[0.02, 0.98], y=[0.06, 0.98]),
        pull=[0.04] + [0.02] * (len(values) - 1),
        marker=dict(colors=colors, line=dict(color=BG, width=2)),
        text=texts, texttemplate="%{text}" if texts else "%{value:,.0f}<br>%{percent}",
        textposition="auto", insidetextorientation="horizontal",
        # Teks gelap di irisan kuning, putih di irisan lain, agar tetap terbaca.
        textfont=dict(size=12, color=[TEXT if _light(c) else "#ffffff" for c in colors], family=FONT),
        outsidetextfont=dict(size=11, color=TEXT_2),
        hovertemplate="<b>%{label}</b><br>%{value:,.0f} (%{percent})<extra></extra>",
    ))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=True, margin=dict(l=10, r=10, t=10, b=10),
                      legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.04, yanchor="top",
                                  font=dict(size=11, color=TEXT_2)),
                      uniformtext=dict(minsize=9, mode="hide"))
    if center:
        fig.add_annotation(text=center, showarrow=False, x=0.5, y=0.52, xref="paper", yref="paper",
                           font=dict(size=13, color=TEXT, family=FONT))
    return fig


def share_pie(parts: dict[str, float], colors: list[str]) -> go.Figure:
    """Pie 3D SoW (A) vs Total Sites: persen = bagian dari jumlah irisan (sama dengan label pie Excel)."""
    total = sum(parts.values())
    texts = [f"{fmt(v)}; {pct(v / total * 100, 2)}" for v in parts.values()]
    fig = pie_3d(list(parts), list(parts.values()), colors, texts=texts)
    fig.data[1].pull = [0, 0.12]
    return fig


def ltoa_pie(s: pd.Series) -> go.Figure:
    """LTOA & Cell Down Contribution: Total TT Impact Service Released per ToCo terpilih."""
    s = s[s > 0]
    if s.empty:
        return empty_fig("Belum ada TT impact service (Site Status Down) pada pilihan ini.", 380)
    total = s.sum()
    texts = [f"{fmt(v)}<br>{k}<br>{pct(v / total * 100, 0)}" for k, v in s.items()]
    return pie_3d(list(s.index), list(s.values), [EXCEL_COLORS[i % len(EXCEL_COLORS)] for i in range(len(s))],
                  texts=texts)


def root_cause_pie(d: pd.DataFrame, top: int = 7, height: int = 400) -> go.Figure:
    """The Root Cause: Count of Huawei TT per kolom Case, Top `top` (persen terhadap jumlah yang tampil)."""
    s = d.loc[d["_has_tt"] & (d["_root"] != "(Kosong)"), "_root"].value_counts().head(top)
    if s.empty:
        return empty_fig("Kolom Case kosong pada pilihan ini.", height)
    total = s.sum()
    texts = [f"{fmt(v)}; {pct(v / total * 100, 0)}" for v in s.values]
    return pie_3d([wrap(k, 28) for k in s.index], list(s.values), EXCEL_COLORS[:len(s)], texts=texts,
                  height=height)


def ranked(d: pd.DataFrame, col: str, top: int = 10, color: str = TOSCA, height: int = 360,
           show_pct: bool = True) -> go.Figure:
    vc = d[col].value_counts().head(top)[::-1]
    if vc.empty:
        return empty_fig("Tidak ada data.", height)
    share = vc / max(len(d), 1) * 100
    fig = go.Figure(go.Bar(
        y=vc.index, x=vc.values, orientation="h", marker_color=color, name="Jumlah TT",
        text=[f"{fmt(v)}  ({pct(s)})" if show_pct else fmt(v) for v, s in zip(vc.values, share)],
        textposition="outside",
        textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:,} TT<extra></extra>"))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=False, margin=dict(t=16, r=70),
                      xaxis=dict(range=[0, vc.max() * 1.3], showgrid=True), yaxis=dict(automargin=True))
    return fig


def pareto(d: pd.DataFrame | None, col: str | None, top: int = 15, values: pd.Series | None = None,
           name: str = "Jumlah TT") -> go.Figure:
    vc = values.sort_values(ascending=False) if values is not None else d[col].value_counts()
    if vc.empty:
        return empty_fig("Tidak ada data.", 400)
    cum = vc.cumsum() / vc.sum() * 100
    vc, cum = vc.head(top), cum.head(top)
    labels = [wrap(i) for i in vc.index]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=labels, y=vc.values, name=name, marker_color=PURPLE, text=bar_text(vc.values),
                         textposition="outside", textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
                         hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,}}<extra></extra>"))
    fig.add_trace(go.Scatter(x=labels, y=cum.values, name="Kumulatif %", yaxis="y2",
                             mode="lines+markers+text", text=bar_text(cum.values, lambda v: pct(v, 0)),
                             textposition="top center", textfont=dict(size=10, color=GOLD), cliponaxis=False,
                             line=dict(color=YELLOW, width=2), marker=dict(size=7),
                             hovertemplate="<b>%{x}</b><br>Kumulatif: %{y:.1f}%<extra></extra>"))
    dark(fig, 400)
    fig.update_layout(yaxis=dict(range=[0, vc.max() * 2.1]),
                      yaxis2=dict(overlaying="y", side="right", range=[-45, 112], showgrid=False,
                                  tickvals=[0, 50, 100], ticksuffix="%", tickfont=dict(color=GOLD)),
                      xaxis=dict(tickangle=-30 if len(vc) > 8 else 0, tickfont=dict(size=10)))
    return fig


def _clean_label(v) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(v))).strip()


def fig_export(fig: go.Figure) -> dict | None:
    """Ambil angka grafik Plotly (bar, garis, pie) jadi spesifikasi data untuk Excel & grafik PPT native."""
    traces = [t for t in fig.data if t.type in ("bar", "scatter", "pie") and t.hoverinfo != "skip"]
    if not traces:
        return None
    pie = next((t for t in traces if t.type == "pie"), None)
    if pie is not None:
        colors = list(pie.marker.colors) if pie.marker.colors is not None else None
        return dict(kind="doughnut" if (pie.hole or 0) > 0 else "pie",
                    categories=[_clean_label(x) for x in pie.labels],
                    series=[dict(name="Jumlah", values=[float(v) for v in pie.values], color=None, colors=colors,
                                 chart=True)], pct=False)
    horiz = getattr(traces[0], "orientation", None) == "h"
    cat_of = (lambda t: t.y) if horiz else (lambda t: t.x)
    val_of = (lambda t: t.x) if horiz else (lambda t: t.y)
    cats = [_clean_label(c) for c in cat_of(traces[0])]
    series = []
    for t in traces:
        if t.type == "scatter" and (t.mode or "") == "text":
            continue
        lookup = dict(zip([_clean_label(c) for c in cat_of(t)], list(val_of(t))))
        vals = [None if lookup.get(c) is None or pd.isna(lookup.get(c)) else float(lookup[c]) for c in cats]
        color = t.marker.color if t.type == "bar" else t.line.color
        many = isinstance(color, (list, tuple))
        series.append(dict(name=t.name or "Nilai", values=vals[::-1] if horiz else vals,
                           color=None if many else color, colors=(list(color)[::-1] if horiz else list(color))
                           if many else None,
                           # garis kumulatif % (skala lain) hanya masuk tabel, tidak digambar di PPT
                           chart=not ("kumulatif" in (t.name or "").lower()
                                      or (t.type == "scatter" and (t.name or "") in ("Grand Total", "MTTR (hari)")))))
    stacked_area = any(t.type == "scatter" and t.stackgroup for t in traces)
    kind = ("bar_h" if horiz else "column_stacked_100" if fig.layout.barnorm == "percent"
            else "column_stacked" if stacked_area
            else "column_stacked" if fig.layout.barmode == "stack" else "column_clustered")
    return dict(kind=kind, categories=cats[::-1] if horiz else cats, series=series,
                pct=fig.layout.yaxis.ticksuffix == "%")


def spec_table(spec: dict) -> pd.DataFrame:
    df = pd.DataFrame({s_["name"]: s_["values"] for s_ in spec["series"]}, index=spec["categories"])
    df.index.name = "Kategori"
    if spec["kind"] in ("pie", "doughnut"):
        total = df.iloc[:, 0].sum()
        df["%"] = (df.iloc[:, 0] / total * 100).round(2) if total else None
    return df


GROUP_COLORS = {"PTI": MAGENTA, "DMT": TOSCA, "CMI": YELLOW, "TBG": RED, "Others": PURPLE}
SLA_COLORS = {"Exceed SLA": MAGENTA, "Meet SLA": TOSCA}
SEV_COLORS = {"CRITICAL": MAGENTA, "MAJOR": TOSCA}
HUB_COLORS = {"BIG HUB SITE": MAGENTA, "MEDIUM HUB SITE": TOSCA}


def pivot_tt(df: pd.DataFrame, row: str, col: str, rows: list, cols: list) -> pd.DataFrame:
    """Pivot Count of Huawei TT (baris x kolom) + Grand Total, urutan baris/kolom ditentukan."""
    ct = df[df["_has_tt"]].groupby([row, col]).size().unstack(fill_value=0)
    ct = ct.reindex(index=rows, columns=cols, fill_value=0)
    ct["Grand Total"] = ct.sum(axis=1)
    return ct


def _labels(vals) -> list[str]:
    return [fmt(v) if v else "" for v in vals]


def sla_area(ct: pd.DataFrame, x: list[str]) -> go.Figure:
    fig = go.Figure()
    base = pd.Series(0, index=ct.index)
    for name in ["Exceed SLA", "Meet SLA"]:
        fig.add_trace(go.Scatter(
            x=x, y=ct[name], name=name, mode="lines", stackgroup="sla", line=dict(width=0.5, color=SLA_COLORS[name]),
            fillcolor=SLA_COLORS[name], hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,}}<extra></extra>"))
        # Label di tengah pita masing-masing (bukan di tepi atas) agar tidak tertimpa label Grand Total.
        fig.add_trace(go.Scatter(
            x=x, y=base + ct[name] / 2, mode="text", text=_labels(ct[name]), textposition="middle center",
            textfont=dict(size=10, color="#ffffff"), showlegend=False, hoverinfo="skip"))
        base = base + ct[name]
    fig.add_trace(go.Scatter(
        x=x, y=ct["Grand Total"], name="Grand Total", mode="lines+markers+text", text=_labels(ct["Grand Total"]),
        textposition="top center", textfont=dict(size=11, color=TEXT), line=dict(color="#5B8DD6", width=2.5),
        marker=dict(size=7), cliponaxis=False, hovertemplate="<b>%{x}</b><br>Grand Total: %{y:,}<extra></extra>"))
    dark(fig, 400)
    fig.update_layout(hovermode="x unified", yaxis=dict(range=[0, (ct["Grand Total"].max() or 1) * 1.2]),
                      xaxis=dict(showgrid=False))
    return fig


def stacked_groups(ct: pd.DataFrame, x: list[str], colors: dict) -> go.Figure:
    fig = go.Figure()
    base = pd.Series(0, index=ct.index)
    for name in [c for c in ct.columns if c != "Grand Total"]:
        fig.add_trace(go.Bar(x=x, y=ct[name], name=name, marker_color=colors.get(name),
                             hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,}}<extra></extra>"))
        # Label ditulis sebagai teks terpisah di tengah tiap potongan: selalu tampil & mendatar,
        # termasuk potongan kecil yang tidak muat untuk label bawaan batang.
        fig.add_trace(go.Scatter(x=x, y=base + ct[name] / 2, mode="text", text=_labels(ct[name]),
                                 textposition="middle center", showlegend=False, hoverinfo="skip",
                                 textfont=dict(size=10, color=TEXT if _light(colors.get(name, "")) else "#ffffff")))
        base = base + ct[name]
    fig.add_trace(go.Scatter(x=x, y=ct["Grand Total"], mode="text", text=_labels(ct["Grand Total"]),
                             textposition="top center", textfont=dict(size=11, color=TEXT), showlegend=False,
                             hoverinfo="skip"))
    dark(fig, 440)
    fig.update_layout(barmode="stack", bargap=0.18,
                      yaxis=dict(range=[0, (ct["Grand Total"].max() or 1) * 1.15]), xaxis=dict(showgrid=False))
    return fig


def clustered_groups(ct: pd.DataFrame, x: list[str], colors: dict, total_line: bool = False,
                     spline: bool = False) -> go.Figure:
    fig = go.Figure()
    series = [c for c in ct.columns if c != "Grand Total"]
    for name in series:
        fig.add_trace(go.Bar(x=x, y=ct[name], name=name, marker_color=colors.get(name),
                             text=_labels(ct[name]), textposition="outside", textangle=0,
                             textfont=dict(size=9 if len(series) > 2 else 10, color=TEXT_2), cliponaxis=False,
                             hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,}}<extra></extra>"))
    bar_max = float(ct[series].to_numpy().max() or 1)
    dark(fig, 400)
    fig.update_layout(barmode="group", bargroupgap=0.05,
                      xaxis=dict(showgrid=False, tickangle=-30 if len(x) > 8 else 0))
    if total_line:
        tot = ct["Grand Total"]
        fig.add_trace(go.Scatter(
            x=x, y=tot, name="Grand Total", yaxis="y2", mode="lines+markers+text", text=_labels(tot),
            textposition="top center", textfont=dict(size=11, color=GOLD), cliponaxis=False,
            line=dict(color=YELLOW, width=3, shape="spline" if spline else "linear"), marker=dict(size=7),
            hovertemplate="<b>%{x}</b><br>Grand Total: %{y:,}<extra></extra>"))
        # Batang di bawah, garis total di atas, agar label tidak bertumpuk.
        t_max = float(tot.max() or 1)
        step = max(1, round(t_max / 2))
        fig.update_layout(yaxis=dict(range=[0, bar_max * 1.8]),
                          yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                      range=[-t_max * 0.9, t_max * 1.15], tickmode="array",
                                      tickvals=[0, step, step * 2] if step * 2 <= t_max * 1.15 else [0, step],
                                      tickfont=dict(color=GOLD)))
    else:
        fig.update_layout(yaxis=dict(range=[0, bar_max * 1.2]))
    return fig


def mttr_critical_table(df: pd.DataFrame) -> pd.DataFrame:
    """Pivot Severity CRITICAL per GROUP: Count of Huawei TT, Sum of Aging Time, MTTR = Sum Aging ÷ TT
    (rumus Excel =E118/D118)."""
    crit = df[df["_has_tt"] & (df["_sev"].str.upper() == "CRITICAL")]
    g = crit.groupby("_grp").agg(tt=("_grp", "size"), aging=("_aging", "sum"))
    g = g.reindex([x for x in GROUP_ORDER if x in g.index])
    g = g.sort_values("tt", ascending=False)
    g["mttr"] = g["aging"] / g["tt"].where(g["tt"] > 0)
    g.index.name = "GROUP"
    return g


def mttr_critical_chart(g: pd.DataFrame) -> go.Figure:
    if g.empty:
        return empty_fig("Tidak ada TT CRITICAL pada filter ini.", 400)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=list(g.index), y=g["tt"], name="TT Critical Access Issue",
                         marker_color=[GROUP_COLORS.get(i, PURPLE) for i in g.index], text=_labels(g["tt"]),
                         textposition="inside", insidetextanchor="end", textangle=0,
                         textfont=dict(size=11, color="#ffffff"),
                         hovertemplate="<b>%{x}</b><br>TT Critical: %{y:,}<extra></extra>"))
    fig.add_trace(go.Scatter(x=list(g.index), y=g["mttr"], name="MTTR (hari)", yaxis="y2",
                             mode="lines+markers+text", text=[days(v) for v in g["mttr"]], textposition="top center",
                             textfont=dict(size=11, color=ORANGE), line=dict(color=ORANGE, width=3, shape="spline"),
                             marker=dict(size=8), cliponaxis=False,
                             hovertemplate="<b>%{x}</b><br>MTTR: %{y:.2f} hari<extra></extra>"))
    dark(fig, 400)
    fig.update_layout(yaxis=dict(range=[0, (g["tt"].max() or 1) * 1.15], title="TT"),
                      yaxis2=dict(overlaying="y", side="right", showgrid=False, rangemode="tozero",
                                  range=[0, (g["mttr"].max() or 1) * 1.25], title="MTTR (hari)",
                                  tickfont=dict(color=ORANGE)),
                      xaxis=dict(showgrid=False))
    return fig


# ---------- Trends [General] / Latest update [Region Basis] / [ToCo Basis] (Access, Stolen, CME) ----------

def span_label(win: list) -> str:
    return f"{win[0]:%b %Y} – {win[-1]:%b %Y}" if len(win) > 1 else f"{win[0]:%b %Y}"


def split_labels(win: list) -> tuple[str | None, str]:
    """Kolom pivot 'bulan-bulan sebelumnya' (mis. Oct 25 - Sep 26) dan 'bulan terakhir' (mis. Oct-26)."""
    prev = f"{win[0]:%b %y} - {win[-2]:%b %y}" if len(win) > 2 else (f"{win[0]:%b-%y}" if len(win) == 2 else None)
    return prev, f"{win[-1]:%b-%y}"


def trend_chart(t: pd.DataFrame) -> go.Figure:
    """Seperti grafik Excel: batang A–D per bulan, E sebagai garis pendek merah."""
    x = t.index.tolist()
    fig = go.Figure()
    for key in "ABCD":
        name, color = SERIES[key]
        fig.add_trace(go.Bar(x=x, y=t[key], name=name, marker_color=color,
                             hovertemplate=f"{name}: %{{y:,.0f}}<extra></extra>"))
    name, color = SERIES["E"]
    fig.add_trace(go.Scatter(x=x, y=t["E"], name=name, mode="markers", line=dict(color=color),
                             marker=dict(symbol="line-ew", size=26, color=color, line=dict(color=color, width=6)),
                             hovertemplate=f"{name}: %{{y:,.0f}}<extra></extra>"))
    dark(fig, 440)
    # Legenda 2 baris (3 + 2 entri) di atas area plot; margin atas cukup agar tidak menimpa batang.
    fig.update_layout(barmode="group", bargap=0.2, bargroupgap=0.04, hovermode="x unified", margin=dict(t=96),
                      legend=dict(x=0, xanchor="left", y=1.0, yanchor="bottom", entrywidth=0.33,
                                  entrywidthmode="fraction"),
                      yaxis=dict(tickformat=",.0f", title=dict(text="TT Released", font=dict(color=MUTED))),
                      xaxis=dict(showgrid=False, tickangle=-30 if len(x) > 10 else 0))
    return fig


def grid_html(df: pd.DataFrame, colors: dict | None = None, pct_rows: tuple = ()) -> str:
    """Tabel data di bawah grafik (seperti data table grafik Excel): baris = seri, kolom = kategori."""
    colors = colors or {}
    th = f"padding:4px 6px;border:1px solid #E3E5EE;background:#F4F5F9;color:{TEXT};font-weight:700;text-align:center"
    td = "padding:4px 6px;border:1px solid #E3E5EE;text-align:center"
    head = f"<tr><th style='{th}'></th>" + "".join(f"<th style='{th}'>{esc(c)}</th>" for c in df.columns) + "</tr>"
    body = ""
    for name, row in df.iterrows():
        sw = (f"<span style='display:inline-block;width:9px;height:9px;margin-right:6px;background:{colors[name]}'>"
              "</span>" if name in colors else "")
        f_ = (lambda v: pct(v, 0)) if name in pct_rows else fmt
        body += (f"<tr><td style='{td};text-align:left;white-space:nowrap;font-weight:600'>{sw}{esc(name)}</td>"
                 + "".join(f"<td style='{td}'>{f_(v)}</td>" for v in row.values) + "</tr>")
    return (f"<div style='overflow-x:auto;font-size:.74rem;color:{TEXT_2};margin:-4px 0 8px'>"
            f"<table style='border-collapse:collapse;width:100%'>{head}{body}</table></div>")


def trend_grid(t: pd.DataFrame) -> str:
    rows = t[list("ABCDE")].rename(columns={k: v[0] for k, v in SERIES.items()}).T
    rows.loc["Progress (%)"] = t["Progress (%)"].values
    rows.loc["Remaining (%)"] = t["Remaining (%)"].values
    return grid_html(rows, {v[0]: v[1] for v in SERIES.values()}, ("Progress (%)", "Remaining (%)"))


def _with_total_row(ct: pd.DataFrame) -> pd.DataFrame:
    ct = ct.copy()
    ct.loc["Grand Total"] = ct.sum()
    return ct.astype(int)


def region_tables(df: pd.DataFrame, win: list) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pivot New Region: (1) total TT/bulan = bulan sebelumnya | bulan terakhir | Grand Total,
    (2) total TT all = Closed | Open | Grand Total. Keduanya dengan baris Grand Total."""
    w = df[df["_period"].isin(win)]
    found = set(w["_region"])
    regions = [r for r in REGION_ORDER if r in found] + sorted(found - set(REGION_ORDER))
    prev, last = split_labels(win)
    month = pd.DataFrame({last: w[w["_period"] == win[-1]].groupby("_region").size()})
    if prev:
        month.insert(0, prev, w[w["_period"] != win[-1]].groupby("_region").size())
    month = month.reindex(regions).fillna(0)
    month["Grand Total"] = month.sum(axis=1)
    status = pd.crosstab(w["_region"], w["_status"]).reindex(index=regions, columns=["Closed", "Open"], fill_value=0)
    status["Grand Total"] = status.sum(axis=1)
    return _with_total_row(month).rename_axis("New Region"), _with_total_row(status).rename_axis("New Region")


def toco_table(df: pd.DataFrame, win: list, col: str, closed_mode: bool = False) -> pd.DataFrame:
    """Pivot per Tower Company: bulan sebelumnya | bulan terakhir | Grand Total (urut terbesar), atau
    (closed_mode, seperti Stolen) total periode | Closed Total."""
    w = df[df["_period"].isin(win) & (df[col] != "(Kosong)")]
    if closed_mode:
        t = pd.DataFrame({f"{win[0]:%b %y} - {win[-1]:%b %y}": w.groupby(col).size(),
                          "Closed Total": w[w["_status"] == "Closed"].groupby(col).size()})
        t = t.fillna(0).sort_values(t.columns[0], ascending=False)
    else:
        prev, last = split_labels(win)
        t = pd.DataFrame({last: w[w["_period"] == win[-1]].groupby(col).size()})
        if prev:
            t.insert(0, prev, w[w["_period"] != win[-1]].groupby(col).size())
        t = t.fillna(0)
        t["Grand Total"] = t.sum(axis=1)
        t = t.sort_values("Grand Total", ascending=False)
    return t.astype(int).rename_axis("Tower Company")


def basis_chart(ct: pd.DataFrame, colors: list[str], height: int = 380, wrap_w: int = 12) -> go.Figure:
    """Batang berkelompok per kategori (region / ToCo), satu seri per kolom pivot."""
    x = [wrap(i, wrap_w) for i in ct.index]
    fig = go.Figure()
    for i, col in enumerate(ct.columns):
        fig.add_trace(go.Bar(x=x, y=ct[col], name=str(col), marker_color=colors[i % len(colors)],
                             text=_labels(ct[col]), textposition="outside", textangle=0,
                             textfont=dict(size=9 if len(ct) > 6 else 10, color=TEXT_2), cliponaxis=False,
                             hovertemplate=f"<b>%{{x}}</b><br>{col}: %{{y:,}}<extra></extra>"))
    dark(fig, height)
    fig.update_layout(barmode="group", bargap=0.22, bargroupgap=0.04,
                      yaxis=dict(range=[0, float(ct.to_numpy().max() or 1) * 1.15], tickformat=",.0f"),
                      xaxis=dict(showgrid=False, tickfont=dict(size=10)))
    return fig


# ---------- Jumlah TT: Achievement PM & Site Type per Region (sheet LIST SITE) ----------

def pm_table(sites: pd.DataFrame) -> pd.DataFrame:
    """ACHIEVEMENT PM per TLP: Total site K1 = COUNTIFS(Owner, TLP), NY PM / Done PM = COUNTIFS(Owner, TLP,
    status, "NY PM"/"DONE PM"), Achievement = Done ÷ Total site K1; + TOTAL, TLP/MO OTHERS, (All Site)."""
    g = pd.DataFrame({"Total site K1": sites.groupby("_toco").size(),
                      "NY PM / Walktest": sites[sites["_pm"] == "NY PM"].groupby("_toco").size(),
                      "Done PM / Walktest": sites[sites["_pm"] == "DONE PM"].groupby("_toco").size()}).fillna(0)
    main = g.drop(index=["OTHERS"], errors="ignore")
    order = [t for t in TOCO_ORDER if t in main.index] + sorted(t for t in main.index if t not in TOCO_ORDER)
    t = main.reindex(order)
    t.loc["TOTAL"] = main.sum()
    t.loc["TLP/MO OTHERS"] = g.loc["OTHERS"] if "OTHERS" in g.index else 0
    t.loc["(All Site)"] = t.loc["TOTAL"] + t.loc["TLP/MO OTHERS"]
    t = t.astype(int)
    t["Achievement (%)"] = t["Done PM / Walktest"] / t["Total site K1"].where(t["Total site K1"] > 0) * 100
    t.index.name = "TLP"
    return t


def _scale_color(v: float) -> str:
    """Skala 3 warna Excel: 0% merah -> 50% kuning -> 100% hijau."""
    if pd.isna(v):
        return "#ffffff"
    stops = [(0, (248, 105, 107)), (50, (255, 235, 132)), (100, (99, 190, 123))]
    v = min(max(v, 0), 100)
    (x0, c0), (x1, c1) = (stops[0], stops[1]) if v <= 50 else (stops[1], stops[2])
    f = (v - x0) / (x1 - x0)
    return "#" + "".join(f"{round(a + (b - a) * f):02X}" for a, b in zip(c0, c1))


def pm_table_html(t: pd.DataFrame, title: str) -> str:
    hd = "background:#26B5A7;color:#fff;font-weight:700;padding:6px 8px;border:1px solid #fff;text-align:center"
    td = "padding:4px 8px;border:1px solid #E3E5EE;text-align:center;font-size:.8rem"
    head = (f"<tr><th style='{hd};background:#FFFF66;color:{TEXT}' colspan='5'>ACHIEVEMENT PM</th></tr>"
            f"<tr><th style='{hd}' rowspan='2'>TLP</th><th style='{hd};background:#FF66CC' colspan='4'>{esc(title)}</th></tr>"
            f"<tr><th style='{hd}'>Total site K1</th><th style='{hd}'>NY PM / Walktest</th>"
            f"<th style='{hd}'>Done PM / Walktest</th><th style='{hd}'>Achievement</th></tr>")
    body = ""
    for name, r in t.iterrows():
        row_bg = {"TOTAL": "#26B5A7", "TLP/MO OTHERS": "#BFBFBF", "(All Site)": "#E2EFDA"}.get(name, "")
        weight = "font-weight:800;" if row_bg else ""
        color = "color:#fff;" if name == "TOTAL" else ""
        cell = f"{td};{weight}{color}" + (f"background:{row_bg};" if row_bg else "")
        ach = "" if name == "(All Site)" else pct(r["Achievement (%)"], 2)
        ach_bg = "" if name == "(All Site)" else f"background:{_scale_color(r['Achievement (%)'])};color:{TEXT};"
        body += (f"<tr><td style='{cell}'>{esc(name)}</td>"
                 + "".join(f"<td style='{cell}'>{fmt(r[c])}</td>"
                           for c in ["Total site K1", "NY PM / Walktest", "Done PM / Walktest"])
                 + f"<td style='{td};{weight}{ach_bg}'>{ach}</td></tr>")
    return ("<div style='overflow-x:auto'><table style='border-collapse:collapse;width:100%;font-family:inherit'>"
            f"{head}{body}</table></div>")


def k1_chart(t: pd.DataFrame) -> go.Figure:
    """Batang Total site K1 per TLP + garis Achievement PM (%) di sumbu kanan."""
    main = t.drop(index=["TOTAL", "TLP/MO OTHERS", "(All Site)"], errors="ignore")
    main = main.sort_values("Total site K1", ascending=False)
    x = [wrap(i) for i in main.index]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x, y=main["Total site K1"], name="Total site K1", marker_color=PURPLE,
                         text=bar_text(main["Total site K1"]), textposition="outside",
                         textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
                         hovertemplate="<b>%{x}</b><br>Total site K1: %{y:,}<extra></extra>"))
    has_pm = (t["NY PM / Walktest"].sum() + t["Done PM / Walktest"].sum()) > 0
    if has_pm:
        # Warna titik & label mengikuti skala conditional tabel (0% merah -> 50% kuning -> 100% hijau).
        ach_colors = [_scale_color(v) for v in main["Achievement (%)"]]
        fig.add_trace(go.Scatter(x=x, y=main["Achievement (%)"], name="Achievement", yaxis="y2",
                                 mode="lines+markers+text", text=[pct(v, 2) for v in main["Achievement (%)"]],
                                 textposition="top center", textfont=dict(size=10, color=TEXT), cliponaxis=False,
                                 line=dict(color="#8a8fa8", width=2),
                                 marker=dict(size=13, color=ach_colors, line=dict(color="#ffffff", width=1.5)),
                                 hovertemplate="<b>%{x}</b><br>Achievement: %{y:.2f}%<extra></extra>"))
    dark(fig, 420)
    fig.update_layout(yaxis=dict(range=[0, float(main["Total site K1"].max() or 1) * (2.0 if has_pm else 1.18)]),
                      yaxis2=dict(overlaying="y", side="right", range=[-60, 112], showgrid=False,
                                  tickvals=[0, 50, 100], ticksuffix="%", tickfont=dict(color=GOLD)),
                      xaxis=dict(tickangle=-30 if len(main) > 8 else 0, tickfont=dict(size=10)))
    return fig


def site_type_chart(sites: pd.DataFrame) -> tuple[go.Figure, pd.DataFrame]:
    """Pivot Count of IOH ID: Region x Site Type (MACRO, MICRO, IBC, ...)."""
    w = sites[(sites["_site"] != "") & (sites["_stype"] != "")]
    if w.empty:
        return empty_fig("Kolom Site Type belum ada / kosong di sheet LIST SITE.", 400), pd.DataFrame()
    ct = pd.crosstab(w["_region"], w["_stype"])
    cols = [c for c in ["MACRO", "MICRO", "IBC"] if c in ct.columns] + sorted(c for c in ct.columns
                                                                            if c not in ("MACRO", "MICRO", "IBC"))
    ct = ct[cols].sort_index()
    fig = go.Figure()
    for i, c in enumerate(cols):
        fig.add_trace(go.Bar(x=list(ct.index), y=ct[c], name=c, marker_color=EXCEL_COLORS[i % len(EXCEL_COLORS)],
                             text=_labels(ct[c]), textposition="outside", textangle=-90,
                             textfont=dict(size=10, color=TEXT_2), cliponaxis=False,
                             hovertemplate=f"<b>%{{x}}</b><br>{c}: %{{y:,}}<extra></extra>"))
    dark(fig, 420)
    fig.update_layout(barmode="group", bargroupgap=0.05,
                      yaxis=dict(range=[0, float(ct.to_numpy().max() or 1) * 1.2]), xaxis=dict(showgrid=False))
    out = ct.copy()
    out["Grand Total"] = out.sum(axis=1)
    return fig, _with_total_row(out).rename_axis("Region")


# TLP yang dinilai di COUPA: nilai kolom TLP2 (spasi dirapikan) -> label kolom tabel.
COUPA_TLP = {"CMI": "CMI", "MITRATEL": "Mitratel", "PTI": "Protelindo", "PTI - STP": "STP", "TBG": "TBG"}
COUPA_KPI = [
    # (no, Question/KPI, bobot, jenis, kategori, kriteria skor 1..4)
    (1, "TT Handling Completion (Severity Emergency)", 0.20, "handling", "CRITICAL",
     ["<85% TT Solved", "85% - 90% TT Solved", "91% - 95% TT Solved", ">95% TT Solved"]),
    (2, "TT Handling Completion (Severity Major)", 0.15, "handling", "MAJOR",
     ["<70% TT Solved", "70% - 80% TT Solved", "81% - 90% TT Solved", ">90% TT Solved"]),
    (3, "TT Handling Completion (Severity Minor)", 0.05, "handling", "MINOR",
     ["<70% TT Solved", "70% - 80% TT Solved", "81% - 90% TT Solved", ">90% TT Solved"]),
    (4, "Time To Resolved for TT Emergency", 0.25, "ttr", "CRITICAL",
     ["<70% TT resolved Meet SLA", "70% - 80% TT resolved Meet SLA", "81% - 90% TT resolved Meet SLA",
      ">90% TT resolved Meet SLA"]),
    (5, "Time To Resolved for TT Major", 0.15, "ttr", "MAJOR",
     ["<60% TT resolved Meet SLA", "60% - 70% TT resolved Meet SLA", "71% - 80% TT resolved Meet SLA",
      ">80% TT resolved Meet SLA"]),
    (6, "Time To Resolved for TT Minor", 0.05, "ttr", "MINOR",
     ["<60% TT resolved Meet SLA", "60% - 70% TT resolved Meet SLA", "71% - 80% TT resolved Meet SLA",
      ">80% TT resolved Meet SLA"]),
    (7, "Site Security", 0.15, "security", None,
     [">25% TT stolen Issue of site rented", "16% - 25% TT stolen Issue of site rented",
      "5% - 15% TT stolen Issue of site rented", "<5% TT stolen Issue of site rented"]),
]


def coupa_score(no: int, v: float) -> int | None:
    """Skor 1–4 sesuai tabel Scoring (batas atas tiap rentang mengikuti angka di Excel)."""
    if v is None or pd.isna(v):
        return None
    if no == 7:  # makin kecil makin baik
        return 4 if v < 5 else 3 if v <= 15 else 2 if v <= 25 else 1
    hi, mid, lo = {1: (95, 90, 85), 2: (90, 80, 70), 3: (90, 80, 70),
                   4: (90, 80, 70), 5: (80, 70, 60), 6: (80, 70, 60)}[no]
    return 4 if v > hi else 3 if v > mid else 2 if v >= lo else 1


def coupa_overview(tt: pd.DataFrame, site: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.Series]:
    """Performances Overview: persentase per KPI x TLP, skor 1–4, dan Result = SUMPRODUCT(skor, bobot).
    - TT Handling Completion = TT Closed ÷ total TT (Count of Site ID) per TLP2 & Kategori Issue.
    - Time To Resolved = Meet SLA ÷ TT Closed per TLP2 & Kategori Issue.
    - Site Security = jumlah TT Problem type "Stolen Case" ÷ Total Site (sheet Site) per TLP2."""
    df = tt.copy()
    df["_tlp"] = df["TLP2"].astype("string").str.split().str.join(" ").str.upper()
    df = df[df["_tlp"].isin(COUPA_TLP) & df["Site ID"].notna()]
    df["_kat"] = df["Kategori Issue"].astype("string").str.strip().str.upper()
    df["_closed"] = df["Status"].astype("string").str.strip().str.lower().eq("closed")
    df["_meet"] = df["SLA"].astype("string").str.strip().str.lower().eq("meet sla")
    df["_stolen"] = df["Problem type"].astype("string").str.strip().str.lower().eq("stolen case")
    totals = {}
    if not site.empty:
        grp = site.iloc[:, 0].astype("string").str.split().str.join(" ").str.upper()
        totals = dict(zip(grp, pd.to_numeric(site.iloc[:, 1], errors="coerce")))
    pc = pd.DataFrame(index=[k[1] for k in COUPA_KPI], columns=list(COUPA_TLP.values()), dtype=float)
    for no, name, _w, kind, kat, _crit in COUPA_KPI:
        for code, label in COUPA_TLP.items():
            g = df[df["_tlp"] == code]
            if kind == "handling":
                sub = g[g["_kat"] == kat]
                v = sub["_closed"].mean() * 100 if len(sub) else float("nan")
            elif kind == "ttr":
                sub = g[(g["_kat"] == kat) & g["_closed"]]
                v = sub["_meet"].mean() * 100 if len(sub) else float("nan")
            else:
                tot = totals.get(code)
                v = g["_stolen"].sum() / tot * 100 if tot else float("nan")
            pc.loc[name, label] = v
    scores = pd.DataFrame({c: [coupa_score(k[0], pc.loc[k[1], c]) for k in COUPA_KPI] for c in pc.columns},
                          index=pc.index)
    weights = pd.Series([k[2] for k in COUPA_KPI], index=pc.index)
    result = scores.mul(weights, axis=0).sum(min_count=1)
    return pc, scores, result


def coupa_result_color(v) -> str:
    return "#16A34A" if v >= 4 else "#FFD100" if v >= 3.5 else RED


def coupa_table_html(pc: pd.DataFrame, scores: pd.DataFrame, result: pd.Series) -> str:
    tlps = list(pc.columns)
    th = "background:#E20074;color:#fff;font-weight:700;padding:6px 8px;border:1px solid #fff;text-align:center"
    th2 = "background:#00A3AD;color:#fff;font-weight:700;padding:6px 8px;border:1px solid #fff;text-align:center"
    td = "padding:6px 8px;border:1px solid #E3E5EE;text-align:center;font-size:.82rem"
    head = (f"<tr><th style='{th}' colspan='2'>KPI: Maintenance</th><th style='{th}' colspan='{len(tlps) + 1}'>Answer"
            f"</th><th style='{th2}' colspan='{len(tlps)}'>Score</th></tr>"
            f"<tr><th style='{th}'>No.</th><th style='{th}'>Question/KPI</th>"
            + "".join(f"<th style='{th}'>{esc(t)} (in %)</th>" for t in tlps)
            + f"<th style='{th}'>Weighting</th>" + "".join(f"<th style='{th2}'>{esc(t)} (in #)</th>" for t in tlps)
            + "</tr>")
    body = ""
    for no, name, w, *_ in COUPA_KPI:
        body += (f"<tr><td style='{td}'>{no}</td><td style='{td};text-align:left'>{esc(name)}</td>"
                 + "".join(f"<td style='{td}'>{pct(pc.loc[name, t], 2)}</td>" for t in tlps)
                 + f"<td style='{td};font-weight:700'>{pct(w * 100, 1)}</td>"
                 + "".join(f"<td style='{td};font-weight:700'>{'' if pd.isna(scores.loc[name, t]) else int(scores.loc[name, t])}"
                           "</td>" for t in tlps) + "</tr>")
    total_w = sum(k[2] for k in COUPA_KPI) * 100
    res = "".join(f"<td style='{td};font-weight:800;background:{coupa_result_color(result[t])};"
                  f"color:{'#1a1c2e' if coupa_result_color(result[t]) == '#FFD100' else '#fff'}'>"
                  f"{days(result[t])}</td>" for t in tlps)
    body += (f"<tr><td style='{td}' colspan='{len(tlps) + 2}'></td><td style='{td};font-weight:700'>{pct(total_w, 0)}</td>"
             f"<td style='{td}' colspan='{len(tlps)}'></td></tr>"
             f"<tr><td style='{td};text-align:right;font-weight:800' colspan='{len(tlps) + 3}'>Result</td>{res}</tr>")
    return ("<div style='overflow-x:auto'><table style='border-collapse:collapse;width:100%;font-family:inherit'>"
            f"{head}{body}</table></div>")


def coupa_scoring_html() -> str:
    th = "background:#16A34A;color:#fff;font-weight:700;padding:6px 8px;border:1px solid #fff;text-align:center"
    td = "padding:6px 8px;border:1px solid #E3E5EE;text-align:center;font-size:.8rem"
    head = (f"<tr><th style='{th}'>No.</th><th style='{th}'>Weighting</th><th style='{th}'>Comment</th>"
            + "".join(f"<th style='{th}'>{i}</th>" for i in range(1, 5)) + "</tr>")
    body = "".join(f"<tr><td style='{td}'>{no}</td><td style='{td}'>{pct(w * 100, 0)}</td>"
                   f"<td style='{td}'>{'Based on TT Handling' if kind == 'handling' else 'Based on TT'}</td>"
                   + "".join(f"<td style='{td}'>{esc(c)}</td>" for c in crit) + "</tr>"
                   for no, _n, w, kind, _k, crit in COUPA_KPI)
    return ("<div style='overflow-x:auto'><table style='border-collapse:collapse;width:100%;font-family:inherit'>"
            f"{head}{body}</table></div>")


COUPA_COLORS = ["#5B9BD5", ORANGE, "#A5A5A5", YELLOW, "#4472C4", "#70AD47", "#264478"]


def coupa_chart(pc: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for i, (no, name, *_r) in enumerate(COUPA_KPI):
        vals = [pc.loc[name, t] for t in pc.columns]
        fig.add_trace(go.Bar(x=[f"{t} (in %)" for t in pc.columns], y=vals, name=name,
                             marker_color=COUPA_COLORS[i % len(COUPA_COLORS)],
                             text=[pct(v, 2) if pd.notna(v) else "" for v in vals],
                             textposition="outside", textangle=-90, textfont=dict(size=10, color=TEXT_2),
                             cliponaxis=False, hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:.2f}}%<extra></extra>"))
    dark(fig, 520, legend_top=False)
    fig.update_layout(barmode="group", bargroupgap=0.04, margin=dict(t=110, b=8),
                      legend=dict(orientation="h", x=0, xanchor="left", y=1.0, yanchor="bottom",
                                  entrywidth=0.33, entrywidthmode="fraction", font=dict(size=10)),
                      yaxis=dict(range=[0, 125], ticksuffix="%", tickformat=".0f"), xaxis=dict(showgrid=False))
    return fig


def _cw_key(cw: str) -> tuple[int, int]:
    m = re.match(r"(\d{4})\D*CW\s*(\d+)", str(cw), flags=re.I)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def chart_card(key: str, badge: str, color: str, title: str, fig, exports: list | None = None,
               table: pd.DataFrame | None = None, pie_labels: str = "value_percent", below: str | None = None) -> None:
    """Kartu grafik. `fig` boleh berupa fungsi: dipanggil di dalam kartu, sehingga widget
    pengatur (mis. dropdown) bisa tampil di antara judul dan grafik.
    Bila `exports` diberikan, angka grafik disimpan untuk Export Excel & PPT; `table` mengganti
    tabel Excel bawaan bila perlu kolom yang lebih jelas; `below` = HTML (mis. tabel data) di bawah grafik."""
    with st.container(key=f"card_{key}"):
        st.markdown(f"<div class='ctitle'><span class='badge' style='--c:{color}'>{esc(badge)}</span>"
                    f"{esc(title)}</div>", unsafe_allow_html=True)
        if callable(fig):
            fig = fig()
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "displayModeBar": False},
                        key=f"fig_{key}")
        if below:
            st.markdown(below, unsafe_allow_html=True)
    if exports is not None:
        spec = fig_export(fig)
        if spec is not None:
            spec["pie_labels"] = pie_labels
            exports.append(dict(title=title, spec=spec, table=table if table is not None else spec_table(spec)))


def section(title: str) -> None:
    st.markdown(f"<div class='section'><span class='bar'></span>{esc(title)}</div>", unsafe_allow_html=True)


# =====================================================================================
# Export PowerPoint
# =====================================================================================

def _num_format(values, pct_: bool) -> str:
    if pct_:
        return '0.0"%"'
    vals = [v for v in values if v is not None]
    return "#,##0" if all(float(v).is_integer() for v in vals) else "#,##0.00"


def _add_native_chart(slide, x, y, w, h, spec: dict) -> None:
    """Grafik PowerPoint asli (bisa diedit di PowerPoint) dari spesifikasi fig_export."""
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.util import Pt

    def rgb(h_: str):
        return RGBColor.from_string(str(h_).lstrip("#")[:6])

    kinds = {"pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT, "bar_h": XL_CHART_TYPE.BAR_CLUSTERED,
             "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED, "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
             "column_stacked_100": XL_CHART_TYPE.COLUMN_STACKED_100}
    drawn = [s_ for s_ in spec["series"] if s_["chart"]]
    cd = CategoryChartData()
    cd.categories = spec["categories"]
    for s_ in drawn:
        cd.add_series(s_["name"], s_["values"], number_format=_num_format(s_["values"], spec.get("pct")))
    chart = slide.shapes.add_chart(kinds[spec["kind"]], x, y, w, h, cd).chart
    chart.font.size, chart.font.name = Pt(11), "Segoe UI"
    chart.font.color.rgb = rgb(TEXT_2)
    chart.has_title = False  # judul sudah ada di slide; cegah judul otomatis dari nama seri
    is_pie = spec["kind"] in ("pie", "doughnut")
    chart.has_legend = is_pie or len(drawn) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(11)
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(10)
    if is_pie:
        # Label ditulis sendiri ("1.390; 51,2%") agar format angka sama dengan dashboard.
        colors = drawn[0]["colors"] or PALETTE
        vals = [v or 0 for v in drawn[0]["values"]]
        total = sum(vals) or 1
        for i, point in enumerate(plot.series[0].points):
            color = colors[i % len(colors)]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = rgb(color)
            share = pct(vals[i] / total * 100, 1)
            label = share if spec.get("pie_labels") == "percent" else f"{fmt(vals[i])}; {share}"
            tf = point.data_label.text_frame
            tf.text = label
            run = tf.paragraphs[0].runs[0]
            run.font.size, run.font.bold = Pt(12), True
            run.font.color.rgb = rgb(TEXT if color in (YELLOW, ORANGE) else "#FFFFFF")
            # Posisi label hanya untuk pie; doughnut tidak mendukung dLblPos (file jadi rusak bila diisi).
            if spec["kind"] == "pie":
                small = vals[i] / total <= 0.06
                point.data_label.position = XL_LABEL_POSITION.OUTSIDE_END if small else XL_LABEL_POSITION.CENTER
                if small:
                    run.font.color.rgb = rgb(TEXT)
        return
    dl.number_format = _num_format([v for s_ in drawn for v in s_["values"]], spec.get("pct"))
    dl.number_format_is_linked = False
    if spec["kind"] in ("bar_h", "column_clustered"):
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.gap_width = 60
    for i, (ser, s_) in enumerate(zip(plot.series, drawn)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(s_["color"] or PALETTE[i % len(PALETTE)])
        if s_["colors"]:
            for j, point in enumerate(ser.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = rgb(s_["colors"][j % len(s_["colors"])])
    if spec["kind"] == "bar_h":
        chart.category_axis.reverse_order = True  # nilai terbesar di atas
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb("#E3E5EE")
    va.tick_labels.font.size = Pt(10)
    chart.category_axis.tick_labels.font.size = Pt(10)


def build_pptx(items: list[dict], meta: dict, kpis: list[tuple[str, str, str]], filter_desc: str) -> bytes:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Emu, Inches, Pt

    def rgb(h: str) -> RGBColor:
        return RGBColor.from_string(h.lstrip("#")[:6])

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    def background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
        bg.fill.solid(); bg.fill.fore_color.rgb = rgb(BG); bg.line.fill.background()
        stripe_w = Emu(int(W / 4))
        for i, c in enumerate([MAGENTA, YELLOW, RED, TOSCA]):
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, stripe_w * i, 0, stripe_w, Inches(0.09))
            s.fill.solid(); s.fill.fore_color.rgb = rgb(c); s.line.fill.background()

    def text(slide, x, y, w, h, value, size, color=TEXT, bold=False, align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        for i, line in enumerate(value if isinstance(value, list) else [value]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = line
            r.font.size, r.font.bold, r.font.name = Pt(size), bold, "Segoe UI"
            r.font.color.rgb = rgb(color)
        return tb

    # 1. Sampul
    s = prs.slides.add_slide(blank)
    background(s)
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.0), Inches(0.12), Inches(2.2))
    accent.fill.solid(); accent.fill.fore_color.rgb = rgb(MAGENTA); accent.line.fill.background()
    text(s, Inches(1.1), Inches(1.9), Inches(11), Inches(1.0), "TLP Dashboard", 40, TEXT, True)
    text(s, Inches(1.1), Inches(2.85), Inches(11), Inches(0.6),
         "Laporan Access Issue · Stolen · TT CME", 20, MAGENTA)
    text(s, Inches(1.1), Inches(3.55), Inches(11), Inches(1.2), [
        f"Data diunggah oleh: {meta.get('uploader', '-')}",
        f"Update data terakhir: {meta.get('timestamp', '-')}  ·  File: {meta.get('filename', '-')}",
    ], 14, TEXT_2)
    text(s, Inches(1.1), Inches(5.3), Inches(11.5), Inches(0.9), f"Filter: {filter_desc}", 11, MUTED)
    text(s, Inches(1.1), Inches(6.6), Inches(11), Inches(0.4),
         f"Dibuat {datetime.now():%d-%m-%Y %H:%M}", 10, MUTED)

    # 2. Ringkasan KPI
    s = prs.slides.add_slide(blank)
    background(s)
    text(s, Inches(0.6), Inches(0.35), Inches(12), Inches(0.7), "Ringkasan Kinerja", 26, TEXT, True)
    gap, top = Inches(0.25), Inches(1.6)
    card_w = Emu(int((W - Inches(1.2) - gap * 2) / 3))
    card_h = Inches(2.2)
    for i, (label, value, sub) in enumerate(kpis):
        x = Inches(0.6) + (card_w + gap) * (i % 3)
        y = top + (card_h + gap) * (i // 3)
        color = [MAGENTA, TOSCA, RED, GOLD, PURPLE, TOSCA][i % 6]
        box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
        box.adjustments[0] = 0.08
        box.fill.solid(); box.fill.fore_color.rgb = rgb("#f6f7fb"); box.line.color.rgb = rgb(color)
        text(s, x + Inches(0.25), y + Inches(0.2), card_w - Inches(0.5), Inches(0.4), label.upper(), 12, TEXT_2, True)
        text(s, x + Inches(0.25), y + Inches(0.65), card_w - Inches(0.5), Inches(0.9), value, 30, color, True)
        text(s, x + Inches(0.25), y + Inches(1.55), card_w - Inches(0.5), Inches(0.5), sub, 12, MUTED)

    # 3+. Satu grafik per slide — grafik PowerPoint asli (tidak butuh Chrome/kaleido, bisa diedit).
    for item in items:
        s = prs.slides.add_slide(blank)
        background(s)
        text(s, Inches(0.6), Inches(0.3), Inches(12), Inches(0.7), item["title"], 22, TEXT, True)
        _add_native_chart(s, Inches(0.6), Inches(1.15), W - Inches(1.2), H - Inches(1.6), item["spec"])

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_excel(items: list[dict], extras: dict[str, pd.DataFrame], meta: dict, filter_desc: str) -> bytes:
    """Satu sheet per grafik berisi tabel angka grafik, ditambah tabel perhitungan pendukung."""
    buf = io.BytesIO()
    used: set[str] = set()

    def sheet_name(title: str) -> str:
        base = re.sub(r"[\[\]:*?/\\]", " ", title).strip()[:28] or "Sheet"
        name, i = base, 2
        while name.lower() in used:
            name, i = f"{base[:25]} ({i})", i + 1
        used.add(name.lower())
        return name

    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        wb = xw.book
        title_fmt = wb.add_format({"bold": True, "font_size": 13, "font_color": MAGENTA})
        info = pd.DataFrame({"Keterangan": ["Laporan", "Data diunggah oleh", "Update data terakhir", "File data",
                                            "Filter", "Dibuat"],
                             "Nilai": ["TLP Dashboard", meta.get("uploader", "-"), meta.get("timestamp", "-"),
                                       meta.get("filename", "-"), filter_desc, f"{datetime.now():%d-%m-%Y %H:%M}"]})
        info.to_excel(xw, sheet_name=sheet_name("Info"), index=False)
        xw.sheets["Info"].set_column(0, 0, 22)
        xw.sheets["Info"].set_column(1, 1, 90)
        for title, table in [(it["title"], it["table"]) for it in items] + list(extras.items()):
            name = sheet_name(title)
            table.to_excel(xw, sheet_name=name, startrow=2)
            ws = xw.sheets[name]
            ws.write(0, 0, title, title_fmt)
            idx_w = max([len(str(table.index.name or ""))] + [len(str(v)) for v in table.index]) + 2
            ws.set_column(0, 0, min(max(idx_w, 12), 60))
            for j, col in enumerate(table.columns, start=1):
                ws.set_column(j, j, min(max(len(str(col)) + 2, 12), 40))
    return buf.getvalue()


# =====================================================================================
# Header: judul, status data live, tombol aksi
# =====================================================================================

source_path, source_meta = resolve_source()
current_sig = source_signature()
if st.session_state.get("loaded_sig") not in (None, current_sig):
    st.toast(f"📥 Data baru diunggah oleh {source_meta.get('uploader', '-')} — dashboard diperbarui.")
st.session_state["loaded_sig"] = current_sig

# Dropdown/slider di dalam tab (ikut di-reset & menentukan ulang isi PPT).
WIDGET_KEYS = ("rpt_weeks", "rpt_range", "rpt_owner", "stl_range", "cme_range", "ltoa_toco", "imp_toco", "rc_group1",
               "stl_owner", "st_owner")
FILTER_KEYS = ("f_cat", "f_year", "f_region", "f_tlp", "f_cause", "f_status", "f_month", "f_sev", "f_backlog", "f_q")

head_l, head_r = st.columns([1.05, 1], vertical_alignment="center")
with head_r:
    b1, b2, b3, b4, b5 = st.columns([1.15, 1.0, 1.0, 1.05, 1.0])
    with b1.popover("📂 Update Data", width="stretch"):
        st.markdown("**Unggah file mingguan terbaru.** File menggantikan data untuk semua pengguna.")
        uploader = st.text_input("Nama pengunggah", key="uploader_name", placeholder="mis. Ayu")
        nonce = st.session_state.get("upload_nonce", 0)
        up = st.file_uploader("File .xlsb", type=["xlsb"], key=f"upload_{nonce}")
        if st.button("⬆️ Upload & Kunci Data", type="primary", width="stretch",
                     disabled=up is None or not uploader.strip()):
            try:
                with st.spinner("Menyimpan & memvalidasi file…"):
                    entry = save_upload(up.getvalue(), uploader.strip(), up.name)
                st.session_state["upload_nonce"] = nonce + 1
                st.toast(f"✅ {entry['filename']} tersimpan — semua pengguna akan ter-refresh.")
                st.rerun()
            except Exception as exc:  # file rusak / lock aktif / sheet tidak lengkap
                st.error(str(exc))
        if not uploader.strip():
            st.caption("Isi nama pengunggah agar tercatat di log.")
        log = read_log()
        if log:
            st.markdown("**Riwayat unggahan**")
            st.dataframe(pd.DataFrame(log[::-1]).head(8)[["timestamp", "uploader", "filename"]],
                         hide_index=True, width="stretch")
    with b2.popover("📘 Panduan", width="stretch"):
        st.markdown(
            "1. Klik **Update Data**, isi nama, pilih file `.xlsb` mingguan, lalu **Upload & Kunci Data**.\n"
            "2. File disimpan ke `shared_data/latest_tlp_data.xlsb`; log pengunggah di `upload_log.json`.\n"
            f"3. Semua pengguna yang membuka dashboard otomatis ter-refresh (cek tiap {REFRESH_EVERY}).\n"
            "4. Sheet wajib: **List Access Issue**, **List Stolen**, **TT CME** (header di baris pertama).\n"
            "   SoW (A) dibaca dari sheet **performance access** bila ada.\n\n"
            "Grafik Trends A–E: TT berstatus *Cancel* tidak dihitung (sama dengan pivot Excel).")
    xlsx_slot = b3.empty()
    ppt_slot = b4.empty()
    if b5.button("↻ Reset Filter", width="stretch"):
        for k in [k for k in st.session_state if str(k).startswith(FILTER_KEYS) or k in WIDGET_KEYS]:
            st.session_state.pop(k, None)
        st.rerun()

with head_l:
    pill = f"<span class='pill'>{esc(source_meta.get('filename', ''))}</span>" if source_path else ""
    st.markdown(f"<div class='hdr'><div class='logo'>TLP</div><div><h1>TLP Dashboard{pill}</h1>"
                f"<div class='sub'>Access Issue · Stolen · TT CME — dibaca dengan pyxlsb</div></div></div>",
                unsafe_allow_html=True)


@st.fragment(run_every=REFRESH_EVERY)
def live_status() -> None:
    """Badge lock terakhir + auto-refresh: bila file bersama berubah, seluruh app dijalankan ulang."""
    _, meta = resolve_source()
    lock = active_lock()
    if lock:
        st.markdown(f"<span class='lock' style='border-color:{YELLOW}'><span class='dot' "
                    f"style='background:{YELLOW}'></span>⏳ Sedang diunggah oleh <b>{esc(lock.get('uploader'))}</b>"
                    "</span>", unsafe_allow_html=True)
    else:
        st.markdown(f"<span class='lock'><span class='dot'></span>🔒 Lock terakhir: "
                    f"<b>{esc(meta.get('uploader', '-'))}</b> · {esc(meta.get('timestamp', '-'))}"
                    f" · {esc(meta.get('filename', '-'))}</span>", unsafe_allow_html=True)
    if source_signature() != st.session_state.get("loaded_sig"):
        st.rerun(scope="app")


live_status()

if source_path is None:
    st.info(f"Belum ada file .xlsb. Letakkan file di folder **{APP_DIR}** atau klik **Update Data**.")
    st.stop()

try:
    prepared, sow_from_file, SITES, COUPA = load_data(f"{current_sig}:v{PREP_VERSION}", str(source_path))
except Exception as exc:  # file rusak / bukan xlsb
    st.error(f"Gagal membaca **{source_meta.get('filename')}**: {exc}")
    st.stop()
if not prepared:
    st.error("Sheet List Access Issue / List Stolen / TT CME tidak ditemukan di file ini.")
    st.stop()

# SoW (A) = Total site K1 dari LIST SITE (bila ada), lalu tabel di sheet performance, lalu nilai cadangan.
SITE_K1 = SITES[SITES["_toco"] != "OTHERS"].groupby("_toco").size() if len(SITES) else pd.Series(dtype=int)
SOW = {**DEFAULT_SOW, **sow_from_file, **{k: float(v) for k, v in SITE_K1.items()}, **load_sow_override()}


# =====================================================================================
# Bar filter horizontal (multiselect)
# =====================================================================================

def options(frames: list[pd.DataFrame], col: str, by_count: bool = True) -> list[str]:
    s = pd.concat([f[col] for f in frames]) if frames else pd.Series(dtype="string")
    s = s[s != "(Kosong)"]
    return [str(v) for v in (s.value_counts().index if by_count else sorted(s.dropna().unique()))]


with st.container(key="filterbar"):
    r1 = st.columns([1.3, 0.9, 1.1, 1.2, 1.5])
    cats_avail = list(prepared)
    f_cat = r1[0].multiselect("Kategori", cats_avail, default=[cats_avail[0]], key="f_cat",
                              format_func=lambda c: f"{CATEGORIES[c]['icon']} {c}", placeholder="Semua kategori")
    cats = f_cat or cats_avail
    frames = [prepared[c] for c in cats]
    base = pd.concat(frames, ignore_index=True)
    f_year = r1[1].multiselect("Years", sorted(options(frames, "_year"), reverse=True), key="f_year",
                               placeholder="Semua Tahun")
    f_region = r1[2].multiselect("Region", options(frames, "_region"), key="f_region", placeholder="Semua Region")
    f_tlp = r1[3].multiselect("TLP (ToCo)", options(frames, "_toco"), key="f_tlp", placeholder="Semua TLP")
    cause_label = CATEGORIES[cats[0]]["cause"] if len(cats) == 1 else "Cause"
    f_cause = r1[4].multiselect(f"Cause · {cause_label}" if len(cats) == 1 else "Cause",
                                options(frames, "_cause"), key="f_cause", placeholder="Semua Cause")
    r2 = st.columns([1.1, 1.3, 1.0, 1.0, 1.6])
    f_status = r2[0].multiselect("Status", ["Open", "Closed", "Cancel"], key="f_status",
                                 placeholder="Open + Closed",
                                 help="Kosong = Open + Closed (tanpa Cancel, sama dengan pivot Summary).")
    months_all = sorted(base["_period"].unique())
    f_month = r2[1].multiselect("Bulan (Open)", [p.strftime("%b_%Y") for p in months_all[::-1]], key="f_month",
                                placeholder="Semua Bulan")
    f_sev = r2[2].multiselect("Severity", options(frames, "_sev"), key="f_sev", placeholder="Semua Severity")
    if "TT CME" in cats:
        bl_opts = options([prepared["TT CME"]], "_backlog")
        f_backlog = r2[3].multiselect("Backlog (CME)", bl_opts, key="f_backlog", placeholder="Semua Backlog",
                                      default=[o for o in bl_opts if o.lower() == "normal tt"])
    else:
        f_backlog = []
        r2[3].multiselect("Backlog (CME)", [], key="f_backlog_off", placeholder="Khusus TT CME", disabled=True)
    f_q = r2[4].text_input("Cari Site / Issue", placeholder="Ketik Site ID, nama, TT, deskripsi…", key="f_q")

FILTERS = dict(years=f_year, regions=f_region, tlp=f_tlp, cause=f_cause, sev=f_sev, backlog=f_backlog,
               q=f_q.strip(), months=f_month, status=f_status)
d = apply_filters(base, FILTERS)

filter_parts = [f"Kategori: {', '.join(cats)}"]
for label, vals in [("Tahun", f_year), ("Region", f_region), ("TLP", f_tlp), ("Cause", f_cause),
                    ("Status", f_status or ["Open", "Closed"]), ("Bulan", f_month), ("Severity", f_sev),
                    ("Backlog", f_backlog)]:
    if vals:
        filter_parts.append(f"{label}: {', '.join(map(str, vals))}")
if FILTERS["q"]:
    filter_parts.append(f"Cari: {FILTERS['q']}")
filter_desc = " · ".join(filter_parts)

export_cols = [c for c in d.columns if not c.startswith("_")]


# =====================================================================================
# Kartu KPI
# =====================================================================================

total = len(d)
closed = d[d["_status"] == "Closed"]
open_ = d[d["_status"] == "Open"]
close_rate = len(closed) / total * 100 if total else float("nan")
mttr = closed["_aging"].mean()
toco_vc = d["_toco"].value_counts()
cause_vc = d["_cause"].value_counts()
top_toco = (toco_vc.index[0], toco_vc.iloc[0]) if len(toco_vc) else ("-", 0)
top_cause = (cause_vc.index[0], cause_vc.iloc[0]) if len(cause_vc) else ("-", 0)
cat_title = cats[0] if len(cats) == 1 else "Gabungan"
KPI_LIST = [
    (f"Total TT {cat_title}", fmt(total), f"Dari {fmt(len(base))} total record"),
    ("Ticket Closed", fmt(len(closed)), f"Close Rate: {pct(close_rate, 2)}"),
    ("Ticket Masih Open", fmt(len(open_)), f"Avg Aging: {days(open_['_aging'].mean())} hari"),
    ("Overall MTTR", f"{days(mttr)} d", f"Dari {fmt(len(closed))} tiket CLOSE"),
    ("Top TLP Volume", str(top_toco[0]), f"{fmt(top_toco[1])} TT ({pct(top_toco[1] / max(total, 1) * 100)})"),
    ("Top Cause", str(top_cause[0]), f"{fmt(top_cause[1])} TT ({pct(top_cause[1] / max(total, 1) * 100)})"),
]
KPI_COLORS = [MAGENTA, TOSCA, RED, GOLD, PURPLE, TOSCA]


def kpi_html(label, value, sub, color, text_value=False) -> str:
    cls = "val txt" if text_value else "val"
    return (f"<div class='kpi' style='--c:{color}'><div class='lbl'>{esc(label)}</div>"
            f"<div class='{cls}' title='{esc(value)}'>{esc(value)}</div><div class='sub'>{esc(sub)}</div></div>")


st.markdown("<div class='kpis'>" + "".join(
    kpi_html(lbl, val, sub, KPI_COLORS[i], i >= 4) for i, (lbl, val, sub) in enumerate(KPI_LIST)
) + "</div>", unsafe_allow_html=True)

if d.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini. Klik **Reset Filter** untuk mengembalikan.")
    st.stop()


# =====================================================================================
# Executive insight
# =====================================================================================

def insights(d: pd.DataFrame) -> list[tuple[str, str, str]]:
    n = max(len(d), 1)
    r_top = d["_region"].value_counts().head(2)
    open_toco = d[d["_status"] == "Open"]["_toco"].value_counts().head(2)
    open_txt = (" (terbanyak di " + " & ".join(f"<b>{esc(k)}</b> {fmt(v)} TT" for k, v in open_toco.items()) + ")"
                if len(open_toco) else "")
    p1 = (" & ".join(f"Region <b>{esc(k)} ({fmt(v)} TT)</b>" for k, v in r_top.items())
          + f" menyumbang <b>{pct(r_top.sum() / n * 100)}</b> dari total tiket. "
          f"Saat ini tersisa <b>{fmt((d['_status'] == 'Open').sum())} tiket OPEN</b>{open_txt}.")

    cl = d[d["_status"] == "Closed"]
    g = cl.groupby("_toco")["_aging"].agg(["mean", "size"]).dropna().sort_values("size", ascending=False).head(4)
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

    tv = d["_toco"].value_counts().head(4)
    p3 = (f"<b>{len(tv)} TLP teratas</b> (" + ", ".join(f"{esc(k)} {fmt(v)}" for k, v in tv.items())
          + f") mendominasi <b>{pct(tv.sum() / n * 100)}</b> dari seluruh tiket.")

    roots = d.loc[d["_root"] != "(Kosong)", "_root"].value_counts().head(3)
    src = roots if len(roots) else d["_cause"].value_counts().head(3)
    p4 = (f"{len(src)} isu teratas — " + ", ".join(f"<b>{esc(k)} ({fmt(v)} TT)</b>" for k, v in src.items())
          + f" — menyumbang <b>{pct(src.sum() / max(len(d), 1) * 100)}</b> dari seluruh tiket.")
    return [("📍 Regional & TLP Open/Close", p1, MAGENTA), ("⏱️ Performa MTTR per TLP", p2, YELLOW),
            ("🏢 Konsentrasi TLP (Pareto)", p3, TOSCA), ("🔎 Root Cause Dominan", p4, RED)]


# =====================================================================================
# Tab navigasi
# =====================================================================================

EXPORTS: list[dict] = []                # angka tiap grafik (untuk Export Excel & PPT)
EXTRA_TABLES: dict[str, pd.DataFrame] = {}  # tabel perhitungan pendukung (hanya Excel)


def tab_filters(cat: str) -> dict:
    """Filter global untuk satu kategori. Cause & Severity hanya berlaku bila kategori itu dipilih (nilainya
    berbeda per sheet); status Open + Closed; TT CME default Backlog = Normal TT (seperti pivot Excel)."""
    f = {**FILTERS, "status": [x for x in FILTERS["status"] if x in ("Open", "Closed")] or ["Open", "Closed"]}
    if cat not in cats:
        f.update(cause=[], sev=[])
    if cat == "TT CME":
        f["backlog"] = FILTERS["backlog"] or ["Normal TT"]
    return f


def trend_base(cat: str) -> pd.DataFrame:
    """Dasar grafik Trends/Latest update: Access = Count of Case Status, Stolen = Count of TT IOH,
    CME = Count of TT Number (baris yang nomor TT-nya terisi)."""
    df = apply_filters(prepared[cat], tab_filters(cat), month=False)
    return df[df["_has_tt"]] if CATEGORIES[cat].get("count_tt") else df


def trend_section(df: pd.DataFrame, key: str, noun: str, noun_toco: str, toco_col: str, toco_colors: list[str],
                  top: int, closed_mode: bool = False, owner_pick: bool = False) -> None:
    """Trends [General] (A–E) + Latest update [Region Basis] (2 grafik) + Latest update [ToCo Basis]."""
    periods = sorted(df["_period"].unique())
    if not periods:
        st.info("Tidak ada data bulan pada filter ini.")
        return
    labels = [p.strftime("%b_%Y") for p in periods]
    win = periods
    if len(labels) > 1:
        lo, hi = st.select_slider("Periode (bulan open)", options=labels, key=f"{key}_range",
                                  value=(labels[max(0, len(labels) - 13)], labels[-1]))
        win = periods[labels.index(lo): labels.index(hi) + 1]
    span = span_label(win)
    t = backlog_summary(df[df["_period"].isin(win)], win)
    chart_card(f"{key}_gen", "GENERAL", MAGENTA, f"Trends of {noun} for period {span} [General]", trend_chart(t),
               EXPORTS, table=t.rename(columns={k: v[0] for k, v in SERIES.items()}), below=trend_grid(t))

    section(f"Latest update of {noun} for period {span} [Region Basis]")
    reg_df = df
    if owner_pick:
        owners = df.loc[df["_period"].isin(win), "_owner"].value_counts()
        owners = owners[owners.index != "(Kosong)"]
        sel = st.multiselect("Site Owner (grafik Region Basis)", list(owners.index), key=f"{key}_owner",
                             placeholder="Semua Site Owner", format_func=lambda o: f"{o} ({fmt(owners.get(o))} TT)",
                             help="Dari kolom Site Owner di sheet List Access Issue.")
        if sel:
            reg_df = df[df["_owner"].isin(sel)]
    month_t, status_t = region_tables(reg_df, win)
    a, b = st.columns(2)
    for col_, key_, tbl, sub, colors in [(a, "reg_m", month_t, "Total TT per Month", [TEAL, RED, GREY]),
                                         (b, "reg_s", status_t, "Total TT All (Closed/Open)", [TEAL, RED, GREY])]:
        with col_:
            chart_card(f"{key}_{key_}", "REGION", TOSCA, f"{noun} [Region Basis] · {sub} · {span}",
                       basis_chart(tbl, colors), EXPORTS, table=tbl,
                       below=grid_html(tbl.T, dict(zip(tbl.columns, colors))))

    tt = toco_table(df, win, toco_col, closed_mode)
    show = tt.head(top)
    chart_card(f"{key}_toco", "TOCO", GOLD, f"Latest update of {noun_toco} for period {span} [ToCo Basis]",
               basis_chart(show, toco_colors, 420, 11), EXPORTS, table=tt,
               below=grid_html(show.T, dict(zip(show.columns, toco_colors))))
    with st.expander(f"Tabel lengkap [ToCo Basis] · {len(tt)} Tower Company"):
        st.dataframe(tt, width="stretch")


tabs = st.tabs(["📊 Executive Summary", "📅 Report TT Access", "📑 COUPA Report", "🔓 Stolen Issue", "🛠️ CME",
                "🏢 Jumlah TT",
                "🔎 Root Cause",
                f"🗂️ Data Explorer ({fmt(len(d))} TT)"])

# ---------- Executive Summary ----------
with tabs[0]:
    st.markdown("<div class='insights'>" + "".join(
        f"<div class='ins' style='--c:{c}'><h4>{esc(t)}</h4><p>{body}</p></div>" for t, body, c in insights(d))
        + "</div>", unsafe_allow_html=True)

    # Seperti sheet tambahan ke TLP: satu Year Open (terbaru, atau tahun terbesar di filter), Case Status
    # Open + Closed. Impact service = Site Status (kolom kedua) "Down".
    years_ai = [y for y in prepared.get("Access Issue", base)["_year"].unique() if str(y).isdigit()]
    yr = max(FILTERS["years"]) if FILTERS["years"] else (max(years_ai) if years_ai else None)
    empty_cols = pd.DataFrame(columns=base.columns)
    access_y = apply_filters(prepared["Access Issue"], {**tab_filters("Access Issue"), "years": [yr]}) \
        if "Access Issue" in prepared else empty_cols
    stolen_y = apply_filters(prepared["Stolen"], {**tab_filters("Stolen"), "years": [yr]}) \
        if "Stolen" in prepared else empty_cols
    stolen_y = stolen_y[stolen_y["_has_tt"]] if len(stolen_y) else stolen_y
    imp_t = impact_table(access_y, stolen_y, SOW)
    rc_base = access_y[access_y["_impact"]] if len(access_y) else access_y
    # Top Cause / Root Cause tab mengikuti semua filter (kolom Problem Type & Case di List Access Issue).
    access_rc = apply_filters(prepared["Access Issue"], FILTERS) if "Access Issue" in prepared else empty_cols

    def g_ltoa() -> go.Figure:
        tt = imp_t["Total TT Impact Service Released"]
        tt = tt[tt > 0].sort_values(ascending=False)
        opts = list(tt.index)
        sel = st.multiselect("ToCo", opts, default=[t for t in ["PTI", "DMT", "TBG", "CMI"] if t in opts],
                             key="ltoa_toco", placeholder="Pilih ToCo…",
                             format_func=lambda t: f"{t} ({fmt(tt.get(t))} TT)",
                             help="ToCo = kolom GROUP1 (List Access Issue). Nilai = Total TT Impact Service Released.")
        return ltoa_pie(tt.reindex(sel).dropna())

    def g_impact() -> go.Figure:
        opts = [t for t in imp_t.index if (imp_t.loc[t, "SoW (A)"] or 0) > 0]
        if not opts:
            return empty_fig("SoW (A) belum tersedia.")
        t = st.selectbox("ToCo", opts, index=opts.index("TBG") if "TBG" in opts else 0, key="imp_toco")
        row = imp_t.loc[t]
        st.caption(f"**{t}** · SoW (A) {fmt(row['SoW (A)'])} · Total TT Impact Service Released "
                   f"{fmt(row['Total TT Impact Service Released'])} · Total Sites Impact Service (B) "
                   f"{fmt(row['Total Sites Impact Service (B)'])} · % (B/A) {pct(row['% (B/A)'], 2)}")
        return share_pie({"SoW (A)": row["SoW (A)"],
                          "Total Sites Impact Service (B)": row["Total Sites Impact Service (B)"]}, [MAGENTA, TOSCA])

    def g_root() -> go.Figure:
        g_vc = rc_base["_toco_raw"].value_counts()
        g_vc = g_vc[g_vc.index != "(Kosong)"]
        sel = st.multiselect("GROUP1", list(g_vc.index), key="rc_group1", placeholder="Semua GROUP1",
                             format_func=lambda g: f"{g} ({fmt(g_vc.get(g))} TT)")
        src = rc_base[rc_base["_toco_raw"].isin(sel)] if sel else rc_base
        blank = int((src["_has_tt"] & (src["_root"] == "(Kosong)")).sum())
        if blank:
            st.caption(f"{fmt(blank)} TT dengan kolom Case kosong tidak dihitung.")
        return root_cause_pie(src, 7)

    # Pie Stolen murni dari sheet List Stolen: dropdown Toco Owner (kolom Toco Owner), Year Open, Status Open +
    # Closed. Filter TLP/Cause/Severity global (berasal dari ToCo Access Issue) tidak dipakai di sini.
    stolen_pie = apply_filters(prepared["Stolen"], {**tab_filters("Stolen"), "years": [yr], "tlp": [], "cause": [],
                                                    "sev": []}) if "Stolen" in prepared else empty_cols
    stolen_pie = stolen_pie[stolen_pie["_has_tt"]] if len(stolen_pie) else stolen_pie

    def g_stolen() -> go.Figure:
        o_vc = stolen_pie["_owner"].value_counts() if len(stolen_pie) else pd.Series(dtype=int)
        o_vc = o_vc[o_vc.index != "(Kosong)"]
        opts = list(o_vc.index)
        if not opts:
            return empty_fig("Tidak ada TT stolen pada filter ini.")
        default = next((o for o in opts if "tower bersama" in o.lower()), opts[0])
        sel = st.selectbox("ToCo Owner", opts, index=opts.index(default), key="stl_owner",
                           format_func=lambda o: f"{o} ({fmt(o_vc.get(o))} TT)",
                           help="Dari kolom Toco Owner di sheet List Stolen.")
        sow = SOW.get(toco_code(sel))  # SoW (A) = Total site milik owner tsb. (LIST SITE / tabel SoW)
        own = stolen_pie[stolen_pie["_owner"] == sel]
        sites = own.loc[own["_site"] != "", "_site"].nunique()
        if not sow:
            return empty_fig(f"SoW (A) untuk {sel} belum diisi — atur di tabel SoW.")
        st.caption(f"Toco Owner **{sel}** · SoW (A) {fmt(sow)} · Total Sites Stolen {fmt(sites)} "
                   f"(IOH ID unik, Count of TT IOH) · {fmt(len(own))} TT IOH")
        return share_pie({"SoW (A)": sow, "Total Sites Stolen": sites}, [MAGENTA, TOSCA])

    section(f"Impact Service, Root Cause & Stolen · Year Open {yr}")
    a, b = st.columns(2)
    with a:
        chart_card("ltoa", "PIE 3D", MAGENTA, f"LTOA & Cell Down Contribution {yr}", g_ltoa, EXPORTS)
    with b:
        chart_card("impact", "PIE 3D", TOSCA, "Access Issue Impact Service", g_impact, EXPORTS, pie_labels="percent")
    a, b = st.columns(2)
    with a:
        chart_card("root", "PIE 3D", GOLD, "The Root Cause · Top 7 (kolom Case)", g_root, EXPORTS)
    with b:
        chart_card("stolen", "PIE 3D", RED, "Stolen Issue", g_stolen, EXPORTS, pie_labels="percent")
    st.caption(f"Year Open {yr}, Case Status Open + Closed. Total TT Impact Service Released = jumlah TT Site Status "
               "Down per ToCo (GROUP1); Total Sites Impact Service (B) = jumlah Site ID unik dari TT tersebut. "
               "The Root Cause = Count of Huawei TT per kolom Case (Site Status Down). Stolen = jumlah IOH ID unik "
               "(Count of TT IOH) per Toco Owner. Persen pie = bagian dari jumlah irisan.")

    EXTRA_TABLES[f"Impact Service & SoW per ToCo {yr}"] = imp_t
    with st.expander(f"📋 Tabel Impact Service & SoW per ToCo · {yr} (SoW bisa disunting)"):
        st.caption("SoW (A) dibaca dari sheet *LIST SITE* / *performance access* bila ada; selebihnya nilai cadangan. "
                   "Ubah angka di kolom SoW lalu klik **Simpan SoW** — berlaku untuk semua pengguna.")
        edit = imp_t.reset_index()
        edited = st.data_editor(
            edit, hide_index=True, width="stretch", key="sow_editor",
            disabled=[c for c in edit.columns if c != "SoW (A)"],
            column_config={"SoW (A)": st.column_config.NumberColumn(format="%d", min_value=0),
                           "MTTR Average (Days)": st.column_config.NumberColumn(format="%.2f"),
                           "% (B/A)": st.column_config.NumberColumn(format="%.2f%%"),
                           "% Stolen / SoW": st.column_config.NumberColumn(format="%.2f%%")})
        if st.button("💾 Simpan SoW"):
            new = {r["ToCo"]: float(r["SoW (A)"]) for _, r in edited.iterrows() if pd.notna(r["SoW (A)"])}
            changed = {k: v for k, v in new.items() if SOW.get(k) != v}
            save_sow_override({**load_sow_override(), **changed})
            st.toast(f"SoW tersimpan ({len(changed)} ToCo diubah).")
            st.rerun()

# ---------- Report TT Access (weekly & monthly, mengikuti pivot Excel) ----------
with tabs[1]:
    if "Access Issue" not in prepared:
        st.info("Sheet List Access Issue tidak tersedia.")
    else:
        # Case Status Open & Closed (Cancel tidak dihitung) + filter global; bulan dipakai hanya untuk monthly.
        rpt = apply_filters(prepared["Access Issue"], {**FILTERS, "status": FILTERS["status"] or ["Open", "Closed"]},
                            month=False)
        section("Weekly Report")
        weeks_all = sorted([w for w in rpt["_cw"].unique() if w], key=_cw_key)
        sel_weeks = st.multiselect("Week (CW Open)", weeks_all[::-1], default=weeks_all[-12:][::-1], key="rpt_weeks",
                                   placeholder="Pilih minggu…")
        weeks = sorted(sel_weeks, key=_cw_key)
        if not weeks:
            st.info("Pilih minimal satu minggu.")
        else:
            wk = rpt[rpt["_cw"].isin(weeks)]
            one_year = len({_cw_key(w)[0] for w in weeks}) == 1
            x_w = [f"CW{_cw_key(w)[1]}" if one_year else w for w in weeks]
            a, b = st.columns(2)
            with a:
                ct_sla = pivot_tt(wk, "_cw", "_sla", weeks, ["Exceed SLA", "Meet SLA"])
                chart_card("rpt_sla", "WEEKLY", MAGENTA, "TT Access based On SLA", sla_area(ct_sla, x_w), EXPORTS,
                           table=ct_sla.rename_axis("CW Open"))
            with b:
                ct_tlp_w = pivot_tt(wk, "_cw", "_grp", weeks, GROUP_ORDER)
                chart_card("rpt_tlp_w", "WEEKLY", TOSCA, "TT Access Based on TLP",
                           stacked_groups(ct_tlp_w, x_w, GROUP_COLORS), EXPORTS, table=ct_tlp_w.rename_axis("CW Open"))

        st.divider()
        section("Monthly Report")
        trend_section(trend_base("Access Issue"), "rpt", "TT access", "TT access", "_owner", [TEAL, YELLOW, MAGENTA],
                      11, owner_pick=True)
        st.caption("Trends/Latest update: Count of Case Status (Open & Closed) per Month Open. A = E bulan sebelumnya, "
                   "B = TT baru, C = A + B, D = C − E, E = akumulasi TT masih Open; Progress = D ÷ C. "
                   "ToCo Basis = kolom Site Owner.")
        st.divider()
        section("Monthly Report · Year Open")
        mon = rpt if FILTERS["years"] else rpt[rpt["_year"] == rpt["_year"].max()]
        if FILTERS["months"]:
            mon = mon[mon["_month"].isin(FILTERS["months"])]
        periods = sorted(mon["_period"].unique())
        if not periods:
            st.info("Tidak ada data bulanan pada filter ini.")
        else:
            mlabels = [p.strftime("%b_%Y") for p in periods]
            span = f"{mlabels[0].split('_')[0]} - {mlabels[-1].replace('_', ' ')}"
            a, b = st.columns(2)
            with a:
                ct_sev = pivot_tt(mon.assign(_sev_u=mon["_sev"].str.upper()), "_period", "_sev_u", periods,
                                  ["CRITICAL", "MAJOR"])
                chart_card("rpt_sev", "MONTHLY", MAGENTA, f"TT Access based on Severity · {span}",
                           clustered_groups(ct_sev, mlabels, SEV_COLORS, total_line=True), EXPORTS,
                           table=ct_sev.set_axis(mlabels).rename_axis("Month Open"))
            with b:
                ct_hub = pivot_tt(mon, "_period", "_sitecat", periods, ["BIG HUB SITE", "MEDIUM HUB SITE"])
                chart_card("rpt_hub", "MONTHLY", TOSCA, f"TT Access based on Hub Site · {span}",
                           clustered_groups(ct_hub, mlabels, HUB_COLORS, total_line=True, spline=True), EXPORTS,
                           table=ct_hub.set_axis(mlabels).rename_axis("Month Open"))
            a, b = st.columns(2)
            with a:
                ct_tlp_m = pivot_tt(mon, "_period", "_grp", periods, GROUP_ORDER)
                chart_card("rpt_tlp_m", "MONTHLY", YELLOW, f"TT Access based on TLP · {span}",
                           clustered_groups(ct_tlp_m, mlabels, GROUP_COLORS), EXPORTS,
                           table=ct_tlp_m.set_axis(mlabels).rename_axis("Month Open"))
            with b:
                mt_crit = mttr_critical_table(mon)
                chart_card("rpt_mttr", "MONTHLY", ORANGE, f"TT Access based on MTTR · {span} (CRITICAL)",
                           mttr_critical_chart(mt_crit), EXPORTS,
                           table=mt_crit.rename(columns={"tt": "TT Critical Access Issue",
                                                         "aging": "Sum of Aging Time", "mttr": "MTTR"}))
            with st.expander("Tabel MTTR CRITICAL per GROUP"):
                st.dataframe(mt_crit.rename(columns={"tt": "Count of Huawei TT", "aging": "Sum of Aging Time",
                                                     "mttr": "MTTR (Sum Aging ÷ TT)"}), width="stretch",
                             column_config={"Sum of Aging Time": st.column_config.NumberColumn(format="%.2f"),
                                            "MTTR (Sum Aging ÷ TT)": st.column_config.NumberColumn(format="%.2f")})
            st.caption("Semua grafik: Count of Huawei TT, Case Status Open & Closed. Monthly memakai Year Open terbaru "
                       "(atau tahun di filter). MTTR di grafik ini = Sum of Aging Time ÷ Count of Huawei TT per GROUP "
                       "untuk Severity CRITICAL (rumus Excel =E118/D118).")

# ---------- COUPA Report (Performances Overview) ----------
with tabs[2]:
    if COUPA["tt"].empty:
        st.info("Sheet **Data TT …** (mis. *Data TT Sept 2026*) tidak ditemukan di file data — COUPA Report tidak "
                "bisa dihitung.")
    else:
        c_pct, c_score, c_result = coupa_overview(COUPA["tt"], COUPA["site"])
        section(f"Performances Overview · {COUPA['tt_sheet']}")
        st.markdown(coupa_table_html(c_pct, c_score, c_result), unsafe_allow_html=True)
        st.caption("Result = SUMPRODUCT(skor, weighting). Warna Result: hijau = 4,00 · kuning ≥ 3,50 · merah < 3,50.")
        chart_card("coupa_chart", "COUPA", MAGENTA, "Performances Overview per TLP (in %)", coupa_chart(c_pct),
                   EXPORTS, table=c_pct.round(2).rename_axis("Question/KPI"))
        section("Scoring")
        st.markdown(coupa_scoring_html(), unsafe_allow_html=True)
        overview = c_pct.round(2).add_suffix(" (in %)").join(c_score.add_suffix(" (in #)"))
        overview.insert(len(c_pct.columns), "Weighting", [k[2] * 100 for k in COUPA_KPI])
        overview.loc["Result"] = [None] * (len(c_pct.columns) + 1) + list(c_result.round(2))
        overview.index.name = "Question/KPI"
        EXTRA_TABLES["COUPA Performances Overview"] = overview
        if COUPA["site"].empty:
            st.warning("Sheet **Site** tidak ditemukan — Site Security tidak bisa dihitung.")
        st.caption(f"Sumber: sheet *{COUPA['tt_sheet']}* dan *Site*. Hanya TLP2 CMI, MITRATEL, PTI (Protelindo), "
                   "PTI - STP, dan TBG. TT Handling Completion = TT Closed ÷ total TT per Kategori Issue; "
                   "Time To Resolved = Meet SLA ÷ TT Closed; Site Security = TT Problem type *Stolen Case* ÷ "
                   "Total Site (sheet Site).")

# ---------- Stolen Issue ----------
with tabs[3]:
    if "Stolen" not in prepared:
        st.info("Sheet List Stolen tidak tersedia.")
    else:
        trend_section(trend_base("Stolen"), "stl", "TT stolen issue", "TT stolen", "_toco_raw", [MAGENTA, TEAL], 9,
                      closed_mode=True)
        st.caption("Sumber: sheet List Stolen, Count of TT IOH per Month Open, Status TT ToCo/OpCo Open & Closed. "
                   "ToCo Basis = kolom Toco Group Name (total periode & jumlah Closed).")

# ---------- CME ----------
with tabs[4]:
    if "TT CME" not in prepared:
        st.info("Sheet TT CME tidak tersedia.")
    else:
        trend_section(trend_base("TT CME"), "cme", "TT CME", "TT CME", "_owner", [TEAL, YELLOW, MAGENTA], 10)
        st.caption("Sumber: sheet TT CME, Count of TT Number per Create Month, STATUS TT TLP Open & Close, "
                   f"Backlog = {', '.join(tab_filters('TT CME')['backlog'])}. ToCo Basis = kolom Owner.")

# ---------- Jumlah TT ----------
with tabs[5]:
    sites_f = SITES
    if f_region:
        sites_f = sites_f[sites_f["_region"].isin(f_region)]
    if f_tlp:
        sites_f = sites_f[sites_f["_toco"].isin(f_tlp)]
    if sites_f.empty:
        st.info("Sheet **LIST SITE** tidak ditemukan / kosong — Total site K1 & Achievement PM tidak bisa dihitung.")
    else:
        pm_t = pm_table(sites_f)
        has_pm = (pm_t["NY PM / Walktest"].sum() + pm_t["Done PM / Walktest"].sum()) > 0
        EXTRA_TABLES["Achievement PM"] = pm_t
        chart_card("pareto", "PVT-3", PURPLE, "Jumlah Site per TLP · Total site K1 & Achievement PM (sheet LIST SITE)",
                   k1_chart(pm_t), EXPORTS, table=pm_t)
        if not has_pm:
            st.caption("Kolom status PM (NY PM / DONE PM) belum ada di sheet LIST SITE file ini — garis Achievement "
                       "dan kolom PM akan terisi otomatis setelah file berisi kolom tersebut diunggah.")
    def g_site_type() -> go.Figure:
        o_vc = SITES["_owner_raw"].value_counts()
        sel = st.multiselect("Owner", list(o_vc.index), key="st_owner", placeholder="(All)",
                             format_func=lambda o: f"{o} ({fmt(o_vc.get(o))} site)")
        src = sites_f[sites_f["_owner_raw"].isin(sel)] if sel else sites_f
        fig, tbl = site_type_chart(src)
        if not tbl.empty:
            EXTRA_TABLES["Site Type per Region"] = tbl
        return fig

    # Grafik berdampingan di atas, tabel ACHIEVEMENT PM selebar halaman di bawah.
    a, b = st.columns(2)
    with a:
        chart_card("tlp_open", "OPEN", MAGENTA, "TT masih OPEN per TLP",
                   ranked(d[d["_status"] == "Open"], "_toco", 10, MAGENTA, 420, show_pct=False))
    with b:
        chart_card("site_type", "SITE", TOSCA, "Site Type per Region · Count of IOH ID", g_site_type, EXPORTS)
    if "Site Type per Region" in EXTRA_TABLES:
        with st.expander("Tabel Site Type per Region"):
            st.dataframe(EXTRA_TABLES["Site Type per Region"], width="stretch")
    if not sites_f.empty:
        with st.container(key="card_pm_table"):
            st.markdown(pm_table_html(pm_t, f"PM K2 {yr}"), unsafe_allow_html=True)
            st.caption("Total site K1 = COUNTIFS(LIST SITE!Owner, TLP); NY PM / Done PM = COUNTIFS(Owner, TLP, "
                       "status PM, \"NY PM\" / \"DONE PM\"); Achievement = Done PM ÷ Total site K1.")

# ---------- Root Cause ----------
with tabs[6]:
    a, b = st.columns(2)
    with a:
        chart_card("root_pie", "TOP 7", GOLD, f"The Root Cause · Top 7 (kolom Case, Site Status Down, {yr})",
                   root_cause_pie(rc_base, 7, 440))
    with b:
        roots = rc_base.loc[rc_base["_has_tt"] & (rc_base["_root"] != "(Kosong)"), "_root"].value_counts()
        tbl = roots.rename("Total").to_frame()
        tbl.index.name = "The Root Cause"
        EXTRA_TABLES["The Root Cause (kolom Case)"] = tbl
        with st.container(key="card_root_table"):
            st.markdown("<div class='ctitle'><span class='badge' style='--c:#B88A00'>TABEL</span>"
                        "The Root Cause</div>", unsafe_allow_html=True)
            st.dataframe(tbl, width="stretch", height=400)
            blank = int((rc_base["_has_tt"] & (rc_base["_root"] == "(Kosong)")).sum())
            if blank:
                st.caption(f"{fmt(blank)} TT tidak memiliki isian root cause (kolom Case kosong).")
    a, b = st.columns(2)
    with a:
        chart_card("cause2", "CAUSE", TOSCA, "Top Cause · Problem Type", ranked(access_rc, "_cause", 12, TOSCA, 420))
    with b:
        chart_card("cause_pareto", "CAUSE", PURPLE, "Pareto Cause · Problem Type", pareto(access_rc, "_cause", 12))
    st.caption(f"The Root Cause = Count of Huawei TT per kolom **Case** (List Access Issue, Site Status Down, "
               f"Year Open {yr}, Case Status Open & Closed). Top Cause & Pareto Cause dari kolom **Problem Type** "
               "(mengikuti semua filter).")

# ---------- Data Explorer ----------
with tabs[7]:
    preferred = []
    for c in cats:
        cfg = CATEGORIES[c]
        preferred += cfg["search"] + [REGION_COL, *cfg["toco"], cfg["status"], cfg["month"], cfg["aging"],
                                      cfg["severity"]]
    default_cols = [c for c in dict.fromkeys(preferred) if c in export_cols]
    cols = st.multiselect("Kolom ditampilkan", export_cols, default=default_cols)
    shown = d[cols].copy()
    for c in shown.columns:
        if shown[c].dtype == object:
            shown[c] = shown[c].astype("string")
    st.dataframe(shown, hide_index=True, width="stretch", height=560)


# =====================================================================================
# Export PowerPoint (grafik dikumpulkan saat dirender di atas)
# =====================================================================================

xlsx_slot.download_button(
    "📊 Export Excel", build_excel(EXPORTS, EXTRA_TABLES, source_meta, filter_desc),
    file_name=f"TLP_Dashboard_Data_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
    help="Tabel angka perhitungan setiap grafik (satu sheet per grafik) sesuai filter saat ini")

ppt_sig = "|".join(str(x) for x in [current_sig, filter_desc] + [st.session_state.get(k) for k in WIDGET_KEYS])
if st.session_state.get("pptx_sig") != ppt_sig:
    st.session_state.pop("pptx_bytes", None)
if "pptx_bytes" in st.session_state:
    ppt_slot.download_button("⬇️ Unduh PPT", st.session_state["pptx_bytes"],
                             file_name=f"TLP_Dashboard_{datetime.now():%Y%m%d_%H%M}.pptx",
                             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                             type="primary", width="stretch")
elif ppt_slot.button("📑 Export PPT", width="stretch", help="Susun presentasi dari grafik pada filter saat ini"):
    try:
        with st.spinner(f"Menyusun {len(EXPORTS)} grafik ke PowerPoint…"):
            st.session_state["pptx_bytes"] = build_pptx(
                EXPORTS, source_meta, [(lbl, val, sub) for lbl, val, sub in KPI_LIST], filter_desc)
            st.session_state["pptx_sig"] = ppt_sig
        st.rerun()
    except Exception as exc:  # data grafik tidak valid, dll.
        st.error(f"Gagal membuat PowerPoint: {exc}")
   
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.viewerBadge_container__1QSob {visibility: hidden;}
[data-testid="stStatusWidget"] {visibility: hidden;}
#stDecoration {display: none;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
