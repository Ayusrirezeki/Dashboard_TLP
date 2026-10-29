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
PREP_VERSION = 2           # naikkan bila kolom turunan di prepare() berubah, agar cache lama tidak dipakai

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
SERIES = {
    "A": ("A: Akumulasi TT open bulan sebelumnya", "#8a8fa8"),
    "B": ("B: TT baru per bulan", MAGENTA),
    "C": ("C: Akumulasi TT open per bulan", YELLOW),
    "D": ("D: TT close per bulan", TOSCA),
    "E": ("E: Sisa TT masih open [C−D]", RED),
}

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
        aging="Aging TT", year="Year Open", site_id="IOH ID", impact=[],
        search=["IOH ID", "Old Site ID", "Site Name", "TT IOH", "Material Stolen"], icon="🔓"),
    "TT CME": dict(
        sheet="TT CME", month="Create Month", status="STATUS TT TLP", toco=["Owner"], owner="Owner",
        cause="Description", cause_rules=CME_RULES, root=[], severity="Severity", aging="Aging TT CLOSED",
        year=None, site_id="Site ID", impact=[], backlog="Backlog",
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
    return pd.DataFrame(body, columns=header).dropna(how="all")


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


@st.cache_data(show_spinner="Membaca file .xlsb (pyxlsb)…", max_entries=3)
def load_workbook(signature: str, path: str) -> tuple[dict[str, pd.DataFrame], dict[str, float], pd.DataFrame]:
    # `signature` (path + waktu ubah + ukuran) menjadi kunci cache: data baru otomatis dibaca ulang.
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
    return sheets, sow, sites


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


def _unify_case(s: pd.Series) -> pd.Series:
    """Samakan ejaan yang hanya beda huruf besar/kecil (Dhost/DHOST/DHost) ke ejaan terbanyak."""
    key = s.str.lower()
    canon = s.groupby(key).agg(lambda v: v.value_counts().index[0])
    return key.map(canon).astype("string")


def _clean_root(s: pd.Series) -> pd.Series:
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.str.replace(r"(?i)^kompesasi", "Kompensasi", regex=True)
    return s.where(s.isin(["(Kosong)"]), s.str.slice(0, 1).str.upper() + s.str.slice(1))


@st.cache_data(show_spinner=False, max_entries=6)
def prepare(signature: str, _df: pd.DataFrame, cat: str) -> pd.DataFrame:
    cfg = CATEGORIES[cat]
    df = _df.copy()
    df["_period"] = to_period(df[cfg["month"]]) if cfg["month"] in df.columns else pd.NaT
    df = df[df["_period"].notna()].copy()
    status = _col(df, cfg["status"], "").str.lower()
    df["_status"] = status.map(lambda s: "Open" if s == "open" else "Cancel" if s.startswith("cancel")
                               else "Closed" if s.startswith("close") else "Lainnya")
    df["_cat"] = cat
    df["_region"] = _col(df, REGION_COL)
    df["_toco"] = _col(df, _first(df, cfg["toco"])).map(toco_code)
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
    impact_col = _first(df, cfg["impact"])
    df["_impact"] = _col(df, impact_col, "").str.lower().eq("down") if impact_col else False
    present = [c for c in cfg["search"] if c in df.columns]
    df["_search"] = (pd.concat([_col(df, c, "") for c in present], axis=1).agg(" ".join, axis=1).str.lower()
                     if present else "")
    return df


@st.cache_data(show_spinner=False, max_entries=3)
def prepare_sites(signature: str, _df: pd.DataFrame) -> pd.DataFrame:
    """LIST SITE: kolom Owner -> grup ToCo; 'TLP/MO OTHERS' dipisah sebagai OTHERS."""
    if _df.empty or "Owner" not in _df.columns:
        return pd.DataFrame(columns=["_owner_raw", "_toco", "_region"])
    df = pd.DataFrame({"_owner_raw": _col(_df, "Owner"), "_region": _col(_df, "Region")})
    others = df["_owner_raw"].str.lower().str.contains("other")
    df["_toco"] = df["_owner_raw"].map(toco_code).where(~others, "OTHERS")
    return df[df["_owner_raw"] != "(Kosong)"]


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
        textfont=dict(size=12, color=[TEXT if c in (YELLOW, ORANGE) else "#ffffff" for c in colors], family=FONT),
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


def impact_by_group(access: pd.DataFrame, owners: list[str] | None = None, top: int = 4) -> pd.Series:
    """Total TT Impact Service Released per grup ToCo (Site Owner dikelompokkan lewat kolom GROUP1,
    mis. Protelindo/STP/IBS/IFORTE -> PTI), hanya `top` grup teratas — seperti tabel di Excel."""
    imp = access[access["_impact"] & access["_status"].isin(["Open", "Closed"])]
    if owners:
        imp = imp[imp["_owner"].isin(owners)]
    s = imp[imp["_toco"] != "(Kosong)"].groupby("_toco").size().sort_values(ascending=False)
    return s.head(top)


def impact_donut(s: pd.Series) -> go.Figure:
    """Donut Top 4 grup; persentase dihitung terhadap jumlah keempat grup (sama dengan pie Excel)."""
    if s.empty:
        return empty_fig("Belum ada TT impact service pada filter ini.")
    total = s.sum()
    texts = [f"{fmt(v)}<br>{k}<br>{pct(v / total * 100, 0)}" for k, v in s.items()]
    return pie_3d(list(s.index), list(s.values), PALETTE[:len(s)], hole=0.45, texts=texts,
                  center=f"<b>{fmt(total)}</b><br><span style='font-size:11px'>TT Impact (Top {len(s)})</span>")


def sow_pie(a: float, b: float, label_b: str, color_b: str) -> go.Figure:
    """Pie SoW (A) vs bagian B; persentase B = B / A seperti kolom % (B/A) di Excel."""
    if not a or pd.isna(a):
        return empty_fig("SoW (A) untuk TLP ini belum diisi — atur di tabel SoW.")
    b = min(b, a)
    p = b / a * 100
    texts = [f"{fmt(a)}; {pct(100 - p, 2)}", f"{fmt(b)}; {pct(p, 2)}"]
    fig = pie_3d([f"SoW (A)", label_b], [a - b, b], [MAGENTA, color_b], texts=texts)
    fig.data[1].customdata = [a, b]
    fig.data[1].hovertemplate = "<b>%{label}</b><br>%{customdata:,.0f} site<extra></extra>"
    fig.data[1].pull = [0, 0.12]
    return fig


def root_cause_pie(d: pd.DataFrame, height: int = 380) -> go.Figure:
    s = top_n(d.loc[d["_root"] != "(Kosong)", "_root"].value_counts(), 5)
    if s.empty:
        return empty_fig("Kolom root cause / header case kosong pada filter ini.")
    total = s.sum()
    texts = [f"{fmt(v)}; {pct(v / total * 100, 0)}" for v in s.values]
    labels = [wrap(k, 28) for k in s.index]
    return pie_3d(labels, list(s.values), PALETTE[:len(s)], texts=texts, height=height)


def summary_chart(t: pd.DataFrame) -> go.Figure:
    x, n = t.index.tolist(), len(t)
    fig = go.Figure()
    for key in ["A", "B", "C", "D"]:
        name, color = SERIES[key]
        fig.add_trace(go.Bar(
            x=x, y=t[key], name=name, marker_color=color, text=bar_text(t[key]),
            # Label vertikal: 4 bar per bulan tidak cukup lebar untuk label mendatar.
            textposition="outside", textangle=-90 if n > 6 else 0, textfont=dict(size=10, color=TEXT_2),
            cliponaxis=False, hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>"))
    name, color = SERIES["E"]
    fig.add_trace(go.Scatter(
        x=x, y=t["E"], name=name, yaxis="y2", mode="lines+markers+text", text=bar_text(t["E"]),
        textposition="top center", textfont=dict(size=11, color=color), cliponaxis=False,
        line=dict(color=color, width=2.5), marker=dict(size=8, color=color, line=dict(color=BG, width=2)),
        hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>"))
    bar_max = float(t[["A", "B", "C", "D"]].to_numpy().max() or 1)
    e_max = float(t["E"].max() or 1)
    step = max(1, round(e_max / 2))
    dark(fig, 520)
    # Bar di ±55% bawah, garis E di ±35% atas (sumbu kanan dimulai negatif) agar label tidak bertabrakan.
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
    for col, color, pos in [("Progress (%)", TOSCA, "inside"), ("Remaining (%)", RED, "outside")]:
        fig.add_trace(go.Bar(x=t.index, y=t[col], name=col.replace(" (%)", ""), marker_color=color,
                             text=bar_text(t[col], pct), textposition=pos, insidetextanchor="middle",
                             textangle=0, textfont=dict(size=10, color="#fff" if pos == "inside" else TEXT_2),
                             hovertemplate=f"<b>%{{x}}</b><br>{col}: %{{y:.1f}}%<extra></extra>"))
    dark(fig, 300)
    fig.update_layout(barmode="stack", uniformtext=dict(minsize=8, mode="hide"),
                      yaxis=dict(range=[0, 112], ticksuffix="%"), xaxis=dict(tickangle=-30, showgrid=False))
    return fig


def open_close_region(d: pd.DataFrame) -> go.Figure:
    if d.empty:
        return empty_fig("Tidak ada data.")
    ct = pd.crosstab(d["_region"], d["_status"]).reindex(columns=["Closed", "Open"], fill_value=0)
    ct = ct.loc[ct.sum(axis=1).sort_values(ascending=False).index]
    fig = go.Figure()
    for col, color, label in [("Closed", TOSCA, "CLOSE"), ("Open", MAGENTA, "OPEN")]:
        fig.add_trace(go.Bar(x=ct.index, y=ct[col], name=label, marker_color=color, text=bar_text(ct[col]),
                             textposition="outside", textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
                             hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}}<extra></extra>"))
    dark(fig, 360)
    fig.update_layout(barmode="group", yaxis=dict(range=[0, (ct.to_numpy().max() or 1) * 1.18]))
    return fig


def mttr_by(d: pd.DataFrame, col: str, top: int = 12, height: int = 360) -> go.Figure:
    closed = d[d["_status"].eq("Closed")]
    g = closed.groupby(col)["_aging"].agg(["mean", "size"]).dropna()
    if g.empty:
        return empty_fig("Belum ada tiket CLOSE pada filter ini.", height)
    g = g.sort_values("size", ascending=False).head(top).sort_values("mean", ascending=False)
    avg = closed["_aging"].mean()
    fig = go.Figure(go.Bar(
        x=[wrap(i) for i in g.index], y=g["mean"], marker_color=[RED if v > avg else TOSCA for v in g["mean"]],
        text=[f"{days(v)} d" for v in g["mean"]], textposition="outside", textfont=dict(size=11, color=TEXT_2),
        cliponaxis=False, customdata=g["size"],
        hovertemplate="<b>%{x}</b><br>MTTR: %{y:.2f} hari<br>%{customdata:,} tiket close<extra></extra>"))
    fig.add_hline(y=avg, line=dict(color=YELLOW, dash="dash", width=1.5),
                  annotation_text=f"Rata-rata {days(avg)} d", annotation_position="top right",
                  annotation_font=dict(color=GOLD, size=11))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=False, margin=dict(t=26),
                      yaxis=dict(range=[0, g["mean"].max() * 1.22], title="hari"),
                      xaxis=dict(tickangle=-30 if len(g) > 8 else 0, tickfont=dict(size=10)))
    return fig


def mttr_table(access: pd.DataFrame) -> pd.DataFrame:
    """Per grup ToCo (impact service: Site Status Down, Case Status Open/Closed):
    tt   = Count of Huawei TT, aging = Sum Aging Time (Days), mttr = tt ÷ aging."""
    imp = access[access["_impact"] & access["_status"].isin(["Open", "Closed"])]
    g = imp.groupby("_toco").agg(tt=("_has_tt", "sum"), aging=("_aging", "sum"))
    g = g.reindex([t for t in TOCO_ORDER if t in g.index])
    g = g[g["tt"] > 0]
    g["mttr"] = g["tt"] / g["aging"].where(g["aging"] > 0)  # aging 0 -> tidak bisa dibagi (kosong)
    return g


def _group_bar(g: pd.DataFrame, col: str, color: str, fmt_fn, avg: float | None, avg_label: str,
               y_title: str, hover: str, height: int) -> go.Figure:
    s = g[col].dropna()
    fig = go.Figure(go.Bar(
        x=[wrap(i) for i in s.index], y=s.values, marker_color=color, name=y_title, text=[fmt_fn(v) for v in s.values],
        textposition="outside", textfont=dict(size=11, color=TEXT_2), cliponaxis=False,
        customdata=list(zip(g.loc[s.index, "tt"], g.loc[s.index, "aging"])), hovertemplate=hover))
    if avg is not None and pd.notna(avg):
        fig.add_hline(y=avg, line=dict(color=YELLOW, dash="dash", width=1.5), annotation_text=avg_label,
                      annotation_position="top right", annotation_font=dict(color=GOLD, size=11))
    dark(fig, height, legend_top=False)
    fig.update_layout(showlegend=False, margin=dict(t=26),
                      yaxis=dict(range=[0, (s.max() if len(s) else 1) * 1.22], title=y_title),
                      xaxis=dict(tickangle=-30 if len(s) > 8 else 0, tickfont=dict(size=10)))
    return fig


HOVER_GROUP = ("<b>%{x}</b><br>Count Huawei TT: %{customdata[0]:,}<br>Sum Aging: %{customdata[1]:,.1f} hari"
               "<extra></extra>")


def mttr_group(access: pd.DataFrame, height: int = 360) -> go.Figure:
    """MTTR = Count of Huawei TT ÷ Sum Aging Time (Days) per grup ToCo."""
    g = mttr_table(access)
    if g["mttr"].dropna().empty:
        return empty_fig("Belum ada TT impact service pada filter ini.", height)
    avg = g["tt"].sum() / g["aging"].sum() if g["aging"].sum() > 0 else None
    return _group_bar(g, "mttr", TOSCA, lambda v: days(v), avg, f"Total {days(avg)}" if avg else "",
                      "TT ÷ Aging", "<b>%{x}</b><br>MTTR (TT ÷ Aging): %{y:.2f}<br>"
                      + HOVER_GROUP.split("<br>", 1)[1], height)


def ranked(d: pd.DataFrame, col: str, top: int = 10, color: str = TOSCA, height: int = 360) -> go.Figure:
    vc = d[col].value_counts().head(top)[::-1]
    if vc.empty:
        return empty_fig("Tidak ada data.", height)
    share = vc / max(len(d), 1) * 100
    fig = go.Figure(go.Bar(
        y=vc.index, x=vc.values, orientation="h", marker_color=color, name="Jumlah TT",
        text=[f"{fmt(v)}  ({pct(s)})" for v, s in zip(vc.values, share)], textposition="outside",
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


def monthly_status(d: pd.DataFrame) -> go.Figure:
    if d.empty:
        return empty_fig("Tidak ada data.")
    ct = pd.crosstab(d["_period"], d["_status"]).reindex(columns=["Closed", "Open"], fill_value=0).sort_index()
    x = [p.strftime("%b_%Y") for p in ct.index]
    fig = go.Figure()
    for col, color, label in [("Closed", TOSCA, "CLOSE"), ("Open", MAGENTA, "OPEN")]:
        fig.add_trace(go.Bar(x=x, y=ct[col], name=label, marker_color=color,
                             text=bar_text(ct[col].where(ct[col] > 0)), textposition="inside",
                             insidetextanchor="middle", textangle=0, textfont=dict(size=10, color="#fff"),
                             hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}}<extra></extra>"))
    total = ct.sum(axis=1)
    fig.add_trace(go.Scatter(x=x, y=total, mode="text", text=bar_text(total), textposition="top center",
                             textfont=dict(size=11, color=TEXT), showlegend=False, hoverinfo="skip"))
    dark(fig, 360)
    fig.update_layout(barmode="stack", uniformtext=dict(minsize=8, mode="hide"),
                      yaxis=dict(range=[0, total.max() * 1.15]),
                      xaxis=dict(tickangle=-30 if len(x) > 8 else 0, showgrid=False))
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
                           chart="kumulatif" not in (t.name or "").lower()))
    kind = ("bar_h" if horiz else "column_stacked_100" if fig.layout.barnorm == "percent"
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


def chart_card(key: str, badge: str, color: str, title: str, fig, exports: list | None = None,
               table: pd.DataFrame | None = None, pie_labels: str = "value_percent") -> None:
    """Kartu grafik. `fig` boleh berupa fungsi: dipanggil di dalam kartu, sehingga widget
    pengatur (mis. dropdown) bisa tampil di antara judul dan grafik.
    Bila `exports` diberikan, angka grafik disimpan untuk Export Excel & PPT; `table` mengganti
    tabel Excel bawaan bila perlu kolom yang lebih jelas."""
    with st.container(key=f"card_{key}"):
        st.markdown(f"<div class='ctitle'><span class='badge' style='--c:{color}'>{esc(badge)}</span>"
                    f"{esc(title)}</div>", unsafe_allow_html=True)
        if callable(fig):
            fig = fig()
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "displayModeBar": False},
                        key=f"fig_{key}")
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
            "Grafik A–E: TT berstatus *Cancel* tidak dihitung (sama dengan pivot Summary).")
    xlsx_slot = b3.empty()
    ppt_slot = b4.empty()
    if b5.button("↻ Reset Filter", width="stretch"):
        for k in [k for k in st.session_state if str(k).startswith(FILTER_KEYS) or k in ("ae_range", "target_tlp", "g1_owner", "g4_owner")]:
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
    sheets, sow_from_file, list_site = load_workbook(current_sig, str(source_path))
except Exception as exc:  # file rusak / bukan xlsb
    st.error(f"Gagal membaca **{source_meta.get('filename')}**: {exc}")
    st.stop()
if not sheets:
    st.error("Sheet List Access Issue / List Stolen / TT CME tidak ditemukan di file ini.")
    st.stop()

prepared = {cat: prepare(f"{current_sig}:{cat}:v{PREP_VERSION}", df, cat) for cat, df in sheets.items()}
SITES = prepare_sites(f"{current_sig}:sites", list_site)
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
dim = apply_filters(base, FILTERS, status=False, month=False)
dim = dim[dim["_status"] != "Cancel"]  # dasar grafik A–E (Cancel tidak dihitung, sama dengan pivot)

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
tabs = st.tabs(["📊 Executive Summary", "📈 Trend A–E", "⏱️ MTTR", "🏢 Jumlah TT", "🔎 Root Cause",
                f"🗂️ Data Explorer ({fmt(len(d))} TT)"])

# ---------- Executive Summary ----------
with tabs[0]:
    st.markdown("<div class='insights'>" + "".join(
        f"<div class='ins' style='--c:{c}'><h4>{esc(t)}</h4><p>{body}</p></div>" for t, body, c in insights(d))
        + "</div>", unsafe_allow_html=True)

    # Impact service & stolen memakai data Access Issue / Stolen dengan filter dimensi yang sama
    # (status Open+Closed sesuai rumus pivot), apa pun kategori yang dipilih.
    access_f = apply_filters(prepared["Access Issue"], FILTERS, status=False) if "Access Issue" in prepared \
        else pd.DataFrame(columns=base.columns)
    stolen_f = apply_filters(prepared["Stolen"], FILTERS, status=False) if "Stolen" in prepared \
        else pd.DataFrame(columns=base.columns)
    imp_t = impact_table(access_f, stolen_f, SOW)
    # Root cause (Grafik 4) selalu dari kolom "Case" di List Access Issue, mengikuti semua filter.
    access_rc = apply_filters(prepared["Access Issue"], FILTERS) if "Access Issue" in prepared \
        else pd.DataFrame(columns=base.columns)

    def grafik1() -> go.Figure:
        imp_rows = access_f[access_f["_impact"] & access_f["_status"].isin(["Open", "Closed"])]
        owner_vc = imp_rows["_owner"].value_counts()
        owner_vc = owner_vc[owner_vc.index != "(Kosong)"]
        groups = imp_rows.groupby("_owner")["_toco"].agg(lambda g: " / ".join(g.value_counts().index))
        sel = st.multiselect("Site Owner", list(owner_vc.index), key="g1_owner", placeholder="Semua Site Owner",
                             format_func=lambda o: f"{o} → {groups.get(o, '-')} ({fmt(owner_vc.get(o))} TT)",
                             help="Site Owner dikelompokkan ke grup ToCo (kolom GROUP1). Grafik menampilkan "
                                  "Top 4 grup dengan Total TT Impact Service Released terbanyak.")
        return impact_donut(impact_by_group(access_f, sel, 4))

    def grafik4() -> go.Figure:
        owner_vc = access_rc["_owner"].value_counts()
        owner_vc = owner_vc[owner_vc.index != "(Kosong)"]
        groups = access_rc.groupby("_owner")["_toco"].agg(lambda g: " / ".join(g.value_counts().index))
        sel = st.multiselect("Site Owner", list(owner_vc.index), key="g4_owner", placeholder="Semua Site Owner",
                             format_func=lambda o: f"{o} → {groups.get(o, '-')} ({fmt(owner_vc.get(o))} TT)",
                             help="Pilih Site Owner untuk melihat Top 5 root cause (kolom Case) miliknya.")
        src = access_rc[access_rc["_owner"].isin(sel)] if sel else access_rc
        return root_cause_pie(src, 380)

    section("Impact Service & SoW per ToCo")
    tlp_opts = [t for t in imp_t.index if pd.notna(imp_t.loc[t, "SoW (A)"])] or list(imp_t.index)
    target = st.selectbox("TLP yang dituju (grafik SoW)", tlp_opts, key="target_tlp",
                          index=tlp_opts.index("TBG") if "TBG" in tlp_opts else 0) if tlp_opts else None
    a, b = st.columns(2)
    with a:
        chart_card("g1", "Grafik 1", MAGENTA, "Impact Service Terbanyak per ToCo · Top 4 (Total TT Released)",
                   grafik1, EXPORTS)
    with b:
        chart_card("g4", "Grafik 4", GOLD, "The Root Cause · Top 5 (kolom Case)", grafik4, EXPORTS)
    a, b = st.columns(2)
    if target is not None:
        row = imp_t.loc[target]
        with a:
            chart_card("g2", "Grafik 2", TOSCA, f"Access Issue Impact Service · {target}",
                       sow_pie(row["SoW (A)"], row["Total Sites Impact Service (B)"],
                               "Total Sites Impact Service (B)", TOSCA), EXPORTS,
                       table=pd.DataFrame({"Nilai": [row["SoW (A)"], row["Total Sites Impact Service (B)"],
                                                     row["Total TT Impact Service Released"], row["% (B/A)"]]},
                                          index=pd.Index(["SoW (A)", "Total Sites Impact Service (B)",
                                                          "Total TT Impact Service Released", "% (B/A)"],
                                                         name=target)),
                       pie_labels="percent")
            st.caption(f"% (B/A) = **{pct(row['% (B/A)'], 2)}** · SoW (A) {fmt(row['SoW (A)'])} site · "
                       f"{fmt(row['Total TT Impact Service Released'])} TT impact service released")
        with b:
            chart_card("g3", "Grafik 3", RED, f"Stolen Issue · {target}",
                       sow_pie(row["SoW (A)"], row["Total Sites Stolen"], "Total Sites Stolen", TOSCA), EXPORTS,
                       table=pd.DataFrame({"Nilai": [row["SoW (A)"], row["Total Sites Stolen"], row["Total TT Stolen"],
                                                     row["% Stolen / SoW"]]},
                                          index=pd.Index(["SoW (A)", "Total Sites Stolen", "Total TT Stolen",
                                                          "% Stolen / SoW"], name=target)),
                       pie_labels="percent")
            st.caption(f"% Stolen / SoW = **{pct(row['% Stolen / SoW'], 2)}** · "
                       f"{fmt(row['Total TT Stolen'])} TT stolen")

    EXTRA_TABLES["Impact Service & SoW per ToCo"] = imp_t
    with st.expander("📋 Tabel Impact Service & SoW per ToCo (SoW bisa disunting)"):
        st.caption("SoW (A) dibaca dari sheet *performance access* bila ada; selebihnya nilai cadangan. "
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

    section("Open/Close & MTTR")
    a, b = st.columns(2)
    with a:
        chart_card("oc_region", "PVT-1", MAGENTA, "Open vs Close Ticket per Region", open_close_region(d), EXPORTS)
    with b:
        mt_export = mttr_table(access_f).rename(columns={"tt": "Count of Huawei TT",
                                                         "aging": "Sum of Aging Time (Days)", "mttr": "MTTR (TT ÷ Aging)"})
        mt_export.index.name = "ToCo"
        chart_card("mttr_tlp", "PVT-2", YELLOW, "MTTR per ToCo · Count Huawei TT ÷ Sum Aging (Days)",
                   mttr_group(access_f), EXPORTS, table=mt_export)
    a, b = st.columns(2)
    with a:
        chart_card("monthly", "PVT-3", TOSCA, "Jumlah TT per Bulan Open (Close/Open)", monthly_status(d), EXPORTS)
    with b:
        chart_card("cause", "PVT-4", RED, "Top Cause · Problem Type", ranked(access_rc, "_cause", 10, MAGENTA), EXPORTS)

# ---------- Trend A–E ----------
with tabs[1]:
    months_sel = sorted(dim["_period"].unique())
    labels_m = [p.strftime("%b_%Y") for p in months_sel]
    if labels_m:
        m_from, m_to = st.select_slider("Periode grafik A–E (bulan open)", options=labels_m,
                                        value=(labels_m[max(0, len(labels_m) - 13)], labels_m[-1]), key="ae_range")
        win = months_sel[labels_m.index(m_from): labels_m.index(m_to) + 1]
        t = backlog_summary(dim[dim["_period"].between(win[0], win[-1])], win)
        chart_card("ae", "A–E", MAGENTA, f"Trend TT {cat_title} · {m_from} – {m_to}", summary_chart(t), EXPORTS,
                   table=t.rename(columns={k: v[0] for k, v in SERIES.items()}))
        chart_card("progress", "%", TOSCA, f"Progress & Remaining TT {cat_title}", progress_chart(t), EXPORTS,
                   table=t[["Progress (%)", "Remaining (%)"]])
        with st.expander("Tabel summary A–E"):
            st.dataframe(t.rename(columns={k: v[0] for k, v in SERIES.items()}), width="stretch",
                         column_config={"Progress (%)": st.column_config.NumberColumn(format="%.1f%%"),
                                        "Remaining (%)": st.column_config.NumberColumn(format="%.1f%%")})
        st.caption("Grafik A–E mengikuti filter di atas kecuali Status & Bulan; TT Cancel tidak dihitung.")
    else:
        st.info("Tidak ada data bulan untuk filter ini.")

# ---------- MTTR ----------
with tabs[2]:
    chart_card("mttr_tlp2", "MTTR", YELLOW, "MTTR per ToCo · Count Huawei TT ÷ Sum Aging (Days)",
               mttr_group(access_f, 400))
    with st.expander("Tabel MTTR per ToCo"):
        mt = mttr_table(access_f).rename(columns={"tt": "Count of Huawei TT", "aging": "Sum of Aging Time (Days)",
                                                  "mttr": "MTTR (TT ÷ Aging)"})
        mt.index.name = "ToCo"
        st.dataframe(mt, width="stretch", column_config={
            "Sum of Aging Time (Days)": st.column_config.NumberColumn(format="%.1f"),
            "MTTR (TT ÷ Aging)": st.column_config.NumberColumn(format="%.2f")})
    st.caption("Per grup ToCo (TT impact service: Site Status Down, Case Status Open/Closed): Total TT = Count of "
               "Huawei TT; MTTR = Count of Huawei TT ÷ Sum of Aging Time (Days). Grup dengan Sum Aging 0 "
               "tidak bisa dibagi sehingga tidak tampil di grafik MTTR.")
    a, b = st.columns(2)
    with a:
        chart_card("mttr_region", "MTTR", ORANGE, "Average MTTR (Days) per Region", mttr_by(d, "_region"))
    with b:
        chart_card("mttr_cause", "MTTR", MAGENTA, "Average MTTR (Days) per Cause", mttr_by(d, "_cause"))

# ---------- Jumlah TT ----------
with tabs[3]:
    sites_f = SITES
    if f_region:
        sites_f = sites_f[sites_f["_region"].isin(f_region)]
    if f_tlp:
        sites_f = sites_f[sites_f["_toco"].isin(f_tlp)]
    k1 = sites_f[sites_f["_toco"] != "OTHERS"].groupby("_toco").size()
    if k1.empty:
        st.info("Sheet **LIST SITE** tidak ditemukan / kosong — grafik Total site K1 tidak bisa dihitung.")
    else:
        chart_card("pareto", "PVT-3", PURPLE, "Jumlah Site per TLP · Total site K1 (sheet LIST SITE)",
                   pareto(None, None, 20, values=k1, name="Total site K1"), EXPORTS)
        with st.expander("Tabel Total site K1 per TLP"):
            order = [t for t in TOCO_ORDER if t in k1.index] + [t for t in k1.index if t not in TOCO_ORDER]
            others = int((sites_f["_toco"] == "OTHERS").sum())
            rows = [(t, int(k1[t])) for t in order]
            rows += [("TOTAL", int(k1.sum())), ("TLP/MO OTHERS", others),
                     ("Total Actual (All Site)", int(k1.sum()) + others)]
            k1_df = pd.DataFrame(rows, columns=["TLP", "Total site K1"])
            EXTRA_TABLES["Total site K1 per TLP"] = k1_df.set_index("TLP")
            st.dataframe(k1_df, hide_index=True, width="stretch")
            st.caption("Total site K1 = COUNTIFS(LIST SITE!Owner, TLP). Kolom status PM (NY PM / Done PM) belum "
                       "ada di sheet LIST SITE, sehingga Achievement PM belum bisa dihitung.")
    a, b = st.columns(2)
    with a:
        chart_card("tlp_open", "OPEN", MAGENTA, "TT masih OPEN per TLP",
                   ranked(d[d["_status"] == "Open"], "_toco", 10, MAGENTA))
    with b:
        chart_card("tlp_sev", "SEV", TOSCA, "Jumlah TT per Severity", ranked(d, "_sev", 8, TOSCA, 300))

# ---------- Root Cause ----------
with tabs[4]:
    a, b = st.columns(2)
    with a:
        chart_card("root_pie", "TOP 5", GOLD, "The Root Cause · Top 5 (kolom Case)", root_cause_pie(access_rc, 440))
    with b:
        roots = access_rc.loc[access_rc["_root"] != "(Kosong)", "_root"].value_counts()
        tbl = roots.rename("Total").to_frame()
        tbl.index.name = "The Root Cause"
        EXTRA_TABLES["The Root Cause (kolom Case)"] = tbl
        with st.container(key="card_root_table"):
            st.markdown("<div class='ctitle'><span class='badge' style='--c:#B88A00'>TABEL</span>"
                        "The Root Cause</div>", unsafe_allow_html=True)
            st.dataframe(tbl, width="stretch", height=400)
            blank = int((access_rc["_root"] == "(Kosong)").sum())
            if blank:
                st.caption(f"{fmt(blank)} TT tidak memiliki isian root cause (kolom Case kosong).")
    a, b = st.columns(2)
    with a:
        chart_card("cause2", "CAUSE", TOSCA, "Top Cause · Problem Type", ranked(access_rc, "_cause", 12, TOSCA, 420))
    with b:
        chart_card("cause_pareto", "CAUSE", PURPLE, "Pareto Cause · Problem Type", pareto(access_rc, "_cause", 12))
    st.caption("The Root Cause diambil dari kolom **Case** di sheet List Access Issue (mengikuti semua filter). "
               "Top Cause & Pareto Cause dihitung dari kolom **Problem Type** di sheet List Access Issue.")

# ---------- Data Explorer ----------
with tabs[5]:
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

ppt_sig = "|".join(str(x) for x in [current_sig, filter_desc] + [st.session_state.get(k) for k in
                                    ("target_tlp", "ae_range", "g1_owner", "g4_owner")])
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
