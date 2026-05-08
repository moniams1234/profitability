import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from copy import deepcopy
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings("ignore")

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Postkalkulacja Profitability",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Main background */
[data-testid="stAppViewContainer"] {background: #F7EFEA;}
[data-testid="stHeader"] {background: transparent;}
[data-testid="stToolbar"] {display: none;}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #4A0000 !important;
    border-right: 2px solid #3A0000;
}
[data-testid="stSidebar"] * {color: #FAE8E0 !important;}
[data-testid="stSidebar"] .stRadio label {color: #FAE8E0 !important;}
[data-testid="stSidebar"] .stSelectbox label {color: #FAE8E0 !important;}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {color: #FF8C66 !important;}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #6B0000, #FF5A1F) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 8px 20px !important;
    font-weight: 600 !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #3A0000, #C91818) !important;
    box-shadow: 0 4px 12px rgba(107,0,0,0.4) !important;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #0FA958, #0d8a47) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 24px !important;
    font-weight: 700 !important;
    font-size: 1.05rem !important;
}

/* Cards */
.card {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(107,0,0,0.08);
    margin-bottom: 16px;
}
.kpi-card {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 10px rgba(107,0,0,0.1);
    border-left: 4px solid #6B0000;
    text-align: center;
}
.kpi-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #6B0000;
}
.kpi-label {
    font-size: 0.78rem;
    color: #6B6B6B;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.alert-card {
    background: #FFF3F0;
    border-left: 4px solid #C91818;
    border-radius: 8px;
    padding: 12px 16px;
    margin: 8px 0;
}
.section-title {
    color: #6B0000;
    font-size: 1.2rem;
    font-weight: 700;
    border-bottom: 2px solid #FF5A1F;
    padding-bottom: 6px;
    margin-bottom: 16px;
}
.file-badge {
    display: inline-block;
    background: #0FA958;
    color: white;
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 0.75rem;
    font-weight: 600;
    margin-left: 8px;
}
.file-badge-miss {
    display: inline-block;
    background: #C91818;
    color: white;
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 0.75rem;
    font-weight: 600;
    margin-left: 8px;
}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    background: #3A0000;
    border-radius: 10px 10px 0 0;
    gap: 2px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    color: #FAE8E0 !important;
    background: transparent !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}
.stTabs [aria-selected="true"] {
    background: #FF5A1F !important;
    color: white !important;
    font-weight: 700 !important;
}

/* Headers */
h1 { color: #3A0000 !important; }
h2 { color: #6B0000 !important; }
h3 { color: #6B0000 !important; }

/* DataFrames */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* Selectbox */
.stSelectbox [data-baseweb="select"] { border-radius: 8px; }

/* Multiselect */
.stMultiSelect [data-baseweb="select"] { border-radius: 8px; }

/* Warning / Info */
.stWarning { border-left: 4px solid #FF5A1F; }
.stInfo { border-left: 4px solid #6B0000; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def norm_cols(df):
    """Normalize column names: strip, collapse whitespace, lower."""
    df.columns = [
        str(c).strip().replace("\n", " ").replace("\r", " ")
        for c in df.columns
    ]
    return df


def find_col(df, candidates):
    """Return the first matching column name (case-insensitive)."""
    lc = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lc:
            return lc[cand.lower()]
    return None


def read_excel_safe(uploaded, sheet_name=0, header="auto"):
    """Read uploaded file, auto-detect header row (skip title rows)."""
    if uploaded is None:
        return None
    try:
        raw = uploaded.read()
        uploaded.seek(0)
        xf = pd.ExcelFile(io.BytesIO(raw))
        sheets = xf.sheet_names
        if isinstance(sheet_name, str) and sheet_name not in sheets:
            sheet_name = sheets[0]
        if header == "auto":
            # scan first 10 rows for header
            probe = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name,
                                  header=None, nrows=10)
            hrow = 0
            for i, row in probe.iterrows():
                non_nan = row.dropna()
                if len(non_nan) >= 3 and not all(
                    str(v).startswith("Export") or str(v).startswith("Zamów")
                    for v in non_nan
                ):
                    hrow = i
                    break
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name,
                               header=hrow)
        else:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name,
                               header=header)
        return norm_cols(df)
    except Exception as e:
        st.warning(f"Błąd wczytywania pliku: {e}")
        return None


def read_post_list(uploaded):
    """post_list has 3 title rows; data starts at row 3 (0-indexed)."""
    if uploaded is None:
        return None
    try:
        raw = uploaded.read()
        uploaded.seek(0)
        df = pd.read_excel(io.BytesIO(raw), header=3)
        return norm_cols(df)
    except Exception as e:
        st.warning(f"Błąd wczytywania Bazy: {e}")
        return None


def batch_label(qty):
    try:
        q = float(qty)
    except (TypeError, ValueError):
        return ""
    if q <= 50: return "0-50"
    if q <= 100: return "51-100"
    if q <= 200: return "101-200"
    if q <= 300: return "201-300"
    if q <= 500: return "301-500"
    if q <= 1000: return "501-1000"
    if q <= 1500: return "1001-1500"
    if q <= 2000: return "1501-2000"
    if q <= 3000: return "2001-3000"
    if q <= 10000: return "3001-10000"
    if q <= 20000: return "10001-20000"
    if q <= 30000: return "20001-30000"
    if q <= 100000: return "30001-100000"
    return "100001-1000000"


def safe_num(val):
    try:
        v = float(val)
        return 0.0 if np.isnan(v) else v
    except (TypeError, ValueError):
        return 0.0


# ─── SESSION STATE ─────────────────────────────────────────────────────────────
if "rates" not in st.session_state:
    st.session_state["rates"] = {}
if "click_costs" not in st.session_state:
    st.session_state["click_costs"] = {
        "HP Indigo 7K Digital Press": {"CMYK": 0.05, "default": 0.05},
    }
if "prepress" not in st.session_state:
    st.session_state["prepress"] = {}
if "other_costs_pct" not in st.session_state:
    st.session_state["other_costs_pct"] = 2.0
if "tpm_threshold" not in st.session_state:
    st.session_state["tpm_threshold"] = 60.0
if "cm_threshold" not in st.session_state:
    st.session_state["cm_threshold"] = 40.0

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Postkalkulacja")
    st.markdown("**Profitability App**")
    st.markdown("---")
    st.markdown("### Nawigacja")
    tab_names = [
        "📂 Upload plików",
        "⚙️ Stawki maszyn",
        "🖨️ Koszty klików",
        "🎨 Prepress",
        "🔧 Parametry",
        "📋 Podgląd Profitability",
        "📈 Podsumowanie",
        "🎯 Kokpit",
        "⬇️ Pobierz XLSX",
    ]
    selected_tab = st.radio("", tab_names, label_visibility="collapsed")
    st.markdown("---")
    st.markdown(
        "<small style='color:#FF8C66'>© Postkalkulacja v1.0</small>",
        unsafe_allow_html=True,
    )

# ─── FILE UPLOAD ─────────────────────────────────────────────────────────────
tab_idx = tab_names.index(selected_tab)

def status_badge(ok):
    if ok:
        return '<span class="file-badge">✓ wczytany</span>'
    return '<span class="file-badge-miss">✗ brak</span>'

if tab_idx == 0:
    st.markdown('<h1>📂 Upload plików</h1>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card"><p>Uploaduj pliki źródłowe. '
        '<b>Baza (post_list)</b> jest wymagana. Pozostałe pliki są opcjonalne.</p></div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Wymagane")
        uf_base = st.file_uploader("📋 Baza / post_list *", type=["xlsx", "xls"], key="uf_base")
        st.markdown("#### Produkcja")
        uf_czasy = st.file_uploader("⏱️ Czasy dla aplikacji", type=["xlsx", "xls"], key="uf_czasy")
        uf_zlec = st.file_uploader("📄 Zlecenia + faktury", type=["xlsx", "xls"], key="uf_zlec")
        uf_fry = st.file_uploader("🧾 Faktury – linie faktury", type=["xlsx", "xls"], key="uf_fry")
    with col2:
        st.markdown("#### Koszty")
        uf_inks = st.file_uploader("🖨️ Kliki / Inks", type=["xlsx", "xls"], key="uf_inks")
        uf_stawki = st.file_uploader("⚙️ Stawki dla aplikacji", type=["xlsx", "xls"], key="uf_stawki")
        uf_farby = st.file_uploader("🎨 Farby podsumowanie (Offset)", type=["xlsx", "xls"], key="uf_farby")

    # Status
    st.markdown("---")
    st.markdown("#### Status plików")
    files_status = {
        "Baza": uf_base,
        "Czasy": uf_czasy,
        "Zlecenia+Faktury": uf_zlec,
        "Faktury linie": uf_fry,
        "Kliki/Inks": uf_inks,
        "Stawki": uf_stawki,
        "Farby": uf_farby,
    }
    scols = st.columns(4)
    for i, (name, f) in enumerate(files_status.items()):
        scols[i % 4].markdown(
            f"**{name}** {status_badge(f is not None)}",
            unsafe_allow_html=True,
        )

# ─── HELPER: BUILD DATA ───────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_stawki_from_file(raw_bytes):
    df = pd.read_excel(io.BytesIO(raw_bytes), header=None)
    # find header row
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row if str(v).strip() != "nan"]
        if "Nazwa maszyny" in vals and "Stawka rbg" in vals:
            df = pd.read_excel(io.BytesIO(raw_bytes), header=i)
            df = norm_cols(df)
            nm = find_col(df, ["Nazwa maszyny"])
            st_col = find_col(df, ["Stawka rbg"])
            if nm and st_col:
                df = df[[nm, st_col]].dropna(subset=[nm])
                df.columns = ["Nazwa maszyny", "Stawka rbg"]
                return df
    return pd.DataFrame(columns=["Nazwa maszyny", "Stawka rbg"])


def get_rates(uf_stawki):
    """Return dict {machine_name: rate} from file or session."""
    rates = dict(st.session_state["rates"])
    if uf_stawki and not rates:
        raw = uf_stawki.read()
        uf_stawki.seek(0)
        df_s = load_stawki_from_file(raw)
        for _, r in df_s.iterrows():
            rates[str(r["Nazwa maszyny"]).strip()] = safe_num(r["Stawka rbg"])
        st.session_state["rates"] = rates
    return rates


# ─── STAWKI MASZYN ────────────────────────────────────────────────────────────
if tab_idx == 1:
    st.markdown('<h1>⚙️ Stawki maszyn</h1>', unsafe_allow_html=True)
    uf_stawki = st.session_state.get("uf_stawki") if "uf_stawki" not in st.session_state else None
    # Reload from sidebar file uploader
    uf_stawki_file = st.session_state.get("uf_stawki")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("Stawki są ładowane automatycznie z pliku **Stawki dla aplikacji**. Możesz je edytować poniżej.")
    
    rates = dict(st.session_state.get("rates", {}))
    
    # Load from file if not yet in session
    uploaded_stawki = st.session_state.get("uf_stawki")
    if uploaded_stawki and not rates:
        raw = uploaded_stawki.read()
        uploaded_stawki.seek(0)
        df_s = load_stawki_from_file(raw)
        for _, r in df_s.iterrows():
            rates[str(r["Nazwa maszyny"]).strip()] = safe_num(r["Stawka rbg"])
        st.session_state["rates"] = rates

    # Add new machine
    with st.expander("➕ Dodaj maszynę"):
        nc1, nc2, nc3 = st.columns([3, 2, 1])
        new_name = nc1.text_input("Nazwa maszyny", key="new_machine_name")
        new_rate = nc2.number_input("Stawka rbg (PLN/h)", value=100.0, min_value=0.0, key="new_machine_rate")
        if nc3.button("Dodaj"):
            if new_name:
                rates[new_name] = new_rate
                st.session_state["rates"] = rates
                st.success(f"Dodano: {new_name}")

    if rates:
        st.markdown("#### Aktualne stawki")
        df_rates = pd.DataFrame(
            [(k, v) for k, v in rates.items()],
            columns=["Nazwa maszyny", "Stawka rbg (PLN/h)"]
        )
        edited = st.data_editor(
            df_rates,
            use_container_width=True,
            num_rows="dynamic",
            key="rates_editor",
        )
        if st.button("💾 Zapisz stawki"):
            new_rates = {}
            for _, row in edited.iterrows():
                if pd.notna(row["Nazwa maszyny"]) and str(row["Nazwa maszyny"]).strip():
                    new_rates[str(row["Nazwa maszyny"]).strip()] = safe_num(row["Stawka rbg (PLN/h)"])
            st.session_state["rates"] = new_rates
            st.success("Stawki zapisane!")
    else:
        st.info("Brak stawek. Uploaduj plik 'Stawki dla aplikacji' lub dodaj ręcznie.")
    st.markdown("</div>", unsafe_allow_html=True)


# ─── KOSZTY KLIKÓW ────────────────────────────────────────────────────────────
if tab_idx == 2:
    st.markdown('<h1>🖨️ Koszty klików</h1>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(
        "Zdefiniuj koszty jednostkowe klików (per separacja) dla każdej maszyny i koloru."
    )
    click_costs = st.session_state.get("click_costs", {})
    
    df_cc = pd.DataFrame([
        {"Maszyna (Press Name)": k, "Kolor": c, "Koszt na separację (PLN)": v}
        for k, colors in click_costs.items()
        for c, v in colors.items()
    ])
    
    edited_cc = st.data_editor(
        df_cc if not df_cc.empty else pd.DataFrame(
            columns=["Maszyna (Press Name)", "Kolor", "Koszt na separację (PLN)"]
        ),
        use_container_width=True,
        num_rows="dynamic",
        key="click_costs_editor",
    )
    if st.button("💾 Zapisz koszty klików"):
        new_cc = {}
        for _, row in edited_cc.iterrows():
            if pd.notna(row.get("Maszyna (Press Name)")) and str(row["Maszyna (Press Name)"]).strip():
                mach = str(row["Maszyna (Press Name)"]).strip()
                color = str(row.get("Kolor", "default")).strip()
                cost = safe_num(row.get("Koszt na separację (PLN)", 0.05))
                if mach not in new_cc:
                    new_cc[mach] = {}
                new_cc[mach][color] = cost
        st.session_state["click_costs"] = new_cc
        st.success("Koszty klików zapisane!")
    st.markdown("</div>", unsafe_allow_html=True)


# ─── PREPRESS ─────────────────────────────────────────────────────────────────
if tab_idx == 3:
    st.markdown('<h1>🎨 Prepress</h1>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(
        "Ustaw stawki Prepress dla każdego klienta. "
        "Domyślna stawka Digital = 10, Offset = 40."
    )
    prepress = st.session_state.get("prepress", {})

    # Global defaults
    col_d, col_o = st.columns(2)
    default_digital = col_d.number_input("Domyślna stawka Digital (PLN)", value=10.0, min_value=0.0, key="pp_digital_default")
    default_offset = col_o.number_input("Domyślna stawka Offset (PLN)", value=40.0, min_value=0.0, key="pp_offset_default")
    st.session_state["pp_digital_default"] = default_digital
    st.session_state["pp_offset_default"] = default_offset

    st.markdown("#### Stawki per klient (opcjonalne)")
    df_pp = pd.DataFrame([
        {"Klient": k, "Stawka Digital": v["digital"], "Stawka Offset": v["offset"]}
        for k, v in prepress.items()
    ])
    edited_pp = st.data_editor(
        df_pp if not df_pp.empty else pd.DataFrame(
            columns=["Klient", "Stawka Digital", "Stawka Offset"]
        ),
        use_container_width=True,
        num_rows="dynamic",
        key="prepress_editor",
    )
    if st.button("💾 Zapisz Prepress"):
        new_pp = {}
        for _, row in edited_pp.iterrows():
            if pd.notna(row.get("Klient")) and str(row["Klient"]).strip():
                new_pp[str(row["Klient"]).strip()] = {
                    "digital": safe_num(row.get("Stawka Digital", default_digital)),
                    "offset": safe_num(row.get("Stawka Offset", default_offset)),
                }
        st.session_state["prepress"] = new_pp
        st.success("Prepress zapisany!")
    st.markdown("</div>", unsafe_allow_html=True)


# ─── PARAMETRY ────────────────────────────────────────────────────────────────
if tab_idx == 4:
    st.markdown('<h1>🔧 Parametry</h1>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    other_pct = c1.number_input(
        "Other costs % (od Sales Value)",
        value=st.session_state["other_costs_pct"],
        min_value=0.0, max_value=100.0, step=0.1, format="%.1f"
    )
    st.session_state["other_costs_pct"] = other_pct

    tpm_thr = c2.number_input(
        "Próg alertu TPM % (poniżej = alert)",
        value=st.session_state["tpm_threshold"],
        min_value=0.0, max_value=100.0, step=1.0, format="%.0f"
    )
    st.session_state["tpm_threshold"] = tpm_thr

    cm_thr = c1.number_input(
        "Próg alertu CM % (poniżej = alert)",
        value=st.session_state["cm_threshold"],
        min_value=0.0, max_value=100.0, step=1.0, format="%.0f"
    )
    st.session_state["cm_threshold"] = cm_thr
    st.markdown("</div>", unsafe_allow_html=True)


# ─── CORE CALCULATION ENGINE ──────────────────────────────────────────────────

def build_profitability(uf_base, uf_czasy, uf_zlec, uf_fry, uf_inks, uf_farby,
                        rates, click_costs, prepress,
                        other_pct, pp_digital_default, pp_offset_default):
    """
    Returns: (df_prof, df_czasy, df_kliki, df_farby_pivot, machine_cols, warnings_list)
    """
    warns = []

    # ── BASE ──────────────────────────────────────────────────────────────────
    if uf_base is None:
        return None, None, None, None, [], ["Brak pliku Baza!"]

    df_base = read_post_list(uf_base)
    if df_base is None:
        return None, None, None, None, [], ["Nie można wczytać Bazy"]

    # Check required columns
    required_base = [
        "Numer", "Zamówienie", "Zamawiana ilość",
        "Papier [16]", "Klej [17]", "Lakiery [20]",
        "Opakowania zbiorcze [24]", "Farby  [19]", "Kliki [48]",
    ]
    # Flexible matching
    col_map = {}
    for req in required_base:
        found = find_col(df_base, [req, req.strip()])
        if found:
            col_map[req] = found
        else:
            warns.append(f"Brakuje kolumny w Bazie: '{req}'")

    df = df_base.copy()

    # Normalize col names (handle double spaces)
    rename_map = {}
    for c in df.columns:
        nc = " ".join(c.split())
        if nc != c:
            rename_map[c] = nc
    df.rename(columns=rename_map, inplace=True)

    def gc(name):
        """Get column value safely."""
        c = find_col(df, [name, name.strip()])
        if c:
            return c
        return None

    # ── ZLECENIE PRODUKCYJNE ──────────────────────────────────────────────────
    numer_col = gc("Numer")
    if numer_col:
        df["Zlecenie produkcyjne"] = df[numer_col].astype(str).str.split("-").str[0].str.strip()
    else:
        df["Zlecenie produkcyjne"] = ""

    zam_col = gc("Zamówienie")
    if zam_col:
        df["Lewy 10"] = df[zam_col].astype(str).str[:10]
    else:
        df["Lewy 10"] = ""

    # ── BATCH ─────────────────────────────────────────────────────────────────
    qty_col = find_col(df, ["Zamawiana ilość", "Zamawiana ilosc"])
    if qty_col:
        df["Batch"] = df[qty_col].apply(batch_label)
    else:
        df["Batch"] = ""

    # ── CZASY ─────────────────────────────────────────────────────────────────
    df_czasy_out = None
    machine_cols = []

    if uf_czasy:
        df_czasy = read_excel_safe(uf_czasy, sheet_name=0, header="auto")
        if df_czasy is not None:
            # normalize
            nzp = find_col(df_czasy, ["Numer zlecenia produkcyjnego"])
            nm_col = find_col(df_czasy, ["Nazwa maszyny"])
            czas_col = find_col(df_czasy, ["Czas czynnosci [min]", "Czas czynności [min]"])

            if nzp and nm_col and czas_col:
                df_czasy["_zp"] = df_czasy[nzp].astype(str).str.strip()
                df_czasy["_machine"] = df_czasy[nm_col].astype(str).str.strip().str.lower()
                df_czasy[czas_col] = pd.to_numeric(df_czasy[czas_col], errors="coerce").fillna(0)

                # Add Koszt pracy column
                def get_rate_for_machine(machine_raw):
                    for k, v in rates.items():
                        if k.lower().strip() in machine_raw or machine_raw in k.lower().strip():
                            return v
                    # fuzzy: partial match
                    for k, v in rates.items():
                        if any(part in machine_raw for part in k.lower().split() if len(part) > 3):
                            return v
                    return 0.0

                df_czasy["_rate"] = df_czasy["_machine"].apply(get_rate_for_machine)
                df_czasy["Koszt pracy"] = df_czasy[czas_col] / 60.0 * df_czasy["_rate"]

                df_czasy_out = df_czasy.copy()

                # Unique machines
                machines = sorted(df_czasy[nm_col].dropna().unique())
                machine_cols = [str(m).strip() for m in machines]

                # Pivot: for each machine, sum Koszt pracy per Zlecenie produkcyjne
                for mach in machines:
                    mach_s = str(mach).strip()
                    mask = df_czasy["_machine"] == mach_s.lower()
                    grp = df_czasy[mask].groupby("_zp")["Koszt pracy"].sum().reset_index()
                    grp.columns = ["_zp_key", mach_s]
                    df = df.merge(grp, left_on="Zlecenie produkcyjne",
                                  right_on="_zp_key", how="left")
                    df[mach_s] = df[mach_s].fillna(0)
                    df.drop(columns=["_zp_key"], errors="ignore", inplace=True)

                # Digital/Offset classification
                def classify_do(zp):
                    zp = str(zp).strip()
                    if not zp:
                        return ""
                    sub = df_czasy[df_czasy["_zp"] == zp]
                    if sub.empty:
                        return ""
                    machines_for_zp = set(sub["_machine"].str.lower().str.strip())
                    hp_names = {"hp 35", "hp 7", "hp 1", "hp35", "hp35k",
                                "hp indigo 7k", "hp indigo", "hp indigo 7"}
                    hd_names = {"heidelberg cx 104", "heidelberg", "cx 104"}
                    is_hp = any(
                        any(h in m for h in hp_names) for m in machines_for_zp
                    )
                    is_hd = any(
                        any(h in m for h in hd_names) for m in machines_for_zp
                    )
                    if is_hp:
                        return "Digital"
                    if is_hd:
                        return "Offset"
                    return "no printing"

                df["Digital/Offset"] = df["Zlecenie produkcyjne"].apply(classify_do)
            else:
                warns.append("Plik Czasy: brak wymaganych kolumn (Numer ZP / Nazwa maszyny / Czas).")
                df["Digital/Offset"] = ""
        else:
            warns.append("Nie można wczytać pliku Czasy.")
            df["Digital/Offset"] = ""
    else:
        warns.append("Brak pliku Czasy – kolumny kosztów maszyn pominięte.")
        df["Digital/Offset"] = ""
        machine_cols = []

    # ── FARBY ─────────────────────────────────────────────────────────────────
    df_farby_pivot = None
    if uf_farby:
        try:
            raw = uf_farby.read()
            uf_farby.seek(0)
            xf = pd.ExcelFile(io.BytesIO(raw))
            pivot_sheet = None
            for sn in xf.sheet_names:
                if "pivot" in sn.lower() or "farb" in sn.lower():
                    pivot_sheet = sn
                    break
            if pivot_sheet:
                raw_piv = pd.read_excel(io.BytesIO(raw), sheet_name=pivot_sheet, header=None, nrows=5)
                hrow_p = 0
                for i, row in raw_piv.iterrows():
                    vals = [str(v).strip() for v in row if str(v).strip() != "nan"]
                    if "Etykiety wierszy" in vals or "Suma koszt farby2" in vals:
                        hrow_p = i
                        break
                df_farby_pivot = pd.read_excel(io.BytesIO(raw), sheet_name=pivot_sheet, header=hrow_p)
                df_farby_pivot = norm_cols(df_farby_pivot)
                ew_col = find_col(df_farby_pivot, ["Etykiety wierszy"])
                if ew_col:
                    df_farby_pivot = df_farby_pivot.dropna(subset=[ew_col])
                    df_farby_pivot["_ew"] = df_farby_pivot[ew_col].astype(str).str.strip()
        except Exception as e:
            warns.append(f"Błąd wczytywania Farby: {e}")

    # Merge Offset inks & Płyta offsetowa
    if df_farby_pivot is not None:
        kf_col = find_col(df_farby_pivot, ["Suma koszt farby2"])
        kp_col = find_col(df_farby_pivot, ["Suma koszt płyty"])
        df["_lewy10"] = df["Lewy 10"].astype(str).str.strip()
        df_fp = df_farby_pivot.copy()
        df_fp["_ew2"] = df_fp["_ew"].str.strip()
        merge_f = df_fp[["_ew2", kf_col, kp_col]].rename(
            columns={"_ew2": "_ew_key", kf_col: "Offset inks", kp_col: "Płyta offsetowa"}
        )
        df = df.merge(merge_f, left_on="_lewy10", right_on="_ew_key", how="left")
        df.drop(columns=["_ew_key", "_lewy10"], errors="ignore", inplace=True)
        df["Offset inks"] = pd.to_numeric(df["Offset inks"], errors="coerce").fillna(0)
        df["Płyta offsetowa"] = pd.to_numeric(df["Płyta offsetowa"], errors="coerce").fillna(0)
    else:
        df["Offset inks"] = 0.0
        df["Płyta offsetowa"] = 0.0
        if uf_farby is None:
            warns.append("Brak pliku Farby – Offset inks i Płyta offsetowa = 0.")

    # ── KLIKI ─────────────────────────────────────────────────────────────────
    df_kliki_out = None
    if uf_inks:
        try:
            raw = uf_inks.read()
            uf_inks.seek(0)
            df_inks = pd.read_excel(io.BytesIO(raw))
            df_inks = norm_cols(df_inks)
            jn_col = find_col(df_inks, ["Job Name"])
            pn_col = find_col(df_inks, ["Press Name"])
            col_col = find_col(df_inks, ["Color"])
            sep_col = find_col(df_inks, ["Separations"])

            if jn_col and sep_col:
                df_inks["Zamówienie"] = df_inks[jn_col].astype(str).str[:10]
                df_inks["Separations_n"] = pd.to_numeric(df_inks[sep_col], errors="coerce").fillna(0)

                def get_click_cost(row):
                    mach = str(row.get(pn_col, "")).strip() if pn_col else ""
                    color = str(row.get(col_col, "")).strip() if col_col else ""
                    if mach in click_costs:
                        cc = click_costs[mach]
                        if color in cc:
                            return cc[color]
                        if "default" in cc:
                            return cc["default"]
                    # fallback: first available
                    for cc in click_costs.values():
                        return list(cc.values())[0] if cc else 0.05
                    return 0.05

                df_inks["Koszt klików"] = df_inks.apply(get_click_cost, axis=1) * df_inks["Separations_n"]
                grp_inks = df_inks.groupby("Zamówienie")["Koszt klików"].sum().reset_index()
                grp_inks.columns = ["_zam_kliki", "Moje Kliki"]
                df = df.merge(grp_inks, left_on="Lewy 10", right_on="_zam_kliki", how="left")
                df.drop(columns=["_zam_kliki"], errors="ignore", inplace=True)
                df["Moje Kliki"] = pd.to_numeric(df["Moje Kliki"], errors="coerce").fillna(0)
                df_kliki_out = df_inks.copy()
            else:
                warns.append("Plik Inks: brak kolumn Job Name lub Separations.")
                df["Moje Kliki"] = 0.0
        except Exception as e:
            warns.append(f"Błąd wczytywania Inks: {e}")
            df["Moje Kliki"] = 0.0
    else:
        df["Moje Kliki"] = 0.0

    kliki_col = find_col(df, ["Kliki [48]"])
    if kliki_col:
        df[kliki_col] = pd.to_numeric(df[kliki_col], errors="coerce").fillna(0)
        df["Kliki final"] = df[[kliki_col, "Moje Kliki"]].max(axis=1)
    else:
        df["Kliki final"] = df["Moje Kliki"]

    # ── SALES VALUE & DATA FAKTURY ────────────────────────────────────────────
    df["Sales Value"] = np.nan
    df["Data faktury"] = pd.NaT

    if uf_zlec:
        df_zlec = read_excel_safe(uf_zlec, sheet_name=0, header="auto")
        if df_zlec is not None:
            nzp_z = find_col(df_zlec, ["Numer zlecenia produkcyjnego"])
            wart_z = find_col(df_zlec, ["Wartosc w linii FV netto"])
            data_z = find_col(df_zlec, ["Data wystawienia FV", "Data wystawienia faktury"])
            if nzp_z and wart_z:
                df_zlec["_nzp"] = df_zlec[nzp_z].astype(str).str.strip()
                grp_z = df_zlec.groupby("_nzp").agg(
                    _sv=(wart_z, "sum"),
                    _dv=(data_z, "first") if data_z else (wart_z, "first")
                ).reset_index()
                df = df.merge(grp_z, left_on="Zlecenie produkcyjne", right_on="_nzp", how="left")
                mask_sv = df["_sv"].notna()
                df.loc[mask_sv, "Sales Value"] = df.loc[mask_sv, "_sv"]
                if data_z:
                    df.loc[mask_sv, "Data faktury"] = pd.to_datetime(
                        df.loc[mask_sv, "_dv"], errors="coerce"
                    )
                df.drop(columns=["_nzp", "_sv", "_dv"], errors="ignore", inplace=True)

    # Fallback: faktury linie
    if uf_fry and df["Sales Value"].isna().any():
        df_fry = read_excel_safe(uf_fry, sheet_name=0, header="auto")
        if df_fry is not None:
            nl_col = find_col(df_fry, ["Nazwa linii faktury"])
            wl_col = find_col(df_fry, ["Wartosc w linii FV netto"])
            il_col = find_col(df_fry, ["Ilosc w linii FV"])
            dl_col = find_col(df_fry, ["Data wystawienia FV"])

            if nl_col and wl_col:
                zam_src = gc("Zamówienie")
                qty_src = find_col(df, ["Zamawiana ilość", "Zamawiana ilosc"])

                for idx in df[df["Sales Value"].isna()].index:
                    zam_val = str(df.at[idx, zam_src] if zam_src else "").strip()
                    # extract fragment after "| "
                    parts = zam_val.split("| ")
                    if len(parts) >= 2:
                        fragment = parts[-1].strip()
                    else:
                        fragment = zam_val[:10]

                    if not fragment:
                        continue
                    mask_fry = df_fry[nl_col].astype(str).str.contains(
                        fragment, case=False, na=False
                    )
                    sub_fry = df_fry[mask_fry]
                    if sub_fry.empty:
                        continue
                    row_fry = sub_fry.iloc[0]
                    wart_fry = safe_num(row_fry[wl_col])
                    ilosc_fry = safe_num(row_fry[il_col]) if il_col else 0
                    qty_prof = safe_num(df.at[idx, qty_src]) if qty_src else 0
                    if ilosc_fry > 0 and qty_prof > 0:
                        sv_calc = (wart_fry / ilosc_fry) * qty_prof
                    else:
                        sv_calc = wart_fry
                    df.at[idx, "Sales Value"] = sv_calc
                    if dl_col:
                        df.at[idx, "Data faktury"] = pd.to_datetime(
                            row_fry[dl_col], errors="coerce"
                        )

    df["Sales Value"] = pd.to_numeric(df["Sales Value"], errors="coerce").fillna(0)
    df["Data faktury"] = pd.to_datetime(df["Data faktury"], errors="coerce")
    df["Miesiąc faktury"] = df["Data faktury"].dt.strftime("%Y-%m")

    # ── PREPRESS COSTS ────────────────────────────────────────────────────────
    klient_col = find_col(df, ["Klient", "Klient ID"])

    def get_prepress_cost(row):
        klient = str(row.get(klient_col, "") if klient_col else "").strip()
        offset_inks = safe_num(row.get("Offset inks", 0))
        if klient in prepress:
            rates_pp = prepress[klient]
            return rates_pp["offset"] if offset_inks > 0 else rates_pp["digital"]
        return pp_offset_default if offset_inks > 0 else pp_digital_default

    df["Prepress costs"] = df.apply(get_prepress_cost, axis=1)

    # ── OTHER COSTS ───────────────────────────────────────────────────────────
    df["Other costs %"] = other_pct / 100.0
    df["Other Materials"] = df["Sales Value"] * df["Other costs %"]

    # ── TOTAL DL ─────────────────────────────────────────────────────────────
    dl_components = machine_cols + ["Prepress costs"]
    for col in dl_components:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Total DL"] = df[dl_components].sum(axis=1)
    if not machine_cols:
        warns.append("Brak pliku Czasy – Total DL = Prepress costs.")

    # ── TOTAL MATERIALS ───────────────────────────────────────────────────────
    mat_cols_candidates = [
        "Papier [16]", "Klej [17]", "Lakiery [20]",
        "Opakowania zbiorcze [24]", "Other Materials",
        "Offset inks", "Płyta offsetowa", "Kliki final"
    ]
    mat_cols = []
    for mc in mat_cols_candidates:
        real = find_col(df, [mc])
        if real:
            df[real] = pd.to_numeric(df[real], errors="coerce").fillna(0)
            mat_cols.append(real)
        else:
            df[mc] = 0.0
            mat_cols.append(mc)
    df["Total Materials"] = df[mat_cols].sum(axis=1)

    # ── TPM & CM ──────────────────────────────────────────────────────────────
    df["TPM"] = df["Sales Value"] - df["Total DL"]
    df["CM"] = df["Sales Value"] - df["Total DL"] - df["Total Materials"]

    # ── COLUMN ORDER ─────────────────────────────────────────────────────────
    end_cols = [
        "Zlecenie produkcyjne", "Lewy 10", "Batch", "Digital/Offset",
        "Total DL", "Total Materials", "Sales Value",
        "Data faktury", "Miesiąc faktury", "TPM", "CM"
    ]
    other_cols = [c for c in df.columns if c not in end_cols]
    final_order = other_cols + end_cols
    df = df[[c for c in final_order if c in df.columns]]

    return df, df_czasy_out, df_kliki_out, df_farby_pivot, machine_cols, warns


# ─── CACHED BUILD ─────────────────────────────────────────────────────────────

def get_data():
    uf_base = st.session_state.get("uf_base")
    uf_czasy = st.session_state.get("uf_czasy")
    uf_zlec = st.session_state.get("uf_zlec")
    uf_fry = st.session_state.get("uf_fry")
    uf_inks = st.session_state.get("uf_inks")
    uf_farby = st.session_state.get("uf_farby")
    uf_stawki = st.session_state.get("uf_stawki")

    rates = get_rates(uf_stawki)
    click_costs = st.session_state.get("click_costs", {})
    prepress = st.session_state.get("prepress", {})
    other_pct = st.session_state.get("other_costs_pct", 2.0)
    pp_digital = st.session_state.get("pp_digital_default", 10.0)
    pp_offset = st.session_state.get("pp_offset_default", 40.0)
    tpm_thr = st.session_state.get("tpm_threshold", 60.0)
    cm_thr = st.session_state.get("cm_threshold", 40.0)

    return build_profitability(
        uf_base, uf_czasy, uf_zlec, uf_fry, uf_inks, uf_farby,
        rates, click_costs, prepress, other_pct, pp_digital, pp_offset
    ), tpm_thr, cm_thr


# ─── PODGLĄD PROFITABILITY ────────────────────────────────────────────────────
if tab_idx == 5:
    st.markdown('<h1>📋 Podgląd Profitability</h1>', unsafe_allow_html=True)
    if not st.session_state.get("uf_base"):
        st.warning("⚠️ Uploaduj plik Baza w zakładce 'Upload plików'.")
    else:
        with st.spinner("Obliczanie..."):
            (df_prof, df_czasy, df_kliki, df_farby_piv,
             machine_cols, warns), tpm_thr, cm_thr = get_data()

        for w in warns:
            st.warning(w)

        if df_prof is not None:
            st.markdown(
                f'<div class="card"><b>Rekordów:</b> {len(df_prof)} | '
                f'<b>Kolumn:</b> {len(df_prof.columns)}</div>',
                unsafe_allow_html=True
            )
            # Month filter
            months = sorted(df_prof["Miesiąc faktury"].dropna().unique())
            sel_months = st.multiselect("Filtruj miesiące", ["Wszystkie"] + months, default=["Wszystkie"])
            if "Wszystkie" not in sel_months and sel_months:
                df_view = df_prof[df_prof["Miesiąc faktury"].isin(sel_months)]
            else:
                df_view = df_prof
            st.dataframe(df_view, use_container_width=True, height=500)


# ─── PODSUMOWANIE ─────────────────────────────────────────────────────────────
if tab_idx == 6:
    st.markdown('<h1>📈 Podsumowanie</h1>', unsafe_allow_html=True)
    if not st.session_state.get("uf_base"):
        st.warning("⚠️ Uploaduj plik Baza.")
    else:
        with st.spinner("Obliczanie..."):
            (df_prof, df_czasy, df_kliki, df_farby_piv,
             machine_cols, warns), tpm_thr, cm_thr = get_data()

        for w in warns:
            st.warning(w)

        if df_prof is not None:
            months = sorted(df_prof["Miesiąc faktury"].dropna().unique())
            sel_months = st.multiselect(
                "Wybierz miesiące do podsumowania",
                months,
                default=months[:1] if months else [],
                key="sum_months"
            )

            if sel_months:
                klient_col = find_col(df_prof, ["Klient", "Klient ID"])

                for month in sel_months:
                    st.markdown(f'<div class="section-title">📅 {month}</div>', unsafe_allow_html=True)
                    df_m = df_prof[df_prof["Miesiąc faktury"] == month].copy()

                    if df_m.empty:
                        st.info(f"Brak danych dla {month}")
                        continue

                    grp_col = klient_col if klient_col else "Zlecenie produkcyjne"
                    df_m[grp_col] = df_m[grp_col].fillna("(brak)")

                    def count_do(series, val):
                        return (series == val).sum()

                    summary_rows = []
                    for klient, grp in df_m.groupby(grp_col):
                        sv = grp["Sales Value"].sum()
                        tpm = grp["TPM"].sum()
                        cm = grp["CM"].sum()
                        n = len(grp)
                        n_dig = count_do(grp["Digital/Offset"], "Digital")
                        n_off = count_do(grp["Digital/Offset"], "Offset")
                        n_nop = count_do(grp["Digital/Offset"], "no printing")
                        summary_rows.append({
                            "Klient": klient,
                            "Miesiąc": month,
                            "Suma sprzedaży": sv,
                            "Suma TPM": tpm,
                            "TPM %": tpm / sv * 100 if sv else 0,
                            "Suma CM": cm,
                            "CM %": cm / sv * 100 if sv else 0,
                            "Liczba zamówień": n,
                            "Digital": n_dig,
                            "Offset": n_off,
                            "No printing": n_nop,
                        })

                    df_sum = pd.DataFrame(summary_rows)

                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.dataframe(
                            df_sum.style.format({
                                "Suma sprzedaży": "{:,.2f}",
                                "Suma TPM": "{:,.2f}",
                                "TPM %": "{:.1f}%",
                                "Suma CM": "{:,.2f}",
                                "CM %": "{:.1f}%",
                            }),
                            use_container_width=True
                        )
                    with col2:
                        # Batch summary
                        if "Batch" in df_m.columns and klient_col:
                            st.markdown("**Zamówienia wg klienta i Batch**")
                            df_batch = df_m.groupby([grp_col, "Batch"]).size().reset_index(name="Liczba")
                            st.dataframe(df_batch, use_container_width=True, height=300)


# ─── KOKPIT ───────────────────────────────────────────────────────────────────
if tab_idx == 7:
    st.markdown('<h1>🎯 Kokpit</h1>', unsafe_allow_html=True)
    if not st.session_state.get("uf_base"):
        st.warning("⚠️ Uploaduj plik Baza.")
    else:
        with st.spinner("Obliczanie..."):
            (df_prof, df_czasy, df_kliki, df_farby_piv,
             machine_cols, warns), tpm_thr, cm_thr = get_data()

        for w in warns:
            st.warning(w)

        if df_prof is not None:
            months = sorted(df_prof["Miesiąc faktury"].dropna().unique())
            sel_months = st.multiselect(
                "Wybierz miesiące",
                months,
                default=months[:1] if months else [],
                key="kokpit_months"
            )

            if not sel_months:
                st.info("Wybierz miesiące.")
            else:
                for month in sel_months:
                    st.markdown(f"---")
                    st.markdown(f'<div class="section-title">📅 Kokpit – {month}</div>',
                                unsafe_allow_html=True)
                    df_m = df_prof[df_prof["Miesiąc faktury"] == month].copy()
                    if df_m.empty:
                        st.info(f"Brak danych dla {month}"); continue

                    klient_col = find_col(df_m, ["Klient", "Klient ID"])
                    grp_col = klient_col if klient_col else "Zlecenie produkcyjne"
                    df_m[grp_col] = df_m[grp_col].fillna("(brak)")

                    total_sv = df_m["Sales Value"].sum()
                    total_tpm = df_m["TPM"].sum()
                    total_cm = df_m["CM"].sum()
                    avg_tpm_pct = (total_tpm / total_sv * 100) if total_sv else 0
                    avg_cm_pct = (total_cm / total_sv * 100) if total_sv else 0
                    n_clients = df_m[grp_col].nunique()
                    n_orders = len(df_m)
                    n_dig = (df_m["Digital/Offset"] == "Digital").sum()
                    n_off = (df_m["Digital/Offset"] == "Offset").sum()
                    n_nop = (df_m["Digital/Offset"] == "no printing").sum()

                    # KPI cards
                    kpi_cols = st.columns(5)
                    kpis = [
                        ("Sprzedaż", f"{total_sv:,.0f} PLN"),
                        ("TPM", f"{total_tpm:,.0f} PLN"),
                        (f"TPM %", f"{avg_tpm_pct:.1f}%"),
                        ("CM", f"{total_cm:,.0f} PLN"),
                        (f"CM %", f"{avg_cm_pct:.1f}%"),
                    ]
                    for i, (label, value) in enumerate(kpis):
                        kpi_cols[i].markdown(
                            f'<div class="kpi-card"><div class="kpi-value">{value}</div>'
                            f'<div class="kpi-label">{label}</div></div>',
                            unsafe_allow_html=True
                        )
                    kpi_cols2 = st.columns(5)
                    kpis2 = [
                        ("Klientów", str(n_clients)),
                        ("Zamówień", str(n_orders)),
                        ("Digital", str(n_dig)),
                        ("Offset", str(n_off)),
                        ("No printing", str(n_nop)),
                    ]
                    for i, (label, value) in enumerate(kpis2):
                        kpi_cols2[i].markdown(
                            f'<div class="kpi-card"><div class="kpi-value">{value}</div>'
                            f'<div class="kpi-label">{label}</div></div>',
                            unsafe_allow_html=True
                        )

                    # Charts
                    grp_kl = df_m.groupby(grp_col).agg(
                        sv=("Sales Value", "sum"),
                        tpm=("TPM", "sum"),
                        cm=("CM", "sum"),
                        n=("Sales Value", "count"),
                    ).reset_index()
                    grp_kl["tpm_pct"] = grp_kl["tpm"] / grp_kl["sv"].replace(0, np.nan) * 100
                    grp_kl["cm_pct"] = grp_kl["cm"] / grp_kl["sv"].replace(0, np.nan) * 100
                    grp_kl.columns = [grp_col if c == grp_col else c for c in grp_kl.columns]

                    BURG = "#6B0000"
                    ORANGE = "#FF5A1F"
                    GREEN = "#0FA958"

                    c1, c2 = st.columns(2)
                    with c1:
                        top5_tpm = grp_kl.nlargest(5, "tpm")
                        fig = px.bar(top5_tpm, x=grp_col, y="tpm",
                                     title="Top 5 klientów wg TPM",
                                     color_discrete_sequence=[BURG])
                        fig.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig, use_container_width=True)
                    with c2:
                        top5_sv = grp_kl.nlargest(5, "sv")
                        fig2 = px.bar(top5_sv, x=grp_col, y="sv",
                                      title="Top 5 klientów wg Sprzedaży",
                                      color_discrete_sequence=[ORANGE])
                        fig2.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig2, use_container_width=True)

                    c3, c4 = st.columns(2)
                    with c3:
                        # Digital/Offset pie
                        do_counts = df_m["Digital/Offset"].value_counts().reset_index()
                        do_counts.columns = ["Typ", "Liczba"]
                        fig3 = px.pie(do_counts, names="Typ", values="Liczba",
                                      title="Digital / Offset / No printing",
                                      color_discrete_sequence=[BURG, ORANGE, "#C91818"])
                        fig3.update_layout(paper_bgcolor="white")
                        st.plotly_chart(fig3, use_container_width=True)
                    with c4:
                        # TPM% per client
                        fig4 = px.bar(grp_kl.sort_values("tpm_pct", ascending=False),
                                      x=grp_col, y="tpm_pct",
                                      title="TPM % wg klientów",
                                      color_discrete_sequence=[GREEN])
                        fig4.add_hline(y=tpm_thr, line_dash="dash", line_color=ORANGE,
                                       annotation_text=f"Próg {tpm_thr:.0f}%")
                        fig4.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig4, use_container_width=True)

                    c5, c6 = st.columns(2)
                    with c5:
                        fig5 = px.bar(grp_kl.sort_values("cm_pct", ascending=False),
                                      x=grp_col, y="cm_pct",
                                      title="CM % wg klientów",
                                      color_discrete_sequence=[BURG])
                        fig5.add_hline(y=cm_thr, line_dash="dash", line_color=ORANGE,
                                       annotation_text=f"Próg {cm_thr:.0f}%")
                        fig5.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig5, use_container_width=True)
                    with c6:
                        # Orders per client
                        fig6 = px.bar(grp_kl.sort_values("n", ascending=False),
                                      x=grp_col, y="n",
                                      title="Liczba zamówień wg klientów",
                                      color_discrete_sequence=[ORANGE])
                        fig6.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig6, use_container_width=True)

                    # Batch x Client heatmap
                    if "Batch" in df_m.columns:
                        df_batch_hm = df_m.groupby([grp_col, "Batch"]).size().reset_index(name="n")
                        pivot_hm = df_batch_hm.pivot(index=grp_col, columns="Batch", values="n").fillna(0)
                        fig7 = px.imshow(pivot_hm, title="Zamówienia: Klient × Batch",
                                         color_continuous_scale=["white", BURG])
                        fig7.update_layout(paper_bgcolor="white")
                        st.plotly_chart(fig7, use_container_width=True)

                    # Alert tables
                    st.markdown("#### ⚠️ Alerty")
                    alert_tpm = grp_kl[grp_kl["tpm_pct"].fillna(0) < tpm_thr]
                    alert_cm = grp_kl[grp_kl["cm_pct"].fillna(0) < cm_thr]
                    ca1, ca2 = st.columns(2)
                    with ca1:
                        st.markdown(f"**TPM % poniżej {tpm_thr:.0f}%**")
                        if not alert_tpm.empty:
                            st.dataframe(
                                alert_tpm[[grp_col, "sv", "tpm", "tpm_pct"]].rename(
                                    columns={"sv": "Sprzedaż", "tpm": "TPM", "tpm_pct": "TPM %"}
                                ).style.format({"Sprzedaż": "{:,.0f}", "TPM": "{:,.0f}", "TPM %": "{:.1f}%"}),
                                use_container_width=True
                            )
                        else:
                            st.success("Brak alertów!")
                    with ca2:
                        st.markdown(f"**CM % poniżej {cm_thr:.0f}%**")
                        if not alert_cm.empty:
                            st.dataframe(
                                alert_cm[[grp_col, "sv", "cm", "cm_pct"]].rename(
                                    columns={"sv": "Sprzedaż", "cm": "CM", "cm_pct": "CM %"}
                                ).style.format({"Sprzedaż": "{:,.0f}", "CM": "{:,.0f}", "CM %": "{:.1f}%"}),
                                use_container_width=True
                            )
                        else:
                            st.success("Brak alertów!")


# ─── POBIERZ XLSX ─────────────────────────────────────────────────────────────

def build_xlsx(df_prof, df_czasy, df_kliki, df_farby_piv,
               machine_cols, rates, click_costs,
               prepress, other_pct, tpm_thr, cm_thr):
    """Build the output XLSX workbook."""
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    # ── Colors ────────────────────────────────────────────────────────────────
    FILL_DL = PatternFill("solid", start_color="C5DCF5")      # light blue
    FILL_MAT = PatternFill("solid", start_color="C5F5D9")     # light green
    FILL_BURG = PatternFill("solid", start_color="6B0000")
    FILL_HEADER = PatternFill("solid", start_color="3A0000")
    FONT_WHITE = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    FONT_BOLD = Font(bold=True, name="Arial", size=10)
    FONT_REG = Font(name="Arial", size=10)
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
    THIN = Side(style="thin", color="CCCCCC")
    THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    fmt_currency = '#,##0.00'
    fmt_pct = '0.0%'
    fmt_date = 'YYYY-MM-DD'
    fmt_ym = 'YYYY-MM'

    HIDDEN_COLS = {
        "Kontrbigi [29]", "Duplex [30]", "Offset [31]",
        "Add on evey box [32]", "Autobottom / ekstra lim [33]",
        "Pantone [34]", "E-flute [35]", "Item number inlay [36]",
        "POD price [37]", "Clames [38]", "Transport [39]",
        "Kooperacja [40]", "B2 price [41]", "B1 price [42]",
        "Energia [43]", "Fix price [44]", "Cena TKW [45]",
        "Click price [46]",
        "Płyty lakierujące [21]", "Matryca Braille- Grawer [22]",
        "Patryca Braille- Wewnatrz [23]",
    }
    # normalize hidden cols
    hidden_norm = {" ".join(h.split()).lower() for h in HIDDEN_COLS}

    DL_GROUP = set(machine_cols + ["Prepress costs", "Total DL"])
    MAT_GROUP = {
        "Papier [16]", "Klej [17]", "Lakiery [20]",
        "Opakowania zbiorcze [24]", "Other Materials",
        "Offset inks", "Płyta offsetowa", "Kliki final", "Total Materials"
    }

    def write_df_sheet(ws, df, title_fill=None, col_fills=None, freeze=True,
                       hidden_set=None, currency_cols=None, pct_cols=None,
                       date_cols=None, ym_cols=None):
        """Write a DataFrame to a worksheet with formatting."""
        cols = list(df.columns)
        # header row
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(row=1, column=ci, value=col)
            col_norm = " ".join(col.split()).lower()
            if col_fills and col in col_fills:
                cell.fill = col_fills[col]
            else:
                cell.fill = FILL_HEADER
            cell.font = FONT_WHITE
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER

        # data rows
        for ri, (_, row) in enumerate(df.iterrows(), 2):
            for ci, col in enumerate(cols, 1):
                val = row[col]
                if pd.isna(val) if not isinstance(val, str) else False:
                    val = None
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font = FONT_REG
                cell.alignment = ALIGN_LEFT
                cell.border = THIN_BORDER
                col_norm = " ".join(col.split()).lower()
                # format
                if currency_cols and col in currency_cols:
                    cell.number_format = fmt_currency
                if pct_cols and col in pct_cols:
                    cell.number_format = fmt_pct
                if date_cols and col in date_cols:
                    cell.number_format = fmt_date
                if ym_cols and col in ym_cols:
                    cell.number_format = '@'  # text
                # light row fill for groups
                if col_fills and col in col_fills:
                    base_fill = col_fills[col]
                    light = PatternFill("solid", start_color=base_fill.fgColor.rgb)
                    cell.fill = light

        # autofit columns
        for ci, col in enumerate(cols, 1):
            max_len = max(
                len(str(col)),
                *[len(str(df.iloc[ri][col]) if not pd.isna(df.iloc[ri][col]) else "") for ri in range(min(len(df), 50))]
            ) if len(df) > 0 else len(str(col))
            ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 10), 40)

        # hide columns
        if hidden_set:
            for ci, col in enumerate(cols, 1):
                col_norm = " ".join(col.split()).lower()
                if col_norm in {" ".join(h.split()).lower() for h in hidden_set}:
                    ws.column_dimensions[get_column_letter(ci)].hidden = True

        # filter & freeze
        ws.auto_filter.ref = ws.dimensions
        if freeze:
            ws.freeze_panes = "A2"

    # ── PROFITABILITY sheet ───────────────────────────────────────────────────
    ws_prof = wb.create_sheet("Profitability")
    df_prof_xlsx = df_prof.copy()

    # Build col_fills for Profitability
    col_fills_prof = {}
    for col in df_prof_xlsx.columns:
        cn = " ".join(col.split())
        if cn in DL_GROUP or col in DL_GROUP:
            col_fills_prof[col] = FILL_DL
        elif cn in MAT_GROUP or col in MAT_GROUP:
            col_fills_prof[col] = FILL_MAT

    currency_cols = {
        c for c in df_prof_xlsx.columns
        if any(k in c for k in ["Wartość", "Wartosc", "koszt", "Koszt", "inks",
                                  "Sales", "TPM", "CM", "Total", "Prepress",
                                  "Materials", "Offset", "Płyta", "Kliki",
                                  "Papier", "Klej", "Lak", "Opak", "Farby"])
    }
    pct_cols = {"Other costs %"}
    date_cols = {"Data faktury"}
    ym_cols = {"Miesiąc faktury"}

    write_df_sheet(
        ws_prof, df_prof_xlsx,
        col_fills=col_fills_prof,
        hidden_set=HIDDEN_COLS,
        currency_cols=currency_cols,
        pct_cols=pct_cols,
        date_cols=date_cols,
        ym_cols=ym_cols,
    )

    # ── CZASY sheet ───────────────────────────────────────────────────────────
    if df_czasy is not None:
        ws_czasy = wb.create_sheet("czasy")
        write_df_sheet(ws_czasy, df_czasy.drop(
            columns=["_zp", "_machine", "_rate"], errors="ignore"
        ))

    # ── KLIKI sheet ───────────────────────────────────────────────────────────
    if df_kliki is not None:
        ws_kliki = wb.create_sheet("Kliki")
        write_df_sheet(ws_kliki, df_kliki.drop(
            columns=["Separations_n"], errors="ignore"
        ))

    # ── FARBY OFFSET sheet ────────────────────────────────────────────────────
    if df_farby_piv is not None:
        ws_farby = wb.create_sheet("Farby Offset")
        df_fp = df_farby_piv.drop(columns=["_ew"], errors="ignore")
        write_df_sheet(ws_farby, df_fp)

    # ── STAWKI sheet ──────────────────────────────────────────────────────────
    ws_st = wb.create_sheet("Stawki")
    df_st = pd.DataFrame(
        [(k, v) for k, v in rates.items()],
        columns=["Nazwa maszyny", "Stawka rbg (PLN/h)"]
    )
    write_df_sheet(ws_st, df_st)

    # ── KOSZTY KLIKÓW sheet ───────────────────────────────────────────────────
    ws_kk = wb.create_sheet("Koszty klików")
    cc_rows = [
        {"Maszyna": k, "Kolor": c, "Koszt (PLN/sep)": v}
        for k, colors in click_costs.items()
        for c, v in colors.items()
    ]
    df_kk = pd.DataFrame(cc_rows) if cc_rows else pd.DataFrame(
        columns=["Maszyna", "Kolor", "Koszt (PLN/sep)"]
    )
    write_df_sheet(ws_kk, df_kk)

    # ── PREPRESS sheet ────────────────────────────────────────────────────────
    ws_pp = wb.create_sheet("Prepress")
    pp_rows = [
        {"Klient": k, "Stawka Digital": v["digital"], "Stawka Offset": v["offset"]}
        for k, v in prepress.items()
    ]
    df_pp = pd.DataFrame(pp_rows) if pp_rows else pd.DataFrame(
        columns=["Klient", "Stawka Digital", "Stawka Offset"]
    )
    write_df_sheet(ws_pp, df_pp)

    # ── PARAMETRY sheet ───────────────────────────────────────────────────────
    ws_par = wb.create_sheet("Parametry")
    df_par = pd.DataFrame([
        {"Parametr": "Other costs %", "Wartość": other_pct / 100},
        {"Parametr": "Próg alertu TPM %", "Wartość": tpm_thr / 100},
        {"Parametr": "Próg alertu CM %", "Wartość": cm_thr / 100},
    ])
    write_df_sheet(ws_par, df_par, pct_cols={"Wartość"})

    # ── PODSUMOWANIE sheets (per month) ───────────────────────────────────────
    months = sorted(df_prof["Miesiąc faktury"].dropna().unique())
    klient_col = find_col(df_prof, ["Klient", "Klient ID"])

    for month in months:
        df_m = df_prof[df_prof["Miesiąc faktury"] == month].copy()
        grp_col = klient_col if klient_col else "Zlecenie produkcyjne"
        if grp_col not in df_m.columns:
            continue
        df_m[grp_col] = df_m[grp_col].fillna("(brak)")

        rows_sum = []
        for klient, grp in df_m.groupby(grp_col):
            sv = grp["Sales Value"].sum()
            tpm = grp["TPM"].sum()
            cm = grp["CM"].sum()
            n = len(grp)
            rows_sum.append({
                "Klient": klient,
                "Miesiąc": month,
                "Suma sprzedaży": sv,
                "Suma TPM": tpm,
                "TPM %": tpm / sv if sv else 0,
                "Suma CM": cm,
                "CM %": cm / sv if sv else 0,
                "Liczba zamówień": n,
                "Digital": (grp["Digital/Offset"] == "Digital").sum(),
                "Offset": (grp["Digital/Offset"] == "Offset").sum(),
                "No printing": (grp["Digital/Offset"] == "no printing").sum(),
            })
        df_sum_m = pd.DataFrame(rows_sum)
        safe_month = month.replace("/", "-").replace(":", "-")
        ws_sum = wb.create_sheet(f"Podsumowanie {safe_month}")
        write_df_sheet(
            ws_sum, df_sum_m,
            currency_cols={"Suma sprzedaży", "Suma TPM", "Suma CM"},
            pct_cols={"TPM %", "CM %"},
        )

        # Batch breakdown
        if "Batch" in df_m.columns:
            ws_batch = wb.create_sheet(f"Batch {safe_month}")
            df_batchs = df_m.groupby([grp_col, "Batch"]).size().reset_index(name="Liczba zamówień")
            write_df_sheet(ws_batch, df_batchs)

    # ── KOKPIT sheet (last selected month or all) ─────────────────────────────
    if months:
        ws_kok = wb.create_sheet("Kokpit")
        all_kpi_rows = []
        for month in months:
            df_m = df_prof[df_prof["Miesiąc faktury"] == month]
            sv = df_m["Sales Value"].sum()
            tpm_v = df_m["TPM"].sum()
            cm_v = df_m["CM"].sum()
            all_kpi_rows.append({
                "Miesiąc": month,
                "Suma sprzedaży": sv,
                "Suma TPM": tpm_v,
                "TPM %": tpm_v / sv if sv else 0,
                "Suma CM": cm_v,
                "CM %": cm_v / sv if sv else 0,
                "Klientów": df_m[klient_col].nunique() if klient_col else 0,
                "Zamówień": len(df_m),
                "Digital": (df_m["Digital/Offset"] == "Digital").sum(),
                "Offset": (df_m["Digital/Offset"] == "Offset").sum(),
                "No printing": (df_m["Digital/Offset"] == "no printing").sum(),
            })
        df_kok = pd.DataFrame(all_kpi_rows)
        write_df_sheet(
            ws_kok, df_kok,
            currency_cols={"Suma sprzedaży", "Suma TPM", "Suma CM"},
            pct_cols={"TPM %", "CM %"},
        )

    # Save to BytesIO
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


if tab_idx == 8:
    st.markdown('<h1>⬇️ Pobierz XLSX</h1>', unsafe_allow_html=True)
    if not st.session_state.get("uf_base"):
        st.warning("⚠️ Uploaduj plik Baza w zakładce 'Upload plików'.")
    else:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(
            "Kliknij **Generuj plik XLSX**, aby obliczyć Profitability i pobrać gotowy plik."
        )

        uf_stawki = st.session_state.get("uf_stawki")
        rates = get_rates(uf_stawki)

        if st.button("🔄 Generuj plik XLSX"):
            with st.spinner("Obliczanie i generowanie pliku..."):
                (df_prof, df_czasy, df_kliki, df_farby_piv,
                 machine_cols, warns), tpm_thr, cm_thr = get_data()

            for w in warns:
                st.warning(w)

            if df_prof is not None:
                with st.spinner("Budowanie XLSX..."):
                    xlsx_buf = build_xlsx(
                        df_prof, df_czasy, df_kliki, df_farby_piv,
                        machine_cols, rates,
                        st.session_state.get("click_costs", {}),
                        st.session_state.get("prepress", {}),
                        st.session_state.get("other_costs_pct", 2.0),
                        tpm_thr, cm_thr,
                    )

                st.success("✅ Plik gotowy do pobrania!")
                st.download_button(
                    label="⬇️ Pobierz Profitability.xlsx",
                    data=xlsx_buf,
                    file_name="Profitability.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("Błąd generowania danych.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("#### Zawartość pliku XLSX")
        st.markdown("""
<div class="card">
<ul>
<li><b>Profitability</b> – arkusz główny z pełną kalkulacją</li>
<li><b>czasy</b> – dane z czasów produkcji + koszt pracy</li>
<li><b>Kliki</b> – dane z inks + koszt klików</li>
<li><b>Farby Offset</b> – Pivot farby z kosztami</li>
<li><b>Stawki</b> – stawki maszyn</li>
<li><b>Koszty klików</b> – tabela kosztów klików</li>
<li><b>Prepress</b> – stawki Prepress per klient</li>
<li><b>Parametry</b> – inne parametry (Other costs %, progi alertów)</li>
<li><b>Podsumowanie YYYY-MM</b> – osobno dla każdego miesiąca</li>
<li><b>Batch YYYY-MM</b> – zestawienie wg Batch per miesiąc</li>
<li><b>Kokpit</b> – KPI dla wszystkich miesięcy</li>
</ul>
</div>
""", unsafe_allow_html=True)
