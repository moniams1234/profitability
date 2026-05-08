import io
import re
import warnings
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Postkalkulacja Profitability",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# CSS / fancy UI
# =========================
st.markdown(
    """
<style>
[data-testid="stAppViewContainer"]{background:#F7EFEA;}
[data-testid="stHeader"]{background:transparent;}
[data-testid="stSidebar"]{background:#4A0000!important;border-right:2px solid #3A0000;}
[data-testid="stSidebar"] *{color:#FAE8E0!important;}
h1,h2,h3{color:#3A0000!important;}
.block-container{padding-top:2rem;}
.hero{
    background:linear-gradient(135deg,#6B0000,#3A0000);
    color:white;border-radius:18px;padding:24px 30px;margin-bottom:24px;
    box-shadow:0 10px 25px rgba(74,0,0,.18);
}
.hero h1{color:white!important;margin:0;font-size:30px;}
.hero p{margin:4px 0 0;color:#FFE0D0;letter-spacing:.08em;text-transform:uppercase;font-size:12px;}
.card{
    background:white;border-radius:16px;padding:20px 22px;margin:12px 0;
    box-shadow:0 6px 18px rgba(74,0,0,.08);border:1px solid rgba(74,0,0,.08);
}
.kpi{
    background:white;border-radius:16px;padding:16px;text-align:center;
    box-shadow:0 5px 16px rgba(74,0,0,.08);border-top:4px solid #FF5A1F;
}
.kpi .v{font-size:26px;font-weight:800;color:#6B0000;}
.kpi .l{font-size:12px;color:#6B6B6B;text-transform:uppercase;letter-spacing:.06em;}
.stButton>button,.stDownloadButton>button{
    background:linear-gradient(135deg,#6B0000,#FF5A1F)!important;color:white!important;
    border:0!important;border-radius:12px!important;font-weight:700!important;
    box-shadow:0 5px 14px rgba(107,0,0,.18)!important;
}
.stTabs [data-baseweb="tab-list"]{gap:6px;}
.stTabs [data-baseweb="tab"]{background:white;border-radius:12px;padding:8px 14px;}
.stTabs [aria-selected="true"]{background:#6B0000!important;color:white!important;}
[data-testid="stDataFrame"]{border-radius:12px;overflow:hidden;}
</style>
""",
    unsafe_allow_html=True,
)

# =========================
# Constants
# =========================
REQUIRED_BASE = [
    "Numer", "Zamówienie", "Zamawiana ilość", "Papier [16]", "Klej [17]",
    "Lakiery [20]", "Opakowania zbiorcze [24]", "Farby [19]", "Kliki [48]"
]
HIDDEN_COLS = [
    "Kontrbigi [29]", "Duplex [30]", "Offset [31]", "Add on evey box [32]",
    "Autobottom / ekstra lim [33]", "Pantone [34]", "E-flute [35]",
    "Item number inlay [36]", "POD price [37]", "Clames [38]", "Transport [39]",
    "Kooperacja [40]", "B2 price [41]", "B1 price [42]", "Energia [43]",
    "Fix price [44]", "Cena TKW [45]", "Click price [46]", "Płyty lakierujące [21]",
    "Matryca Braille- Grawer [22]", "Patryca Braille- Wewnatrz [23]"
]
BATCHES = [
    (50, "0-50"), (100, "51-100"), (200, "101-200"), (300, "201-300"),
    (500, "301-500"), (1000, "501-1000"), (1500, "1001-1500"),
    (2000, "1501-2000"), (3000, "2001-3000"), (10000, "3001-10000"),
    (20000, "10001-20000"), (30000, "20001-30000"),
    (100000, "30001-100000"), (1000000, "100001-1000000")
]

# =========================
# Helpers
# =========================
def norm_text(x) -> str:
    return re.sub(r"\s+", " ", str(x).replace("\n", " ").replace("\r", " ")).strip()

def norm_key(x) -> str:
    x = norm_text(x).lower()
    x = x.replace("ą","a").replace("ć","c").replace("ę","e").replace("ł","l").replace("ń","n").replace("ó","o").replace("ś","s").replace("ż","z").replace("ź","z")
    return x

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [norm_text(c) for c in df.columns]
    return df

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    lookup = {norm_key(c): c for c in df.columns}
    for cand in candidates:
        if norm_key(cand) in lookup:
            return lookup[norm_key(cand)]
    # contains fallback
    for cand in candidates:
        nk = norm_key(cand)
        for k, v in lookup.items():
            if nk == k or nk in k or k in nk:
                return v
    return None

def to_num(s):
    if isinstance(s, pd.Series):
        return pd.to_numeric(
            s.astype(str).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce",
        ).fillna(0.0)
    try:
        return float(str(s).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0

def safe_read_excel(uploaded, sheet_name=0, header="auto") -> Optional[pd.DataFrame]:
    if uploaded is None:
        return None
    try:
        raw = uploaded.getvalue()
        xl = pd.ExcelFile(io.BytesIO(raw))
        if isinstance(sheet_name, str) and sheet_name not in xl.sheet_names:
            # choose similar sheet
            lower = {s.lower(): s for s in xl.sheet_names}
            found = None
            for s in xl.sheet_names:
                if sheet_name.lower() in s.lower() or s.lower() in sheet_name.lower():
                    found = s
                    break
            sheet_name = found or xl.sheet_names[0]
        if header == "auto":
            probe = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=None, nrows=20)
            hrow = 0
            best = 0
            for i, row in probe.iterrows():
                vals = [norm_text(v) for v in row.dropna().tolist()]
                score = sum(1 for v in vals if len(v) > 1)
                score += sum(4 for v in vals if any(k in norm_key(v) for k in ["numer", "klient", "wartosc", "faktury", "zamowienie", "nazwa", "czas"]))
                if score > best:
                    best = score
                    hrow = i
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=hrow)
        else:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=header)
        df = normalize_columns(df)
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.warning(f"Nie można wczytać pliku: {e}")
        return None

def read_base(uploaded) -> Optional[pd.DataFrame]:
    if uploaded is None:
        return None
    # Try common post_list header row 3 first, then auto.
    for h in [3, "auto", 0]:
        df = safe_read_excel(uploaded, 0, h)
        if df is not None and find_col(df, ["Numer"]) and find_col(df, ["Zamówienie"]):
            return df
    return safe_read_excel(uploaded, 0, "auto")

def batch_label(q):
    q = to_num(q)
    if q <= 0:
        return ""
    for limit, label in BATCHES:
        if q <= limit:
            return label
    return "100001-1000000"

def order_fragment(zamowienie: str) -> str:
    """Extract e.g. KOH... | 260220_Digital [K...] -> 260220_Digital."""
    s = norm_text(zamowienie)
    if "| " in s:
        s = s.split("| ", 1)[1]
    elif "|" in s:
        s = s.split("|", 1)[1].strip()
    if " [" in s:
        s = s.split(" [", 1)[0]
    if "|" in s:
        s = s.split("|", 1)[0].strip()
    return s.strip()

def get_month(dt):
    if pd.isna(dt):
        return ""
    try:
        return pd.to_datetime(dt).strftime("%Y-%m")
    except Exception:
        return ""

def load_rates(uploaded) -> pd.DataFrame:
    df = safe_read_excel(uploaded, 0, "auto") if uploaded else None
    if df is None:
        return pd.DataFrame(columns=["Nazwa maszyny", "Stawka rbg"])
    m = find_col(df, ["Nazwa maszyny", "Maszyna"])
    r = find_col(df, ["Stawka rbg", "Stawka", "rate"])
    if not m or not r:
        return pd.DataFrame(columns=["Nazwa maszyny", "Stawka rbg"])
    out = df[[m, r]].copy()
    out.columns = ["Nazwa maszyny", "Stawka rbg"]
    out["Stawka rbg"] = to_num(out["Stawka rbg"])
    out = out.dropna(subset=["Nazwa maszyny"])
    return out

def load_click_costs(uploaded) -> pd.DataFrame:
    df = safe_read_excel(uploaded, 0, "auto") if uploaded else None
    if df is None or df.empty:
        return pd.DataFrame([
            {"Press Name": "HP 35", "Color": "default", "Koszt jednostkowy": 0.05},
            {"Press Name": "HP 7", "Color": "default", "Koszt jednostkowy": 0.05},
            {"Press Name": "HP 1", "Color": "default", "Koszt jednostkowy": 0.05},
        ])
    p = find_col(df, ["Press Name", "Maszyna", "Nazwa maszyny"])
    c = find_col(df, ["Color", "Kolor"])
    k = find_col(df, ["Koszt jednostkowy", "Koszt", "Click cost", "Cena"])
    if not p:
        p = df.columns[0]
    if not c:
        df["Color"] = "default"; c = "Color"
    if not k:
        last = df.columns[-1]
        k = last
    out = df[[p, c, k]].copy()
    out.columns = ["Press Name", "Color", "Koszt jednostkowy"]
    out["Koszt jednostkowy"] = to_num(out["Koszt jednostkowy"])
    return out.dropna(subset=["Press Name"])

def machine_rate(machine: str, rates_df: pd.DataFrame) -> float:
    if rates_df is None or rates_df.empty:
        return 0.0
    mk = norm_key(machine)
    for _, r in rates_df.iterrows():
        rk = norm_key(r["Nazwa maszyny"])
        if rk and (rk in mk or mk in rk):
            return to_num(r["Stawka rbg"])
    return 0.0

def click_unit_cost(press, color, click_df: pd.DataFrame) -> float:
    if click_df is None or click_df.empty:
        return 0.05
    pk = norm_key(press)
    ck = norm_key(color)
    for _, r in click_df.iterrows():
        rp = norm_key(r["Press Name"])
        rc = norm_key(r["Color"])
        if rp and (rp in pk or pk in rp) and (rc == ck or rc == "default" or not rc):
            return to_num(r["Koszt jednostkowy"])
    return to_num(click_df.iloc[0]["Koszt jednostkowy"]) if len(click_df) else 0.05

def classify_printing(zp: str, czasy: Optional[pd.DataFrame], zp_col: Optional[str], mach_col: Optional[str]) -> str:
    if czasy is None or not zp_col or not mach_col:
        return ""
    zp = norm_text(zp)
    sub = czasy[czasy[zp_col].astype(str).map(norm_text) == zp]
    if sub.empty:
        return ""
    machines = " | ".join(sub[mach_col].astype(str).map(norm_key).tolist())
    has_hp = any(h in machines for h in ["hp 35", "hp35", "hp 7", "hp7", "hp 1", "hp1"])
    has_heid = "heidelberg cx 104" in machines or "cx 104" in machines
    if has_hp:
        return "Digital"
    if has_heid:
        return "Offset"
    return "no printing"

# =========================
# Core calculation
# =========================
def build_profitability(
    files: Dict[str, object],
    rates_df: pd.DataFrame,
    click_df: pd.DataFrame,
    prepress_df: pd.DataFrame,
    other_pct: float,
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame], List[str], List[str]]:
    warnings_list = []
    aux = {}
    base = read_base(files.get("base"))
    if base is None or base.empty:
        raise ValueError("Brak poprawnego pliku Baza / post_list.")

    # required column warnings
    for c in REQUIRED_BASE:
        if not find_col(base, [c]):
            warnings_list.append(f"Baza: brakuje kolumny „{c}” – brakujące wartości zostaną przyjęte jako 0/puste, jeśli możliwe.")

    df = base.copy()

    numer = find_col(df, ["Numer"])
    zam = find_col(df, ["Zamówienie"])
    qty = find_col(df, ["Zamawiana ilość", "Zamawiana ilosc"])
    klient = find_col(df, ["Klient", "Klient ID"])

    df["Zlecenie produkcyjne"] = df[numer].astype(str).str.split("-", n=1).str[0].str.strip() if numer else ""
    df["Lewy 10"] = df[zam].astype(str).str[:10] if zam else ""
    df["Batch"] = df[qty].apply(batch_label) if qty else ""

    # CZASY
    machine_cols = []
    czasy = safe_read_excel(files.get("czasy"), 0, "auto") if files.get("czasy") else None
    if czasy is not None:
        zp_c = find_col(czasy, ["Numer zlecenia produkcyjnego"])
        mach_c = find_col(czasy, ["Nazwa maszyny"])
        time_c = find_col(czasy, ["Czas czynnosci [min]", "Czas czynności [min]"])
        if zp_c and mach_c and time_c:
            czasy["_zp"] = czasy[zp_c].astype(str).map(norm_text)
            czasy["_machine"] = czasy[mach_c].astype(str).map(norm_text)
            czasy["_minutes"] = to_num(czasy[time_c])
            czasy["_rate"] = czasy["_machine"].apply(lambda x: machine_rate(x, rates_df))
            czasy["Koszt pracy"] = czasy["_minutes"] / 60 * czasy["_rate"]
            aux["czasy"] = czasy
            # pivot costs
            pivot = czasy.pivot_table(index="_zp", columns="_machine", values="Koszt pracy", aggfunc="sum", fill_value=0)
            pivot = pivot.reset_index()
            machine_cols = [c for c in pivot.columns if c != "_zp"]
            df = df.merge(pivot, left_on="Zlecenie produkcyjne", right_on="_zp", how="left").drop(columns=["_zp"], errors="ignore")
            for c in machine_cols:
                df[c] = to_num(df[c])
            df["Digital/Offset"] = df["Zlecenie produkcyjne"].apply(lambda z: classify_printing(z, czasy, zp_c, mach_c))
        else:
            warnings_list.append("Czasy: brak kolumn Numer zlecenia produkcyjnego / Nazwa maszyny / Czas czynnosci [min].")
            df["Digital/Offset"] = ""
    else:
        warnings_list.append("Brak pliku Czasy – Total DL nie zawiera kosztów pracy maszyn.")
        df["Digital/Offset"] = ""

    # FARBY
    df["Offset inks"] = 0.0
    df["Płyta offsetowa"] = 0.0
    if files.get("farby"):
        try:
            raw = files["farby"].getvalue()
            xl = pd.ExcelFile(io.BytesIO(raw))
            sheet = next((s for s in xl.sheet_names if "pivot" in s.lower() or "farb" in s.lower()), xl.sheet_names[0])
            farby = safe_read_excel(files["farby"], sheet, "auto")
            if farby is not None:
                key = find_col(farby, ["Etykieta wiersza", "Etykiety wierszy"])
                ink = find_col(farby, ["Suma koszt farby2", "Koszt farby2"])
                plyta = find_col(farby, ["Suma koszt płyty", "Suma koszt plyty", "Koszt płyty"])
                if key:
                    aux["Pivot farby"] = farby
                    tmp = farby.copy()
                    tmp["_key"] = tmp[key].astype(str).str.strip()
                    tmp["_ink"] = to_num(tmp[ink]) if ink else 0
                    tmp["_plyta"] = to_num(tmp[plyta]) if plyta else 0
                    tmp = tmp.groupby("_key")[["_ink", "_plyta"]].sum().reset_index()
                    df = df.merge(tmp, left_on="Lewy 10", right_on="_key", how="left")
                    df["Offset inks"] = to_num(df["_ink"])
                    df["Płyta offsetowa"] = to_num(df["_plyta"])
                    df = df.drop(columns=["_key", "_ink", "_plyta"], errors="ignore")
        except Exception as e:
            warnings_list.append(f"Farby: nie udało się wczytać pliku ({e}).")
    else:
        warnings_list.append("Brak pliku Farby – Offset inks i Płyta offsetowa = 0.")

    # KLIKI
    df["Moje Kliki"] = 0.0
    if files.get("inks"):
        inks = safe_read_excel(files.get("inks"), 0, "auto")
        if inks is not None:
            job = find_col(inks, ["Job Name"])
            press = find_col(inks, ["Press Name"])
            color = find_col(inks, ["Color"])
            sep = find_col(inks, ["Separations"])
            if job and sep:
                inks["Zamówienie"] = inks[job].astype(str).str[:10]
                inks["_sep"] = to_num(inks[sep])
                inks["Koszt klików"] = inks.apply(lambda r: r["_sep"] * click_unit_cost(r.get(press, ""), r.get(color, ""), click_df), axis=1)
                aux["Kliki"] = inks
                kg = inks.groupby("Zamówienie")["Koszt klików"].sum().reset_index()
                df = df.merge(kg.rename(columns={"Koszt klików":"Moje Kliki"}), left_on="Lewy 10", right_on="Zamówienie", how="left", suffixes=("", "_k"))
                df = df.drop(columns=["Zamówienie_k"], errors="ignore")
                df["Moje Kliki"] = to_num(df["Moje Kliki"])
            else:
                warnings_list.append("Kliki/Inks: brak kolumn Job Name lub Separations.")
    else:
        warnings_list.append("Brak pliku Kliki/Inks – Moje Kliki = 0.")

    kliki_col = find_col(df, ["Kliki [48]"])
    df["Kliki final"] = np.maximum(to_num(df[kliki_col]) if kliki_col else 0, to_num(df["Moje Kliki"]))

    # SALES VALUE AND DATE - first zlec
    df["Sales Value"] = np.nan
    df["Data faktury"] = pd.NaT
    zlec = safe_read_excel(files.get("zlec"), 0, "auto") if files.get("zlec") else None
    if zlec is not None:
        aux["zlec + faktury"] = zlec
        z_zp = find_col(zlec, ["Numer Zlecenia produkcyjnego", "Numer zlecenia produkcyjnego"])
        z_val = find_col(zlec, ["Wartosc w linii FV netto", "Wartość w linii FV netto"])
        z_dt = find_col(zlec, ["Data wystawienia faktury", "Data wystawienia FV"])
        if z_zp and z_val:
            tmp = zlec.copy()
            tmp["_zp"] = tmp[z_zp].astype(str).map(norm_text)
            tmp["_val"] = to_num(tmp[z_val])
            tmp["_dt"] = pd.to_datetime(tmp[z_dt], errors="coerce") if z_dt else pd.NaT
            tmp = tmp.groupby("_zp").agg({"_val":"sum","_dt":"first"}).reset_index()
            df = df.merge(tmp, left_on="Zlecenie produkcyjne", right_on="_zp", how="left")
            found = df["_val"].notna()
            df.loc[found, "Sales Value"] = df.loc[found, "_val"]
            df.loc[found, "Data faktury"] = df.loc[found, "_dt"]
            df = df.drop(columns=["_zp","_val","_dt"], errors="ignore")
        else:
            warnings_list.append("Zlecenia + faktury: brak kolumn do Sales Value/Data faktury.")
    else:
        warnings_list.append("Brak pliku Zlecenia + faktury – Sales Value będzie szukane w Faktury linie, jeśli dostępne.")

    # fallback fry lines
    fry = safe_read_excel(files.get("fry"), 0, "auto") if files.get("fry") else None
    if fry is not None:
        aux["faktury linie"] = fry
        nl = find_col(fry, ["Nazwa linii faktury", "Nazwa lini faktury"])
        val = find_col(fry, ["Wartosc w linii FV netto", "Wartość w linii FV netto"])
        il = find_col(fry, ["Ilosc w linii FV", "Ilość w linii FV"])
        dt = find_col(fry, ["Data wystawienia FV", "Data wystawienia faktury"])
        if nl and val:
            needs = df["Sales Value"].isna() | (to_num(df["Sales Value"]) == 0)
            for idx in df[needs].index:
                frag = order_fragment(df.at[idx, zam]) if zam else ""
                if not frag:
                    continue
                mask = fry[nl].astype(str).str.contains(re.escape(frag), case=False, na=False)
                if not mask.any():
                    continue
                row = fry.loc[mask].iloc[0]
                amount = to_num(row[val])
                inv_qty = to_num(row[il]) if il else 0
                order_qty = to_num(df.at[idx, qty]) if qty else 0
                df.at[idx, "Sales Value"] = (amount / inv_qty * order_qty) if inv_qty and order_qty else amount
                if dt:
                    df.at[idx, "Data faktury"] = pd.to_datetime(row[dt], errors="coerce")
        else:
            warnings_list.append("Faktury linie: brak kolumn Nazwa linii faktury / Wartosc w linii FV netto.")
    else:
        warnings_list.append("Brak pliku Faktury linie – fallback Sales Value/Data faktury niedostępny.")

    df["Sales Value"] = to_num(df["Sales Value"])
    df["Data faktury"] = pd.to_datetime(df["Data faktury"], errors="coerce")
    df["Miesiąc faktury"] = df["Data faktury"].apply(get_month)

    # PREPRESS
    if prepress_df is None or prepress_df.empty:
        prepress_df = pd.DataFrame(columns=["Klient", "Stawka Prepress Digital", "Stawka Prepress Offset"])
    def pp(row):
        k = norm_text(row.get(klient, "")) if klient else ""
        offset_inks = to_num(row.get("Offset inks", 0))
        sub = prepress_df[prepress_df["Klient"].astype(str).map(norm_text) == k] if "Klient" in prepress_df.columns else pd.DataFrame()
        if not sub.empty:
            return to_num(sub.iloc[0]["Stawka Prepress Offset" if offset_inks > 0 else "Stawka Prepress Digital"])
        return 40.0 if offset_inks > 0 else 10.0
    df["Prepress costs"] = df.apply(pp, axis=1)

    # OTHER MATERIALS
    df["Other Materials"] = df["Sales Value"] * (other_pct / 100)

    # Total DL
    for c in machine_cols:
        df[c] = to_num(df[c])
    df["Prepress costs"] = to_num(df["Prepress costs"])
    df["Total DL"] = df[machine_cols + ["Prepress costs"]].sum(axis=1) if machine_cols else df["Prepress costs"]

    # Total Materials
    material_cols = ["Papier [16]", "Klej [17]", "Lakiery [20]", "Opakowania zbiorcze [24]"]
    real_material_cols = []
    for mc in material_cols:
        rc = find_col(df, [mc])
        if rc:
            df[rc] = to_num(df[rc])
            real_material_cols.append(rc)
        else:
            df[mc] = 0.0
            real_material_cols.append(mc)
    all_mat = real_material_cols + ["Other Materials", "Offset inks", "Płyta offsetowa", "Kliki final"]
    df["Total Materials"] = df[all_mat].sum(axis=1)

    # margins
    df["TPM"] = df["Sales Value"] - df["Total DL"]
    df["CM"] = df["Sales Value"] - df["Total DL"] - df["Total Materials"]

    # final order: keep originals, then derived, move final block before TPM
    final_block = ["Total DL", "Total Materials", "Sales Value", "Data faktury", "Miesiąc faktury", "TPM", "CM"]
    derived = ["Zlecenie produkcyjne", "Lewy 10", "Batch", "Digital/Offset", "Moje Kliki", "Kliki final", "Offset inks", "Płyta offsetowa", "Prepress costs", "Other Materials"]
    ordered = [c for c in df.columns if c not in derived + final_block] + derived + final_block
    df = df[[c for c in ordered if c in df.columns]]
    return df, aux, warnings_list, machine_cols

def make_summary(df: pd.DataFrame, selected_month: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    data = df.copy()
    if selected_month and selected_month != "Wszystkie":
        data = data[data["Miesiąc faktury"] == selected_month]
    klient = find_col(data, ["Klient", "Klient ID"]) or "Zlecenie produkcyjne"
    if klient not in data.columns:
        data[klient] = "(brak)"
    rows = []
    for k, g in data.groupby(klient, dropna=False):
        sv = g["Sales Value"].sum()
        tpm = g["TPM"].sum()
        cm = g["CM"].sum()
        rows.append({
            "Klient": k,
            "Miesiąc faktury": selected_month,
            "Suma sprzedaży": sv,
            "Suma TPM": tpm,
            "TPM %": tpm / sv if sv else 0,
            "Suma CM": cm,
            "CM %": cm / sv if sv else 0,
            "Liczba zamówień": len(g),
            "Liczba Digital": (g["Digital/Offset"] == "Digital").sum(),
            "Liczba Offset": (g["Digital/Offset"] == "Offset").sum(),
            "Liczba no printing": (g["Digital/Offset"] == "no printing").sum(),
        })
    summary = pd.DataFrame(rows)
    if "Batch" in data.columns:
        batch = data.groupby([klient, "Batch"], dropna=False).size().reset_index(name="Liczba zamówień").rename(columns={klient:"Klient"})
    else:
        batch = pd.DataFrame(columns=["Klient", "Batch", "Liczba zamówień"])
    return summary, batch

# =========================
# XLSX export
# =========================
def write_df(ws, df: pd.DataFrame, fills: Dict[str, PatternFill] = None, hidden: List[str] = None):
    fills = fills or {}
    hidden_norm = {norm_key(x) for x in (hidden or [])}
    header_fill = PatternFill("solid", fgColor="3A0000")
    white = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cidx, col in enumerate(df.columns, 1):
        cell = ws.cell(1, cidx, col)
        cell.fill = fills.get(col, header_fill)
        cell.font = white
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for ridx, row in enumerate(df.itertuples(index=False), 2):
        for cidx, value in enumerate(row, 1):
            col = df.columns[cidx-1]
            if pd.isna(value):
                value = None
            cell = ws.cell(ridx, cidx, value)
            cell.border = border
            if col in fills:
                light_color = "EAF3FF" if fills[col].fgColor.rgb in ["00C5DCF5","C5DCF5"] else "EAFBEF"
                cell.fill = PatternFill("solid", fgColor=light_color)
            if any(k in norm_key(col) for k in ["sales", "suma", "total", "koszt", "cost", "materials", "tpm", "cm", "wartosc", "papier", "klej", "lakiery", "offset", "plyta", "kliki"]):
                cell.number_format = '#,##0.00'
            if "%" in col:
                cell.number_format = '0.0%'
            if "data faktury" == norm_key(col):
                cell.number_format = 'yyyy-mm-dd'
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cidx, col in enumerate(df.columns, 1):
        letter = get_column_letter(cidx)
        width = max(10, min(42, max(len(str(col)) + 2, 12)))
        sample = df[col].head(100).astype(str).map(len).max() if len(df) else 0
        ws.column_dimensions[letter].width = max(width, min(42, sample + 2))
        if norm_key(col) in hidden_norm:
            ws.column_dimensions[letter].hidden = True

def build_xlsx(df: pd.DataFrame, aux: Dict[str,pd.DataFrame], machine_cols: List[str], rates_df, click_df, prepress_df, summary, batch, params) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Profitability"

    fill_dl = PatternFill("solid", fgColor="C5DCF5")
    fill_mat = PatternFill("solid", fgColor="C5F5D9")
    dl_cols = set(machine_cols + ["Prepress costs", "Total DL"])
    mat_cols = set(["Papier [16]", "Klej [17]", "Lakiery [20]", "Opakowania zbiorcze [24]", "Other Materials", "Offset inks", "Płyta offsetowa", "Kliki final", "Total Materials"])
    fills = {c: fill_dl for c in df.columns if c in dl_cols}
    fills.update({c: fill_mat for c in df.columns if c in mat_cols or norm_key(c) in {norm_key(x) for x in mat_cols}})
    write_df(ws, df, fills=fills, hidden=HIDDEN_COLS)

    for name, data in aux.items():
        safe_name = name[:31]
        wsx = wb.create_sheet(safe_name)
        write_df(wsx, data.drop(columns=[c for c in data.columns if str(c).startswith("_")], errors="ignore"))

    ws_rates = wb.create_sheet("Stawki")
    write_df(ws_rates, rates_df if rates_df is not None else pd.DataFrame(columns=["Nazwa maszyny","Stawka rbg"]))

    ws_click = wb.create_sheet("Koszty klików")
    write_df(ws_click, click_df if click_df is not None else pd.DataFrame(columns=["Press Name","Color","Koszt jednostkowy"]))

    ws_pp = wb.create_sheet("Prepress")
    write_df(ws_pp, prepress_df if prepress_df is not None else pd.DataFrame(columns=["Klient","Stawka Prepress Digital","Stawka Prepress Offset"]))

    ws_par = wb.create_sheet("Parametry")
    write_df(ws_par, pd.DataFrame([params]))

    ws_sum = wb.create_sheet("Podsumowanie")
    write_df(ws_sum, summary)

    ws_batch = wb.create_sheet("Batch summary")
    write_df(ws_batch, batch)

    ws_k = wb.create_sheet("Kokpit")
    kpis = pd.DataFrame({
        "KPI": ["Suma sprzedaży","Suma TPM","Średni TPM %","Suma CM","Średni CM %","Liczba klientów","Liczba zamówień","Liczba Digital","Liczba Offset","Liczba no printing"],
        "Wartość": [
            df["Sales Value"].sum(), df["TPM"].sum(), df["TPM"].sum()/df["Sales Value"].sum() if df["Sales Value"].sum() else 0,
            df["CM"].sum(), df["CM"].sum()/df["Sales Value"].sum() if df["Sales Value"].sum() else 0,
            summary["Klient"].nunique() if not summary.empty else 0, len(df),
            (df["Digital/Offset"]=="Digital").sum(), (df["Digital/Offset"]=="Offset").sum(), (df["Digital/Offset"]=="no printing").sum()
        ]
    })
    write_df(ws_k, kpis)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

# =========================
# App state/defaults
# =========================
if "rates_df" not in st.session_state:
    st.session_state.rates_df = pd.DataFrame(columns=["Nazwa maszyny", "Stawka rbg"])
if "click_df" not in st.session_state:
    st.session_state.click_df = load_click_costs(None)
if "prepress_df" not in st.session_state:
    st.session_state.prepress_df = pd.DataFrame(columns=["Klient", "Stawka Prepress Digital", "Stawka Prepress Offset"])

st.sidebar.markdown("## 📊 Postkalkulacja")
st.sidebar.markdown("**Profitability App**")
st.sidebar.markdown("---")
tabs = [
    "📂 Upload plików", "⚙️ Stawki maszyn", "🖨️ Koszty klików", "🎨 Prepress",
    "🔧 Parametry", "📋 Podgląd Profitability", "📈 Podsumowanie", "🎯 Kokpit", "⬇️ Pobierz XLSX"
]
selected = st.sidebar.radio("Nawigacja", tabs, label_visibility="collapsed")

st.markdown('<div class="hero"><h1>📊 Postkalkulacja Profitability</h1><p>Interim CFO — kalkulacja rentowności zleceń</p></div>', unsafe_allow_html=True)

# Uploads in session
if "files" not in st.session_state:
    st.session_state.files = {}

if selected == tabs[0]:
    st.markdown('<div class="card">Uploaduj pliki źródłowe. Wymagana jest tylko Baza / post_list.</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.files["base"] = st.file_uploader("Baza / post_list — wymagany", type=["xlsx","xls"], key="base")
        st.session_state.files["czasy"] = st.file_uploader("Czasy dla aplikacji", type=["xlsx","xls"], key="czasy")
        st.session_state.files["zlec"] = st.file_uploader("Zlecenia + faktury", type=["xlsx","xls"], key="zlec")
        st.session_state.files["fry"] = st.file_uploader("Faktury linie faktury", type=["xlsx","xls"], key="fry")
    with c2:
        st.session_state.files["inks"] = st.file_uploader("Kliki / inks", type=["xlsx","xls"], key="inks")
        st.session_state.files["click_costs"] = st.file_uploader("Koszty klików", type=["xlsx","xls"], key="click_costs")
        st.session_state.files["stawki"] = st.file_uploader("Stawki dla aplikacji", type=["xlsx","xls"], key="stawki")
        st.session_state.files["farby"] = st.file_uploader("Farby podsumowanie", type=["xlsx","xls"], key="farby")
    if st.session_state.files.get("stawki"):
        st.session_state.rates_df = load_rates(st.session_state.files["stawki"])
    if st.session_state.files.get("click_costs"):
        st.session_state.click_df = load_click_costs(st.session_state.files["click_costs"])
    st.success("Pliki wczytane do sesji. Przejdź do kolejnych zakładek.")

elif selected == tabs[1]:
    st.subheader("Stawki maszyn")
    st.session_state.rates_df = st.data_editor(st.session_state.rates_df, num_rows="dynamic", use_container_width=True)
elif selected == tabs[2]:
    st.subheader("Koszty klików")
    st.session_state.click_df = st.data_editor(st.session_state.click_df, num_rows="dynamic", use_container_width=True)
elif selected == tabs[3]:
    st.subheader("Prepress")
    st.caption("Domyślnie Digital = 10, Offset = 40. Możesz dodać stawki per klient.")
    st.session_state.prepress_df = st.data_editor(st.session_state.prepress_df, num_rows="dynamic", use_container_width=True)
elif selected == tabs[4]:
    st.subheader("Parametry")
    st.session_state.other_pct = st.number_input("Other costs %", min_value=0.0, max_value=100.0, value=float(st.session_state.get("other_pct",2.0)), step=0.1)
    st.session_state.tpm_thr = st.number_input("Próg alertu TPM %", min_value=0.0, max_value=100.0, value=float(st.session_state.get("tpm_thr",60.0)), step=1.0)
    st.session_state.cm_thr = st.number_input("Próg alertu CM %", min_value=0.0, max_value=100.0, value=float(st.session_state.get("cm_thr",40.0)), step=1.0)

def calculate_or_warn():
    if not st.session_state.files.get("base"):
        st.warning("Najpierw uploaduj plik Baza / post_list.")
        return None
    try:
        df, aux, warns, machines = build_profitability(
            st.session_state.files,
            st.session_state.rates_df,
            st.session_state.click_df,
            st.session_state.prepress_df,
            float(st.session_state.get("other_pct", 2.0)),
        )
        for w in warns:
            st.warning(w)
        return df, aux, warns, machines
    except Exception as e:
        st.error(str(e))
        return None

if selected == tabs[5]:
    res = calculate_or_warn()
    if res:
        df, aux, warns, machines = res
        months = ["Wszystkie"] + sorted([m for m in df["Miesiąc faktury"].dropna().unique() if m])
        chosen = st.multiselect("Filtr miesiąca", months, default=["Wszystkie"])
        view = df if "Wszystkie" in chosen or not chosen else df[df["Miesiąc faktury"].isin(chosen)]
        st.dataframe(view, use_container_width=True, height=560)

elif selected == tabs[6]:
    res = calculate_or_warn()
    if res:
        df, aux, warns, machines = res
        months = ["Wszystkie"] + sorted([m for m in df["Miesiąc faktury"].dropna().unique() if m])
        month = st.selectbox("Miesiąc raportu", months)
        summary, batch = make_summary(df, month)
        st.markdown("### Podsumowanie klientów")
        st.dataframe(summary, use_container_width=True)
        st.markdown("### Liczba zamówień: klient × batch")
        st.dataframe(batch, use_container_width=True)

elif selected == tabs[7]:
    res = calculate_or_warn()
    if res:
        df, aux, warns, machines = res
        months = ["Wszystkie"] + sorted([m for m in df["Miesiąc faktury"].dropna().unique() if m])
        month = st.selectbox("Miesiąc kokpitu", months)
        data = df if month == "Wszystkie" else df[df["Miesiąc faktury"] == month]
        summary, batch = make_summary(df, month)
        sv, tpm, cm = data["Sales Value"].sum(), data["TPM"].sum(), data["CM"].sum()
        kpis = [
            ("Sprzedaż", f"{sv:,.0f}"), ("TPM", f"{tpm:,.0f}"), ("TPM %", f"{(tpm/sv*100 if sv else 0):.1f}%"),
            ("CM", f"{cm:,.0f}"), ("CM %", f"{(cm/sv*100 if sv else 0):.1f}%"),
            ("Zamówienia", f"{len(data):,}"), ("Digital", str((data["Digital/Offset"]=="Digital").sum())), ("Offset", str((data["Digital/Offset"]=="Offset").sum()))
        ]
        cols = st.columns(4)
        for i, (l, v) in enumerate(kpis):
            cols[i % 4].markdown(f'<div class="kpi"><div class="v">{v}</div><div class="l">{l}</div></div>', unsafe_allow_html=True)
        klient = find_col(data, ["Klient","Klient ID"]) or "Zlecenie produkcyjne"
        grp = data.groupby(klient).agg({"Sales Value":"sum","TPM":"sum","CM":"sum"}).reset_index()
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.bar(grp.nlargest(5,"TPM"), x=klient, y="TPM", title="Top 5 klientów wg TPM"), use_container_width=True)
            st.plotly_chart(px.bar(data["Digital/Offset"].value_counts().reset_index(), x="Digital/Offset", y="count", title="Digital/Offset/no printing"), use_container_width=True)
        with c2:
            st.plotly_chart(px.bar(grp.nlargest(5,"CM"), x=klient, y="CM", title="Top 5 klientów wg CM"), use_container_width=True)
            if not batch.empty:
                st.plotly_chart(px.bar(batch, x="Klient", y="Liczba zamówień", color="Batch", title="Zamówienia wg klienta i batch"), use_container_width=True)
        st.markdown("### Alerty")
        if not summary.empty:
            alerts = summary[(summary["TPM %"] < st.session_state.get("tpm_thr",60)/100) | (summary["CM %"] < st.session_state.get("cm_thr",40)/100)]
            st.dataframe(alerts, use_container_width=True)

elif selected == tabs[8]:
    res = calculate_or_warn()
    if res:
        df, aux, warns, machines = res
        months = ["Wszystkie"] + sorted([m for m in df["Miesiąc faktury"].dropna().unique() if m])
        month = st.selectbox("Miesiąc do arkusza Podsumowanie/Kokpit", months)
        summary, batch = make_summary(df, month)
        params = {
            "Other costs %": st.session_state.get("other_pct",2.0)/100,
            "Próg TPM %": st.session_state.get("tpm_thr",60.0)/100,
            "Próg CM %": st.session_state.get("cm_thr",40.0)/100,
            "Miesiąc raportu": month
        }
        xlsx = build_xlsx(df, aux, machines, st.session_state.rates_df, st.session_state.click_df, st.session_state.prepress_df, summary, batch, params)
        st.download_button(
            "⬇️ Pobierz plik XLSX",
            data=xlsx,
            file_name="postkalkulacja_profitability.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Plik gotowy do pobrania.")
