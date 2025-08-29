import os
import io
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import streamlit as st
from pathlib import Path

# --- config logo ---
LOGO = "assets/logo.png"            # si ton fichier est .png
if not Path(LOGO).exists():
    LOGO = "assets/logo.jpg"        # fallback si tu as mis un .jpg

st.set_page_config(page_title="Analyse comptable", page_icon=LOGO, layout="wide")

# en-tête avec logo + titre
col_logo, col_title = st.columns([1, 20])
with col_logo:
    st.image(LOGO, width=36)
with col_title:
    st.markdown("# Tableau de bord d’analyse comptable")

# ---------------- Utils ----------------
def file_mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0

def normalize_compte_col(df: pd.DataFrame) -> pd.DataFrame:
    if "compte" in df.columns:
        df["compte"] = (
            df["compte"]
            .astype("string")
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)   # supprime 601.0 -> 601
        )
    return df

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip().lower() for c in df.columns]
    return df

def eur(x, decimals: int = 2, symbol: bool = True) -> str:
    """Format fr : 12 345,67 € ; gère NaN/None."""
    try:
        if pd.isna(x):
            return "—"
    except Exception:
        pass
    s = f"{float(x):,.{decimals}f}".replace(",", " ").replace(".", ",")
    return f"{s} €" if symbol else s

def add_watermark(fig, text="© 2025 Arlès Fanampindrainy — analyse-comptable.streamlit.app"):
    """Ajoute un filigrane discret en bas à droite de la figure."""
    fig.text(0.995, 0.005, text, ha="right", va="bottom", fontsize=7, color="0.5", alpha=0.6)

# ---------------- Chargement (cache bust sur mtime) ----------------
brut_path = "factures_comptables_brutes.xlsx"
nettoye_path = "factures_comptables_nettoyees.xlsx"
brut_mtime = file_mtime(brut_path)
nettoye_mtime = file_mtime(nettoye_path)

@st.cache_data
def load_data(_brut_mtime: float, _nettoye_mtime: float) -> pd.DataFrame:
    if os.path.exists(nettoye_path):
        df = pd.read_excel(nettoye_path)
    elif os.path.exists(brut_path):
        df = pd.read_excel(brut_path)
    else:
        return pd.DataFrame()

    df = standardize_columns(df)

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    if "montant" in df.columns:
        df["montant"] = pd.to_numeric(df["montant"], errors="coerce")

    # Si brut, recrée quelques colonnes utiles
    if "type" not in df.columns and "montant" in df.columns:
        df["type"] = df["montant"].apply(lambda x: "Entrée" if pd.notna(x) and x >= 0 else "Sortie")
    if "mois" not in df.columns and "date" in df.columns:
        df["mois"] = df["date"].dt.to_period("M").dt.to_timestamp()
    if "solde_cumulé" not in df.columns and {"date", "montant"}.issubset(df.columns):
        df = df.sort_values("date")
        df["solde_cumulé"] = df["montant"].cumsum()

    df = normalize_compte_col(df)
    return df

df = load_data(brut_mtime, nettoye_mtime)

if df.empty:
    st.warning(
        "Aucun fichier Excel trouvé à la racine du repo. "
        "Ajoute `factures_comptables_nettoyees.xlsx` ou `factures_comptables_brutes.xlsx`."
    )
    st.stop()

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Réglages")
    comptes = ["(Tous)"] + sorted([str(x) for x in df.get("compte", pd.Series(dtype=str)).dropna().unique()])
    type_options = ["(Tous)"] + sorted(df.get("type", pd.Series(["Entrée","Sortie"])).dropna().unique().tolist())
    top_n = st.slider("Top N comptes", min_value=3, max_value=15, value=8, step=1)
    debug_mode = st.checkbox("Afficher debug", value=False)

# ---------------- Filtres (en-tête) ----------------
col1, col2, col3 = st.columns(3)
with col1:
    compte_sel = st.selectbox("Compte", comptes)
with col2:
    type_sel = st.selectbox("Type d'opération", type_options)
with col3:
    if "date" in df.columns and df["date"].notna().any():
        dmin = pd.to_datetime(df["date"].min())
        dmax = pd.to_datetime(df["date"].max())
        start, end = st.slider(
            "Période", min_value=dmin.to_pydatetime(), max_value=dmax.to_pydatetime(),
            value=(dmin.to_pydatetime(), dmax.to_pydatetime())
        )
    else:
        start = end = None

mask = pd.Series(True, index=df.index)
if compte_sel != "(Tous)" and "compte" in df.columns:
    mask &= df["compte"].astype(str) == compte_sel
if type_sel != "(Tous)" and "type" in df.columns:
    mask &= df["type"] == type_sel
if start and end and "date" in df.columns:
    mask &= (df["date"] >= pd.to_datetime(start)) & (df["date"] <= pd.to_datetime(end))

dff = df.loc[mask].copy()

# ---------------- KPIs (format € FR) ----------------
colA, colB, colC = st.columns(3)
total = dff.get("montant", pd.Series([0.0])).sum()
entrees = dff.loc[dff.get("type","")=="Entrée","montant"].sum() if "type" in dff.columns else float("nan")
sorties = dff.loc[dff.get("type","")=="Sortie","montant"].sum() if "type" in dff.columns else float("nan")
colA.metric("Résultat filtré", eur(total))
colB.metric("Total entrées", eur(entrees))
colC.metric("Total sorties", eur(sorties))

# ---------------- Exports ----------------
exp_col1, exp_col2 = st.columns(2)
with exp_col1:
    csv_bytes = dff.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Télécharger les données filtrées (CSV)",
        data=csv_bytes,
        file_name=f"donnees_filtrees_{datetime.now():%Y%m%d}.csv",
        mime="text/csv",
    )
with exp_col2:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        dff.to_excel(writer, index=False, sheet_name="Filtre")
    st.download_button(
        "⬇️ Télécharger les données filtrées (Excel)",
        data=out.getvalue(),
        file_name=f"donnees_filtrees_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# ---------- Graphiques (compacts, 2 colonnes) ----------
gcol1, gcol2 = st.columns(2)

# 1) Solde cumulé (gauche)
with gcol1:
    if {"date", "montant"}.issubset(dff.columns) and dff["date"].notna().any():
        dff = dff.sort_values("date")
        dff["solde_cumulé_view"] = dff["montant"].cumsum()
        st.markdown("#### Solde cumulé (filtré)")
        fig1, ax1 = plt.subplots(figsize=(6, 2.6))
        ax1.plot(dff["date"], dff["solde_cumulé_view"])
        ax1.set_xlabel("Date", fontsize=10)
        ax1.set_ylabel("Solde cumulé (€)", fontsize=10)
        ax1.tick_params(axis="both", labelsize=9)
        ax1.yaxis.set_major_formatter(FuncFormatter(lambda y, _: eur(y, decimals=0, symbol=False)))
        fig1.tight_layout()
        add_watermark(fig1)
        st.pyplot(fig1, use_container_width=True)

# 2) Top comptes (droite)
with gcol2:
    if {"compte", "montant"}.issubset(dff.columns):
        st.markdown("#### Top comptes")
        dff["compte"] = dff["compte"].astype("string")
        top = (dff.groupby("compte", dropna=True)["montant"]
                 .sum().sort_values(ascending=False).head(top_n))
        if not top.empty:
            fig2, ax2 = plt.subplots(figsize=(6, 2.6))
            top.plot(kind="bar", ax=ax2)
            ax2.set_xlabel("Compte", fontsize=10)
            ax2.set_ylabel("Montant total (€)", fontsize=10)
            ax2.tick_params(axis="x", labelrotation=0, labelsize=9)
            ax2.tick_params(axis="y", labelsize=9)
            ax2.yaxis.set_major_formatter(FuncFormatter(lambda y, _: eur(y, decimals=0, symbol=False)))
            ax2.margins(y=0.15)  # marge en haut pour éviter que l'étiquette touche la bordure
            for p in ax2.patches:
                v = p.get_height()
                ax2.annotate(
                    eur(v, decimals=0),
                    (p.get_x() + p.get_width()/2, v),
                    ha="center",
                    va="bottom" if v >= 0 else "top",
                    fontsize=9,
                    xytext=(0, 2), textcoords="offset points",
                    clip_on=False,
                )
            fig2.tight_layout()
            add_watermark(fig2)
            st.pyplot(fig2, use_container_width=True)

# ---------------- Debug ----------------
if debug_mode:
    with st.expander("🔎 Debug"):
        st.write("Colonnes:", list(df.columns))
        st.write("Types:", df.dtypes)
        st.write(dff.head(10))

# --- Copyright / footer visible ---
st.markdown(
    """
    <style>
    .footer-copyright {
        position: fixed; left: 0; bottom: 0; width: 100%;
        text-align: center; font-size: 12px; color: #666;
        background: rgba(255,255,255,0.75); padding: 6px 0; z-index: 9999;
        backdrop-filter: blur(2px);
    }
    </style>
    <div class="footer-copyright">
      © 2025 Arlès Fanampindrainy — analyse-comptable.streamlit.app — Reproduction interdite sans attribution.
    </div>
    """,
    unsafe_allow_html=True,
)
