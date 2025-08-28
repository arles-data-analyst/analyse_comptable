import os
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# --- Réglages page ---
st.set_page_config(page_title="Analyse comptable", page_icon="📊", layout="wide")
st.title("📊 Analyse comptable – Démo interactive")

# ---------- Utils ----------
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
            .str.replace(r"\.0$", "", regex=True)   # supprime le .0 terminal
        )
    return df

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip().lower() for c in df.columns]
    return df

# ---------- Chargement (avec bust de cache si Excel change) ----------
brut_path = "factures_comptables_brutes.xlsx"
nettoye_path = "factures_comptables_nettoyees.xlsx"
brut_mtime = file_mtime(brut_path)
nettoye_mtime = file_mtime(nettoye_path)

@st.cache_data
def load_data(_brut_mtime: float, _nettoye_mtime: float) -> pd.DataFrame:
    # Priorité au fichier nettoyé s'il existe
    if os.path.exists(nettoye_path):
        df = pd.read_excel(nettoye_path)
    elif os.path.exists(brut_path):
        df = pd.read_excel(brut_path)
    else:
        return pd.DataFrame()

    df = standardize_columns(df)

    # Types de base
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    if "montant" in df.columns:
        df["montant"] = pd.to_numeric(df["montant"], errors="coerce")

    # Si le fichier est brut, recrée vite-fait 2/3 colonnes utiles
    if "type" not in df.columns and "montant" in df.columns:
        df["type"] = df["montant"].apply(lambda x: "Entrée" if pd.notna(x) and x >= 0 else "Sortie")
    if "mois" not in df.columns and "date" in df.columns:
        df["mois"] = df["date"].dt.to_period("M").dt.to_timestamp()
    if "solde_cumulé" not in df.columns and {"date", "montant"}.issubset(df.columns):
        df = df.sort_values("date")
        df["solde_cumulé"] = df["montant"].cumsum()

    # Comptes sans .0
    df = normalize_compte_col(df)
    return df

df = load_data(brut_mtime, nettoye_mtime)

if df.empty:
    st.warning("Aucun fichier Excel trouvé à la racine du repo. Ajoute : "
               "`factures_comptables_nettoyees.xlsx` ou `factures_comptables_brutes.xlsx`.")
    st.stop()

# ---------- Filtres ----------
col1, col2, col3 = st.columns(3)
with col1:
    comptes = ["(Tous)"] + sorted([str(x) for x in df.get("compte", pd.Series(dtype=str)).dropna().unique()])
    compte_sel = st.selectbox("Compte", comptes)
with col2:
    types = ["(Tous)"] + sorted(df.get("type", pd.Series(["Entrée", "Sortie"])).dropna().unique().tolist())
    type_sel = st.selectbox("Type d'opération", types)
with col3:
    if "date" in df.columns and df["date"].notna().any():
        dmin = pd.to_datetime(df["date"].min())
        dmax = pd.to_datetime(df["date"].max())
        start, end = st.slider("Période",
                               min_value=dmin.to_pydatetime(),
                               max_value=dmax.to_pydatetime(),
                               value=(dmin.to_pydatetime(), dmax.to_pydatetime()))
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

# ---------- KPIs ----------
colA, colB, colC = st.columns(3)
total = dff.get("montant", pd.Series([0.0])).sum()
entrees = dff.loc[dff.get("type","")=="Entrée","montant"].sum() if "type" in dff.columns else float("nan")
sorties = dff.loc[dff.get("type","")=="Sortie","montant"].sum() if "type" in dff.columns else float("nan")
colA.metric("Résultat filtré", f"{total:,.2f} €".replace(",", " ").replace(".", ","))
colB.metric("Total entrées", f"{entrees:,.2f} €".replace(",", " ").replace(".", ","))
colC.metric("Total sorties", f"{sorties:,.2f} €".replace(",", " ").replace(".", ","))

st.divider()

# ---------- Graphiques (format compacts) ----------
# 1) Évolution du solde cumulé
if {"date", "montant"}.issubset(dff.columns) and dff["date"].notna().any():
    dff = dff.sort_values("date")
    dff["solde_cumulé_view"] = dff["montant"].cumsum()
    st.subheader("Évolution du solde cumulé (filtré)")
    fig1, ax1 = plt.subplots(figsize=(9, 3))
    ax1.plot(dff["date"], dff["solde_cumulé_view"])
    ax1.set_xlabel("Date")
    ax1.set_ylabel("Solde cumulé (€)")
    fig1.tight_layout()
    st.pyplot(fig1, use_container_width=True)

# 2) Top comptes
if {"compte", "montant"}.issubset(dff.columns):
    st.subheader("Top comptes (par somme des montants)")
    dff["compte"] = dff["compte"].astype("string")
    top = (dff.groupby("compte", dropna=True)["montant"]
             .sum()
             .sort_values(ascending=False)
             .head(8))
    if not top.empty:
        fig2, ax2 = plt.subplots(figsize=(8, 3))
        top.plot(kind="bar", ax=ax2)
        ax2.set_xlabel("Compte")
        ax2.set_ylabel("Montant total (€)")
        fig2.tight_layout()
        st.pyplot(fig2, use_container_width=True)

# ---------- Debug (optionnel) ----------
with st.expander("🔎 Debug (optionnel)"):
    st.write("Colonnes:", list(df.columns))
    st.write("Types:", df.dtypes)
    st.write(dff.head(10))
