import os
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Analyse comptable", page_icon="📊", layout="wide")
st.title("📊 Analyse comptable – Démo interactive")

@st.cache_data
def load_data():
    if os.path.exists("factures_comptables_nettoyees.xlsx"):
        df = pd.read_excel("factures_comptables_nettoyees.xlsx")
        # normalise colonnes possibles
        df.columns = [c.strip().lower() for c in df.columns]
        if "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"])
    else:
        df = pd.read_excel("factures_comptables_brutes.xlsx")
        df.columns = [c.strip().lower() for c in df.columns]
        # conversions minimales
        if "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
        if "montant" in df.columns:
            df["montant"] = pd.to_numeric(df["montant"], errors="coerce")
        # type simple (entrées/sorties)
        if "montant" in df.columns:
            df["type"] = df["montant"].apply(lambda x: "Entrée" if pd.notna(x) and x >= 0 else "Sortie")
        # mois pour agrégations
        if "date" in df.columns:
            df["mois"] = df["date"].dt.to_period("M").dt.to_timestamp()
        # solde cumulé (ordre chronologique)
        if "date" in df.columns and "montant" in df.columns:
            df = df.sort_values("date")
            df["solde_cumulé"] = df["montant"].cumsum()
    return df

df = load_data()
if df.empty:
    st.warning("Aucune donnée chargée. Vérifie le fichier Excel à la racine.")
    st.stop()

# Filtres
col1, col2, col3 = st.columns(3)
with col1:
    comptes = ["(Tous)"] + sorted([str(x) for x in df.get("compte", pd.Series(dtype=str)).dropna().unique()])
    compte_sel = st.selectbox("Compte", comptes)
with col2:
    types = ["(Tous)"] + sorted(df.get("type", pd.Series(["Entrée","Sortie"])).dropna().unique().tolist())
    type_sel = st.selectbox("Type d'opération", types)
with col3:
    if "date" in df.columns:
        dmin, dmax = pd.to_datetime(df["date"].min()), pd.to_datetime(df["date"].max())
        start, end = st.slider("Période", min_value=dmin.to_pydatetime(), max_value=dmax.to_pydatetime(),
                               value=(dmin.to_pydatetime(), dmax.to_pydatetime()))
    else:
        start = end = None

# Appliquer filtres
mask = pd.Series([True]*len(df))
if compte_sel != "(Tous)" and "compte" in df.columns:
    mask &= df["compte"].astype(str) == compte_sel
if type_sel != "(Tous)" and "type" in df.columns:
    mask &= df["type"] == type_sel
if start and end and "date" in df.columns:
    mask &= (df["date"] >= pd.to_datetime(start)) & (df["date"] <= pd.to_datetime(end))

dff = df.loc[mask].copy()

# KPIs
colA, colB, colC = st.columns(3)
total = dff.get("montant", pd.Series([0])).sum()
entrees = dff.loc[dff.get("type","")=="Entrée","montant"].sum() if "type" in dff.columns else float("nan")
sorties = dff.loc[dff.get("type","")=="Sortie","montant"].sum() if "type" in dff.columns else float("nan")
colA.metric("Résultat filtré", f"{total:,.2f} €")
colB.metric("Total entrées", f"{entrees:,.2f} €")
colC.metric("Total sorties", f"{sorties:,.2f} €")

st.divider()

# Graph 1 : Évolution du solde cumulé
if "date" in dff.columns and "montant" in dff.columns:
    dff = dff.sort_values("date")
    dff["solde_cumulé_view"] = dff["montant"].cumsum()
    st.subheader("Évolution du solde cumulé (filtré)")
    fig1, ax1 = plt.subplots()
    ax1.plot(dff["date"], dff["solde_cumulé_view"])
    ax1.set_xlabel("Date"); ax1.set_ylabel("Solde cumulé (€)")
    st.pyplot(fig1)

# Graph 2 : Top comptes (agrégat)
if "compte" in dff.columns and "montant" in dff.columns:
    st.subheader("Top comptes (par somme des montants)")
    top = (dff.groupby("compte", dropna=True)["montant"]
             .sum()
             .sort_values(ascending=False)
             .head(10))
    if not top.empty:
        fig2, ax2 = plt.subplots()
        top.plot(kind="bar", ax=ax2)
        ax2.set_xlabel("Compte"); ax2.set_ylabel("Montant total (€)")
        st.pyplot(fig2)

st.caption("Demo Streamlit – basée sur le fichier Excel du projet.")
