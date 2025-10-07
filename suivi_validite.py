import pandas as pd
import streamlit as st
from datetime import datetime

# === CONFIGURATION ===
EXCEL_PATH = r"C:\Users\Administrateur.PC-WOODCITY2\Documents\Suivi_Validité.xlsm"

# === FONCTIONS ===
def charger_donnees():
    try:
        df = pd.read_excel(EXCEL_PATH)
        # Nettoyage des colonnes attendues
        df.columns = [col.strip() for col in df.columns]
        # Vérifie si les colonnes existent
        required_cols = ["Nom du fichier", "Date de validité", "Alerte avant (jours)"]
        for col in required_cols:
            if col not in df.columns:
                st.error(f"Colonne manquante : {col}")
                return None

        # Conversion des dates
        df["Date de validité"] = pd.to_datetime(df["Date de validité"], errors="coerce")
        today = datetime.now().date()
        df["Jours restants"] = (df["Date de validité"].dt.date - today).dt.days

        # Détermination du statut
        def statut(row):
            if pd.isnull(row["Date de validité"]):
                return "❌ Date invalide"
            if row["Jours restants"] < 0:
                return "🔴 Expiré"
            elif row["Jours restants"] <= row["Alerte avant (jours)"]:
                return "🟠 À surveiller"
            else:
                return "🟢 OK"

        df["Statut"] = df.apply(statut, axis=1)
        return df

    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier : {e}")
        return None

# === INTERFACE STREAMLIT ===
st.set_page_config(page_title="Suivi Validité Fichiers", layout="wide")

st.title("📁 Suivi de validité des fichiers")
st.caption("Système d’alerte automatique basé sur ton fichier Excel")

if st.button("🔄 Recharger les données"):
    st.experimental_rerun()

df = charger_donnees()

if df is not None:
    # Filtres
    filtre = st.selectbox("Afficher :", ["Tous", "OK", "À surveiller", "Expiré"])
    if filtre == "OK":
        df = df[df["Statut"].str.contains("🟢")]
    elif filtre == "À surveiller":
        df = df[df["Statut"].str.contains("🟠")]
    elif filtre == "Expiré":
        df = df[df["Statut"].str.contains("🔴")]

    # Mise en forme
    def couleur_ligne(val):
        if "🔴" in val:
            return "background-color: #ffcccc"
        elif "🟠" in val:
            return "background-color: #fff3cd"
        elif "🟢" in val:
            return "background-color: #d4edda"
        return ""

    st.dataframe(
        df.style.applymap(couleur_ligne, subset=["Statut"]),
        use_container_width=True
    )
else:
    st.warning("⚠️ Impossible de charger les données. Vérifie le chemin ou le format du fichier.")
