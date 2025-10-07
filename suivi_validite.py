import os
import pandas as pd
from datetime import datetime

# === CONFIGURATION ===
EXCEL_PATH = r"C:\Users\Administrateur.PC-WOODCITY2\Documents\Suivi_Validité.xlsm"

def charger_donnees(path):
    """Charge les données du fichier Excel et calcule les jours restants"""
    print(f"📂 Lecture du fichier : {path}")
    if not os.path.exists(path):
        print("❌ Fichier introuvable à ce chemin (probablement sur GitHub Actions).")
        return None

    try:
        df = pd.read_excel(path)
        df.columns = [col.strip() for col in df.columns]

        required_cols = ["Nom du fichier", "Date de validité", "Alerte avant (jours)"]
        for col in required_cols:
            if col not in df.columns:
                print(f"⚠️ Colonne manquante : {col}")
                return None

        df["Date de validité"] = pd.to_datetime(df["Date de validité"], errors="coerce")
        today = datetime.now().date()
        df["Jours restants"] = (df["Date de validité"].dt.date - today).dt.days

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
        print(f"❌ Erreur lors du chargement du fichier : {e}")
        return None


def afficher_resultats(df):
    """Affiche un résumé des statuts dans la console"""
    print("\n===== RAPPORT DE VALIDITÉ =====")
    print(df[["Nom du fichier", "Date de validité", "Jours restants", "Statut"]])
    print("===============================\n")

    alertes = df[df["Statut"].str.contains("🔴|🟠")]
    if alertes.empty:
        print("✅ Aucun fichier en alerte.")
    else:
        print("⚠️ Fichiers en alerte :")
        for _, row in alertes.iterrows():
            print(f"- {row['Nom du fichier']} → {row['Statut']} ({row['Jours restants']} jours restants)")


if __name__ == "__main__":
    print("🚀 Script démarré")
    print("Répertoire de travail :", os.getcwd())

    df = charger_donnees(EXCEL_PATH)
    if df is not None:
        afficher_resultats(df)
    else:
        print("⚠️ Aucune donnée analysée.")

    print("✅ Vérification terminée avec succès.")
