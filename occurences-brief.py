import streamlit as st
import pandas as pd
from io import BytesIO
import re

# 📌 Titre de l'application
st.title("📊 Consolidation Mot‑Clé Principal + Occurrences (keywords only)")

st.markdown("""
**Mode d'emploi** :
1️⃣ Téléversez un classeur **.xlsx** comportant **une feuille par cluster**.  
   • Les cellules **A5** et **B5** forment le *mot‑clé principal* (ex. « eti » et « définition »).  
   • Les mots‑clés secondaires peuvent être dans n'importe quelle cellule, séparés par `|`.  
2️⃣ Cliquez sur **« Fusionner & Télécharger »**.  
3️⃣ Le fichier exporté contient **deux colonnes** :  
   • **Mot Clé Principal**  
   • **Occurrences** : uniquement des mots‑clés (texte), dédupliqués, séparés par ` | `.
""")

# ➡️ Import du fichier
xlsx_file = st.file_uploader("📂 Charger votre fichier Excel", type=["xlsx"])

# ➡️ Fonction utilitaire pour savoir si un token est purement numérique/pourcentage, etc.
def is_keyword(token: str) -> bool:
    # rejet si uniquement chiffres, espace ou ponctuation
    token_stripped = token.strip()
    if not token_stripped:
        return False
    # pattern nombre (entier ou décimal) éventuellement suivi d'un %
    if re.fullmatch(r"[0-9]+([.,][0-9]+)?%?", token_stripped):
        return False
    return True

if xlsx_file and st.button("🚀 Fusionner & Télécharger"):
    try:
        sheets_dict = pd.read_excel(xlsx_file, sheet_name=None, header=None, engine="openpyxl")

        lignes = []
        for _, df in sheets_dict.items():
            # ---- Mot‑clé principal ----
            try:
                mot_cle_principal = f"{df.iloc[4, 0]} {df.iloc[4, 1]}".strip()
            except Exception:
                mot_cle_principal = ""

            # ---- Récupération & nettoyage des tokens ----
            tokens = []
            for cell in df.values.flatten():
                if pd.isna(cell):
                    continue
                for part in str(cell).split("|"):
                    part = part.strip()
                    if part and is_keyword(part):
                        tokens.append(part)

            if tokens:
                unique_tokens = list(dict.fromkeys(tokens))
                lignes.append({
                    "Mot Clé Principal": mot_cle_principal,
                    "Occurrences": " | ".join(unique_tokens)
                })

        if not lignes:
            st.warning("⚠️ Aucun mot‑clé valide trouvé dans le classeur.")
        else:
            result_df = pd.DataFrame(lignes, columns=["Mot Clé Principal", "Occurrences"])
            st.subheader("🔎 Aperçu consolidation (keywords only)")
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            # ---- Export en mémoire ----
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, sheet_name="Consolidé", index=False, header=True)
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le XLSX consolidé",
                data=buffer,
                file_name="motcles_consolides.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
