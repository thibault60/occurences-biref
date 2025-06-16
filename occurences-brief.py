import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📊 Consolidateur d'occurrences de mots‑clés ➡️ Export XLSX")

# 📝 Mode d'emploi
st.markdown("""
Collez ou téléversez vos lignes d'occurrences : chaque ligne contient plusieurs mots‑clés séparés par `|`.
Le script comptabilise la fréquence de chaque mot‑clé sur l'ensemble des lignes, puis génère un fichier **XLSX** prêt à télécharger.
""")

# 🔹 Widgets d'entrée
occ_text = st.text_area("✂️ Collez vos occurrences ici", height=250)
uploaded_file = st.file_uploader("📂 …ou téléchargez un fichier .txt contenant vos occurrences", type=["txt"])

# 🔹 Détermination de la source de données
input_data = ""
if uploaded_file is not None:
    input_data = uploaded_file.getvalue().decode("utf-8")
else:
    input_data = occ_text

# 🔹 Traitement au clic
if input_data and st.button("🚀 Générer le tableau XLSX"):
    # 1️⃣ Pré‑traitement : découpe des lignes non vides
    lines = [line for line in input_data.splitlines() if line.strip()]

    # 2️⃣ Extraction / nettoyage des tokens
    tokens = []
    for line in lines:
        # scinder par « | », retirer les espaces superflus et conserver les non‑vides
        parts = [part.strip() for part in line.split("|") if part.strip()]
        tokens.extend(parts)

    # 3️⃣ Comptage des occurrences
    if tokens:
        counts = (
            pd.Series(tokens)
            .value_counts()  # fréquence descendante
            .reset_index(names=["Occurrences"])  # renommage colonne comptage
            .rename(columns={"index": "Mot Clé"})
        )

        # 4️⃣ Aperçu dans l'app
        st.subheader("🔎 Aperçu du tableau consolidé")
        st.dataframe(counts, use_container_width=True)

        # 5️⃣ Export en mémoire (BytesIO) puis téléchargement
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            counts.to_excel(writer, index=False, sheet_name="Occurrences")
        buffer.seek(0)

        st.download_button(
            label="📥 Télécharger le XLSX consolidé",
            data=buffer,
            file_name="occurrences.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Aucun mot‑clé détecté dans l'entrée.")
