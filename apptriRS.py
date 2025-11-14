import streamlit as st
import pandas as pd
import io

# ------------------------------
# CONFIGURATION DE LA PAGE
# ------------------------------
st.set_page_config(
    page_title="Revue Systématique Ixodes – Navigate & Analyse",
    page_icon="🕷️",
    layout="wide"
)

# ------------------------------
# CHEMIN DU FICHIER EXCEL
# ------------------------------
FILE_PATH = "Revue_systematique_resultats.xlsx"


# ------------------------------
# CHARGEMENT DU FICHIER EXCEL
# ------------------------------
@st.cache_data
def load_excel(path):
    xls = pd.ExcelFile(path)
    data = {sheet: pd.read_excel(path, sheet) for sheet in xls.sheet_names}
    return data, xls.sheet_names

data, sheet_names = load_excel(FILE_PATH)

# ------------------------------
# HEADER
# ------------------------------
st.title("🕷️ Revue Systématique Ixodes ricinus – Application d’exploration")
st.markdown("")  # gardé mais valide (chaîne vide)

# =====================================================================
# 1. TÉLÉCHARGEMENT GLOBAL
# =====================================================================
st.subheader("📥 Télécharger le fichier complet")

with open(FILE_PATH, "rb") as f:
    st.download_button(
        label="📦 Télécharger `Revue_systematique_complete_5.xlsx`",
        data=f.read(),
        file_name="Revue_systematique_complete_5.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")

# =====================================================================
# 2. NAVIGATION PAR FEUILLET (GRANDE FENÊTRE)
# =====================================================================
st.header("📑 Feuilleter les feuillets")

selected_sheet = st.selectbox("Choisir un feuillet à afficher :", sheet_names)

df = data[selected_sheet]

st.write(f"### 📘 Feuillet sélectionné : `{selected_sheet}`")

# Affichage en grande hauteur et pleine largeur
st.dataframe(df, use_container_width=True, height=700)

# Téléchargement du feuillet actuel
output_sheet = io.BytesIO()
with pd.ExcelWriter(output_sheet, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)

st.download_button(
    label=f"⬇️ Télécharger le feuillet `{selected_sheet}`",
    data=output_sheet.getvalue(),
    file_name=f"{selected_sheet}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")

# =====================================================================
# 3. MODULE AVANCÉ : FINAL_ARTICLES_AND_VARIABLES
# =====================================================================
st.header("🔬 Exploration avancée du feuillet `Final_articles_and_variables`")

df_final = data["Final_articles_and_variables"].copy()

# --------------------------
# 3.1. FILTRAGE & RECHERCHE
# --------------------------
st.subheader("🔍 Explorer, filtrer et extraire les articles")

df_filtered = df_final.copy()

# Filtres par colonnes
st.write("### 🎛️ Filtres par colonnes")

filter_cols = st.multiselect(
    "Choisir des colonnes à filtrer (optionnel) :",
    options=df_final.columns.tolist()
)

for col in filter_cols:
    unique_vals = sorted(df_final[col].dropna().unique().tolist())
    selected_vals = st.multiselect(
        f"Valeurs à retenir pour `{col}` :",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered[col].isin(selected_vals)]

st.write(f"### 📄 Résultats filtrés ({df_filtered.shape[0]} lignes)")
st.dataframe(df_filtered, use_container_width=True, height=600)

# --------------------------
# 3.2. TÉLÉCHARGEMENT DES RÉSULTATS FILTRÉS
# --------------------------
st.write("#### ⬇️ Exporter les résultats filtrés")

output_filtered = io.BytesIO()
with pd.ExcelWriter(output_filtered, engine="openpyxl") as writer:
    df_filtered.to_excel(writer, index=False, sheet_name="Filtered_Final")

st.download_button(
    label="💾 Télécharger les résultats filtrés (`Final_articles_and_variables_filtered.xlsx`)",
    data=output_filtered.getvalue(),
    file_name="Final_articles_and_variables_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")

# =====================================================================
# 3.3. STATISTIQUES DE DISTRIBUTION / REDONDANCE PAR VALEUR
# =====================================================================
st.subheader("📊 Statistiques de distribution par colonne (sur les résultats filtrés)")

if df_filtered.empty:
    st.warning("Aucun résultat filtré pour l’instant. Ajuste la recherche ou les filtres ci-dessus.")
else:
    # Choix de la colonne à analyser (par défaut 'Category' si présente)
    default_col = "Category" if "Category" in df_filtered.columns else df_filtered.columns[0]
    col_to_analyse = st.selectbox(
        "Choisir une colonne pour voir la distribution des valeurs :",
        options=df_filtered.columns.tolist(),
        index=list(df_filtered.columns).index(default_col)
    )

    # Extraction de la série et calcul de la distribution
    series = df_filtered[col_to_analyse].dropna()
    total_non_null = len(series)
    total_rows = len(df_filtered)

    if total_non_null == 0:
        st.warning(f"Aucune valeur non nulle dans la colonne `{col_to_analyse}` pour les résultats filtrés.")
    else:
        dist_df = (
            series.value_counts()
            .reset_index()
        )
        dist_df.columns = [col_to_analyse, "N"]

        dist_df["% parmi non nuls"] = dist_df["N"] / total_non_null * 100
        dist_df["% parmi toutes les lignes filtrées"] = dist_df["N"] / total_rows * 100

        # Option pour limiter aux top modalités
        max_modalities = st.slider(
            "Nombre maximum de modalités à afficher (triées par fréquence décroissante) :",
            min_value=5,
            max_value=min(50, dist_df.shape[0]),
            value=min(20, dist_df.shape[0])
        )

        dist_display = dist_df.head(max_modalities)

        st.write(
            f"### 📊 Distribution de la colonne `{col_to_analyse}` "
            f"(sur {total_rows} lignes filtrées, {total_non_null} non nulles)"
        )
        st.dataframe(dist_display, use_container_width=True, height=500)

        # Bar chart sur les N
        st.write("#### 🔎 Visualisation des effectifs (Top modalités)")
        chart_data = dist_display.set_index(col_to_analyse)["N"]
        st.bar_chart(chart_data)

        # Téléchargement des stats de distribution
        dist_output = io.BytesIO()
        with pd.ExcelWriter(dist_output, engine="openpyxl") as writer:
            dist_df.to_excel(writer, index=False, sheet_name=f"Distribution_{col_to_analyse}")

        st.download_button(
            label=f"📊 Télécharger la distribution complète de `{col_to_analyse}` (`Distribution_{col_to_analyse}.xlsx`)",
            data=dist_output.getvalue(),
            file_name=f"Distribution_{col_to_analyse}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")
st.write()

