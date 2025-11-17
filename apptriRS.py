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
        label="📦 Télécharger `Revue_systematique_fichier.xlsx`",
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
# NETTOYAGE LÉGER DES DONNÉES
# --------------------------
# - suppression des espaces en début/fin de chaîne
# - harmonisation de certaines modalités (ALL/All, temperature/Temperature, etc.)

str_cols = df_final.select_dtypes(include=["object"]).columns
for col in str_cols:
    df_final[col] = df_final[col].apply(
        lambda x: x.strip() if isinstance(x, str) else x
    )

# Harmonisation de Life_stage (ALL -> All)
if "Life_stage" in df_final.columns:
    df_final["Life_stage"] = df_final["Life_stage"].replace({"ALL": "All"})

# Harmonisation de la variable "temperature" / "Temperature"
if "Variable_real" in df_final.columns:
    df_final["Variable_real"] = df_final["Variable_real"].replace(
        {"temperature": "Temperature"}
    )

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

    # Extraction de la série pour la colonne choisie
    series = df_filtered[col_to_analyse]
    total_rows = len(df_filtered)
    total_non_null = series.notna().sum()

    if total_rows == 0:
        st.warning(f"Aucune ligne dans les résultats filtrés pour la colonne `{col_to_analyse}`.")
    else:
        # Comptage des modalités en incluant les NA
        counts_all = series.value_counts(dropna=False)

        dist_df = counts_all.reset_index()
        dist_df.columns = [col_to_analyse, "N"]

        # Pourcentage parmi toutes les lignes filtrées
        dist_df["% parmi toutes les lignes filtrées"] = dist_df["N"] / total_rows * 100

        # Pourcentage parmi les non-nuls (NA n'ont pas de % parmi non nuls)
        if total_non_null > 0:
            def pct_non_null(row):
                if pd.isna(row[col_to_analyse]):
                    return None
                return row["N"] / total_non_null * 100

            dist_df["% parmi non nuls"] = dist_df.apply(pct_non_null, axis=1)
        else:
            dist_df["% parmi non nuls"] = None

        # Option pour limiter aux top modalités (correction du slider)
        nb_modalites = dist_df.shape[0]
        max_for_slider = min(50, nb_modalites)
        min_for_slider = 1
        default_val = min(20, max_for_slider)

        max_modalities = st.slider(
            "Nombre maximum de modalités à afficher (triées par fréquence décroissante) :",
            min_value=min_for_slider,
            max_value=max_for_slider,
            value=default_val
        )

        # Tri par fréquence décroissante et sélection des top modalités
        dist_df = dist_df.sort_values("N", ascending=False)
        dist_display = dist_df.head(max_modalities).copy()

        # Remplacement de la modalité NA par un libellé explicite
        if dist_display[col_to_analyse].isna().any():
            dist_display[col_to_analyse] = dist_display[col_to_analyse].astype(object)
            dist_display.loc[
                dist_display[col_to_analyse].isna(),
                col_to_analyse
            ] = "NA / manquant"

        st.write(
            f"### 📊 Distribution de la colonne `{col_to_analyse}` "
            f"(sur {total_rows} lignes filtrées, {total_non_null} non nulles)"
        )
        st.dataframe(dist_display, use_container_width=True, height=500)

        # Bar chart sur les N
        st.write("#### 🔎 Visualisation des effectifs (Top modalités)")
        chart_data = dist_display.set_index(col_to_analyse)["N"]
        st.bar_chart(chart_data)

        # Téléchargement des stats de distribution complètes (avec toutes les modalités)
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
