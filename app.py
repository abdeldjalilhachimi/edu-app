"""
app.py

Streamlit UI entry point for the Excel Multi-File Processor.
All business logic is delegated to the modules/ package.
No data manipulation occurs in this file.
"""

import streamlit as st

from modules.validator import validate_all_files
from modules.cleaner import clean_all_dataframes
from modules.calculator import run_calculation
from modules.exporter import create_output_excel

# ─────────────────────────────────────────────────────────────────────────────
# Page configuration
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Excel Multi-File Processor",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# Session state initialisation
# ─────────────────────────────────────────────────────────────────────────────

if "result_bytes" not in st.session_state:
    st.session_state["result_bytes"] = None
if "error_message" not in st.session_state:
    st.session_state["error_message"] = None
if "processing_info" not in st.session_state:
    st.session_state["processing_info"] = None

# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────

st.title("📊 Excel Multi-File Processor")
st.markdown(
    "Traitez un fichier principal et plusieurs fichiers additionnels Excel. "
    "Pour chaque employé (NOM/PRENOM/NUMCPT) trouvé dans les fichiers additionnels, "
    "le **BRUTSS** additionnel est ajouté à la ligne correspondante du fichier principal. "
    "Les lignes sans correspondance restent inchangées."
)
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# Upload section
# ─────────────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Fichier Principal")
    st.caption("Un seul fichier .xlsx contenant les colonnes : BRUTSS, NOM, PRENOM, NUMCPT")
    main_file = st.file_uploader(
        label="Fichier principal",
        type=["xlsx"],
        accept_multiple_files=False,
        key="main_file_uploader",
        label_visibility="collapsed",
    )
    if main_file:
        st.success(f"✅ {main_file.name} ({main_file.size / 1024:.1f} Ko)")

with col2:
    st.subheader("📂 Fichiers Additionnels")
    st.caption("Un ou plusieurs fichiers .xlsx avec les mêmes colonnes obligatoires")
    additional_files = st.file_uploader(
        label="Fichiers additionnels",
        type=["xlsx"],
        accept_multiple_files=True,
        key="additional_files_uploader",
        label_visibility="collapsed",
    )
    if additional_files:
        for f in additional_files:
            st.success(f"✅ {f.name} ({f.size / 1024:.1f} Ko)")

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# Process button
# ─────────────────────────────────────────────────────────────────────────────

files_ready = main_file is not None and bool(additional_files)

if not files_ready:
    missing = []
    if main_file is None:
        missing.append("fichier principal")
    if not additional_files:
        missing.append("fichier(s) additionnel(s)")
    st.info(f"ℹ️ Veuillez téléverser le(s) {' et le(s) '.join(missing)} pour continuer.")

process_clicked = st.button(
    "▶  Traiter les fichiers",
    type="primary",
    disabled=not files_ready,
    use_container_width=False,
)

# ─────────────────────────────────────────────────────────────────────────────
# Processing pipeline
# ─────────────────────────────────────────────────────────────────────────────

if process_clicked and files_ready:
    # Clear previous results
    st.session_state["result_bytes"] = None
    st.session_state["error_message"] = None
    st.session_state["processing_info"] = None

    main_filename = main_file.name
    additional_filenames = [f.name for f in additional_files]

    try:
        with st.spinner("Validation et traitement en cours…"):

            # Step 1 — Validate all files
            main_df, additional_dfs, dropped_counts = validate_all_files(
                main_file=main_file,
                main_filename=main_filename,
                additional_files=additional_files,
                additional_filenames=additional_filenames,
            )

            # Step 2 — Clean and normalize BRUTSS; build composite keys
            main_clean, additional_clean = clean_all_dataframes(
                main_df=main_df,
                additional_dfs=additional_dfs,
                main_filename=main_filename,
                additional_filenames=additional_filenames,
            )

            # Step 3 — Key-based merge: match rows + update BRUTSS
            result = run_calculation(
                main_df=main_clean,
                additional_dfs=additional_clean,
            )

            # Step 4 — Build output Excel bytes
            output_bytes = create_output_excel(
                updated_main_df=result.updated_main_df,
                duplicates=result.duplicates,
                stats=result.stats,
            )

        # Store results in session state so they persist on re-render
        st.session_state["result_bytes"] = output_bytes
        st.session_state["processing_info"] = {
            "main_rows": result.stats["total"],
            "additional_count": len(additional_files),
            "duplicate_count": result.stats["duplicate_count"],
            "brutss_total": result.stats["brutss_total"],
            "dropped_counts": dropped_counts,
        }

    except ValueError as e:
        st.session_state["error_message"] = str(e)

# ─────────────────────────────────────────────────────────────────────────────
# Results / Error display
# ─────────────────────────────────────────────────────────────────────────────

if st.session_state["error_message"]:
    st.divider()
    st.error("❌ Erreur de traitement")
    st.code(st.session_state["error_message"], language=None)
    st.caption("Corrigez le problème dans le fichier source et re-téléversez.")

if st.session_state["result_bytes"]:
    st.divider()
    info = st.session_state["processing_info"]

    st.success(
        f"✅ Traitement terminé — "
        f"{info['additional_count']} fichier(s) additionnel(s) traité(s), "
        f"{info['main_rows']} ligne(s) mises à jour."
    )

    # Warn about any auto-dropped rows (blank / total rows)
    total_dropped = sum(info["dropped_counts"].values())
    if total_dropped > 0:
        details = "; ".join(
            f"{fname}: {n} ligne(s)"
            for fname, n in info["dropped_counts"].items()
            if n > 0
        )
        st.warning(
            f"⚠️ {total_dropped} ligne(s) ignorée(s) automatiquement "
            f"(lignes vides ou lignes de totaux sans NOM/PRENOM/NUMCPT) — {details}"
        )

    # Processing summary metrics
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("Total BRUTSS final", f"{info['brutss_total']:,.2f}".replace(",", " ").replace(".", ","))
    with col_b:
        st.metric("Lignes dans le fichier principal", info["main_rows"])
    with col_c:
        dup_label = f"{info['duplicate_count']}" if info["duplicate_count"] else "Aucune"
        st.metric("Correspondances trouvées", dup_label)

    st.download_button(
        label="⬇  Télécharger le résultat (output.xlsx)",
        data=st.session_state["result_bytes"],
        file_name="output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=False,
    )

# ─────────────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────────────

st.divider()
st.caption(
    "Colonnes requises dans chaque fichier : **BRUTSS**, **NOM**, **PRENOM**, **NUMCPT**  |  "
    "Formats numériques supportés : `1 234,56` · `1,234.56` · `1.234,56` · `1234.56`"
)
