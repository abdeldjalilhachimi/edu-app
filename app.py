"""
app.py

Streamlit UI entry point for the Excel Multi-File Processor.
All business logic is delegated to the modules/ package.
No data manipulation occurs in this file.

Tab 1: Multi-file merge (original feature)
Tab 2: Trimestrial BRUTSS consolidation (3 monthly files → 1 output)
Tab 3: Payroll calculation (RETSS / PARTSS / NETPAI)
Tab 4: Annual declaration (4 trimestrial files → 1 annual output)
"""

import os
import streamlit as st
import pandas as pd

from modules.validator import validate_all_files
from modules.cleaner import clean_all_dataframes
from modules.calculator import run_calculation
from modules.exporter import create_output_excel

from modules.trimestrial_parser import parse_monthly_file, validate_trimestrial_files
from modules.trimestrial_merger import merge_trimestrial
from modules.trimestrial_exporter import create_trimestrial_excel, _format_french
from modules.trimestrial_types import filename_to_label

from modules.payroll_parser import parse_payroll_file
from modules.payroll_calculator import calculate_payroll
from modules.payroll_exporter import create_payroll_excel, _format_french_payroll

from modules.annual_parser import parse_quarterly_file
from modules.annual_merger import merge_annual
from modules.annual_exporter import create_annual_excel, _format_french_annual

from modules.txt_converter import convert_xlsx_to_txt

from modules.demo_guard import (
    is_demo_expired,
    is_unlocked,
    get_remaining_downloads,
    increment_downloads,
    try_activate,
    CONTACT_EMAIL,
    MAX_FREE_DOWNLOADS,
)

# ─────────────────────────────────────────────────────────────────────────────
# Page configuration
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Excel Multi-File Processor",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# Session state initialisation (single loop instead of 12 individual checks)
# ─────────────────────────────────────────────────────────────────────────────

_SESSION_DEFAULTS = {
    # Tab 1
    "result_bytes": None,
    "error_message": None,
    "processing_info": None,
    "processing_tab1": False,
    # Tab 2
    "trim_result_bytes": None,
    "trim_error_message": None,
    "trim_processing_info": None,
    "processing_tab2": False,
    # Tab 3
    "pay_result_bytes": None,
    "pay_error_message": None,
    "pay_processing_info": None,
    "processing_tab3": False,
    # Tab 4
    "ann_result_bytes": None,
    "ann_error_message": None,
    "ann_processing_info": None,
    "processing_tab4": False,
    # Tab 5
    "txt_result_bytes": None,
    "txt_error_message": None,
    "txt_processing_info": None,
    "processing_tab5": False,
}
for _key, _default in _SESSION_DEFAULTS.items():
    if _key not in st.session_state:
        st.session_state[_key] = _default

# ─────────────────────────────────────────────────────────────────────────────
# Demo protection — activation dialog
# ─────────────────────────────────────────────────────────────────────────────

@st.dialog("Activation requise")
def _show_activation_dialog():
    """Blocking modal shown when the demo period has expired."""
    st.markdown("### :lock: Version d'essai terminée")
    st.markdown(
        f"Vous avez utilisé vos **{MAX_FREE_DOWNLOADS} téléchargements gratuits**."
    )
    st.divider()
    st.markdown("Entrez le code d'activation pour débloquer l'application :")

    code = st.text_input(
        "Code d'activation",
        type="password",
        placeholder="Entrez le code ici...",
        key="activation_code_input",
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Activer", type="primary", use_container_width=True):
            if code and try_activate(code):
                st.balloons()
                st.success("Application activée avec succès !")
                st.rerun()
            else:
                st.error("Code incorrect. Veuillez réessayer.")

    st.divider()
    st.markdown(
        f"Pour obtenir le code d'activation, contactez :\n\n"
        f"**:envelope: {CONTACT_EMAIL}**"
    )


# ── Gate: block the app if demo expired ─────────────────────────────────────

if is_demo_expired():
    _show_activation_dialog()
    st.stop()

# ── Demo badge: show remaining downloads (only in demo mode) ────────────────

if not is_unlocked():
    _remaining = get_remaining_downloads()
    st.info(
        f":hourglass_flowing_sand: **Version d'essai** — "
        f"{_remaining}/{MAX_FREE_DOWNLOADS} téléchargement(s) restant(s)",
        icon="ℹ️",
    )

# ─────────────────────────────────────────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Fusion Multi-Fichiers",
    "📅 Consolidation Trimestrielle",
    "💰 Calcul Paie (RETSS/PARTSS)",
    "📋 Déclaration Annuelle",
    "📄 Export TXT",
])

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — Multi-file merge (original feature)
# ═════════════════════════════════════════════════════════════════════════════

with tab1:

    st.markdown(
        "Traitez un fichier principal et plusieurs fichiers additionnels Excel. "
        "Pour chaque employé (NOM/PRENOM/NUMCPT) trouvé dans les fichiers additionnels, "
        "le **BRUTSS** additionnel est ajouté à la ligne correspondante du fichier principal. "
        "Les lignes sans correspondance restent inchangées."
    )
    st.divider()

    # ── Upload section ───────────────────────────────────────────────────────

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
            disabled=st.session_state["processing_tab1"],
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
            disabled=st.session_state["processing_tab1"],
        )
        if additional_files:
            for f in additional_files:
                st.success(f"✅ {f.name} ({f.size / 1024:.1f} Ko)")

    st.divider()

    # ── Options ──────────────────────────────────────────────────────────────

    include_internal = st.checkbox(
        "📑 Inclure les feuilles internes comme données additionnelles",
        value=False,
        help="Si coché, les feuilles supplémentaires dans chaque fichier (ex: RAP1, RAP2...) "
             "seront automatiquement traitées comme des fichiers additionnels.",
        disabled=st.session_state["processing_tab1"],
    )

    anref_year = st.number_input(
        "📅 ANREF (année de référence)",
        min_value=2020,
        max_value=2099,
        value=2025,
        step=1,
        disabled=st.session_state["processing_tab1"],
    )

    st.divider()

    # ── Process button ───────────────────────────────────────────────────────

    files_ready = main_file is not None

    if not files_ready:
        st.info("ℹ️ Veuillez téléverser le fichier principal pour continuer. Les fichiers additionnels sont optionnels.")

    process_clicked = st.button(
        "▶  Traiter les fichiers",
        type="primary",
        disabled=not files_ready or st.session_state["processing_tab1"],
        use_container_width=False,
        key="btn_tab1",
    )

    # ── Processing pipeline ──────────────────────────────────────────────────

    # Pass 1: button click → set flag + rerun so widgets render as disabled
    if process_clicked and files_ready:
        st.session_state["result_bytes"] = None
        st.session_state["error_message"] = None
        st.session_state["processing_info"] = None
        st.session_state["processing_tab1"] = True
        st.rerun()

    # Pass 2: flag is True → widgets are disabled, now run the actual pipeline
    if st.session_state["processing_tab1"] and main_file is not None:
        main_filename = main_file.name
        add_files = additional_files if additional_files else []
        add_filenames = [f.name for f in add_files]

        try:
            with st.spinner("⏳ Traitement en cours…"):

                main_df, additional_dfs, additional_names, dropped_counts = validate_all_files(
                    main_file=main_file,
                    main_filename=main_filename,
                    additional_files=add_files,
                    additional_filenames=add_filenames,
                    include_internal_sheets=include_internal,
                )

                main_clean, additional_clean = clean_all_dataframes(
                    main_df=main_df,
                    additional_dfs=additional_dfs,
                    main_filename=main_filename,
                    additional_filenames=additional_names,
                )

                result = run_calculation(
                    main_df=main_clean,
                    additional_dfs=additional_clean,
                )

                output_bytes = create_output_excel(
                    updated_main_df=result.updated_main_df,
                    duplicates=result.duplicates,
                    stats=result.stats,
                    anref_year=int(anref_year),
                )

            # Store results in session state so they persist on re-render
            st.session_state["result_bytes"] = output_bytes
            st.session_state["processing_info"] = {
                "main_rows": result.stats["total"],
                "additional_count": len(additional_dfs),
                "duplicate_count": result.stats["duplicate_count"],
                "added_count": result.stats["added_count"],
                "brutss_total": result.stats["brutss_total"],
                "dropped_counts": dropped_counts,
            }

        except ValueError as e:
            st.session_state["error_message"] = str(e)
        finally:
            st.session_state["processing_tab1"] = False

    # ── Results / Error display ──────────────────────────────────────────────

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
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            st.metric("Total BRUTSS final", f"{info['brutss_total']:,.2f}".replace(",", " ").replace(".", ","))
        with col_b:
            st.metric("Lignes dans le résultat", info["main_rows"])
        with col_c:
            dup_label = f"{info['duplicate_count']}" if info["duplicate_count"] else "Aucune"
            st.metric("Correspondances trouvées", dup_label)
        with col_d:
            added_label = f"{info['added_count']}" if info.get("added_count") else "Aucune"
            st.metric("Nouvelles lignes ajoutées", added_label)

        st.download_button(
            label="⬇  Télécharger le résultat (output.xlsx)",
            data=st.session_state["result_bytes"],
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
            key="download_tab1",
            on_click=increment_downloads,
        )

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — Trimestrial BRUTSS consolidation
# ═════════════════════════════════════════════════════════════════════════════

with tab2:

    st.markdown(
        "Consolidez le **BRUTSS** de trois fichiers mensuels en un seul fichier trimestriel. "
        "Chaque employé (NUMCPT+NOM+PRENOM) apparaît une seule fois avec le BRUTSS de chaque mois "
        "et le total. Si un employé est absent d'un mois, son BRUTSS est mis à 0."
    )
    st.divider()

    # ── Upload section — 3 columns for 3 months ─────────────────────────────

    col_m1, col_m2, col_m3 = st.columns(3)

    with col_m1:
        st.subheader("📅 Mois 1")
        st.caption("Fichier .xlsx — colonnes : BRUTSS, NOM, PRENOM, NUMCPT")
        file_m1 = st.file_uploader(
            label="Mois 1",
            type=["xlsx"],
            accept_multiple_files=False,
            key="trim_file_m1",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab2"],
        )
        if file_m1:
            st.success(f"✅ {file_m1.name} ({file_m1.size / 1024:.1f} Ko)")

    with col_m2:
        st.subheader("📅 Mois 2")
        st.caption("Fichier .xlsx — colonnes : BRUTSS, NOM, PRENOM, NUMCPT")
        file_m2 = st.file_uploader(
            label="Mois 2",
            type=["xlsx"],
            accept_multiple_files=False,
            key="trim_file_m2",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab2"],
        )
        if file_m2:
            st.success(f"✅ {file_m2.name} ({file_m2.size / 1024:.1f} Ko)")

    with col_m3:
        st.subheader("📅 Mois 3")
        st.caption("Fichier .xlsx — colonnes : BRUTSS, NOM, PRENOM, NUMCPT")
        file_m3 = st.file_uploader(
            label="Mois 3",
            type=["xlsx"],
            accept_multiple_files=False,
            key="trim_file_m3",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab2"],
        )
        if file_m3:
            st.success(f"✅ {file_m3.name} ({file_m3.size / 1024:.1f} Ko)")

    st.divider()

    trim_anref_year = st.number_input(
        "📅 ANREF (année de référence)",
        min_value=2020,
        max_value=2099,
        value=2025,
        step=1,
        disabled=st.session_state["processing_tab2"],
        key="trim_anref_year",
    )

    st.divider()

    # ── Process button ───────────────────────────────────────────────────────

    trim_ready = file_m1 is not None and file_m2 is not None and file_m3 is not None

    if not trim_ready:
        missing = []
        if file_m1 is None:
            missing.append("Mois 1")
        if file_m2 is None:
            missing.append("Mois 2")
        if file_m3 is None:
            missing.append("Mois 3")
        st.info(f"ℹ️ Veuillez téléverser le(s) fichier(s) manquant(s) : {', '.join(missing)}.")

    trim_clicked = st.button(
        "▶  Consolider le trimestre",
        type="primary",
        disabled=not trim_ready or st.session_state["processing_tab2"],
        use_container_width=False,
        key="btn_tab2",
    )

    # ── Processing pipeline ──────────────────────────────────────────────────

    # Pass 1: button click → set flag + rerun so widgets render as disabled
    if trim_clicked and trim_ready:
        st.session_state["trim_result_bytes"] = None
        st.session_state["trim_error_message"] = None
        st.session_state["trim_processing_info"] = None
        st.session_state["processing_tab2"] = True
        st.rerun()

    # Pass 2: flag is True → widgets are disabled, now run the actual pipeline
    if st.session_state["processing_tab2"] and file_m1 is not None and file_m2 is not None and file_m3 is not None:
        try:
            with st.spinner("⏳ Consolidation en cours…"):

                validate_trimestrial_files(file_m1, file_m2, file_m3)

                month1, empty1 = parse_monthly_file(file_m1, file_m1.name)
                month2, empty2 = parse_monthly_file(file_m2, file_m2.name)
                month3, empty3 = parse_monthly_file(file_m3, file_m3.name)

                trim_result = merge_trimestrial(
                    month1, month2, month3,
                    file_m1.name, file_m2.name, file_m3.name,
                )

                # Collect empty-BRUTSS info per file (only non-empty lists)
                empty_brutss_per_file = {}
                for fname, emp_list in [
                    (file_m1.name, empty1),
                    (file_m2.name, empty2),
                    (file_m3.name, empty3),
                ]:
                    if emp_list:
                        empty_brutss_per_file[fname] = emp_list

                trim_bytes = create_trimestrial_excel(
                    trim_result,
                    anref_year=int(trim_anref_year),
                    empty_brutss_per_file=empty_brutss_per_file,
                )

            # Store results in session state
            st.session_state["trim_result_bytes"] = trim_bytes
            st.session_state["trim_processing_info"] = {
                "unique_count": trim_result.stats.unique_count,
                "file_labels": trim_result.stats.file_labels,
                "monthly_totals_cents": trim_result.stats.monthly_totals_cents,
                "grand_total_cents": trim_result.stats.grand_total_cents,
                "missing_per_file": [
                    [(e.numcpt_raw, e.nom, e.prenom) for e in missing_list]
                    for missing_list in trim_result.missing_per_file
                ],
            }

        except ValueError as e:
            st.session_state["trim_error_message"] = str(e)
        finally:
            st.session_state["processing_tab2"] = False

    # ── Results / Error display ──────────────────────────────────────────────

    if st.session_state["trim_error_message"]:
        st.divider()
        st.error("❌ Erreur de traitement")
        st.code(st.session_state["trim_error_message"], language=None)
        st.caption("Corrigez le problème dans le fichier source et re-téléversez.")

    if st.session_state["trim_result_bytes"]:
        st.divider()
        info = st.session_state["trim_processing_info"]

        st.success(
            f"✅ Consolidation terminée — "
            f"{info['unique_count']} employé(s) unique(s) consolidé(s)."
        )

        # Summary metrics — dynamic columns based on filenames
        file_labels = info["file_labels"]
        monthly_totals = info["monthly_totals_cents"]

        metric_cols = st.columns(2 + len(file_labels))
        with metric_cols[0]:
            st.metric("Employés uniques", info["unique_count"])
        for i, (label, total_cents) in enumerate(zip(file_labels, monthly_totals)):
            with metric_cols[1 + i]:
                st.metric(f"Total {label}", _format_french(total_cents))
        with metric_cols[-1]:
            st.metric("Total Trimestre", _format_french(info["grand_total_cents"]))

        st.download_button(
            label="⬇  Télécharger la consolidation (trimestre.xlsx)",
            data=st.session_state["trim_result_bytes"],
            file_name="trimestre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
            key="download_tab2",
            on_click=increment_downloads,
        )

        # ── Absences detection — employees missing per file ──────────────
        missing_per_file = info["missing_per_file"]
        total_missing = sum(len(m) for m in missing_per_file)

        if total_missing > 0:
            st.divider()
            st.subheader("🔍 Détection des absences")
            st.caption(
                "Employés présents dans au moins un fichier mais absents d'un autre. "
                "Le détail complet est aussi disponible dans l'onglet « Absences » du fichier Excel."
            )

            for i, (label, missing_list) in enumerate(zip(file_labels, missing_per_file)):
                count = len(missing_list)
                if count == 0:
                    st.success(f"✅ **{label}** — Aucun absent (tous les employés sont présents)")
                else:
                    with st.expander(f"⚠️ **{label}** — {count} employé(s) absent(s)", expanded=False):
                        missing_df = pd.DataFrame(
                            missing_list,
                            columns=["NUMCPT", "NOM", "PRENOM"],
                        )
                        missing_df.index = missing_df.index + 1  # 1-based index
                        st.dataframe(missing_df, use_container_width=True, hide_index=False)
        else:
            st.divider()
            st.success("✅ Aucune absence détectée — tous les employés sont présents dans les 3 fichiers.")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 — Payroll calculation (RETSS / PARTSS / NETPAI)
# ═════════════════════════════════════════════════════════════════════════════

with tab3:

    st.markdown(
        "Calculez les cotisations **RETSS** (9 %) et **PARTSS** à partir du **BRUTSS** "
        "de chaque employé. Le **NETPAI** est lu directement depuis le fichier.\n\n"
        "- **Titulaires** : PARTSS = 25 %\n"
        "- **Non Titulaires** (handicap) : PARTSS = 12,5 %"
    )
    st.divider()

    # ── Upload section ───────────────────────────────────────────────────────

    st.subheader("📁 Fichier de paie")
    st.caption("Un fichier .xlsx contenant les colonnes : BRUTSS, NETPAI, NOM, PRENOM, NUMCPT")
    pay_file = st.file_uploader(
        label="Fichier de paie",
        type=["xlsx"],
        accept_multiple_files=False,
        key="pay_file_uploader",
        label_visibility="collapsed",
        disabled=st.session_state["processing_tab3"],
    )
    if pay_file:
        st.success(f"✅ {pay_file.name} ({pay_file.size / 1024:.1f} Ko)")

    st.divider()

    # ── Process button ───────────────────────────────────────────────────────

    pay_ready = pay_file is not None

    if not pay_ready:
        st.info("ℹ️ Veuillez téléverser un fichier de paie pour continuer.")

    pay_clicked = st.button(
        "▶  Calculer la paie",
        type="primary",
        disabled=not pay_ready or st.session_state["processing_tab3"],
        use_container_width=False,
        key="btn_tab3",
    )

    # ── Processing pipeline ──────────────────────────────────────────────────

    # Pass 1: button click → set flag + rerun so widgets render as disabled
    if pay_clicked and pay_ready:
        st.session_state["pay_result_bytes"] = None
        st.session_state["pay_error_message"] = None
        st.session_state["pay_processing_info"] = None
        st.session_state["processing_tab3"] = True
        st.rerun()

    # Pass 2: flag is True → widgets are disabled, now run the actual pipeline
    if st.session_state["processing_tab3"] and pay_file is not None:
        try:
            with st.spinner("⏳ Calcul de la paie en cours…"):

                employees, pay_empty_brutss = parse_payroll_file(pay_file, pay_file.name)

                pay_result = calculate_payroll(employees)

                pay_bytes = create_payroll_excel(pay_result, empty_brutss=pay_empty_brutss)

            # Store results in session state
            st.session_state["pay_result_bytes"] = pay_bytes
            st.session_state["pay_processing_info"] = {
                "total_employees": pay_result.stats.total_employees,
                "confirmed_count": pay_result.stats.confirmed_count,
                "non_confirmed_count": pay_result.stats.non_confirmed_count,
                "grand_brutss_cents": pay_result.stats.grand_brutss_cents,
                "grand_retss_cents": pay_result.stats.grand_retss_cents,
                "grand_partss_cents": pay_result.stats.grand_partss_cents,
                "grand_netpai_cents": pay_result.stats.grand_netpai_cents,
                "confirmed_partss_cents": pay_result.stats.confirmed_partss_cents,
                "non_confirmed_partss_cents": pay_result.stats.non_confirmed_partss_cents,
            }

        except ValueError as e:
            st.session_state["pay_error_message"] = str(e)
        finally:
            st.session_state["processing_tab3"] = False

    # ── Results / Error display ──────────────────────────────────────────────

    if st.session_state["pay_error_message"]:
        st.divider()
        st.error("❌ Erreur de traitement")
        st.code(st.session_state["pay_error_message"], language=None)
        st.caption("Corrigez le problème dans le fichier source et re-téléversez.")

    if st.session_state["pay_result_bytes"]:
        st.divider()
        info = st.session_state["pay_processing_info"]

        st.success(
            f"✅ Calcul terminé — "
            f"{info['total_employees']} employé(s) traité(s) "
            f"({info['confirmed_count']} titulaires, {info['non_confirmed_count']} non titulaires)."
        )

        # Summary metrics
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            st.metric("Total employés", info["total_employees"])
        with col_b:
            st.metric("Total BRUTSS", _format_french_payroll(info["grand_brutss_cents"]))
        with col_c:
            st.metric("Total RETSS (9%)", _format_french_payroll(info["grand_retss_cents"]))
        with col_d:
            st.metric("Total NETPAI", _format_french_payroll(info["grand_netpai_cents"]))

        col_e, col_f, col_g, col_h = st.columns(4)
        with col_e:
            st.metric("Titulaires", info["confirmed_count"])
        with col_f:
            st.metric("PARTSS Titulaires (25%)", _format_french_payroll(info["confirmed_partss_cents"]))
        with col_g:
            st.metric("Non Titulaires", info["non_confirmed_count"])
        with col_h:
            st.metric("PARTSS Non Tit. (12,5%)", _format_french_payroll(info["non_confirmed_partss_cents"]))

        st.download_button(
            label="⬇  Télécharger le calcul de paie (paie.xlsx)",
            data=st.session_state["pay_result_bytes"],
            file_name="paie.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
            key="download_tab3",
            on_click=increment_downloads,
        )

# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 — Annual declaration (4 trimestrial files)
# ═════════════════════════════════════════════════════════════════════════════

with tab4:

    st.markdown(
        "Consolidez **4 fichiers trimestriels** en une seule **déclaration annuelle**. "
        "Chaque fichier doit contenir la colonne **BRUTSS_TOTAL** (fichier de sortie de l'onglet Consolidation Trimestrielle). "
        "Le résultat affiche le BRUTSS de chaque trimestre et le **total annuel** par employé."
    )
    st.divider()

    # ── Upload section — 4 columns for 4 quarters ──────────────────────────

    col_q1, col_q2 = st.columns(2)
    col_q3, col_q4 = st.columns(2)

    with col_q1:
        st.subheader("📅 1er Trimestre")
        st.caption("Fichier .xlsx avec colonne BRUTSS_TOTAL")
        file_q1 = st.file_uploader(
            label="1er Trimestre",
            type=["xlsx"],
            accept_multiple_files=False,
            key="ann_file_q1",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab4"],
        )
        if file_q1:
            st.success(f"✅ {file_q1.name} ({file_q1.size / 1024:.1f} Ko)")

    with col_q2:
        st.subheader("📅 2ème Trimestre")
        st.caption("Fichier .xlsx avec colonne BRUTSS_TOTAL")
        file_q2 = st.file_uploader(
            label="2ème Trimestre",
            type=["xlsx"],
            accept_multiple_files=False,
            key="ann_file_q2",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab4"],
        )
        if file_q2:
            st.success(f"✅ {file_q2.name} ({file_q2.size / 1024:.1f} Ko)")

    with col_q3:
        st.subheader("📅 3ème Trimestre")
        st.caption("Fichier .xlsx avec colonne BRUTSS_TOTAL")
        file_q3 = st.file_uploader(
            label="3ème Trimestre",
            type=["xlsx"],
            accept_multiple_files=False,
            key="ann_file_q3",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab4"],
        )
        if file_q3:
            st.success(f"✅ {file_q3.name} ({file_q3.size / 1024:.1f} Ko)")

    with col_q4:
        st.subheader("📅 4ème Trimestre")
        st.caption("Fichier .xlsx avec colonne BRUTSS_TOTAL")
        file_q4 = st.file_uploader(
            label="4ème Trimestre",
            type=["xlsx"],
            accept_multiple_files=False,
            key="ann_file_q4",
            label_visibility="collapsed",
            disabled=st.session_state["processing_tab4"],
        )
        if file_q4:
            st.success(f"✅ {file_q4.name} ({file_q4.size / 1024:.1f} Ko)")

    st.divider()

    ann_anref_year = st.number_input(
        "📅 ANREF (année de référence)",
        min_value=2020,
        max_value=2099,
        value=2025,
        step=1,
        disabled=st.session_state["processing_tab4"],
        key="ann_anref_year",
    )

    st.divider()

    # ── Process button ───────────────────────────────────────────────────────

    ann_ready = all(f is not None for f in [file_q1, file_q2, file_q3, file_q4])

    if not ann_ready:
        missing = []
        if file_q1 is None:
            missing.append("1er Trimestre")
        if file_q2 is None:
            missing.append("2ème Trimestre")
        if file_q3 is None:
            missing.append("3ème Trimestre")
        if file_q4 is None:
            missing.append("4ème Trimestre")
        st.info(f"ℹ️ Veuillez téléverser le(s) fichier(s) manquant(s) : {', '.join(missing)}.")

    ann_clicked = st.button(
        "▶  Générer la déclaration annuelle",
        type="primary",
        disabled=not ann_ready or st.session_state["processing_tab4"],
        use_container_width=False,
        key="btn_tab4",
    )

    # ── Processing pipeline ──────────────────────────────────────────────────

    # Pass 1: button click → set flag + rerun
    if ann_clicked and ann_ready:
        st.session_state["ann_result_bytes"] = None
        st.session_state["ann_error_message"] = None
        st.session_state["ann_processing_info"] = None
        st.session_state["processing_tab4"] = True
        st.rerun()

    # Pass 2: flag is True → widgets are disabled, now run pipeline
    if st.session_state["processing_tab4"] and all(
        f is not None for f in [file_q1, file_q2, file_q3, file_q4]
    ):
        try:
            with st.spinner("⏳ Génération de la déclaration annuelle en cours…"):

                q1_data, empty_q1 = parse_quarterly_file(file_q1, file_q1.name)
                q2_data, empty_q2 = parse_quarterly_file(file_q2, file_q2.name)
                q3_data, empty_q3 = parse_quarterly_file(file_q3, file_q3.name)
                q4_data, empty_q4 = parse_quarterly_file(file_q4, file_q4.name)

                ann_result = merge_annual(
                    q1_data, q2_data, q3_data, q4_data,
                    file_q1.name, file_q2.name, file_q3.name, file_q4.name,
                )

                # Collect empty-BRUTSS info per file
                ann_empty_brutss = {}
                for fname, emp_list in [
                    (file_q1.name, empty_q1),
                    (file_q2.name, empty_q2),
                    (file_q3.name, empty_q3),
                    (file_q4.name, empty_q4),
                ]:
                    if emp_list:
                        ann_empty_brutss[fname] = emp_list

                ann_bytes = create_annual_excel(
                    ann_result,
                    anref_year=int(ann_anref_year),
                    empty_brutss_per_file=ann_empty_brutss,
                )

            # Store results
            st.session_state["ann_result_bytes"] = ann_bytes
            st.session_state["ann_processing_info"] = {
                "unique_count": ann_result.stats.unique_count,
                "file_labels": ann_result.stats.file_labels,
                "quarterly_totals_cents": ann_result.stats.quarterly_totals_cents,
                "grand_total_cents": ann_result.stats.grand_total_cents,
                "missing_per_file": [
                    [(e.numcpt_raw, e.nom, e.prenom) for e in missing_list]
                    for missing_list in ann_result.missing_per_file
                ],
            }

        except ValueError as e:
            st.session_state["ann_error_message"] = str(e)
        finally:
            st.session_state["processing_tab4"] = False

    # ── Results / Error display ──────────────────────────────────────────────

    if st.session_state["ann_error_message"]:
        st.divider()
        st.error("❌ Erreur de traitement")
        st.code(st.session_state["ann_error_message"], language=None)
        st.caption("Corrigez le problème dans le fichier source et re-téléversez.")

    if st.session_state["ann_result_bytes"]:
        st.divider()
        info = st.session_state["ann_processing_info"]

        st.success(
            f"✅ Déclaration annuelle générée — "
            f"{info['unique_count']} employé(s) unique(s) consolidé(s)."
        )

        # Summary metrics
        file_labels = info["file_labels"]
        quarterly_totals = info["quarterly_totals_cents"]

        metric_cols = st.columns(2 + len(file_labels))
        with metric_cols[0]:
            st.metric("Employés uniques", info["unique_count"])
        for i, (label, total_cents) in enumerate(zip(file_labels, quarterly_totals)):
            with metric_cols[1 + i]:
                st.metric(f"Total {label}", _format_french_annual(total_cents))
        with metric_cols[-1]:
            st.metric("Total Annuel", _format_french_annual(info["grand_total_cents"]))

        st.download_button(
            label="⬇  Télécharger la déclaration annuelle (annuel.xlsx)",
            data=st.session_state["ann_result_bytes"],
            file_name="annuel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
            key="download_tab4",
            on_click=increment_downloads,
        )

        # Absences detection
        missing_per_file = info["missing_per_file"]
        total_missing = sum(len(m) for m in missing_per_file)

        if total_missing > 0:
            st.divider()
            st.subheader("🔍 Détection des absences")
            st.caption(
                "Employés présents dans au moins un trimestre mais absents d'un autre. "
                "Le détail complet est aussi disponible dans l'onglet « Absences » du fichier Excel."
            )

            for label, missing_list in zip(file_labels, missing_per_file):
                count = len(missing_list)
                if count == 0:
                    st.success(f"✅ **{label}** — Aucun absent")
                else:
                    with st.expander(f"⚠️ **{label}** — {count} employé(s) absent(s)", expanded=False):
                        missing_df = pd.DataFrame(
                            missing_list,
                            columns=["NUMCPT", "NOM", "PRENOM"],
                        )
                        missing_df.index = missing_df.index + 1
                        st.dataframe(missing_df, use_container_width=True, hide_index=False)
        else:
            st.divider()
            st.success("✅ Aucune absence détectée — tous les employés sont présents dans les 4 trimestres.")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 — Export TXT (annual xlsx → pipe-delimited txt)
# ═════════════════════════════════════════════════════════════════════════════

with tab5:

    st.markdown(
        "Convertissez un fichier de **déclaration annuelle** (.xlsx) en fichier texte (.txt). "
        "Les colonnes trimestrielles (BRUTSS par trimestre) sont supprimées — seul le **total annuel** "
        "(BRUTSS_ANNUEL) est conservé avec toutes les autres colonnes. "
        "Format de sortie : délimité par **|** (pipe), encodage UTF-8."
    )
    st.divider()

    # ── Upload section ───────────────────────────────────────────────────────

    st.subheader("📁 Fichier de déclaration annuelle")
    st.caption("Un fichier .xlsx contenant la colonne BRUTSS_ANNUEL (fichier de sortie de l'onglet Déclaration Annuelle)")
    txt_file = st.file_uploader(
        label="Fichier annuel",
        type=["xlsx"],
        accept_multiple_files=False,
        key="txt_file_uploader",
        label_visibility="collapsed",
        disabled=st.session_state["processing_tab5"],
    )
    if txt_file:
        st.success(f"✅ {txt_file.name} ({txt_file.size / 1024:.1f} Ko)")

    st.divider()

    # ── Process button ───────────────────────────────────────────────────────

    txt_ready = txt_file is not None

    if not txt_ready:
        st.info("ℹ️ Veuillez téléverser un fichier de déclaration annuelle pour continuer.")

    txt_clicked = st.button(
        "▶  Convertir en TXT",
        type="primary",
        disabled=not txt_ready or st.session_state["processing_tab5"],
        use_container_width=False,
        key="btn_tab5",
    )

    # ── Processing pipeline ──────────────────────────────────────────────────

    # Pass 1: button click → set flag + rerun
    if txt_clicked and txt_ready:
        st.session_state["txt_result_bytes"] = None
        st.session_state["txt_error_message"] = None
        st.session_state["txt_processing_info"] = None
        st.session_state["processing_tab5"] = True
        st.rerun()

    # Pass 2: flag is True → run pipeline
    if st.session_state["processing_tab5"] and txt_file is not None:
        try:
            with st.spinner("⏳ Conversion en cours…"):

                txt_bytes, txt_info = convert_xlsx_to_txt(txt_file, txt_file.name)

            st.session_state["txt_result_bytes"] = txt_bytes
            st.session_state["txt_processing_info"] = txt_info

        except ValueError as e:
            st.session_state["txt_error_message"] = str(e)
        finally:
            st.session_state["processing_tab5"] = False

    # ── Results / Error display ──────────────────────────────────────────────

    if st.session_state["txt_error_message"]:
        st.divider()
        st.error("❌ Erreur de traitement")
        st.code(st.session_state["txt_error_message"], language=None)
        st.caption("Corrigez le problème dans le fichier source et re-téléversez.")

    if st.session_state["txt_result_bytes"]:
        st.divider()
        info = st.session_state["txt_processing_info"]

        st.success(
            f"✅ Conversion terminée — "
            f"{info['total_rows']} ligne(s), {info['columns_kept']} colonne(s) conservée(s)."
        )

        if info["columns_dropped_count"] > 0:
            st.info(
                f"ℹ️ {info['columns_dropped_count']} colonne(s) trimestrielle(s) supprimée(s) : "
                f"{', '.join(info['columns_dropped'])}"
            )

        # Summary metrics
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Lignes", info["total_rows"])
        with col_b:
            st.metric("Colonnes conservées", info["columns_kept"])
        with col_c:
            st.metric("Colonnes supprimées", info["columns_dropped_count"])

        # Preview
        preview_text = st.session_state["txt_result_bytes"].decode("utf-8")
        preview_lines = preview_text.strip().split("\n")
        if len(preview_lines) > 6:
            preview_display = "\n".join(preview_lines[:6]) + f"\n... ({len(preview_lines) - 6} lignes supplémentaires)"
        else:
            preview_display = preview_text.strip()

        st.subheader("👁 Aperçu du fichier TXT")
        st.code(preview_display, language=None)

        # Generate output filename from input
        txt_output_name = os.path.splitext(txt_file.name)[0] + ".txt"

        st.download_button(
            label=f"⬇  Télécharger ({txt_output_name})",
            data=st.session_state["txt_result_bytes"],
            file_name=txt_output_name,
            mime="text/plain",
            type="primary",
            use_container_width=False,
            key="download_tab5",
            on_click=increment_downloads,
        )

# ─────────────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────────────

st.divider()

st.title("🧑🏻‍💻 Create by Abdeldjalil Hachimi ")
