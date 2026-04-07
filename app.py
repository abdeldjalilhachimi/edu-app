"""
app.py

Streamlit UI entry point for the Excel Multi-File Processor.
All business logic is delegated to the modules/ package.
No data manipulation occurs in this file.

Tab 1: Multi-file merge (original feature)
Tab 2: Trimestrial BRUTSS consolidation (3 monthly files → 1 output)
Tab 3: Payroll calculation (RETSS / PARTSS / NETPAI)
"""

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
}
for _key, _default in _SESSION_DEFAULTS.items():
    if _key not in st.session_state:
        st.session_state[_key] = _default

# ─────────────────────────────────────────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs([
    "📊 Fusion Multi-Fichiers",
    "📅 Consolidation Trimestrielle",
    "💰 Calcul Paie (RETSS/PARTSS)",
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

    files_ready = main_file is not None and (bool(additional_files) or include_internal)

    if not files_ready:
        missing = []
        if main_file is None:
            missing.append("fichier principal")
        if not additional_files and not include_internal:
            missing.append("fichier(s) additionnel(s)")
        st.info(f"ℹ️ Veuillez téléverser le(s) {' et le(s) '.join(missing)} pour continuer.")

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

                month1 = parse_monthly_file(file_m1, file_m1.name)
                month2 = parse_monthly_file(file_m2, file_m2.name)
                month3 = parse_monthly_file(file_m3, file_m3.name)

                trim_result = merge_trimestrial(
                    month1, month2, month3,
                    file_m1.name, file_m2.name, file_m3.name,
                )

                trim_bytes = create_trimestrial_excel(trim_result, anref_year=int(trim_anref_year))

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

                employees = parse_payroll_file(pay_file, pay_file.name)

                pay_result = calculate_payroll(employees)

                pay_bytes = create_payroll_excel(pay_result)

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
        )

# ─────────────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────────────

st.divider()

st.title("🧑🏻‍💻 Create by Abdeldjalil Hachimi ")
