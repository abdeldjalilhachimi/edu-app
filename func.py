def run_calculation(
    main_df: pd.DataFrame,
    additional_dfs: list,
) -> CalculationResult:
    """
    Key-based merge between main file and additional files.

    Rules:
    - If key exists in both main and additional -> BRUTSS = main + additional
    - If key exists only in main -> keep unchanged
    - If key exists only in additional -> append as new row
    """
    lookup = build_additional_lookup(additional_dfs)

    duplicates = []
    updated_brutss = []
    main_keys = set()

    for _, row in main_df.iterrows():
        key = row[KEY_COLUMN]
        main_keys.add(key)

        brutss_main = float(row[BRUTSS_COLUMN])
        additional_brutss = lookup.get(key)

        if additional_brutss is not None and key != "||":
            brutss_sum = brutss_main + additional_brutss
            updated_brutss.append(brutss_sum)

            duplicates.append(DuplicateMatch(
                numcpt=str(row.get("NUMCPT", "")).strip(),
                nom=str(row.get("NOM", "")).strip(),
                prenom=str(row.get("PRENOM", "")).strip(),
                brutss_main=brutss_main,
                brutss_additional=additional_brutss,
                brutss_sum=brutss_sum,
            ))
        else:
            updated_brutss.append(brutss_main)

    result_df = main_df.copy()
    result_df[BRUTSS_COLUMN] = updated_brutss

    # Append rows موجودة فقط في additional files
    additional_only_rows = []

    for df in additional_dfs:
        for _, row in df.iterrows():
            key = row[KEY_COLUMN]

            if key == "||" or key in main_keys:
                continue

            # avoid duplicate append for same additional-only key
            if any(r[KEY_COLUMN] == key for r in additional_only_rows):
                continue

            new_row = row.copy()

            # ensure BRUTSS is the summed value from lookup
            new_row[BRUTSS_COLUMN] = lookup[key]

            additional_only_rows.append(new_row)

    if additional_only_rows:
        additional_only_df = pd.DataFrame(additional_only_rows)
        result_df = pd.concat([result_df, additional_only_df], ignore_index=True)

    brutss_total = math.fsum(result_df[BRUTSS_COLUMN].astype(float).tolist())

    stats = {
        "total": len(result_df),
        "duplicate_count": len(duplicates),
        "brutss_total": brutss_total,
    }

    return CalculationResult(
        updated_main_df=result_df,
        duplicates=duplicates,
        stats=stats,
    )
