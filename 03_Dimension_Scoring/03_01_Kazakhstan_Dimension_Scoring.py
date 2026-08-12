#!/usr/bin/env python3
"""
03_01 Kazakhstan Dimension Scoring
Deterministic, reproducible version.

Fixed input:
    BDSF_ALL/02_Scoring_Matrix/
    02_01_Kazakhstan_Cybersecurity_Scored_Dataset.csv

Outputs:
    BDSF_ALL/03_Dimension_Scoring/
    03_01_Kazakhstan_dimensions.csv
"""

from pathlib import Path
import sys
import numpy as np
import pandas as pd

EXPECTED_N = 176

DIMENSION_MAPPING = {
    "D1": ["IND01", "IND02", "IND03"],
    "D2": ["IND04", "IND05", "IND06"],
    "D3": ["IND07", "IND08"],
    "D4": ["IND09", "IND10"],
}

EXPECTED_INDICATORS = [f"IND{i:02d}" for i in range(1, 11)]
REQUIRED_COLUMNS = ["Respondent_ID"] + EXPECTED_INDICATORS

OUTPUT_COLUMNS = [
    "Respondent_ID",
    "D1", "D2", "D3", "D4",
    "BDSF_equal",
    "BDSF_indicator",
]


def fail(message: str) -> None:
    raise ValueError(message)


def get_paths():
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent

    input_path = (
        project_root
        / "02_Scoring_Matrix"
        / "02_01_Kazakhstan_Cybersecurity_Scored_Dataset.csv"
    )

    dimensions_output = script_dir / "03_01_Kazakhstan_dimensions.csv"

    return project_root, input_path, dimensions_output


def audit_mapping() -> None:
    """Every expected indicator must occur exactly once."""
    mapped = [
        indicator
        for indicators in DIMENSION_MAPPING.values()
        for indicator in indicators
    ]

    duplicates = sorted({
        indicator for indicator in mapped
        if mapped.count(indicator) > 1
    })
    missing = sorted(set(EXPECTED_INDICATORS) - set(mapped))
    unexpected = sorted(set(mapped) - set(EXPECTED_INDICATORS))

    if duplicates:
        fail(
            "IND -> D mapping audit failed. "
            f"Indicators assigned more than once: {duplicates}"
        )

    if missing:
        fail(
            "IND -> D mapping audit failed. "
            f"Expected indicators without dimension: {missing}"
        )

    if unexpected:
        fail(
            "IND -> D mapping audit failed. "
            f"Unexpected indicators in mapping: {unexpected}"
        )

    if len(mapped) != len(EXPECTED_INDICATORS):
        fail("IND -> D mapping audit failed: indicator count mismatch.")


def load_input(input_path: Path) -> pd.DataFrame:
    """
    Load ONLY the explicitly defined analysis input.
    No fallback search is allowed.
    """
    if not input_path.exists():
        fail(
            "Deterministic input file not found:\n"
            f"{input_path}\n\n"
            "No alternative Kazakhstan file will be selected automatically."
        )

    try:
        return pd.read_csv(input_path)
    except Exception as exc:
        fail(f"Could not read input CSV: {input_path}\nReason: {exc}")


def audit_input(df: pd.DataFrame) -> pd.DataFrame:
    missing_columns = [
        column for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]
    if missing_columns:
        fail(f"Input audit failed. Missing columns: {missing_columns}")

    if len(df) != EXPECTED_N:
        fail(
            f"Input audit failed. Expected N={EXPECTED_N}, "
            f"found N={len(df)}"
        )

    if df["Respondent_ID"].isna().any():
        fail("Input audit failed. Missing Respondent_ID detected.")

    duplicate_mask = df["Respondent_ID"].duplicated(keep=False)
    if duplicate_mask.any():
        duplicate_ids = (
            df.loc[duplicate_mask, "Respondent_ID"]
            .astype(str)
            .tolist()
        )
        fail(
            "Input audit failed. Duplicate Respondent_ID detected: "
            f"{duplicate_ids}"
        )

    # Work on a copy so validation does not mutate an external object.
    df = df.copy()

    df[EXPECTED_INDICATORS] = df[EXPECTED_INDICATORS].apply(
        pd.to_numeric,
        errors="coerce",
    )

    missing_ind_count = int(
        df[EXPECTED_INDICATORS].isna().sum().sum()
    )
    if missing_ind_count:
        fail(
            "Input audit failed. Missing/non-numeric IND values: "
            f"{missing_ind_count}"
        )

    values = df[EXPECTED_INDICATORS].to_numpy(dtype=float)

    if not np.isfinite(values).all():
        fail("Input audit failed. Inf/-Inf detected in IND values.")

    valid_range = df[EXPECTED_INDICATORS].apply(
        lambda column: column.between(0, 100)
    )
    bad_rows = ~valid_range.all(axis=1)

    if bad_rows.any():
        ids = df.loc[bad_rows, "Respondent_ID"].astype(str).tolist()
        fail(
            "Input audit failed. IND score outside [0,100] for "
            f"Respondent_ID: {ids}"
        )

    return df


def calculate_dimensions(df: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame({
        "Respondent_ID": df["Respondent_ID"].copy()
    })

    for dimension, indicators in DIMENSION_MAPPING.items():
        result[dimension] = df[indicators].mean(axis=1)

    # Primary integrated score: equal weighting of conceptual dimensions.
    result["BDSF_equal"] = result[
        ["D1", "D2", "D3", "D4"]
    ].mean(axis=1)

    # Sensitivity score: equal weighting of all indicators.
    result["BDSF_indicator"] = df[
        EXPECTED_INDICATORS
    ].mean(axis=1)

    return result[OUTPUT_COLUMNS]


def audit_output(result: pd.DataFrame) -> None:
    dimensions = ["D1", "D2", "D3", "D4"]
    score_columns = dimensions + [
        "BDSF_equal",
        "BDSF_indicator",
    ]

    if result[dimensions].isna().any().any():
        fail("Output audit failed. Missing dimension score detected.")

    if result[
        ["BDSF_equal", "BDSF_indicator"]
    ].isna().any().any():
        fail("Output audit failed. Missing BDSF score detected.")

    values = result[score_columns].to_numpy(dtype=float)
    if not np.isfinite(values).all():
        fail("Output audit failed. Inf/-Inf detected.")

    for column in score_columns:
        bad = ~result[column].between(0, 100)
        if bad.any():
            ids = (
                result.loc[bad, "Respondent_ID"]
                .astype(str)
                .tolist()
            )
            fail(
                f"Output audit failed. {column} outside [0,100] "
                f"for Respondent_ID: {ids}"
            )

    if result["Respondent_ID"].duplicated().any():
        fail("Output audit failed. Duplicate Respondent_ID detected.")

    if len(result) != EXPECTED_N:
        fail(
            f"Output audit failed. Expected N={EXPECTED_N}, "
            f"found N={len(result)}"
        )


def main() -> None:
    project_root, input_path, dimensions_output = get_paths()

    # 1. Validate the hard-coded conceptual mapping.
    audit_mapping()

    # 2. Deterministically load the one approved analysis input.
    df = load_input(input_path)

    # 3. Validate respondent-level IND data.
    df = audit_input(df)

    # 4. Calculate dimensions and integrated scores.
    result = calculate_dimensions(df)

    # 5. Validate calculated scores.
    audit_output(result)

    # 6. Write output only after all audits pass.
    result.to_csv(
        dimensions_output,
        index=False,
        float_format="%.6f",
    )

    print("=" * 72)
    print("03_01 KAZAKHSTAN DIMENSION SCORING: PASS")
    print("=" * 72)
    print(f"Project root     : {project_root}")
    print(f"Fixed input      : {input_path}")
    print(f"Dimensions output: {dimensions_output}")
    print(f"N                 : {len(result)}")
    print()

    print("Fixed IND -> Dimension mapping:")
    for dimension, indicators in DIMENSION_MAPPING.items():
        print(
            f"  {dimension} = mean({', '.join(indicators)})"
        )

    print()
    print("Audit:")
    print("  Deterministic input path     : PASS")
    print("  IND -> D unique assignment   : PASS")
    print("  All expected IND mapped      : PASS")
    print("  Duplicate Respondent_ID      : 0")
    print("  Missing D                    : 0")
    print("  Missing BDSF                 : 0")
    print("  Inf/-Inf                     : 0")
    print("  Score range [0,100]          : PASS")
    print(f"  Expected N={EXPECTED_N}               : PASS")
    print("=" * 72)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)
