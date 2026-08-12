#!/usr/bin/env python3
"""
04_01 Kazakhstan Descriptive Analysis

Fixed input:
    BDSF_ALL/03_Dimension_Scoring/03_01_Kazakhstan_dimensions.csv

Outputs:
    BDSF_ALL/04_Descriptive_Analysis/04_01_Kazakhstan_descriptive.xlsx
    BDSF_ALL/04_Descriptive_Analysis/04_01_Kazakhstan_Figure_A_Boxplot.png
    BDSF_ALL/04_Descriptive_Analysis/04_01_Kazakhstan_Figure_B_BDSF_Distribution.png
    BDSF_ALL/04_Descriptive_Analysis/04_01_Kazakhstan_Figure_C_Mean_Profile_95CI.png

Statistics for D1-D4 and BDSF_equal:
    N, Mean, SD, Median, Q1, Q3, Min, Max,
    95% CI of Mean, Floor %, Ceiling %, Missing N, Skewness.

95% CI:
    Mean +/- t_(0.975, n-1) * SD / sqrt(n)
"""

from pathlib import Path
import sys
import math

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import t, skew
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


EXPECTED_N = 176
SCORE_COLUMNS = ["D1", "D2", "D3", "D4", "BDSF_equal"]
DIMENSION_COLUMNS = ["D1", "D2", "D3", "D4"]

FLOOR_THRESHOLD = 0.0
CEILING_THRESHOLD = 100.0

# Audit thresholds used only to flag potentially strong distributional effects.
ABS_SKEW_WARNING = 1.0
FLOOR_CEILING_WARNING_PCT = 15.0


def fail(message: str) -> None:
    raise ValueError(message)


def get_paths():
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent

    input_path = (
        project_root
        / "03_Dimension_Scoring"
        / "03_01_Kazakhstan_dimensions.csv"
    )

    output_xlsx = script_dir / "04_01_Kazakhstan_descriptive.xlsx"
    figure_a = script_dir / "04_01_Kazakhstan_Figure_A_Boxplot.png"
    figure_b = script_dir / "04_01_Kazakhstan_Figure_B_BDSF_Distribution.png"
    figure_c = script_dir / "04_01_Kazakhstan_Figure_C_Mean_Profile_95CI.png"

    return project_root, input_path, output_xlsx, figure_a, figure_b, figure_c


def load_input(input_path: Path) -> pd.DataFrame:
    """Load only the fixed official 03-stage output. No fallback search."""
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
    required = ["Respondent_ID"] + SCORE_COLUMNS

    missing_columns = [c for c in required if c not in df.columns]
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
            df.loc[duplicate_mask, "Respondent_ID"].astype(str).tolist()
        )
        fail(
            "Input audit failed. Duplicate Respondent_ID detected: "
            f"{duplicate_ids}"
        )

    df = df.copy()

    df[SCORE_COLUMNS] = df[SCORE_COLUMNS].apply(
        pd.to_numeric, errors="coerce"
    )

    values = df[SCORE_COLUMNS].to_numpy(dtype=float)

    if np.isinf(values).any():
        fail("Input audit failed. Inf/-Inf detected in score columns.")

    # Missing is reported descriptively, but a completely missing variable is invalid.
    all_missing = [c for c in SCORE_COLUMNS if df[c].notna().sum() == 0]
    if all_missing:
        fail(f"Input audit failed. Entirely missing variables: {all_missing}")

    # Impossible values are fatal; missing values are handled with available-case N.
    impossible = pd.DataFrame(
        {
            c: df[c].notna() & ~df[c].between(0, 100)
            for c in SCORE_COLUMNS
        }
    )
    bad_rows = impossible.any(axis=1)

    if bad_rows.any():
        ids = df.loc[bad_rows, "Respondent_ID"].astype(str).tolist()
        fail(
            "Input audit failed. Score outside [0,100] for "
            f"Respondent_ID: {ids}"
        )

    return df


def describe_variable(series: pd.Series) -> dict:
    clean = series.dropna().astype(float)
    n = int(clean.size)
    missing_n = int(series.isna().sum())

    if n < 2:
        fail(f"At least 2 non-missing observations required for {series.name}.")

    mean = float(clean.mean())
    sd = float(clean.std(ddof=1))
    median = float(clean.median())
    q1 = float(clean.quantile(0.25))
    q3 = float(clean.quantile(0.75))
    minimum = float(clean.min())
    maximum = float(clean.max())

    se = sd / math.sqrt(n)
    t_crit = float(t.ppf(0.975, df=n - 1))
    ci_low = mean - t_crit * se
    ci_high = mean + t_crit * se

    floor_pct = float((clean == FLOOR_THRESHOLD).mean() * 100.0)
    ceiling_pct = float((clean == CEILING_THRESHOLD).mean() * 100.0)

    # Bias-corrected Fisher-Pearson sample skewness.
    skewness = float(skew(clean, bias=False)) if n >= 3 else np.nan

    return {
        "Variable": series.name,
        "N": n,
        "Mean": mean,
        "SD": sd,
        "Median": median,
        "Q1": q1,
        "Q3": q3,
        "Min": minimum,
        "Max": maximum,
        "CI_95_Lower": ci_low,
        "CI_95_Upper": ci_high,
        "Floor_pct": floor_pct,
        "Ceiling_pct": ceiling_pct,
        "Missing_N": missing_n,
        "Skewness": skewness,
    }


def calculate_descriptives(df: pd.DataFrame) -> pd.DataFrame:
    rows = [describe_variable(df[column]) for column in SCORE_COLUMNS]
    return pd.DataFrame(rows)


def build_quality_audit(df: pd.DataFrame, stats: pd.DataFrame) -> pd.DataFrame:
    rows = [
        ["Expected N", EXPECTED_N, len(df), "PASS" if len(df) == EXPECTED_N else "FAIL"],
        ["Duplicate Respondent_ID", 0, int(df["Respondent_ID"].duplicated().sum()),
         "PASS" if not df["Respondent_ID"].duplicated().any() else "FAIL"],
        ["Inf/-Inf", 0, int(np.isinf(df[SCORE_COLUMNS].to_numpy(dtype=float)).sum()),
         "PASS" if not np.isinf(df[SCORE_COLUMNS].to_numpy(dtype=float)).any() else "FAIL"],
        ["Impossible values outside [0,100]", 0,
         int(sum((df[c].notna() & ~df[c].between(0, 100)).sum() for c in SCORE_COLUMNS)),
         "PASS"],
        ["Total missing score cells", 0, int(df[SCORE_COLUMNS].isna().sum().sum()),
         "PASS" if not df[SCORE_COLUMNS].isna().any().any() else "REVIEW"],
    ]

    for _, row in stats.iterrows():
        skew_flag = (
            "REVIEW"
            if pd.notna(row["Skewness"]) and abs(row["Skewness"]) >= ABS_SKEW_WARNING
            else "PASS"
        )
        rows.append([
            f"{row['Variable']} | abs(skewness) >= {ABS_SKEW_WARNING}",
            f"< {ABS_SKEW_WARNING}",
            abs(row["Skewness"]) if pd.notna(row["Skewness"]) else np.nan,
            skew_flag,
        ])

        floor_flag = (
            "REVIEW"
            if row["Floor_pct"] >= FLOOR_CEILING_WARNING_PCT
            else "PASS"
        )
        rows.append([
            f"{row['Variable']} | floor effect >= {FLOOR_CEILING_WARNING_PCT:.0f}%",
            f"< {FLOOR_CEILING_WARNING_PCT:.0f}%",
            row["Floor_pct"],
            floor_flag,
        ])

        ceiling_flag = (
            "REVIEW"
            if row["Ceiling_pct"] >= FLOOR_CEILING_WARNING_PCT
            else "PASS"
        )
        rows.append([
            f"{row['Variable']} | ceiling effect >= {FLOOR_CEILING_WARNING_PCT:.0f}%",
            f"< {FLOOR_CEILING_WARNING_PCT:.0f}%",
            row["Ceiling_pct"],
            ceiling_flag,
        ])

    return pd.DataFrame(
        rows,
        columns=["Check", "Expected/Rule", "Observed", "Status"],
    )


def create_figure_a(df: pd.DataFrame, output_path: Path) -> None:
    data = [df[c].dropna().to_numpy() for c in DIMENSION_COLUMNS]

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.boxplot(data, tick_labels=DIMENSION_COLUMNS)
    ax.set_title("Kazakhstan: D1-D4 Score Distributions")
    ax.set_xlabel("Dimension")
    ax.set_ylabel("Score (0-100)")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def create_figure_b(df: pd.DataFrame, output_path: Path) -> None:
    values = df["BDSF_equal"].dropna().to_numpy()

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.hist(values, bins="auto", edgecolor="black", alpha=0.8)
    ax.axvline(np.mean(values), linestyle="--", linewidth=1.5, label="Mean")
    ax.axvline(np.median(values), linestyle=":", linewidth=1.5, label="Median")
    ax.set_title("Kazakhstan: BDSF_equal Distribution")
    ax.set_xlabel("BDSF_equal Score")
    ax.set_ylabel("Frequency")
    ax.set_xlim(0, 100)
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def create_figure_c(stats: pd.DataFrame, output_path: Path) -> None:
    profile = stats[stats["Variable"].isin(DIMENSION_COLUMNS)].copy()
    profile["Variable"] = pd.Categorical(
        profile["Variable"],
        categories=DIMENSION_COLUMNS,
        ordered=True,
    )
    profile = profile.sort_values("Variable")

    means = profile["Mean"].to_numpy()
    lower = means - profile["CI_95_Lower"].to_numpy()
    upper = profile["CI_95_Upper"].to_numpy() - means
    yerr = np.vstack([lower, upper])

    x = np.arange(len(DIMENSION_COLUMNS))

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.errorbar(
        x,
        means,
        yerr=yerr,
        fmt="o",
        capsize=5,
        linewidth=1.5,
    )
    ax.set_xticks(x, DIMENSION_COLUMNS)
    ax.set_title("Kazakhstan: D1-D4 Means with 95% CI")
    ax.set_xlabel("Dimension")
    ax.set_ylabel("Mean Score (0-100)")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def write_excel(stats: pd.DataFrame, audit: pd.DataFrame, output_path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Descriptive_Statistics"

    headers = list(stats.columns)
    ws.append(headers)

    for row in stats.itertuples(index=False, name=None):
        ws.append(list(row))

    audit_ws = wb.create_sheet("Quality_Audit")
    audit_ws.append(list(audit.columns))
    for row in audit.itertuples(index=False, name=None):
        audit_ws.append(list(row))

    notes_ws = wb.create_sheet("Method_Notes")
    notes = [
        ["Item", "Definition"],
        ["Input", "03_01_Kazakhstan_dimensions.csv"],
        ["Variables", "D1, D2, D3, D4, BDSF_equal"],
        ["SD", "Sample standard deviation (ddof=1)"],
        ["95% CI", "Student t interval: Mean +/- t(0.975, N-1) * SD / sqrt(N)"],
        ["Floor effect", "Percentage of non-missing scores == 0"],
        ["Ceiling effect", "Percentage of non-missing scores == 100"],
        ["Skewness", "Bias-corrected Fisher-Pearson sample skewness"],
        ["Skewness review rule", f"|skewness| >= {ABS_SKEW_WARNING}"],
        ["Floor/Ceiling review rule", f">= {FLOOR_CEILING_WARNING_PCT:.0f}%"],
    ]
    for row in notes:
        notes_ws.append(row)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for sheet in [ws, audit_ws, notes_ws]:
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False

        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for column_cells in sheet.columns:
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in column_cells
            )
            width = min(max(max_length + 2, 12), 45)
            sheet.column_dimensions[
                get_column_letter(column_cells[0].column)
            ].width = width

    # Numeric formatting.
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column == 2 or cell.column == 13:  # N, Missing_N
                cell.number_format = "0"
            elif cell.column in (11, 12):  # Floor/Ceiling percentages
                cell.number_format = "0.00"
            elif cell.column > 2:
                cell.number_format = "0.000"

    wb.save(output_path)


def main() -> None:
    (
        project_root,
        input_path,
        output_xlsx,
        figure_a,
        figure_b,
        figure_c,
    ) = get_paths()

    df = load_input(input_path)
    df = audit_input(df)

    stats = calculate_descriptives(df)
    audit = build_quality_audit(df, stats)

    create_figure_a(df, figure_a)
    create_figure_b(df, figure_b)
    create_figure_c(stats, figure_c)

    write_excel(stats, audit, output_xlsx)

    print("=" * 76)
    print("04_01 KAZAKHSTAN DESCRIPTIVE ANALYSIS: PASS")
    print("=" * 76)
    print(f"Project root : {project_root}")
    print(f"Fixed input  : {input_path}")
    print(f"N            : {len(df)}")
    print()
    print("Outputs:")
    print(f"  {output_xlsx}")
    print(f"  {figure_a}")
    print(f"  {figure_b}")
    print(f"  {figure_c}")
    print()
    print("Pre-figure audit:")
    print(f"  Expected N={EXPECTED_N}              : PASS")
    print("  Duplicate Respondent_ID     : 0")
    print("  Impossible values [0,100]   : 0")
    print(f"  Missing score cells         : {int(df[SCORE_COLUMNS].isna().sum().sum())}")
    print("  Inf/-Inf                    : 0")
    print()
    print("Distribution diagnostics:")
    for _, row in stats.iterrows():
        print(
            f"  {row['Variable']}: "
            f"N={int(row['N'])}, "
            f"skew={row['Skewness']:.3f}, "
            f"floor={row['Floor_pct']:.2f}%, "
            f"ceiling={row['Ceiling_pct']:.2f}%"
        )
    print("=" * 76)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)
