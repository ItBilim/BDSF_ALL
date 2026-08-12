#!/usr/bin/env python3
"""
07_03 India External Validation

The India cleaned dataset contains coded Q1-Q63 variables, while the official
scoring matrix documents only questions used in BDSF construction. For unused
questions, no questionnaire/answer metadata is available in the deterministic
inputs supplied to this stage.

Therefore this script performs the required leakage/candidate audit but does
NOT infer the meaning of unused numeric question codes and does NOT select a
criterion based on statistical significance.

If no independently interpretable external criterion can be established from
the supplied metadata, the scientifically correct result is:

    Independent external criterion unavailable.

Deterministic inputs:
    BDSF_ALL/01_Data_Preprocessing/01_03_India_cleaned.csv
    BDSF_ALL/02_Scoring_Matrix/02_03_India_Cybersecurity_Scoring_Matrix.xlsx
    BDSF_ALL/03_Dimension_Scoring/03_03_India_dimensions.csv

Output:
    BDSF_ALL/07_External_Validation/07_03_India_external_validation.xlsx
"""

from pathlib import Path
import sys
import re

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


EXPECTED_N = 586
STATUS_TEXT = "Independent external criterion unavailable."


def fail(message: str) -> None:
    raise ValueError(message)


def get_paths():
    script_dir = Path(__file__).resolve().parent
    root = script_dir.parent

    cleaned = root / "01_Data_Preprocessing" / "01_03_India_cleaned.csv"
    matrix = (
        root / "02_Scoring_Matrix"
        / "02_03_India_Cybersecurity_Scoring_Matrix.xlsx"
    )
    dimensions = (
        root / "03_Dimension_Scoring"
        / "03_03_India_dimensions.csv"
    )
    output = script_dir / "07_03_India_external_validation.xlsx"

    return root, cleaned, matrix, dimensions, output


def require_file(path: Path, label: str) -> None:
    if not path.exists():
        fail(
            f"{label} not found:\n{path}\n"
            "No alternative file will be selected automatically."
        )


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )
    return df


def load_inputs(cleaned_path, matrix_path, dimensions_path):
    require_file(cleaned_path, "India cleaned dataset")
    require_file(matrix_path, "India official scoring matrix")
    require_file(dimensions_path, "India official dimension dataset")

    cleaned = normalize_headers(pd.read_csv(cleaned_path))
    matrix = normalize_headers(
        pd.read_excel(matrix_path, sheet_name="Scoring")
    )
    dimensions = normalize_headers(pd.read_csv(dimensions_path))

    return cleaned, matrix, dimensions


def base_question(question_code):
    """
    Convert derived codes such as Q13a and Q9a to their source questions
    Q13 and Q9 for leakage auditing.
    """
    text = str(question_code).strip()
    match = re.match(r"^(Q\d+)", text)
    return match.group(1) if match else text


def audit_scoring_matrix(matrix):
    required = {
        "Indicator Code",
        "Question Code",
        "Question Type",
        "Survey Question Text (Original Language)",
    }
    missing = sorted(required - set(matrix.columns))
    if missing:
        fail(
            f"Scoring matrix schema error. Missing columns: {missing}. "
            f"Actual columns: {matrix.columns.tolist()}"
        )

    expected_indicators = {f"IND{i:02d}" for i in range(1, 20)}
    observed_indicators = set(
        matrix["Indicator Code"].astype(str).str.strip()
    )

    if observed_indicators != expected_indicators:
        fail(
            "Indicator-set audit failed. "
            f"Missing={sorted(expected_indicators-observed_indicators)}, "
            f"Extra={sorted(observed_indicators-expected_indicators)}"
        )

    if matrix["Indicator Code"].duplicated().any():
        fail("Duplicate Indicator Code in scoring matrix.")

    used_source_questions = {
        base_question(q)
        for q in matrix["Question Code"].dropna()
    }

    leakage_rows = []
    for _, row in matrix.iterrows():
        leakage_rows.append({
            "Indicator": row["Indicator Code"],
            "Matrix_Question_Code": row["Question Code"],
            "Source_Question": base_question(row["Question Code"]),
            "Question_Type": row["Question Type"],
            "Used_in_BDSF": "Yes",
            "Question_Text": row[
                "Survey Question Text (Original Language)"
            ],
        })

    return used_source_questions, pd.DataFrame(leakage_rows)


def audit_and_align(cleaned, dimensions):
    for name, df in [("cleaned", cleaned), ("dimensions", dimensions)]:
        if "Respondent_ID" not in df.columns:
            fail(f"{name}: Respondent_ID missing.")
        if df["Respondent_ID"].isna().any():
            fail(f"{name}: missing Respondent_ID.")
        if df["Respondent_ID"].duplicated().any():
            fail(f"{name}: duplicate Respondent_ID.")

    if len(dimensions) != EXPECTED_N:
        fail(
            f"Official dimensions expected N={EXPECTED_N}, "
            f"found N={len(dimensions)}"
        )

    if "BDSF_equal" not in dimensions.columns:
        fail("Dimensions dataset missing BDSF_equal.")

    dimensions = dimensions.copy()
    dimensions["BDSF_equal"] = pd.to_numeric(
        dimensions["BDSF_equal"], errors="coerce"
    )

    if dimensions["BDSF_equal"].isna().any():
        fail("Missing/non-numeric BDSF_equal.")
    if not np.isfinite(
        dimensions["BDSF_equal"].to_numpy(float)
    ).all():
        fail("Inf/-Inf in BDSF_equal.")
    if not dimensions["BDSF_equal"].between(0, 100).all():
        fail("BDSF_equal outside [0,100].")

    aligned = dimensions[["Respondent_ID", "BDSF_equal"]].merge(
        cleaned,
        on="Respondent_ID",
        how="left",
        validate="one_to_one",
    )

    if len(aligned) != EXPECTED_N:
        fail(
            f"Respondent_ID alignment expected N={EXPECTED_N}, "
            f"found N={len(aligned)}"
        )

    return aligned


def build_candidate_audit(cleaned, used_source_questions):
    """
    Inventory all cleaned Q variables.

    Unused variables are NOT declared eligible merely because they are outside
    BDSF. Without questionnaire metadata their construct meaning, measurement
    scale, category ordering, and theoretical direction cannot be established
    reproducibly.
    """
    q_columns = [
        c for c in cleaned.columns
        if re.fullmatch(r"Q\d+", str(c))
    ]

    rows = []
    for q in sorted(
        q_columns,
        key=lambda x: int(x[1:]),
    ):
        used = q in used_source_questions

        if used:
            eligibility = "No"
            reason = (
                "Used directly or as source of a derived indicator in BDSF; "
                "external validation would cause data leakage."
            )
            metadata_status = "Documented in scoring matrix"
        else:
            eligibility = "Not assessable"
            reason = (
                "Not used in BDSF, but the supplied deterministic inputs do "
                "not provide questionnaire/answer metadata for this unused "
                "question. The construct, response ordering and expected "
                "direction cannot be established without guessing."
            )
            metadata_status = (
                "External-question metadata unavailable in supplied inputs"
            )

        rows.append({
            "Country": "India",
            "External_variable": q,
            "Used_in_BDSF": "Yes" if used else "No",
            "Questionnaire_Metadata_Available": metadata_status,
            "Why_theoretically_relevant": (
                "Cannot be established reproducibly from supplied metadata"
                if not used
                else "Not applicable because variable is part of BDSF"
            ),
            "Expected_direction": (
                "Cannot be pre-specified from supplied metadata"
                if not used
                else "Not applicable"
            ),
            "Eligible_for_external_validation": eligibility,
            "Reason": reason,
        })

    return pd.DataFrame(rows)


def build_result(candidate_audit):
    eligible = candidate_audit[
        candidate_audit[
            "Eligible_for_external_validation"
        ] == "Yes"
    ]

    if not eligible.empty:
        fail(
            "Internal logic error: an eligible criterion exists but no "
            "pre-specified statistical analysis was defined."
        )

    return pd.DataFrame([{
        "Country": "India",
        "External_Validation_Status": STATUS_TEXT,
        "Statistical_Analysis_Performed": "No",
        "Reason": (
            "Unused cleaned variables exist, but their questionnaire/answer "
            "metadata is not available in the supplied deterministic inputs. "
            "Selecting or ordering a variable would require unsupported "
            "assumptions and could introduce post-hoc selection bias."
        ),
        "Data_Leakage_Avoided": "Yes",
        "Artificial_Variable_Created": "No",
        "Variable_Selected_by_p_value": "No",
    }])


def build_quality_audit(
    cleaned,
    dimensions,
    aligned,
    candidate_audit,
):
    used_n = int(
        (candidate_audit["Used_in_BDSF"] == "Yes").sum()
    )
    unused_n = int(
        (candidate_audit["Used_in_BDSF"] == "No").sum()
    )

    return pd.DataFrame([
        [
            "Official dimension N",
            EXPECTED_N,
            len(dimensions),
            "PASS",
        ],
        [
            "Aligned validation N",
            EXPECTED_N,
            len(aligned),
            "PASS",
        ],
        [
            "Duplicate Respondent_ID dimensions",
            0,
            int(dimensions["Respondent_ID"].duplicated().sum()),
            "PASS",
        ],
        [
            "Duplicate Respondent_ID cleaned",
            0,
            int(cleaned["Respondent_ID"].duplicated().sum()),
            "PASS",
        ],
        [
            "BDSF_equal within [0,100]",
            "Required",
            "Verified",
            "PASS",
        ],
        [
            "Cleaned Q variables used in BDSF",
            "Audited",
            used_n,
            "PASS",
        ],
        [
            "Cleaned Q variables outside BDSF",
            "Audited",
            unused_n,
            "PASS",
        ],
        [
            "Unused-variable questionnaire metadata",
            "Required before criterion selection",
            "Unavailable in supplied deterministic inputs",
            "REVIEW - no external test performed",
        ],
        [
            "External variable selected by p-value",
            "No",
            "No",
            "PASS",
        ],
        [
            "Artificial external variable",
            "No",
            "No",
            "PASS",
        ],
        [
            "Unsupported ordinal coding",
            "No",
            "No",
            "PASS",
        ],
    ], columns=[
        "Check",
        "Expected/Rule",
        "Observed",
        "Status",
    ])


def write_dataframe(ws, df):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))


def format_workbook(wb):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for column_cells in ws.columns:
            max_length = max(
                len(str(cell.value))
                if cell.value is not None
                else 0
                for cell in column_cells
            )
            ws.column_dimensions[
                get_column_letter(column_cells[0].column)
            ].width = min(max(max_length + 2, 12), 70)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.0000"


def write_excel(
    output_path,
    result,
    candidate_audit,
    leakage_audit,
    quality_audit,
):
    wb = Workbook()

    ws = wb.active
    ws.title = "External_Validation"
    write_dataframe(ws, result)

    ws = wb.create_sheet("Candidate_Audit")
    write_dataframe(ws, candidate_audit)

    ws = wb.create_sheet("BDSF_Leakage_Audit")
    write_dataframe(ws, leakage_audit)

    notes = pd.DataFrame([
        [
            "External-validation status",
            STATUS_TEXT,
        ],
        [
            "Reason",
            "No unused question can be interpreted as an independent external "
            "criterion from the supplied cleaned codes and scoring-matrix "
            "metadata alone.",
        ],
        [
            "Leakage rule",
            "Any question used directly or as the source of a derived IND is "
            "ineligible for external validation.",
        ],
        [
            "Selection rule",
            "External variables must be selected from theory and documented "
            "questionnaire meaning before inferential testing.",
        ],
        [
            "Statistical analysis",
            "Not performed because no independently interpretable criterion "
            "was available.",
        ],
        [
            "Artificial variable",
            "Not created.",
        ],
        [
            "Post-hoc p-value selection",
            "Not performed.",
        ],
        [
            "What would enable analysis",
            "A questionnaire/question-answer metadata file documenting the "
            "meaning and response scale of unused India Q variables.",
        ],
    ], columns=["Item", "Definition"])

    ws = wb.create_sheet("Method_Notes")
    write_dataframe(ws, notes)

    ws = wb.create_sheet("Quality_Audit")
    write_dataframe(ws, quality_audit)

    format_workbook(wb)
    wb.save(output_path)


def main():
    (
        root,
        cleaned_path,
        matrix_path,
        dimensions_path,
        output_path,
    ) = get_paths()

    cleaned, matrix, dimensions = load_inputs(
        cleaned_path,
        matrix_path,
        dimensions_path,
    )

    used_source_questions, leakage_audit = audit_scoring_matrix(
        matrix
    )

    aligned = audit_and_align(cleaned, dimensions)

    candidate_audit = build_candidate_audit(
        cleaned,
        used_source_questions,
    )

    result = build_result(candidate_audit)

    qa = build_quality_audit(
        cleaned,
        dimensions,
        aligned,
        candidate_audit,
    )

    write_excel(
        output_path,
        result,
        candidate_audit,
        leakage_audit,
        qa,
    )

    print("=" * 82)
    print("07_03 INDIA EXTERNAL VALIDATION: PASS")
    print("=" * 82)
    print(f"Project root : {root}")
    print(f"Cleaned input: {cleaned_path}")
    print(f"Matrix input : {matrix_path}")
    print(f"BDSF input   : {dimensions_path}")
    print(f"Output       : {output_path}")
    print(f"Cleaned N    : {len(cleaned)}")
    print(f"Official N   : {len(dimensions)}")
    print(f"Aligned N    : {len(aligned)}")
    print()
    print(STATUS_TEXT)
    print(
        "No statistical external-validation test was performed because "
        "unused-question metadata is unavailable."
    )
    print(
        "No artificial variable, arbitrary ordinal coding, or p-value-based "
        "criterion selection was used."
    )
    print("=" * 82)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)
