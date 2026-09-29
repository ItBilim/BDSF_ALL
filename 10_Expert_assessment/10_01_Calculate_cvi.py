import math
import re
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================
# FILES
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent

INPUT_FILE = SCRIPT_DIR / "10_01_Expert_evaluation_results.xlsx"

# Support downloaded/versioned filenames such as
# "10_01_Expert_evaluation_results(2).xlsx" when the canonical
# filename is not present.
if not INPUT_FILE.exists():
    input_candidates = sorted(
        SCRIPT_DIR.glob("10_01_Expert_evaluation_results*.xlsx")
    )

    if len(input_candidates) == 1:
        INPUT_FILE = input_candidates[0]

# All CSV results are saved in the SAME folder as the Excel file.
OUTPUT_DIR = SCRIPT_DIR


# ============================================================
# OFFICIAL 4-POINT EXPERT SCALE
# ============================================================

TEXT_TO_SCORE = {
    "не подходит": 1,
    "требует серьезной доработки": 2,
    "подходит, но требует незначительной доработки": 3,
    "полностью соответствует": 4,
}


# ============================================================
# FUNCTIONS
# ============================================================

def normalize_text(value):
    if pd.isna(value):
        return None

    text = str(value).strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)

    return text


def convert_to_expert_score(value):
    """
    Convert a response to the official expert rating 1-4.

    Valid:
        1, 2, 3, 4
        or the corresponding text labels.

    Invalid:
        0, 25, 50, 75, 100,
        empty values,
        ambiguous values,
        any other values.
    """

    if pd.isna(value):
        return None

    # Integer
    if isinstance(value, (int, np.integer)):
        value = int(value)
        return value if value in (1, 2, 3, 4) else None

    # Float
    if isinstance(value, (float, np.floating)):
        if float(value).is_integer():
            value = int(value)
            return value if value in (1, 2, 3, 4) else None
        return None

    # Text
    text = normalize_text(value)

    if text in TEXT_TO_SCORE:
        return TEXT_TO_SCORE[text]

    # Numeric text: "1", "2", "3", "4"
    try:
        number = float(text)

        if number.is_integer() and int(number) in (1, 2, 3, 4):
            return int(number)

    except (ValueError, TypeError):
        pass

    return None


def parse_item_metadata(column_name):
    """
    Extract Dimension, Indicator and Question from item header.
    """

    text = str(column_name)

    dimension = ""
    indicator = ""
    question = ""

    dimension_match = re.search(
        r"\b(D\d+)\b",
        text,
        flags=re.IGNORECASE
    )

    indicator_match = re.search(
        r"\b(IND\d+)\b",
        text,
        flags=re.IGNORECASE
    )

    question_match = re.search(
        r"\b(Q\d+)\b",
        text,
        flags=re.IGNORECASE
    )

    if dimension_match:
        dimension = dimension_match.group(1).upper()

    if indicator_match:
        indicator = indicator_match.group(1).upper()

    if question_match:
        question = question_match.group(1).upper()

    return dimension, indicator, question


def calculate_modified_kappa(i_cvi, n_experts, n_agreement):
    """
    Modified kappa for content validity.

    Pc = C(N,A) * 0.5^N

    kappa* = (I-CVI - Pc) / (1 - Pc)
    """

    if n_experts <= 0:
        return np.nan, np.nan

    pc = math.comb(n_experts, n_agreement) * (0.5 ** n_experts)

    if pc == 1:
        return np.nan, pc

    kappa = (i_cvi - pc) / (1 - pc)

    return kappa, pc


# ============================================================
# CHECK INPUT FILE
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"File not found:\n{INPUT_FILE}"
    )


# ============================================================
# READ EXCEL
# ============================================================

df = pd.read_excel(INPUT_FILE)

print("=" * 70)
print("BDSF EXPERT CONTENT VALIDITY ANALYSIS")
print("=" * 70)

print(f"Input file: {INPUT_FILE}")
print(f"Total rows: {len(df)}")
print(f"Total columns: {len(df.columns)}")


# ============================================================
# IDENTIFY 43 EVALUATION ITEMS
# ============================================================

# The BDSF file has 8 metadata columns followed by 43
# expert-evaluation columns.

if len(df.columns) < 51:
    raise ValueError(
        f"The Excel file contains only {len(df.columns)} columns. "
        f"Expected at least 51 columns."
    )

item_columns = list(df.columns[8:])

if len(item_columns) != 43:
    raise ValueError(
        f"Expected 43 evaluation items, "
        f"but found {len(item_columns)}."
    )


# ============================================================
# VALIDATE ALL EXPERT RESPONSES
# ============================================================

expert_records = []

for row_index, row in df.iterrows():

    scores = []
    invalid_values = []

    for column in item_columns:

        original_value = row[column]

        score = convert_to_expert_score(original_value)

        scores.append(score)

        if score is None:
            invalid_values.append(
                f"{column}: {original_value}"
            )

    valid_items = sum(
        score is not None
        for score in scores
    )

    invalid_items = 43 - valid_items

    included = invalid_items == 0

    expert_id = f"Expert_{row_index + 1}"

    if included:
        reason = (
            "Included: all 43 items contain valid "
            "expert ratings from 1 to 4."
        )
    else:
        reason = (
            "Excluded: at least one item contains "
            "an invalid, empty, ambiguous, or non-1-4 response."
        )

    expert_records.append({
        "row_index": row_index,
        "expert_id": expert_id,
        "scores": scores,
        "included": included,
        "valid_items": valid_items,
        "invalid_items": invalid_items,
        "invalid_values": " | ".join(invalid_values),
        "reason": reason,
    })


# ============================================================
# SELECT VALID EXPERTS
# ============================================================

valid_experts = [
    record
    for record in expert_records
    if record["included"]
]

excluded_experts = [
    record
    for record in expert_records
    if not record["included"]
]

N = len(valid_experts)

if N == 0:
    raise ValueError(
        "No valid expert responses were found."
    )


# ============================================================
# VALID EXPERT SCORE MATRIX
# ============================================================

score_matrix = np.array(
    [
        record["scores"]
        for record in valid_experts
    ],
    dtype=int
)

if score_matrix.shape != (N, 43):
    raise ValueError(
        f"Unexpected score matrix shape: "
        f"{score_matrix.shape}. "
        f"Expected ({N}, 43)."
    )


# ============================================================
# ITEM-LEVEL CVI
# ============================================================

item_results = []
binary_results = []

for item_index, column in enumerate(item_columns):

    scores = score_matrix[:, item_index]

    # Agreement = rating 3 or 4.
    agreement = scores >= 3

    n_agreement = int(
        np.sum(agreement)
    )

    i_cvi = n_agreement / N

    kappa, pc = calculate_modified_kappa(
        i_cvi=i_cvi,
        n_experts=N,
        n_agreement=n_agreement
    )

    dimension, indicator, question = parse_item_metadata(
        column
    )

    # D_i = number of experts assigning rating 3:
    # "relevant, but requires minor improvement".
    d_i = int(np.sum(scores == 3))

    ratings_4_item = int(np.sum(scores == 4))

    # The I-CVI rule takes precedence when content relevance is
    # insufficient. Among items meeting the I-CVI threshold, two or
    # more rating-3 responses trigger wording review and re-evaluation.
    if i_cvi < 0.78:
        additional_review = "Yes"
        decision = "Review content validity"
    elif d_i >= 2:
        additional_review = "Yes"
        decision = "Review wording and re-evaluate"
    else:
        additional_review = "No"
        decision = "Retain"

    result = {
        "Item": item_index + 1,
        "Dimension": dimension,
        "Indicator": indicator,
        "Question": question,
    }

    # Dynamic number of experts.
    for expert_position, record in enumerate(valid_experts):
        result[record["expert_id"]] = scores[expert_position]

    result.update({
        "N_experts": N,
        "Agreement_3_or_4": n_agreement,
        "Ratings_3_Item": d_i,
        "Ratings_4_Item": ratings_4_item,
        "D_i": d_i,
        "I_CVI": round(i_cvi, 4),
        "Pc": round(pc, 6),
        "Modified_Kappa": round(kappa, 6),
        "Additional_Review_Required": additional_review,
        "Decision": decision,
    })

    item_results.append(result)

    # Binary matrix:
    # 1 = agreement (3/4)
    # 0 = disagreement (1/2)

    binary_row = {
        "Item": item_index + 1
    }

    for expert_position, record in enumerate(valid_experts):
        binary_row[record["expert_id"]] = int(
            agreement[expert_position]
        )

    binary_results.append(binary_row)


item_results_df = pd.DataFrame(item_results)
binary_matrix_df = pd.DataFrame(binary_results)


# ============================================================
# SCALE-LEVEL CVI
# ============================================================

i_cvi_values = (
    item_results_df["I_CVI"].to_numpy()
)

s_cvi_ave = float(
    np.mean(i_cvi_values)
)

s_cvi_ua = float(
    np.mean(i_cvi_values == 1.0)
)

kappa_values = (
    item_results_df["Modified_Kappa"]
    .to_numpy()
)

di_values = item_results_df["D_i"].to_numpy(dtype=int)

items_di_0 = int(np.sum(di_values == 0))
items_di_1 = int(np.sum(di_values == 1))
items_di_ge_2 = int(np.sum(di_values >= 2))

additional_review_items = (
    item_results_df.loc[
        item_results_df["Additional_Review_Required"] == "Yes",
        "Item"
    ]
    .astype(str)
    .tolist()
)


# ============================================================
# RATING COUNTS
# ============================================================

rating_counts = {
    4: int(np.sum(score_matrix == 4)),
    3: int(np.sum(score_matrix == 3)),
    2: int(np.sum(score_matrix == 2)),
    1: int(np.sum(score_matrix == 1)),
}


# ============================================================
# SUMMARY
# ============================================================

summary_df = pd.DataFrame(
    [
        ["Total_responses", len(df)],
        ["Valid_experts", N],
        ["Excluded_responses", len(excluded_experts)],
        ["Number_of_items", 43],
        [
            "I_CVI_min",
            round(float(np.min(i_cvi_values)), 4)
        ],
        [
            "I_CVI_max",
            round(float(np.max(i_cvi_values)), 4)
        ],
        [
            "I_CVI_mean",
            round(float(np.mean(i_cvi_values)), 4)
        ],
        [
            "S_CVI_Ave",
            round(s_cvi_ave, 4)
        ],
        [
            "S_CVI_UA",
            round(s_cvi_ua, 4)
        ],
        [
            "Modified_Kappa_min",
            round(float(np.min(kappa_values)), 6)
        ],
        [
            "Modified_Kappa_max",
            round(float(np.max(kappa_values)), 6)
        ],
        ["D_i_min", int(np.min(di_values))],
        ["D_i_max", int(np.max(di_values))],
        ["Items_D_i_0", items_di_0],
        ["Items_D_i_1", items_di_1],
        ["Items_D_i_ge_2", items_di_ge_2],
        [
            "Additional_review_items",
            ", ".join(additional_review_items)
        ],
        ["Ratings_4", rating_counts[4]],
        ["Ratings_3", rating_counts[3]],
        ["Ratings_2", rating_counts[2]],
        ["Ratings_1", rating_counts[1]],
    ],
    columns=["Metric", "Value"]
)


# ============================================================
# EXPERT INCLUSION AUDIT
# ============================================================

audit_rows = []

for record in expert_records:

    audit_rows.append({
        "Response_ID": record["expert_id"],
        "Included_in_CVI": (
            "Yes"
            if record["included"]
            else "No"
        ),
        "Valid_items": record["valid_items"],
        "Invalid_items": record["invalid_items"],
        "Invalid_values": record["invalid_values"],
        "Exclusion_reason": record["reason"],
    })

audit_df = pd.DataFrame(audit_rows)


# ============================================================
# SAVE RESULTS IN THE SAME FOLDER
# ============================================================

item_output = OUTPUT_DIR / "CVI_item_results.csv"
summary_output = OUTPUT_DIR / "CVI_summary.csv"
audit_output = OUTPUT_DIR / "CVI_expert_inclusion_audit.csv"
binary_output = OUTPUT_DIR / "CVI_binary_matrix.csv"

item_results_df.to_csv(
    item_output,
    index=False,
    encoding="utf-8-sig"
)

summary_df.to_csv(
    summary_output,
    index=False,
    encoding="utf-8-sig"
)

audit_df.to_csv(
    audit_output,
    index=False,
    encoding="utf-8-sig"
)

binary_matrix_df.to_csv(
    binary_output,
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# CONSOLE REPORT
# ============================================================

print()
print("RESULTS")
print("-" * 70)

print(
    f"Total responses:       {len(df)}"
)

print(
    f"Valid experts:         {N}"
)

print(
    f"Excluded responses:    {len(excluded_experts)}"
)

print(
    f"Evaluation items:      {len(item_columns)}"
)

print()

print(
    f"I-CVI minimum:         "
    f"{np.min(i_cvi_values):.4f}"
)

print(
    f"I-CVI maximum:         "
    f"{np.max(i_cvi_values):.4f}"
)

print(
    f"I-CVI mean:            "
    f"{np.mean(i_cvi_values):.4f}"
)

print(
    f"S-CVI/Ave:             "
    f"{s_cvi_ave:.4f}"
)

print(
    f"S-CVI/UA:              "
    f"{s_cvi_ua:.4f}"
)

print()

print(
    f"Modified Kappa min:    "
    f"{np.min(kappa_values):.6f}"
)

print(
    f"Modified Kappa max:    "
    f"{np.max(kappa_values):.6f}"
)

print()

print(
    f"Ratings 4:             "
    f"{rating_counts[4]}"
)

print(
    f"Ratings 3:             "
    f"{rating_counts[3]}"
)

print(
    f"Ratings 2:             "
    f"{rating_counts[2]}"
)

print(
    f"Ratings 1:             "
    f"{rating_counts[1]}"
)

print()

print(
    f"Items with D_i = 0:    "
    f"{items_di_0}"
)

print(
    f"Items with D_i = 1:    "
    f"{items_di_1}"
)

print(
    f"Items with D_i >= 2:   "
    f"{items_di_ge_2}"
)

print(
    f"Additional review:     "
    f"{', '.join(additional_review_items) or 'None'}"
)

print()
print("Expert inclusion audit")
print("-" * 70)

for record in expert_records:

    status = (
        "INCLUDED"
        if record["included"]
        else "EXCLUDED"
    )

    print(
        f"{record['expert_id']}: "
        f"{status}; "
        f"valid_items={record['valid_items']}; "
        f"invalid_items={record['invalid_items']}"
    )

print()
print("CSV files created in:")
print(OUTPUT_DIR.resolve())

print()
print("Files:")
print(item_output.name)
print(summary_output.name)
print(audit_output.name)
print(binary_output.name)

print("=" * 70)
