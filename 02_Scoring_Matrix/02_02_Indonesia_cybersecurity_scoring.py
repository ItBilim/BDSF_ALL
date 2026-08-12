from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd


DEFAULT_CLEANED = Path("01_Data_Preprocessing/01_02_Indonesia_cleaned.csv")
DEFAULT_DICTIONARY = Path("01_Data_Preprocessing/01_02_Indonesia_questions_answers.csv")
DEFAULT_MATRIX = Path("02_Scoring_Matrix/02_02_Indonesia_Cybersecurity_Scoring_Matrix.xlsx")
DEFAULT_OUTPUT = Path("02_Scoring_Matrix/02_02_Indonesia_Cybersecurity_Scored_Dataset.csv")

DERIVED_MULTI_SELECT = {}
MULTI_SELECT_ALIASES = {}

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Score Indonesia cybersecurity survey responses using the scoring matrix."
    )
    parser.add_argument("--cleaned", type=Path, help="Cleaned respondents CSV")
    parser.add_argument("--dictionary", type=Path, help="Questions/answers dictionary CSV")
    parser.add_argument("--matrix", type=Path, help="Scoring matrix XLSX")
    parser.add_argument("--output", type=Path, help="Output scored CSV")
    return parser.parse_args()


def find_project_root(start: Path) -> Path:
    for candidate in [start, *start.parents]:
        if (candidate / "01_Data_Preprocessing").is_dir() and (candidate / "02_Scoring_Matrix").is_dir():
            return candidate
    return start


def resolve_path(value: Path | None, default: Path, project_root: Path) -> Path:
    path = value if value is not None else default
    return path if path.is_absolute() else project_root / path


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("ё", "е").replace("Ё", "Е")
    text = re.sub(r"[«»„“”\"']", "", text)
    text = re.sub(r"[‐‑‒–—―]", "-", text)
    text = re.sub(r"\s+", " ", text).strip(" ,;.")
    return text.casefold()


def normalize_answer_id(value: object) -> str:
    if pd.isna(value):
        return ""
    try:
        number = float(value)
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def require_columns(frame: pd.DataFrame, columns: list[str], file_label: str) -> None:
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise ValueError(f"{file_label}: missing required columns: {missing}")


def read_inputs(cleaned_path: Path, dictionary_path: Path, matrix_path: Path):
    for path in (cleaned_path, dictionary_path, matrix_path):
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

    cleaned = pd.read_csv(cleaned_path)
    dictionary = pd.read_csv(dictionary_path)
    matrix = pd.read_excel(matrix_path, sheet_name="Scoring")

    require_columns(cleaned, ["Respondent_ID"], cleaned_path.name)
    require_columns(dictionary, ["Question_ID", "Answer_ID", "Answer_Text"], dictionary_path.name)
    require_columns(matrix, ["Indicator Code", "Question Code", "Question Type"], matrix_path.name)
    return cleaned, dictionary, matrix


def build_answer_lookup(dictionary: pd.DataFrame) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    for row in dictionary.itertuples(index=False):
        q = str(row.Question_ID).strip()
        a = normalize_answer_id(row.Answer_ID)
        lookup.setdefault(q, {})[a] = str(row.Answer_Text).strip()
    return lookup


def extract_scoring_options(matrix_row: pd.Series) -> list[tuple[str, float]]:
    options = []
    n = 1
    while f"Answer {n}" in matrix_row.index:
        answer = matrix_row[f"Answer {n}"]
        score = matrix_row.get(f"Score {n}")
        if pd.notna(answer) and pd.notna(score):
            options.append((str(answer).strip(), float(score)))
        n += 1
    if not options:
        raise ValueError(f"{matrix_row['Question Code']}: no Answer/Score pairs in matrix")
    return options


def answer_text_from_id(question_code: str, answer_id: object, lookup: dict[str, dict[str, str]]) -> str:
    question_lookup = lookup.get(question_code)
    if not question_lookup:
        raise ValueError(f"Dictionary does not contain {question_code}")
    key = normalize_answer_id(answer_id)
    if key not in question_lookup:
        raise ValueError(f"{question_code}: Answer_ID={answer_id!r} not found in dictionary")
    return question_lookup[key]


def score_single_answer(question_code: str, answer_text: str, options: list[tuple[str, float]]) -> float:
    scoring = {normalize_text(answer): score for answer, score in options}
    key = normalize_text(answer_text)
    if key not in scoring:
        raise ValueError(f"{question_code}: answer not found in scoring matrix: {answer_text!r}")
    return scoring[key]


def score_multi_answer(question_code: str, answer_text: str, options: list[tuple[str, float]]) -> float:
    selected_parts = {
        normalize_text(part) for part in re.split(r"[,;]", answer_text) if normalize_text(part)
    }
    aliases = MULTI_SELECT_ALIASES.get(question_code, {})
    matched = set()
    total = 0.0

    for option_text, score in options:
        option_key = normalize_text(option_text)
        option_aliases = {normalize_text(a) for a in aliases.get(option_text, ())}
        exact = option_key in selected_parts
        alias = any(a in part for a in option_aliases for part in selected_parts)
        if exact or alias:
            matched.add(option_key)
            total += score

    if not matched:
        raise ValueError(f"{question_code}: multi-select answer did not match matrix: {answer_text!r}")
    return min(total, 100.0)


def score_derived_binary(
    derived_code: str,
    source_answer_text: str,
    options: list[tuple[str, float]],
) -> float:
    rule = DERIVED_MULTI_SELECT[derived_code]
    selected = normalize_text(rule["needle"]) in normalize_text(source_answer_text)

    # New matrix defines option 1 = not selected, option 2 = selected.
    if len(options) < 2:
        raise ValueError(f"{derived_code}: derived binary indicator needs two scoring options")
    return options[1][1] if selected else options[0][1]


def score_dataset(cleaned: pd.DataFrame, dictionary: pd.DataFrame, matrix: pd.DataFrame):
    matrix = matrix.copy()
    matrix["Question Code"] = matrix["Question Code"].astype(str).str.strip()
    matrix["Indicator Code"] = matrix["Indicator Code"].astype(str).str.strip()

    question_codes = matrix["Question Code"].tolist()
    indicator_codes = matrix["Indicator Code"].tolist()
    if len(question_codes) != len(set(question_codes)):
        raise ValueError("Scoring matrix contains duplicate Question Code")
    if len(indicator_codes) != len(set(indicator_codes)):
        raise ValueError("Scoring matrix contains duplicate Indicator Code")

    # Physical source columns needed from respondents. Q7a/Q7b are derived from Q7.
    source_questions = []
    for code in question_codes:
        source = DERIVED_MULTI_SELECT.get(code, {}).get("source", code)
        if source not in source_questions:
            source_questions.append(source)
    require_columns(cleaned, source_questions, "Cleaned dataset")

    # Remove a respondent only when a required SOURCE answer is missing.
    answer_cells = cleaned[source_questions].replace(r"^\s*$", pd.NA, regex=True)
    complete_mask = answer_cells.notna().all(axis=1)
    removed_count = int((~complete_mask).sum())
    complete = cleaned.loc[complete_mask, ["Respondent_ID", *source_questions]].copy()

    lookup = build_answer_lookup(dictionary)
    result = pd.DataFrame({"Respondent_ID": complete["Respondent_ID"].astype(str).to_numpy()})

    for _, row in matrix.iterrows():
        question_code = row["Question Code"]
        indicator_code = row["Indicator Code"]
        question_type = normalize_text(row["Question Type"])
        options = extract_scoring_options(row)
        scores = []

        if question_type == "derived_binary_from_multi_select":
            if question_code not in DERIVED_MULTI_SELECT:
                raise ValueError(f"No derivation rule configured for {question_code}")
            source_code = DERIVED_MULTI_SELECT[question_code]["source"]
            for answer_id in complete[source_code]:
                source_text = answer_text_from_id(source_code, answer_id, lookup)
                scores.append(score_derived_binary(question_code, source_text, options))
        else:
            for answer_id in complete[question_code]:
                answer_text = answer_text_from_id(question_code, answer_id, lookup)
                if question_type == "multi_select":
                    score = score_multi_answer(question_code, answer_text, options)
                else:
                    score = score_single_answer(question_code, answer_text, options)
                scores.append(score)

        numeric = pd.Series(scores)
        if not numeric.empty and (numeric % 1 == 0).all():
            numeric = numeric.astype("int64")
        result[indicator_code] = numeric.to_numpy()

    return result, removed_count


def main() -> None:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    project_root = find_project_root(script_dir)

    cleaned_path = resolve_path(args.cleaned, DEFAULT_CLEANED, project_root)
    dictionary_path = resolve_path(args.dictionary, DEFAULT_DICTIONARY, project_root)
    matrix_path = resolve_path(args.matrix, DEFAULT_MATRIX, project_root)
    output_path = resolve_path(args.output, DEFAULT_OUTPUT, project_root)

    cleaned, dictionary, matrix = read_inputs(cleaned_path, dictionary_path, matrix_path)
    result, removed_count = score_dataset(cleaned, dictionary, matrix)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(output_path, index=False, encoding="utf-8-sig")

    print(f"All respondents: {len(cleaned)}")
    print(f"Removed because required answers are incomplete: {removed_count}")
    print(f"Respondents in result: {len(result)}")
    print(f"Scored indicators: {len(result.columns) - 1}")
    print(f"Result saved: {output_path}")


if __name__ == "__main__":
    main()
