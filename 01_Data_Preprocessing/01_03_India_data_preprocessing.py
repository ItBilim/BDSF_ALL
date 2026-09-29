import re
import pandas as pd

INPUT_FILE = "01_03_India.xlsx"

MATRIX_OUTPUT = "01_03_India_cleaned.csv"
DICTIONARY_OUTPUT = "01_03_India_questions_answers.csv"

RESPONDENT_PREFIX = "IN_R"


def clean_text(value):
    if pd.isna(value):
        return pd.NA

    if isinstance(value, str):
        value = re.sub(r"[\r\n\t]+", " ", value)
        value = re.sub(r"\s+", " ", value).strip()

        if value == "":
            return pd.NA

    return value


df = pd.read_excel(
    INPUT_FILE,
    sheet_name=0
)

df = df.dropna(
    axis=0,
    how="all"
)

df = df.dropna(
    axis=1,
    how="all"
)

df.columns = [
    clean_text(column)
    for column in df.columns
]

df = df.apply(
    lambda column: column.map(clean_text)
)

time_date_columns = {
    "timestamp",
    "time stamp",
    "time",
    "date",
    "datetime",
    "submission time",
    "submission date",
    "created at",
    "updated at",
    "временная метка",
    "отметка времени",
    "время",
    "дата"
}

columns_to_remove = [
    column
    for column in df.columns
    if str(column).strip().lower() in time_date_columns
]

df = df.drop(
    columns=columns_to_remove,
    errors="ignore"
)

df = df.replace(
    r"^\s*$",
    pd.NA,
    regex=True
)

df = df.reset_index(drop=True)

original_questions = list(df.columns)

question_ids = [
    f"Q{i}"
    for i in range(
        1,
        len(original_questions) + 1
    )
]

matrix_df = pd.DataFrame()

dictionary_rows = []

for question_id, question_text in zip(
    question_ids,
    original_questions
):
    answers = df[question_text].tolist()

    unique_answers = [
        answer
        for answer in dict.fromkeys(answers)
        if not pd.isna(answer)
    ]

    answer_mapping = {
        answer: answer_id
        for answer_id, answer in enumerate(
            unique_answers,
            start=1
        )
    }

    matrix_df[question_id] = [
        answer_mapping.get(answer, pd.NA)
        if not pd.isna(answer)
        else pd.NA
        for answer in answers
    ]

    for answer in unique_answers:
        dictionary_rows.append({
            "Question_ID": question_id,
            "Question_Text": question_text,
            "Answer_ID": answer_mapping[answer],
            "Answer_Text": answer
        })

matrix_df.insert(
    0,
    "Respondent_ID",
    [
        f"{RESPONDENT_PREFIX}{i:03d}"
        for i in range(
            1,
            len(matrix_df) + 1
        )
    ]
)

dictionary_df = pd.DataFrame(
    dictionary_rows,
    columns=[
        "Question_ID",
        "Question_Text",
        "Answer_ID",
        "Answer_Text"
    ]
)

matrix_df.to_csv(
    MATRIX_OUTPUT,
    index=False,
    encoding="utf-8-sig"
)

dictionary_df.to_csv(
    DICTIONARY_OUTPUT,
    index=False,
    encoding="utf-8-sig"
)

print(f"Respondents: {len(matrix_df)}")
print(f"Questions: {len(question_ids)}")
print(f"Missing answers: {matrix_df.isna().sum().sum()}")
print(f"Matrix saved: {MATRIX_OUTPUT}")
print(f"Dictionary saved: {DICTIONARY_OUTPUT}")