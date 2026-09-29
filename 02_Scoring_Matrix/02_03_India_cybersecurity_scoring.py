from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd

DEFAULT_CLEANED = Path('01_Data_Preprocessing/01_03_India_cleaned.csv')
DEFAULT_DICTIONARY = Path('01_Data_Preprocessing/01_03_India_questions_answers.csv')
DEFAULT_MATRIX = Path('02_Scoring_Matrix/02_03_India_Cybersecurity_Scoring_Matrix.xlsx')
DEFAULT_OUTPUT = Path('02_Scoring_Matrix/02_03_India_Cybersecurity_Scored_Dataset.csv')


def parse_args():
    p = argparse.ArgumentParser(description='India cybersecurity scoring')
    p.add_argument('--cleaned', type=Path)
    p.add_argument('--dictionary', type=Path)
    p.add_argument('--matrix', type=Path)
    p.add_argument('--output', type=Path)
    return p.parse_args()


def find_project_root(start: Path) -> Path:
    for p in [start, *start.parents]:
        if (p / '01_Data_Preprocessing').is_dir() and (p / '02_Scoring_Matrix').is_dir():
            return p
    return start


def resolve_path(value, default, root):
    p = value if value is not None else default
    return p if p.is_absolute() else root / p


def norm(v) -> str:
    s = unicodedata.normalize('NFKC', str(v))
    s = re.sub(r'[\u2010-\u2015]', '-', s)
    s = re.sub(r'[\u201c\u201d\u2018\u2019"\']', '', s)
    s = re.sub(r'\s+', ' ', s).strip(' ,;.').casefold()
    return s


def norm_id(v) -> str:
    if pd.isna(v):
        return ''
    try:
        x = float(v)
        if x.is_integer():
            return str(int(x))
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def base_question(code: str) -> str:
    m = re.fullmatch(r'(Q\d+)[a-z]+', code.strip(), flags=re.I)
    return m.group(1) if m else code.strip()


def build_lookup(dictionary):
    out = {}
    for r in dictionary.itertuples(index=False):
        out.setdefault(str(r.Question_ID).strip(), {})[norm_id(r.Answer_ID)] = str(r.Answer_Text).strip()
    return out


def options(row):
    out = []
    n = 1
    while f'Answer {n}' in row.index:
        a, s = row.get(f'Answer {n}'), row.get(f'Score {n}')
        if pd.notna(a) and pd.notna(s):
            out.append((str(a).strip(), float(s)))
        n += 1
    return out


def direct_score(text, opts, code):
    mp = {norm(a): s for a, s in opts}
    k = norm(text)
    if k not in mp:
        raise ValueError(f'{code}: answer not found in matrix: {text!r}')
    return mp[k]


def contains(text: str, token: str) -> bool:
    return norm(token) in norm(text)


def derived_score(qcode, qtype, text, opts):
    # Q13a-f: binary indicators derived from the multi-select Q13 answer.
    if qtype == 'derived_binary_from_multi_select' and qcode.lower().startswith('q13'):
        targets = {
            'q13a': 'Encryption',
            'q13b': 'Anti-virus',
            'q13c': 'Firewall',
            'q13d': 'Software update',
            'q13e': 'Backup',
            'q13f': 'Authentication (e.g. password, PIN)',
        }
        target = targets[qcode.lower()]
        return 100.0 if contains(text, target) else 0.0

    # Q9a: public Wi-Fi exposure is reverse binary.
    if qtype == 'derived_reverse_binary_exposure':
        target = 'Public Wi-Fi (e.g. in coffee shop)'
        return 0.0 if contains(text, target) else 100.0

    # Q26: number of protected device categories; None/I don't know -> 0.
    if qtype == 'derived_coverage_ordinal':
        t = norm(text)
        if t in {norm('None of the above'), norm("I don't know")}:
            return 0.0
        devices = ['Desktop', 'Laptop', 'Smartphone', 'Tablet']
        count = sum(contains(text, d) for d in devices)
        return float(min(count, 4) * 25)

    # Q28/Q29: only "I do not feel..." -> 0; any other nonblank source -> 100.
    if qtype == 'derived_binary_from_multi_select':
        only_none = norm('I do not feel that I keep myself updated')
        parts = [norm(x) for x in text.split(',') if norm(x)]
        other = [x for x in parts if x != only_none]
        return 100.0 if other else 0.0

    return direct_score(text, opts, qcode)


def score_dataset(cleaned, dictionary, matrix):
    lookup = build_lookup(dictionary)
    required = []
    for code in matrix['Question Code'].astype(str):
        b = base_question(code)
        if b not in required:
            required.append(b)
    missing = [q for q in required if q not in cleaned.columns]
    if missing:
        raise ValueError(f'Cleaned dataset missing columns: {missing}')

    cells = cleaned[required].replace(r'^\s*$', pd.NA, regex=True)
    mask = cells.notna().all(axis=1)
    removed = int((~mask).sum())
    complete = cleaned.loc[mask, ['Respondent_ID', *required]].copy()
    result = pd.DataFrame({'Respondent_ID': complete['Respondent_ID'].astype(str)})

    for _, row in matrix.iterrows():
        qcode = str(row['Question Code']).strip()
        bq = base_question(qcode)
        indicator = str(row['Indicator Code']).strip()
        qtype = norm(row['Question Type']).replace(' ', '_')
        opts = options(row)
        if bq not in lookup:
            raise ValueError(f'{bq}: question missing from dictionary')

        vals = []
        for answer_id in complete[bq]:
            key = norm_id(answer_id)
            if key not in lookup[bq]:
                raise ValueError(f'{bq}: Answer_ID={answer_id!r} missing from dictionary')
            text = lookup[bq][key]
            vals.append(derived_score(qcode, qtype, text, opts))
        s = pd.Series(vals)
        if (s % 1 == 0).all():
            s = s.astype('int64')
        result[indicator] = s.to_numpy()
    return result, removed


def main():
    args = parse_args()
    root = find_project_root(Path(__file__).resolve().parent)
    cleaned_path = resolve_path(args.cleaned, DEFAULT_CLEANED, root)
    dictionary_path = resolve_path(args.dictionary, DEFAULT_DICTIONARY, root)
    matrix_path = resolve_path(args.matrix, DEFAULT_MATRIX, root)
    output_path = resolve_path(args.output, DEFAULT_OUTPUT, root)

    cleaned = pd.read_csv(cleaned_path)
    dictionary = pd.read_csv(dictionary_path)
    matrix = pd.read_excel(matrix_path, sheet_name='Scoring')
    result, removed = score_dataset(cleaned, dictionary, matrix)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f'All respondents: {len(cleaned)}')
    print(f'Removed incomplete respondents: {removed}')
    print(f'Remaining respondents: {len(result)}')
    print(f'Scored indicators: {len(result.columns)-1}')
    print(f'Saved: {output_path}')


if __name__ == '__main__':
    main()
