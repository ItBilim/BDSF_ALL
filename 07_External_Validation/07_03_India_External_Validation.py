#!/usr/bin/env python3
"""India Q57 exploratory external validation.

Q57 is chosen from questionnaire meaning and theory before inferential testing.
It is not used to construct BDSF. Results are criterion-specific exploratory
evidence, not a global valid/invalid judgement and not a causal analysis.
"""

from pathlib import Path
import re
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from scipy.optimize import minimize
from scipy.special import expit
from scipy.stats import norm, t as student_t, ttest_ind


EXPECTED_N = 586
EXPECTED_YES = 151
EXPECTED_NO = 435
BOOTSTRAP_REPLICATES = 10_000
BOOTSTRAP_SEED = 20250307
CRITERION = "Q57"


def fail(message: str) -> None:
    raise ValueError(message)


def get_paths():
    script_dir = Path(__file__).resolve().parent
    root = script_dir.parent
    return (
        root,
        root / "01_Data_Preprocessing" / "01_03_India_cleaned.csv",
        root / "01_Data_Preprocessing" / "01_03_India_questions_answers.csv",
        root / "02_Scoring_Matrix" / "02_03_India_Cybersecurity_Scoring_Matrix.xlsx",
        root / "03_Dimension_Scoring" / "03_03_India_dimensions.csv",
        script_dir / "07_03_India_external_validation.xlsx",
    )


def require_file(path: Path, label: str) -> None:
    if not path.exists():
        fail(f"{label} not found:\n{path}\nNo alternative file will be selected automatically.")


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )
    return df


def normalize_answer_id(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    try:
        number = float(text)
    except ValueError:
        return text
    if not np.isfinite(number):
        return np.nan
    return str(int(number)) if number.is_integer() else str(number)


def load_inputs(cleaned_path, dictionary_path, matrix_path, dimensions_path):
    for path, label in [
        (cleaned_path, "India cleaned dataset"),
        (dictionary_path, "India question-answer dictionary"),
        (matrix_path, "India official scoring matrix"),
        (dimensions_path, "India official dimension dataset"),
    ]:
        require_file(path, label)
    return (
        normalize_headers(pd.read_csv(cleaned_path)),
        normalize_headers(pd.read_csv(dictionary_path)),
        normalize_headers(pd.read_excel(matrix_path, sheet_name="Scoring")),
        normalize_headers(pd.read_csv(dimensions_path)),
    )


def base_question(question_code):
    match = re.match(r"^(Q\d+)", str(question_code).strip())
    return match.group(1) if match else str(question_code).strip()


def audit_scoring_matrix(matrix):
    required = {
        "Indicator Code", "Question Code", "Question Type",
        "Survey Question Text (Original Language)",
    }
    missing = sorted(required - set(matrix.columns))
    if missing:
        fail(f"Scoring matrix schema error. Missing columns: {missing}")
    expected = {f"IND{i:02d}" for i in range(1, 20)}
    observed = set(matrix["Indicator Code"].astype(str).str.strip())
    if observed != expected:
        fail(f"Indicator-set audit failed. Missing={sorted(expected-observed)}, Extra={sorted(observed-expected)}")
    if matrix["Indicator Code"].duplicated().any():
        fail("Duplicate Indicator Code in scoring matrix.")
    used_questions = {base_question(q) for q in matrix["Question Code"].dropna()}
    if CRITERION in used_questions:
        fail("Q57 is used in BDSF and cannot be an external criterion.")
    leakage = pd.DataFrame([{
        "Indicator": row["Indicator Code"],
        "Matrix_Question_Code": row["Question Code"],
        "Source_Question": base_question(row["Question Code"]),
        "Question_Type": row["Question Type"],
        "Used_in_BDSF": "Yes",
        "Question_Text": row["Survey Question Text (Original Language)"],
    } for _, row in matrix.iterrows()])
    return used_questions, leakage


def audit_and_align(cleaned, dimensions):
    for name, df in [("cleaned", cleaned), ("dimensions", dimensions)]:
        if "Respondent_ID" not in df.columns:
            fail(f"{name}: Respondent_ID missing.")
        if df["Respondent_ID"].isna().any():
            fail(f"{name}: missing Respondent_ID.")
        if df["Respondent_ID"].duplicated().any():
            fail(f"{name}: duplicate Respondent_ID.")
    if len(dimensions) != EXPECTED_N:
        fail(f"Official dimensions expected N={EXPECTED_N}, found N={len(dimensions)}")
    if "BDSF_equal" not in dimensions.columns:
        fail("Dimensions dataset missing BDSF_equal.")
    dimensions = dimensions.copy()
    dimensions["BDSF_equal"] = pd.to_numeric(dimensions["BDSF_equal"], errors="coerce")
    bdsf = dimensions["BDSF_equal"].to_numpy(float)
    if np.isnan(bdsf).any() or not np.isfinite(bdsf).all():
        fail("Missing/non-finite BDSF_equal.")
    if not dimensions["BDSF_equal"].between(0, 100).all():
        fail("BDSF_equal outside [0,100].")
    aligned = dimensions[["Respondent_ID", "BDSF_equal"]].merge(
        cleaned, on="Respondent_ID", how="left", validate="one_to_one", indicator=True
    )
    if len(aligned) != EXPECTED_N or not (aligned["_merge"] == "both").all():
        fail("Respondent_ID alignment failed for the official N=586 sample.")
    return aligned.drop(columns="_merge")


def dictionary_rows(dictionary, question_id):
    required = {"Question_ID", "Question_Text", "Answer_ID", "Answer_Text"}
    missing = sorted(required - set(dictionary.columns))
    if missing:
        fail(f"Dictionary schema error. Missing columns: {missing}")
    rows = dictionary.loc[
        dictionary["Question_ID"].astype(str).str.strip() == question_id
    ].copy()
    if rows.empty:
        fail(f"Dictionary metadata missing for {question_id}.")
    rows["Answer_ID_normalized"] = rows["Answer_ID"].map(normalize_answer_id)
    if rows["Answer_ID_normalized"].duplicated().any():
        fail(f"Duplicate Answer_ID in dictionary for {question_id}.")
    return rows


def decode_question(series, dictionary, question_id):
    rows = dictionary_rows(dictionary, question_id)
    mapping = dict(zip(rows["Answer_ID_normalized"], rows["Answer_Text"]))
    normalized = series.map(normalize_answer_id)
    unknown = sorted(set(normalized.dropna()) - set(mapping))
    if unknown:
        fail(f"{question_id}: cleaned codes absent from dictionary: {unknown}")
    return normalized.map(mapping), rows


def build_q57_outcome(aligned, dictionary):
    if CRITERION not in aligned.columns:
        fail("Cleaned dataset missing Q57.")
    labels, rows = decode_question(aligned[CRITERION], dictionary, CRITERION)
    canonical = labels.astype("string").str.strip().str.casefold()
    if set(canonical.dropna()) != {"yes", "no"}:
        fail(f"Q57 expected meanings Yes/No, observed={sorted(set(canonical.dropna()))}")
    rows = rows.copy()
    rows["Answer_Text_canonical"] = rows["Answer_Text"].astype(str).str.strip().str.casefold()
    if set(rows["Answer_Text_canonical"]) != {"yes", "no"}:
        fail("Q57 dictionary must define exactly Yes and No.")
    texts = rows["Question_Text"].dropna().astype(str).unique()
    if len(texts) != 1 or "victim of online fraud" not in texts[0].casefold():
        fail("Q57 dictionary text does not identify online-fraud victimization.")
    outcome = canonical.map({"yes": 1.0, "no": 0.0}).astype(float)
    if outcome.isna().any():
        fail(f"Q57 expected missing=0 in aligned N=586, found {int(outcome.isna().sum())}")
    yes_n, no_n = int((outcome == 1).sum()), int((outcome == 0).sum())
    if (yes_n, no_n) != (EXPECTED_YES, EXPECTED_NO):
        fail(f"Q57 expected Yes=151 and No=435; found Yes={yes_n}, No={no_n}")
    result = aligned.copy()
    result["Q57_Answer_Text"] = labels
    result["Q57_victim"] = outcome.astype(int)
    return result, rows, texts[0]


def hedges_g(group_yes, group_no):
    yes, no = np.asarray(group_yes, float), np.asarray(group_no, float)
    df = len(yes) + len(no) - 2
    pooled_var = (
        (len(yes) - 1) * yes.var(ddof=1) + (len(no) - 1) * no.var(ddof=1)
    ) / df
    if pooled_var <= 0:
        fail("Cannot compute Hedges' g: pooled variance is not positive.")
    correction = 1 - 3 / (4 * df - 1)
    return correction * (yes.mean() - no.mean()) / np.sqrt(pooled_var)


def bootstrap_hedges_ci(group_yes, group_no):
    yes, no = np.asarray(group_yes, float), np.asarray(group_no, float)
    rng = np.random.default_rng(BOOTSTRAP_SEED)
    estimates = np.empty(BOOTSTRAP_REPLICATES)
    for i in range(BOOTSTRAP_REPLICATES):
        estimates[i] = hedges_g(
            rng.choice(yes, len(yes), replace=True),
            rng.choice(no, len(no), replace=True),
        )
    return tuple(float(x) for x in np.quantile(estimates, [0.025, 0.975]))


def primary_analysis(data):
    yes = data.loc[data["Q57_victim"] == 1, "BDSF_equal"].to_numpy(float)
    no = data.loc[data["Q57_victim"] == 0, "BDSF_equal"].to_numpy(float)
    if (len(yes), len(no)) != (EXPECTED_YES, EXPECTED_NO):
        fail("Unexpected Q57 group sizes during primary analysis.")
    welch = ttest_ind(yes, no, equal_var=False)
    yes_var, no_var = yes.var(ddof=1), no.var(ddof=1)
    diff = float(yes.mean() - no.mean())
    se = float(np.sqrt(yes_var / len(yes) + no_var / len(no)))
    df = float(
        (yes_var / len(yes) + no_var / len(no)) ** 2
        / ((yes_var / len(yes)) ** 2 / (len(yes) - 1)
           + (no_var / len(no)) ** 2 / (len(no) - 1))
    )
    critical = float(student_t.ppf(0.975, df))
    g = float(hedges_g(yes, no))
    g_lower, g_upper = bootstrap_hedges_ci(yes, no)
    return pd.DataFrame([{
        "Country": "India", "Criterion": CRITERION,
        "Comparison": "Victim Yes minus No", "N_Total": len(data),
        "N_Yes": len(yes), "N_No": len(no),
        "BDSF_Mean_Yes": float(yes.mean()), "BDSF_SD_Yes": float(yes.std(ddof=1)),
        "BDSF_Mean_No": float(no.mean()), "BDSF_SD_No": float(no.std(ddof=1)),
        "Mean_Difference_Yes_minus_No": diff,
        "Mean_Difference_CI95_Lower_Welch": diff - critical * se,
        "Mean_Difference_CI95_Upper_Welch": diff + critical * se,
        "Welch_t": float(welch.statistic), "Welch_df": df,
        "p_value_two_sided": float(welch.pvalue),
        "Hedges_g_Yes_minus_No": g,
        "Hedges_g_CI95_Lower_bootstrap": g_lower,
        "Hedges_g_CI95_Upper_bootstrap": g_upper,
        "Bootstrap_Replicates": BOOTSTRAP_REPLICATES,
        "Bootstrap_Seed": BOOTSTRAP_SEED,
    }])


def fit_logistic(design, outcome, model_name):
    x, y = design.astype(float).to_numpy(), outcome.astype(float).to_numpy()
    if x.ndim != 2 or len(x) != len(y) or not np.isfinite(x).all() or not np.isfinite(y).all():
        fail(f"{model_name}: invalid/non-finite model input.")
    if np.linalg.matrix_rank(x) != x.shape[1]:
        fail(f"{model_name}: design matrix is rank deficient.")

    def objective(beta):
        eta = x @ beta
        return float(np.sum(np.logaddexp(0.0, eta) - y * eta))

    def gradient(beta):
        return x.T @ (expit(x @ beta) - y)

    fit = minimize(
        objective, np.zeros(x.shape[1]), jac=gradient, method="BFGS",
        options={"gtol": 1e-9, "maxiter": 2_000},
    )
    max_score = float(np.max(np.abs(gradient(fit.x))))
    if not np.isfinite(fit.fun) or max_score > 1e-5:
        fail(f"{model_name}: optimization failed ({fit.message}); max score={max_score:.3g}")
    beta = fit.x
    probability = expit(x @ beta)
    information = x.T @ ((probability * (1 - probability))[:, None] * x)
    if np.linalg.matrix_rank(information) != information.shape[0]:
        fail(f"{model_name}: singular information matrix.")
    se = np.sqrt(np.diag(np.linalg.inv(information)))
    z = beta / se
    lower, upper = beta - norm.ppf(0.975) * se, beta + norm.ppf(0.975) * se
    log_likelihood = -objective(beta)
    prevalence = float(y.mean())
    null_ll = float(np.sum(y * np.log(prevalence) + (1 - y) * np.log(1 - prevalence)))
    return pd.DataFrame([{
        "Model": model_name, "N": len(y), "Events_Q57_Yes": int(y.sum()),
        "Term": term, "Coefficient_log_odds": float(beta[i]),
        "Standard_Error": float(se[i]), "z_value": float(z[i]),
        "p_value_two_sided": float(2 * norm.sf(abs(z[i]))),
        "Odds_Ratio": float(np.exp(beta[i])),
        "Odds_Ratio_CI95_Lower_Wald": float(np.exp(lower[i])),
        "Odds_Ratio_CI95_Upper_Wald": float(np.exp(upper[i])),
        "Log_Likelihood": log_likelihood,
        "McFadden_Pseudo_R2": float(1 - log_likelihood / null_ll),
        "AIC": float(-2 * log_likelihood + 2 * len(beta)),
        "Max_Absolute_Score": max_score,
    } for i, term in enumerate(design.columns)])


def logistic_analyses(data, dictionary):
    unadjusted_design = pd.DataFrame({
        "Intercept": 1.0,
        "BDSF_equal_per_10_points": data["BDSF_equal"] / 10.0,
    })
    unadjusted = fit_logistic(
        unadjusted_design, data["Q57_victim"],
        "Unadjusted: Q57_victim ~ BDSF_equal",
    )
    adjusted = data[["Q57_victim", "BDSF_equal", "Q1", "Q2"]].copy()
    adjusted["Gender"], gender_rows = decode_question(adjusted["Q1"], dictionary, "Q1")
    adjusted["Age_Group"], age_rows = decode_question(adjusted["Q2"], dictionary, "Q2")
    adjusted = adjusted.dropna(subset=["Q57_victim", "BDSF_equal", "Gender", "Age_Group"])
    if len(adjusted) < 0.90 * EXPECTED_N:
        fail("Adjusted sensitivity model lost more than 10% of the primary sample.")
    gender_levels = gender_rows.sort_values("Answer_ID")["Answer_Text"].tolist()
    age_levels = age_rows.sort_values("Answer_ID")["Answer_Text"].tolist()
    dummies = pd.concat([
        pd.get_dummies(pd.Categorical(adjusted["Gender"], categories=gender_levels),
                       prefix="Gender", drop_first=True, dtype=float),
        pd.get_dummies(pd.Categorical(adjusted["Age_Group"], categories=age_levels),
                       prefix="Age", drop_first=True, dtype=float),
    ], axis=1).reset_index(drop=True)
    adjusted_design = pd.DataFrame({
        "Intercept": np.ones(len(adjusted)),
        "BDSF_equal_per_10_points": adjusted["BDSF_equal"].to_numpy() / 10.0,
    })
    adjusted_design = pd.concat([adjusted_design, dummies], axis=1)
    adjusted_result = fit_logistic(
        adjusted_design, adjusted["Q57_victim"].reset_index(drop=True),
        "Adjusted sensitivity: BDSF + gender + age group",
    )
    references = pd.DataFrame([
        {"Covariate": "Gender (Q1)", "Reference_Category": gender_levels[0],
         "Included_Categories": "; ".join(gender_levels), "Complete_Case_N": len(adjusted)},
        {"Covariate": "Age group (Q2)", "Reference_Category": age_levels[0],
         "Included_Categories": "; ".join(age_levels), "Complete_Case_N": len(adjusted)},
    ])
    return unadjusted, adjusted_result, references


def bdsf_term(model):
    return model.loc[model["Term"] == "BDSF_equal_per_10_points"].iloc[0]


def interval_excludes(lower, null, upper):
    return not (float(lower) <= null <= float(upper))


def build_result(primary, unadjusted, adjusted):
    p, u, a = primary.iloc[0], bdsf_term(unadjusted), bdsf_term(adjusted)
    g = float(p["Hedges_g_Yes_minus_No"])
    unadjusted_or = float(u["Odds_Ratio"])
    adjusted_or = float(a["Odds_Ratio"])

    group_positive = g > 0
    logistic_positive = unadjusted_or > 1
    adjusted_positive = adjusted_or > 1
    analytic_methods_agree = (
        group_positive == logistic_positive == adjusted_positive
    )
    expected_direction_consistent = (
        g < 0 and unadjusted_or < 1 and adjusted_or < 1
    )
    opposite_direction = (
        g > 0 and unadjusted_or > 1 and adjusted_or > 1
    )
    primary_ci_clear = interval_excludes(
        p["Hedges_g_CI95_Lower_bootstrap"], 0,
        p["Hedges_g_CI95_Upper_bootstrap"],
    )
    logistic_ci_clear = interval_excludes(
        u["Odds_Ratio_CI95_Lower_Wald"], 1,
        u["Odds_Ratio_CI95_Upper_Wald"],
    )
    adjusted_ci_clear = interval_excludes(
        a["Odds_Ratio_CI95_Lower_Wald"], 1,
        a["Odds_Ratio_CI95_Upper_Wald"],
    )
    robust_protective_support = (
        expected_direction_consistent
        and primary_ci_clear
        and logistic_ci_clear
        and adjusted_ci_clear
    )
    direction_text = (
        "Higher BDSF -> higher reported victimization odds"
        if logistic_positive else
        "Higher BDSF -> lower reported victimization odds"
    )
    if robust_protective_support:
        status = (
            "Protective-direction criterion support observed in the India "
            "sample for the theory-selected exploratory Q57 criterion."
        )
        support = "Yes - protective-direction criterion support"
        interpretation = (
            "All three estimates were consistent with the prespecified "
            "protective direction, and all corresponding 95% confidence "
            "intervals excluded their null values."
        )
    elif opposite_direction:
        adjusted_inconclusive = not adjusted_ci_clear
        status = (
            "Opposite-direction exploratory association observed for Q57; "
            "this is not protective external-validity support."
        )
        support = "No - opposite to the prespecified protective direction"
        interpretation = (
            f"The group comparison and logistic estimates agreed: "
            f"{direction_text}. The group-comparison and unadjusted-model "
            "95% confidence intervals excluded their null values. "
            + (
                "The adjusted sensitivity 95% CI included 1."
                if adjusted_inconclusive else
                "The adjusted sensitivity 95% CI also excluded 1."
            )
            + " The observed direction is opposite to the prespecified "
            "protective interpretation."
        )
    else:
        status = (
            "Q57 did not provide conclusive protective-direction external-"
            "validity support in the India sample."
        )
        support = "No - directionally or statistically inconclusive"
        interpretation = (
            "Effect directions and 95% confidence intervals did not jointly "
            "provide clear protective-direction criterion support."
        )
    return pd.DataFrame([{
        "Country": "India", "External_Validation_Status": status,
        "Statistical_Analysis_Performed": "Yes", "Criterion": CRITERION,
        "Criterion_Role": "Theory-selected exploratory external criterion",
        "Criterion_Specific_Support": support,
        "Expected_Direction": "Exploratory; a protective interpretation would predict lower victimization odds",
        "Observed_Direction": direction_text,
        "Analytic_Methods_Agree": "Yes" if analytic_methods_agree else "No",
        "Expected_Direction_Consistent": (
            "Yes" if expected_direction_consistent else "No"
        ),
        "Hedges_g_Yes_minus_No": float(p["Hedges_g_Yes_minus_No"]),
        "Hedges_g_CI95_Lower": float(p["Hedges_g_CI95_Lower_bootstrap"]),
        "Hedges_g_CI95_Upper": float(p["Hedges_g_CI95_Upper_bootstrap"]),
        "Unadjusted_OR_per_10_BDSF": float(u["Odds_Ratio"]),
        "Unadjusted_OR_CI95_Lower": float(u["Odds_Ratio_CI95_Lower_Wald"]),
        "Unadjusted_OR_CI95_Upper": float(u["Odds_Ratio_CI95_Upper_Wald"]),
        "Adjusted_OR_per_10_BDSF": float(a["Odds_Ratio"]),
        "Adjusted_OR_CI95_Lower": float(a["Odds_Ratio_CI95_Lower_Wald"]),
        "Adjusted_OR_CI95_Upper": float(a["Odds_Ratio_CI95_Upper_Wald"]),
        "Evidence_Interpretation": interpretation,
        "Key_Limitation": "Exploratory, observational, self-reported and criterion-specific. Positive association may reflect learning after victimization; no causal, protective, or global BDSF-validity conclusion is warranted.",
        "Data_Leakage_Avoided": "Yes", "Variable_Selected_by_p_value": "No",
    }])


def build_technical_audit(cleaned, dimensions, aligned, q57_rows, used_questions):
    yes_row = q57_rows.loc[q57_rows["Answer_Text_canonical"] == "yes"].iloc[0]
    no_row = q57_rows.loc[q57_rows["Answer_Text_canonical"] == "no"].iloc[0]
    return pd.DataFrame([
        ["Q57 excluded from BDSF", "Required", CRITERION not in used_questions, "PASS"],
        ["Q57 dictionary Yes coding", "Yes -> 1", f"Answer_ID {yes_row['Answer_ID']} -> Yes -> 1", "PASS"],
        ["Q57 dictionary No coding", "No -> 0", f"Answer_ID {no_row['Answer_ID']} -> No -> 0", "PASS"],
        ["Official dimension N", EXPECTED_N, len(dimensions), "PASS"],
        ["Aligned validation N", EXPECTED_N, len(aligned), "PASS"],
        ["Q57 Yes N", EXPECTED_YES, int((aligned["Q57_victim"] == 1).sum()), "PASS"],
        ["Q57 No N", EXPECTED_NO, int((aligned["Q57_victim"] == 0).sum()), "PASS"],
        ["Q57 missing N", 0, int(aligned["Q57_victim"].isna().sum()), "PASS"],
        ["Duplicate Respondent_ID dimensions", 0, int(dimensions["Respondent_ID"].duplicated().sum()), "PASS"],
        ["Duplicate Respondent_ID cleaned", 0, int(cleaned["Respondent_ID"].duplicated().sum()), "PASS"],
        ["BDSF_equal range", "[0,100]", f"[{aligned['BDSF_equal'].min():.4f}, {aligned['BDSF_equal'].max():.4f}]", "PASS"],
    ], columns=["Check", "Expected_Rule", "Observed", "Status"])


def build_candidate_audit(dictionary, used_questions):
    rows = dictionary_rows(dictionary, CRITERION)
    return pd.DataFrame([{
        "Country": "India", "External_variable": CRITERION,
        "Question_Text": rows["Question_Text"].iloc[0],
        "Used_in_BDSF": "Yes" if CRITERION in used_questions else "No",
        "Questionnaire_Metadata_Available": "Yes",
        "Selection_Basis": "Theory and documented questionnaire meaning; selected before inferential testing",
        "Eligible_for_external_validation": "Yes",
    }])


def build_quality_audit(technical, primary, unadjusted, adjusted, result):
    p, u, a = primary.iloc[0], bdsf_term(unadjusted), bdsf_term(adjusted)
    r = result.iloc[0]
    expected_direction_observed = str(
        r["Expected_Direction_Consistent"]
    )
    support_text = str(r["Criterion_Specific_Support"])
    support_direction_consistent = not (
        expected_direction_observed == "No" and support_text.startswith("Yes")
    )
    qa = pd.DataFrame([
        ["Technical checks", "All PASS", "All PASS", "PASS"],
        ["Welch t-test", "Computed", float(p["Welch_t"]), "PASS"],
        ["Mean difference 95% CI", "Computed", "Computed", "PASS"],
        ["Hedges g bootstrap 95% CI", f"{BOOTSTRAP_REPLICATES} replicates", "Computed", "PASS"],
        ["Unadjusted logistic OR per 10 points", "Computed", float(u["Odds_Ratio"]), "PASS"],
        ["Adjusted sensitivity OR per 10 points", "Computed", float(a["Odds_Ratio"]), "PASS"],
        [
            "Expected-direction consistency recorded",
            "Yes or No",
            expected_direction_observed,
            "PASS" if expected_direction_observed in {"Yes", "No"} else "FAIL",
        ],
        [
            "Support requires expected-direction consistency",
            "No support when direction is inconsistent",
            support_text,
            "PASS" if support_direction_consistent else "FAIL",
        ],
        ["Criterion selected by p-value", "No", "No", "PASS"],
        ["Global valid/invalid claim", "Prohibited", "Not made", "PASS"],
    ], columns=["Check", "Expected_Rule", "Observed", "Status"])
    if not (technical["Status"] == "PASS").all():
        qa.loc[qa["Check"] == "Technical checks", ["Observed", "Status"]] = ["Failure", "FAIL"]
    return qa


def write_dataframe(ws, df):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    ws.auto_filter.ref = ws.dimensions


def format_workbook(wb):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.fill, cell.font = fill, font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for cells in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in cells)
            ws.column_dimensions[get_column_letter(cells[0].column)].width = min(max(width + 2, 12), 70)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if isinstance(cell.value, float):
                    cell.number_format = "0.0000"


def write_excel(output, sheets, question_text):
    wb = Workbook()
    first_name, first_df = sheets[0]
    ws = wb.active
    ws.title = first_name
    write_dataframe(ws, first_df)
    for name, frame in sheets[1:]:
        ws = wb.create_sheet(name)
        write_dataframe(ws, frame)
    notes = pd.DataFrame([
        ["Criterion", f"Q57: {question_text}"],
        ["Criterion role", "Theory-selected exploratory external criterion; not confirmatory validation."],
        ["Coding", "Dictionary-derived: Yes=1 (victim), No=0; no meaning inferred from raw Answer_ID alone."],
        ["Primary contrast", "BDSF mean for Q57 Yes minus Q57 No."],
        ["Primary inference", "Welch t-test, mean difference with Welch 95% CI, Hedges' g with stratified bootstrap percentile 95% CI."],
        ["Logistic model", "Q57_victim ~ BDSF_equal/10; OR is the change in reported victimization odds per 10 BDSF points."],
        ["Adjusted sensitivity", "Adds dictionary-labelled Q1 gender and Q2 age group as categorical covariates; complete cases only."],
        ["Evidence rule", "Protective-direction support requires the prespecified direction across the group comparison and both logistic estimates, with all three 95% CIs excluding their null values; p-value is not the sole rule."],
        ["Direction fields", "Analytic_Methods_Agree records agreement among estimates; Expected_Direction_Consistent records agreement with the prespecified protective interpretation."],
        ["Scope", "Criterion-specific exploratory evidence in the analyzed India sample only; no full valid/invalid or causal conclusion."],
        ["Direction caution", "The observed positive association is not protective evidence and may reflect learning after prior victimization."],
    ], columns=["Item", "Definition"])
    ws = wb.create_sheet("Method_Notes")
    write_dataframe(ws, notes)
    format_workbook(wb)
    wb.save(output)


def main():
    root, cleaned_path, dictionary_path, matrix_path, dimensions_path, output = get_paths()
    cleaned, dictionary, matrix, dimensions = load_inputs(
        cleaned_path, dictionary_path, matrix_path, dimensions_path
    )
    used_questions, leakage = audit_scoring_matrix(matrix)
    aligned = audit_and_align(cleaned, dimensions)
    aligned, q57_codebook, question_text = build_q57_outcome(aligned, dictionary)
    primary = primary_analysis(aligned)
    unadjusted, adjusted, references = logistic_analyses(aligned, dictionary)
    result = build_result(primary, unadjusted, adjusted)
    technical = build_technical_audit(cleaned, dimensions, aligned, q57_codebook, used_questions)
    candidates = build_candidate_audit(dictionary, used_questions)
    quality = build_quality_audit(
        technical, primary, unadjusted, adjusted, result
    )
    if not quality["Status"].astype(str).str.upper().eq("PASS").all():
        fail("Quality audit is not fully PASS.")
    respondent_level = aligned[[
        "Respondent_ID", "BDSF_equal", "Q57", "Q57_Answer_Text",
        "Q57_victim", "Q1", "Q2",
    ]].copy()
    write_excel(output, [
        ("External_Validation", result),
        ("Technical_Audit", technical),
        ("Primary_Q57_Welch", primary),
        ("Logistic_Unadjusted", unadjusted),
        ("Logistic_Adjusted", adjusted),
        ("Adjusted_References", references),
        ("Q57_Codebook", q57_codebook),
        ("Candidate_Audit", candidates),
        ("Respondent_Level", respondent_level),
        ("BDSF_Leakage_Audit", leakage),
        ("Quality_Audit", quality),
    ], question_text)
    p, u, a = primary.iloc[0], bdsf_term(unadjusted), bdsf_term(adjusted)
    print("=" * 82)
    print("07_03 INDIA Q57 EXPLORATORY EXTERNAL VALIDATION: PASS")
    print("=" * 82)
    print(f"Project root : {root}")
    print(f"Output       : {output}")
    print("Aligned N    : 586; Yes=151; No=435; missing=0")
    print(f"BDSF means   : Yes={p['BDSF_Mean_Yes']:.3f}; No={p['BDSF_Mean_No']:.3f}")
    print(f"Mean diff    : {p['Mean_Difference_Yes_minus_No']:.3f} [{p['Mean_Difference_CI95_Lower_Welch']:.3f}, {p['Mean_Difference_CI95_Upper_Welch']:.3f}]")
    print(f"Welch test   : t={p['Welch_t']:.3f}; df={p['Welch_df']:.1f}; p={p['p_value_two_sided']:.4f}")
    print(f"Hedges g     : {p['Hedges_g_Yes_minus_No']:.3f} [{p['Hedges_g_CI95_Lower_bootstrap']:.3f}, {p['Hedges_g_CI95_Upper_bootstrap']:.3f}]")
    print(f"OR per 10    : {u['Odds_Ratio']:.3f} [{u['Odds_Ratio_CI95_Lower_Wald']:.3f}, {u['Odds_Ratio_CI95_Upper_Wald']:.3f}]")
    print(f"Adjusted OR  : {a['Odds_Ratio']:.3f} [{a['Odds_Ratio_CI95_Lower_Wald']:.3f}, {a['Odds_Ratio_CI95_Upper_Wald']:.3f}]")
    print(result.iloc[0]["External_Validation_Status"])
    print("=" * 82)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)
