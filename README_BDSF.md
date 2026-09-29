# BDSF — Behavioural Digital Safety Framework

A reproducible multi-stage analysis pipeline for harmonising heterogeneous cybersecurity/digital-safety survey data into a common behavioural scoring framework.

This repository implements the full **Behavioural Digital Safety Framework (BDSF)** workflow for three national survey datasets: **Kazakhstan, Indonesia, and India**. The pipeline converts raw questionnaire responses into harmonised 0–100 indicator scores, aggregates them into four conceptual dimensions, evaluates alternative scoring and aggregation specifications, performs criterion-specific external validation, produces cross-country descriptive profiles, and calculates expert-based content-validity indices.

> **Interpretation boundary:** BDSF is used here as a harmonisation and scoring framework. The national survey instruments are heterogeneous and measurement invariance/psychometric equivalence across countries has not been established. Therefore, country means in this repository describe the analysed samples and methodological sensitivity; they should **not** be interpreted as a ranking of countries by digital-safety behaviour.

---

## 1. Repository overview

The project is organised as a deterministic 10-stage pipeline containing **24 executable Python scripts**.

```text
BDSF_ALL/
├── 00_Run_All.py
│
├── 01_Data_Preprocessing/
│   ├── 01_01_Kazakhstan_data_preprocessing.py
│   ├── 01_02_Indonesia_data_preprocessing.py
│   └── 01_03_India_data_preprocessing.py
│
├── 02_Scoring_Matrix/
│   ├── 02_01_Kazakhstan_cybersecurity_scoring.py
│   ├── 02_02_Indonesia_cybersecurity_scoring.py
│   └── 02_03_India_cybersecurity_scoring.py
│
├── 03_Dimension_Scoring/
│   ├── 03_01_Kazakhstan_Dimension_Scoring.py
│   ├── 03_02_Indonesia_Dimension_Scoring.py
│   └── 03_03_India_Dimension_Scoring.py
│
├── 04_Descriptive_Analysis/
├── 05_Scoring_Robustness/
├── 06_Weighting_Robustness/
├── 07_External_Validation/
├── 08_Cross_Country_Analysis/
├── 09_Final_Summary/
└── 10_Expert_assessment/
```

The root runner `00_Run_All.py` executes all stages sequentially and stops immediately if any required script returns a non-zero exit status.

---

## 2. Scientific objective

The pipeline addresses the following methodological problem:

> How can heterogeneous national cybersecurity-behaviour questionnaires be mapped to a common behavioural framework while preserving the original item semantics, scoring logic, and country-specific survey structure?

The implementation follows a hierarchy:

```text
Raw survey response
        ↓
Question-level response coding
        ↓
BDSF indicator score (0–100)
        ↓
Four conceptual dimensions (D1–D4)
        ↓
Composite BDSF score
        ↓
Robustness + external validation + expert content validation
```

The framework therefore does not require the three countries to have identical questionnaires or identical numbers of indicators. Instead, country-specific items are mapped to the same four higher-level conceptual dimensions.

---

## 3. Data included in the current repository

| Country | Respondents (N) | BDSF indicators | Stage-03 output |
|---|---:|---:|---|
| Kazakhstan | 176 | 10 | `03_01_Kazakhstan_dimensions.csv` |
| Indonesia | 342 | 14 | `03_02_Indonesia_dimensions.csv` |
| India | 586 | 19 | `03_03_India_dimensions.csv` |
| **Combined descriptive sample** | **1,104** | country-specific | Stage 08 |

The number of indicators differs by survey. This is intentional and reflects the heterogeneous source instruments.

---

## 4. BDSF dimension mapping

### Kazakhstan

```text
D1 = IND01, IND02, IND03
D2 = IND04, IND05, IND06
D3 = IND07, IND08
D4 = IND09, IND10
```

### Indonesia

```text
D1 = IND01, IND02, IND03, IND04
D2 = IND05, IND06, IND07
D3 = IND08, IND09, IND10, IND11
D4 = IND12, IND13, IND14
```

### India

```text
D1 = IND01, IND02, IND03
D2 = IND04, IND05, IND06, IND07
D3 = IND08, IND09, IND10, IND11, IND12, IND13, IND14, IND15, IND16
D4 = IND17, IND18, IND19
```

Each Stage-03 script explicitly audits that every expected indicator is assigned **exactly once** and that no indicator is missing, duplicated, or unexpectedly introduced.

---

## 5. Composite score definitions

All indicators and dimensions are represented on a **0–100** scale.

For respondent \(i\), dimension \(d\) is calculated as the arithmetic mean of the indicators assigned to that dimension:

\[
D_{i,d}=\frac{1}{m_d}\sum_{j \in d} IND_{i,j}
\]

where \(m_d\) is the number of indicators mapped to dimension \(d\).

### Primary composite: `BDSF_equal`

The primary BDSF score gives equal weight to the four conceptual dimensions:

\[
BDSF_{equal,i}=\frac{D_{i,1}+D_{i,2}+D_{i,3}+D_{i,4}}{4}
\]

### Sensitivity composite: `BDSF_indicator`

The alternative aggregation gives equal weight to all country-specific indicators:

\[
BDSF_{indicator,i}=\frac{1}{J}\sum_{j=1}^{J} IND_{i,j}
\]

where \(J\) is 10 for Kazakhstan, 14 for Indonesia, and 19 for India.

The distinction is important: `BDSF_equal` weights the **four dimensions equally**, whereas `BDSF_indicator` weights **all indicators equally**. When dimensions contain different numbers of indicators, the two scores need not be identical.

---

## 6. Pipeline stages

### Stage 01 — Data preprocessing

Directory: `01_Data_Preprocessing/`

Purpose:

- load the original national survey file;
- remove completely empty rows and columns;
- normalise text formatting;
- remove recognised timestamp/date fields;
- generate stable respondent identifiers;
- encode observed answer categories using question-specific answer IDs;
- export a clean response matrix and a question/answer dictionary.

Outputs for each country:

```text
*_cleaned.csv
*_questions_answers.csv
```

Respondent ID prefixes:

```text
Kazakhstan → KZ_R...
Indonesia  → ID_R...
India      → IN_R...
```

### Stage 02 — Indicator scoring

Directory: `02_Scoring_Matrix/`

Purpose:

- map cleaned questionnaire responses to BDSF indicators;
- apply country-specific scoring rules;
- support direct, multi-answer, and derived indicator logic where defined;
- generate the official 0–100 indicator dataset;
- retain an auditable scoring matrix.

Outputs:

```text
02_01_Kazakhstan_Cybersecurity_Scored_Dataset.csv
02_02_Indonesia_Cybersecurity_Scored_Dataset.csv
02_03_India_Cybersecurity_Scored_Dataset.csv
```

and country-specific `*_Cybersecurity_Scoring_Matrix.xlsx` files.

### Stage 03 — Dimension and composite scoring

Directory: `03_Dimension_Scoring/`

Purpose:

- audit the IND→dimension mapping;
- verify expected respondent count and required indicator columns;
- calculate D1–D4;
- calculate `BDSF_equal`;
- calculate `BDSF_indicator`;
- reject missing, infinite, or out-of-range scores.

Output schema:

```text
Respondent_ID
D1
D2
D3
D4
BDSF_equal
BDSF_indicator
```

### Stage 04 — Descriptive analysis

Directory: `04_Descriptive_Analysis/`

Purpose:

- calculate descriptive statistics for D1–D4 and `BDSF_equal`;
- inspect distributional characteristics;
- calculate confidence intervals;
- audit floor/ceiling effects and skewness;
- export publication-oriented figures.

Generated figures per country:

```text
Figure_A_Boxplot.png
Figure_B_BDSF_Distribution.png
Figure_C_Mean_Profile_95CI.png
```

The current Stage-04 quality audits pass all implemented checks for:

- expected N;
- duplicate respondent IDs;
- missing values;
- infinite values;
- scores outside [0,100];
- predefined skewness diagnostics;
- predefined floor/ceiling diagnostics.

### Stage 05 — Scoring robustness

Directory: `05_Scoring_Robustness/`

Purpose:

Evaluate sensitivity to an alternative **equal-spacing ordinal coding** specification.

Important interpretation:

- **Kazakhstan:** Scheme B changes relevant ordinal coding and can therefore be compared with the official Scheme A.
- **Indonesia and India:** the official relevant ordinal scoring already uses equal spacing. In these datasets, Scheme A and Scheme B are mathematically identical; therefore this analysis is an audit of equivalence, not independent evidence from a genuinely different scoring model.

Current Kazakhstan result for `BDSF_equal`:

```text
Pearson r       = 0.9989
Spearman rho    = 0.9980
Mean Δ (B−A)    = -0.5642
Median Δ (B−A)  = -0.5556
```

### Stage 06 — Weighting robustness

Directory: `06_Weighting_Robustness/`

Purpose:

Compare:

```text
BDSF_equal       = equal weight to D1–D4
BDSF_indicator   = equal weight to all indicators
```

Current results:

| Country | Pearson r | Spearman rho | Mean Δ (`indicator − equal`) | Mean absolute respondent difference |
|---|---:|---:|---:|---:|
| Kazakhstan | 0.9864 | 0.9850 | +0.5003 | 1.8355 |
| Indonesia | 0.9990 | 0.9986 | -0.0838 | 0.5093 |
| India | 0.9614 | 0.9533 | +1.0347 | 3.2323 |

These statistics quantify sensitivity; the code does not use an arbitrary correlation threshold to declare the framework “valid” or “invalid”.

### Stage 07 — External validation

Directory: `07_External_Validation/`

Purpose:

Evaluate **criterion-specific** associations between BDSF and external questionnaire variables that are not used to construct BDSF.

The scripts explicitly audit criterion leakage before analysis.

#### Kazakhstan

Primary criterion: `Q6` — previous cyber victimisation.

```text
Hedges g = -0.299
95% CI   = [-0.629, 0.015]
p         = 0.0715
```

Interpretation used by the pipeline: directionally plausible but inconclusive because the confidence interval includes zero.

A secondary exploratory association with `Q11` is also reported.

#### Indonesia

Primary criterion: `Q85`.

```text
Spearman rho = 0.525
95% bootstrap CI = [0.438, 0.604]
p ≈ 1.37 × 10^-25
```

This is treated as criterion-specific observational support, not as universal predictive validity.

#### India

Exploratory criterion: `Q57`.

```text
Hedges g = 0.190
95% CI   = [0.009, 0.372]

Unadjusted OR per +10 BDSF points = 1.153
95% CI                            = [1.003, 1.326]

Adjusted OR per +10 BDSF points   = 1.139
95% CI                            = [0.988, 1.314]
```

The observed direction is opposite to the prespecified protective direction. The repository therefore records this as an **opposite-direction exploratory association**, not as protective external-validity support. The adjusted association is inconclusive.

### Stage 08 — Cross-country descriptive profile

Directory: `08_Cross_Country_Analysis/`

Purpose:

- combine country-level D1–D4 and `BDSF_equal` outputs;
- audit score direction and formula consistency;
- produce sample-level cross-country descriptive summaries;
- generate a profile figure with confidence intervals.

The script explicitly avoids an IND-level cross-country comparison and does not generate a “most cybersecure country” conclusion.

Current combined N:

```text
1,104 respondents
```

### Stage 09 — Final summary

Directory: `09_Final_Summary/`

Outputs:

```text
09_Final_BDSF_Results.csv
09_Final_BDSF_Results.xlsx
```

Current descriptive means:

| Country | N | D1 | D2 | D3 | D4 | BDSF mean |
|---|---:|---:|---:|---:|---:|---:|
| Kazakhstan | 176 | 56.02 | 71.80 | 50.26 | 67.56 | 61.41 |
| Indonesia | 342 | 65.28 | 63.94 | 62.81 | 66.50 | 64.63 |
| India | 586 | 51.59 | 78.67 | 62.94 | 56.31 | 62.38 |

These values describe the analysed datasets. They are **not** evidence that one country has a higher population-level level of digital-safety behaviour than another.

### Stage 10 — Expert content-validity assessment

Directory: `10_Expert_assessment/`

Purpose:

- validate expert responses against the official 4-point relevance scale;
- exclude incomplete/invalid expert-response rows from CVI calculations;
- calculate item-level CVI (`I-CVI`);
- calculate scale-level CVI (`S-CVI/Ave`, `S-CVI/UA`);
- calculate modified kappa;
- retain a transparent expert inclusion audit;
- identify items receiving rating 3 (“minor revision”) from multiple experts for additional review.

Current audit:

```text
Total submitted responses = 7
Valid experts             = 4
Excluded responses        = 3
Evaluated items           = 43

I-CVI range               = 1.000–1.000
S-CVI/Ave                 = 1.000
S-CVI/UA                  = 1.000
Modified kappa range      = 1.000–1.000

Ratings = 4               = 152
Ratings = 3               = 20
Ratings = 2               = 0
Ratings = 1               = 0

Additional review items   = 3, 6, 14
```

The perfect CVI values apply to the **four expert responses that passed the script's completeness/validity criteria**. Three other submitted rows were excluded because they did not contain valid 1–4 ratings for all 43 items. This distinction should be retained when reporting the result.

---

## 7. Software requirements

The code uses only standard Python modules plus the following third-party packages:

```text
pandas
numpy
scipy
matplotlib
openpyxl
```

Recommended environment:

```text
Python 3.10+
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it on Windows:

```bash
.venv\Scripts\activate
```

Activate it on Linux/macOS:

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install pandas numpy scipy matplotlib openpyxl
```

---

## 8. Running the complete pipeline

From the project root:

```bash
python 00_Run_All.py
```

The runner executes the 24 scripts in dependency order.

Execution logic:

```text
01 preprocessing
      ↓
02 indicator scoring
      ↓
03 dimension/composite scoring
      ↓
04 descriptive analysis
      ↓
05 scoring sensitivity
      ↓
06 weighting sensitivity
      ↓
07 external validation
      ↓
08 cross-country descriptive profile
      ↓
09 final summary
      ↓
10 expert content-validity analysis
```

Each script is executed with its own directory as the working directory because several scripts intentionally use local relative filenames.

If any script fails, `00_Run_All.py` stops immediately and reports the failing stage and exit code.

---

## 9. Running a single stage

Example:

```bash
cd 06_Weighting_Robustness
python 06_01_Kazakhstan_Weighting_Robustness.py
```

Most later-stage scripts resolve inputs relative to the project structure, so the directory names should be preserved.

---

## 10. Reproducibility and audit safeguards

The repository contains multiple explicit fail-fast checks. Depending on the stage, these include:

- expected respondent count;
- duplicate `Respondent_ID` detection;
- required-column validation;
- unique IND→dimension mapping;
- score-range validation `[0,100]`;
- missing-value checks;
- `Inf/-Inf` checks;
- respondent alignment by ID;
- verification that external criteria are not BDSF construction variables;
- verification of `BDSF_equal = mean(D1,D2,D3,D4)`;
- source-file existence checks;
- quality-audit worksheets in generated Excel files.

The supplied result workbooks currently report PASS for all implemented technical quality checks.

A verification run on the supplied repository copy successfully reproduced the later external-validation, cross-country, final-summary, and CVI stages; the full runner was also observed progressing successfully through the preceding stages before the execution environment time limit interrupted the aggregate run. This is an environment execution-limit note, not a scientific pipeline failure.

---

## 11. Important methodological limitations

### 11.1 No established cross-national measurement invariance

The three national datasets originate from different questionnaires. The common BDSF dimensions provide a standards/framework-based harmonisation layer, but this alone does not establish measurement invariance or psychometric equivalence.

Therefore:

- raw country means should not be interpreted as directly interchangeable latent-trait estimates;
- cross-country descriptive results should not be converted into country rankings;
- no population-level causal conclusion follows from the descriptive differences.

### 11.2 Scoring robustness is asymmetric across countries

The equal-spacing sensitivity analysis is substantively informative for Kazakhstan because its alternative specification differs from the official scoring for relevant ordinal indicators.

For Indonesia and India, the original ordinal coding already uses equal spacing. Consequently, `Scheme A = Scheme B` is a mathematical equivalence check rather than evidence from an independent alternative coding model.

### 11.3 External validation is criterion-specific

The external variables differ across national source surveys. The Stage-07 analyses therefore test country-specific criterion relationships, not a single common external gold standard.

The results should not be described as demonstrating universal predictive validity of BDSF.

### 11.4 Observational data do not establish causality

All external-validation analyses are observational. Significant associations do not show that a higher BDSF score causes a given real-world outcome.

### 11.5 Expert CVI effective N is four

Although seven response rows are present in the expert-assessment workbook, the CVI script includes only four experts after validating that all 43 ratings are present and valid on the 1–4 scale. Any publication report should state the effective expert N used for CVI calculations.

---

## 12. Key generated outputs

### Core respondent-level scores

```text
03_Dimension_Scoring/
├── 03_01_Kazakhstan_dimensions.csv
├── 03_02_Indonesia_dimensions.csv
└── 03_03_India_dimensions.csv
```

### Descriptive outputs

```text
04_Descriptive_Analysis/
├── *_descriptive.xlsx
├── *_Figure_A_Boxplot.png
├── *_Figure_B_BDSF_Distribution.png
└── *_Figure_C_Mean_Profile_95CI.png
```

### Robustness outputs

```text
05_Scoring_Robustness/*.xlsx
06_Weighting_Robustness/*.xlsx
06_Weighting_Robustness/*_scatter.png
```

### External validation

```text
07_External_Validation/*.xlsx
```

### Cross-country descriptive profile

```text
08_Cross_Country_Analysis/08_Cross_Country_Profile.xlsx
08_Cross_Country_Analysis/08_Cross_Country_Profile.png
```

### Final integrated summary

```text
09_Final_Summary/09_Final_BDSF_Results.csv
09_Final_Summary/09_Final_BDSF_Results.xlsx
```

### Expert validity outputs

```text
10_Expert_assessment/CVI_item_results.csv
10_Expert_assessment/CVI_summary.csv
10_Expert_assessment/CVI_expert_inclusion_audit.csv
10_Expert_assessment/CVI_binary_matrix.csv
```

---

## 13. Recommended reporting language

For scientific reporting, use formulations such as:

> “The BDSF harmonised heterogeneous national survey indicators into four common behavioural dimensions on a 0–100 scale.”

> “Weighting sensitivity was evaluated by comparing equal dimension weighting with equal indicator weighting.”

> “External validation was criterion-specific because different external variables were available in the national surveys.”

> “Cross-country means are reported descriptively for the analysed samples and are not interpreted as direct psychometric rankings because measurement invariance was not established.”

Avoid formulations that imply:

- one country is definitively “more cybersecure” than another;
- the framework has universal predictive validity;
- observational associations establish causal protection;
- Indonesia/India provide independent alternative-scoring robustness when Scheme A and Scheme B are mathematically identical.

---

## 14. Suggested citation

If this repository accompanies a manuscript, replace the placeholder below with the final bibliographic record and DOI:

```text
[Authors]. Behavioural Digital Safety Framework (BDSF):
A standards-based approach for harmonising heterogeneous cybersecurity survey data.
[Journal / Year / DOI]
```

For code citation, a versioned release (for example, Zenodo/GitHub release with DOI) is recommended so that the exact pipeline revision used in the publication can be reconstructed.

---

## 15. Data governance

Before making the repository public, verify that each raw national dataset may legally and ethically be redistributed. If redistribution is restricted, publish only:

- code;
- scoring matrices that do not reveal protected data;
- synthetic/example input schemas;
- aggregate outputs permitted by the source-data licence;
- instructions for obtaining the original datasets from their authorised sources.

Do not expose direct or quasi-identifiers in a public repository.

---

## 16. License

No explicit software licence was identified in the supplied project archive. Before public release, add a `LICENSE` file and ensure that the chosen software licence is compatible with the licences/terms of the underlying datasets.

---

## 17. Project status

Current supplied outputs indicate that the implemented pipeline stages and their internal quality audits complete successfully on the included data. The primary remaining scientific constraints are interpretive rather than computational: cross-national measurement equivalence has not been established, external validation is criterion-specific, and observational analyses should not be interpreted causally.
