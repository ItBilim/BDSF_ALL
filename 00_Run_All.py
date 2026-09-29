"""Run the complete BDSF pipeline in dependency order.

Usage:
    python 00_Run_All.py

The runner preserves the project directory structure and executes each stage
sequentially. It stops immediately if any script fails.
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

PIPELINE = [
    # 01 - Data preprocessing
    "01_Data_Preprocessing/01_01_Kazakhstan_data_preprocessing.py",
    "01_Data_Preprocessing/01_02_Indonesia_data_preprocessing.py",
    "01_Data_Preprocessing/01_03_India_data_preprocessing.py",

    # 02 - Indicator scoring
    "02_Scoring_Matrix/02_01_Kazakhstan_cybersecurity_scoring.py",
    "02_Scoring_Matrix/02_02_Indonesia_cybersecurity_scoring.py",
    "02_Scoring_Matrix/02_03_India_cybersecurity_scoring.py",

    # 03 - Dimension scoring
    "03_Dimension_Scoring/03_01_Kazakhstan_Dimension_Scoring.py",
    "03_Dimension_Scoring/03_02_Indonesia_Dimension_Scoring.py",
    "03_Dimension_Scoring/03_03_India_Dimension_Scoring.py",

    # 04 - Descriptive analysis
    "04_Descriptive_Analysis/04_01_Kazakhstan_Descriptive_Analysis.py",
    "04_Descriptive_Analysis/04_02_Indonesia_Descriptive_Analysis.py",
    "04_Descriptive_Analysis/04_03_India_Descriptive_Analysis.py",

    # 05 - Scoring robustness
    "05_Scoring_Robustness/05_01_Kazakhstan_Scoring_Robustness.py",
    "05_Scoring_Robustness/05_02_Indonesia_Scoring_Robustness.py",
    "05_Scoring_Robustness/05_03_India_Scoring_Robustness.py",

    # 06 - Weighting robustness
    "06_Weighting_Robustness/06_01_Kazakhstan_Weighting_Robustness.py",
    "06_Weighting_Robustness/06_02_Indonesia_Weighting_Robustness.py",
    "06_Weighting_Robustness/06_03_India_Weighting_Robustness.py",

    # 07 - External validation
    "07_External_Validation/07_01_Kazakhstan_External_Validation.py",
    "07_External_Validation/07_02_Indonesia_External_Validation.py",
    "07_External_Validation/07_03_India_External_Validation.py",

    # 08 - Cross-country analysis
    "08_Cross_Country_Analysis/08_Cross_Country_Profile.py",

    # 09 - Final summary
    "09_Final_Summary/09_Final_BDSF_Results.py",

    # 10 - Expert assessment
    "10_Expert_assessment/10_01_Calculate_cvi.py",
]


def run_script(relative_path: str, number: int, total: int) -> None:
    script = PROJECT_ROOT / relative_path
    if not script.is_file():
        raise FileNotFoundError(f"Required script not found: {script}")

    print("\n" + "=" * 78, flush=True)
    print(f"[{number:02d}/{total:02d}] RUNNING: {relative_path}", flush=True)
    print("=" * 78, flush=True)

    started = time.perf_counter()

    # Important: several original scripts use filenames relative to their own
    # directory. Therefore each script is executed with its folder as cwd.
    completed = subprocess.run(
        [sys.executable, script.name],
        cwd=script.parent,
        check=False,
    )

    elapsed = time.perf_counter() - started
    if completed.returncode != 0:
        raise RuntimeError(
            f"Pipeline stopped: {relative_path} failed "
            f"with exit code {completed.returncode}."
        )

    print(f"[OK] Completed in {elapsed:.2f} s", flush=True)


def main() -> int:
    total = len(PIPELINE)
    print("BDSF FULL PIPELINE", flush=True)
    print(f"Project root : {PROJECT_ROOT}", flush=True)
    print(f"Python       : {sys.executable}", flush=True)
    print(f"Scripts      : {total}", flush=True)

    pipeline_started = time.perf_counter()

    try:
        for number, relative_path in enumerate(PIPELINE, start=1):
            run_script(relative_path, number, total)
    except (FileNotFoundError, RuntimeError) as exc:
        print("\n" + "!" * 78, file=sys.stderr, flush=True)
        print(f"BDSF PIPELINE FAILED\n{exc}", file=sys.stderr, flush=True)
        print("!" * 78, file=sys.stderr, flush=True)
        return 1

    elapsed = time.perf_counter() - pipeline_started
    print("\n" + "=" * 78, flush=True)
    print("BDSF PIPELINE COMPLETED SUCCESSFULLY", flush=True)
    print(f"Total execution time: {elapsed:.2f} s", flush=True)
    print("=" * 78, flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
