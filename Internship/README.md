# CDC COVID-19 Mortality Analytics and Automated Reporting

## Overview

This repository contains a six-stage data science internship project built around the CDC/NCHS **Provisional COVID-19 Deaths by Sex and Age** dataset. The project addresses the practical problem of turning a large aggregate mortality file—with stacked time grains, overlapping demographic categories, privacy-suppressed cells, partial periods, and strongly skewed counts—into reproducible descriptive and predictive outputs.

The workflow:

1. Acquires, validates, cleans, and preprocesses the CDC extract.
2. Performs grain-safe exploratory data analysis and visualization.
3. Creates scale-free jurisdiction profiles and applies K-Means clustering.
4. Evaluates a one-month-ahead gradient-boosted regression model with chronological validation.
5. Trains a compact TensorFlow residual network anchored to a previous-month baseline.
6. Generates six formatted Word reports and separate submission descriptions.

The repository includes generated figures, result tables, sanitized reports, and numerical summaries so the main findings can be reviewed without rerunning model training. The raw CDC CSV is not committed; reproducible download instructions are provided in `data/README.md`.

## Project structure

```text
Internship/
├── README.md
├── requirements.txt
├── .gitignore
├── build_analysis.py
├── build_deep_learning.py
├── build_reports.py
├── analysis_summary.json
├── analysis_summary_pre_deep.json
├── assets/                 # Generated PNG figures used by the Word reports
├── data/
│   └── README.md           # Official source and local dataset setup
├── tables/                 # Generated CSV summaries, assignments, and predictions
└── deliverables/           # Sanitized Word reports and submission descriptions
```

The exploratory model-search scripts from the working directory are intentionally excluded. They were scratch experiments rather than required pipeline stages.

## Data

Dataset: **Provisional COVID-19 Deaths by Sex and Age**

Publisher: CDC, National Center for Health Statistics

Official sources:

- Dataset page: https://data.cdc.gov/National-Center-for-Health-Statistics/Provisional-COVID-19-Deaths-by-Sex-and-Age/9bhg-hcku
- CSV download: https://data.cdc.gov/api/views/9bhg-hcku/rows.csv?accessType=DOWNLOAD

Download the CSV and place it at this exact repository-relative path:

```text
data/Provisional COVID-19 Deaths by Sex and Age.csv
```

The data file is ignored by Git to keep the repository lean and to preserve a clear link to the official source. See `data/README.md` for attribution, filename, schema, and update-status notes.

## Installation

Python 3.11 or later is recommended. Create and activate a virtual environment, then install only the project dependencies.

### Linux or macOS

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### Windows PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

TensorFlow is used only by the deep-learning stage. Exact floating-point values may vary slightly across operating systems, processors, or TensorFlow builds even though the scripts set deterministic seeds where supported.

## Usage

Run the scripts from the repository root and in the following order:

```bash
python build_analysis.py
python build_deep_learning.py
python build_reports.py
```

The scripts do not accept command-line arguments.

### 1. `build_analysis.py`

Reads the locally downloaded CDC CSV, cleans the source, validates data-quality relationships, creates EDA and clustering outputs, evaluates the gradient-boosted model, and writes:

- Week 1–4 figures to `assets/`
- Analysis tables to `tables/`
- `analysis_summary_pre_deep.json`
- A local compressed model-array intermediate, which is ignored by Git

### 2. `build_deep_learning.py`

Uses the portable project path, reuses the cleaning/feature logic from `build_analysis.py`, trains the regularized and unregularized TensorFlow models, and writes:

- Week 5 figures to `assets/`
- `tables/deep_test_predictions.csv`
- `analysis_summary.json`

It depends on `analysis_summary_pre_deep.json`, so run `build_analysis.py` first when reproducing the project from raw data.

### 3. `build_reports.py`

Reads `analysis_summary.json` and the exact figure filenames in `assets/`, then creates the six `.docx` reports and six submission-description text files in `deliverables/`.

It does not read the CSV files in `tables/`; those files are transparent supporting outputs for review.

## Reproducibility and methodology notes

- The raw source is treated as immutable.
- Privacy-suppressed counts remain missing in the clean analytical representation; the midpoint value 5 is used only for selected lagged model predictors where the source establishes an interval of 1–9.
- Monthly, yearly, and cumulative grains are never summed together.
- Age analyses use a documented mutually exclusive age-band sequence.
- The national total and overlapping New York City geography are excluded from the 52-jurisdiction modeling matrix.
- Model validation is chronological rather than random.
- K-Means uses fixed random seeds and repeated initialization.
- The neural model uses fixed seeds, L2 regularization, dropout, Huber loss, and early stopping.
- The previous-month forecast remains the baseline for model acceptance.

A full rebuild should use the script order shown above. Generated outputs already committed to the repository correspond to the included JSON summaries and report narrative.

## Outputs

### `assets/`

Contains publication-ready PNG files for missingness, outlier context, national trends, age/sex profiles, state composition, cluster selection and PCA, supervised-model errors, neural architecture, learning curves, and forecasts.

### `tables/`

Contains machine-readable results including:

- Missingness counts and percentages
- Age and jurisdiction summaries
- Cluster features, assignments, centroids, and selection metrics
- Supervised cross-validation metrics and holdout predictions
- Deep-learning test metadata and predictions

### `deliverables/`

Contains six sanitized Word reports and six standalone submission descriptions. The document packages were checked for comments, tracked revisions, embedded signatures, email addresses, and personal author metadata. Generated author metadata is set to the neutral value `Internship Project`.

## Privacy and repository hygiene

Private/local files are intentionally excluded. The repository does not include credentials, API keys, authentication tokens, personal contact information, virtual environments, caches, logs, Office lock files, raw model arrays, or the raw CDC CSV.

Before publishing a modified report, inspect the Word document again if you add your name, student ID, email address, signature, comments, or tracked changes. Git history is persistent; sensitive information should be removed before it is committed.

## Limitations

The source is aggregate, provisional, and historical, ending in September 2023. It does not include population denominators, vaccination, hospitalization, variant, or individual-level clinical data. The clustering is descriptive, and the forecasts are a methodological demonstration rather than a current public-health surveillance product.
