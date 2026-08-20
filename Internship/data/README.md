# Dataset setup

## Dataset

**Provisional COVID-19 Deaths by Sex and Age**

Publisher: Centers for Disease Control and Prevention (CDC), National Center for Health Statistics (NCHS)

The project uses this historical aggregate mortality dataset for data-quality analysis, exploratory visualization, jurisdiction clustering, one-month-ahead supervised forecasting, and residual neural-network evaluation.

The raw CSV is intentionally not committed to this repository. The supplied extract is approximately 24 MB, is reproducible from the official source, and should remain separate from generated project outputs.

## Official source

- Dataset page: https://data.cdc.gov/National-Center-for-Health-Statistics/Provisional-COVID-19-Deaths-by-Sex-and-Age/9bhg-hcku
- Official CSV download: https://data.cdc.gov/api/views/9bhg-hcku/rows.csv?accessType=DOWNLOAD

The publisher states that this dataset stopped receiving updates after 27 September 2023 and points users to CDC WONDER for similar current data. The repository treats the file as a historical analysis dataset, not a current surveillance feed.

## Required local filename and location

Download the official CSV and save it exactly as:

```text
data/Provisional COVID-19 Deaths by Sex and Age.csv
```

From the repository root, the expected path is therefore:

```text
Internship/data/Provisional COVID-19 Deaths by Sex and Age.csv
```

Do not rename the 16 source headers. The cleaning script expects the CDC header names and parses source dates in `MM/DD/YYYY` format.

The `.gitignore` excludes `data/*.csv`, so downloading the dataset locally will not accidentally add it to a Git commit.

## Attribution and reuse

Use of the data should retain attribution to CDC/NCHS and link to the official dataset page above. The CDC page identifies the material as a U.S. Government work. Users are responsible for checking the current source terms and documentation before redistributing a copy.
