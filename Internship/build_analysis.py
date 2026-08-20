from __future__ import annotations

import json
import math
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from scipy.cluster.hierarchy import dendrogram, linkage
from sklearn.cluster import KMeans
from sklearn.compose import ColumnTransformer
from sklearn.decomposition import PCA
from sklearn.ensemble import HistGradientBoostingRegressor, RandomForestRegressor
from sklearn.metrics import (
    mean_absolute_error,
    mean_squared_error,
    r2_score,
    silhouette_score,
)
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

warnings.filterwarnings("ignore", category=FutureWarning)
sns.set_theme(style="whitegrid", context="notebook")

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
TABLES = ROOT / "tables"
ASSETS.mkdir(parents=True, exist_ok=True)
TABLES.mkdir(parents=True, exist_ok=True)
DATA_PATH = ROOT / "data" / "Provisional COVID-19 Deaths by Sex and Age.csv"

if not DATA_PATH.exists():
    raise FileNotFoundError(
        f"Required dataset not found at {DATA_PATH}. "
        "Follow the download instructions in data/README.md."
    )

DATE_COLS = ["Data As Of", "Start Date", "End Date"]
COUNT_COLS_ORIGINAL = [
    "COVID-19 Deaths",
    "Total Deaths",
    "Pneumonia Deaths",
    "Pneumonia and COVID-19 Deaths",
    "Influenza Deaths",
    "Pneumonia, Influenza, or COVID-19 Deaths",
]
RENAME = {
    "Data As Of": "data_as_of",
    "Start Date": "start_date",
    "End Date": "end_date",
    "Group": "group",
    "Year": "year",
    "Month": "month",
    "State": "state",
    "Sex": "sex",
    "Age Group": "age_group",
    "COVID-19 Deaths": "covid_deaths",
    "Total Deaths": "total_deaths",
    "Pneumonia Deaths": "pneumonia_deaths",
    "Pneumonia and COVID-19 Deaths": "pneumonia_covid_deaths",
    "Influenza Deaths": "influenza_deaths",
    "Pneumonia, Influenza, or COVID-19 Deaths": "pic_deaths",
    "Footnote": "footnote",
}
COUNT_COLS = [
    "covid_deaths",
    "total_deaths",
    "pneumonia_deaths",
    "pneumonia_covid_deaths",
    "influenza_deaths",
    "pic_deaths",
]
EXCLUSIVE_AGE_GROUPS = [
    "Under 1 year", "1-4 years", "5-14 years", "15-24 years",
    "25-34 years", "35-44 years", "45-54 years", "55-64 years",
    "65-74 years", "75-84 years", "85 years and over",
]
JURISDICTIONS = None


def savefig(name: str):
    path = ASSETS / name
    plt.tight_layout()
    plt.savefig(path, dpi=220, bbox_inches="tight", facecolor="white")
    plt.close()
    return str(path)


def load_clean():
    raw = pd.read_csv(DATA_PATH, dtype=str, low_memory=False)
    df = raw.rename(columns=RENAME).copy()

    for c in ["data_as_of", "start_date", "end_date"]:
        df[c] = pd.to_datetime(df[c], format="%m/%d/%Y", errors="raise")
    for c in ["year", "month"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
    for c in COUNT_COLS:
        df[c] = pd.to_numeric(
            df[c].str.replace(",", "", regex=False), errors="coerce"
        ).astype("Float64")
        df[f"{c}_suppressed"] = df[c].isna() & df["footnote"].notna()

    df["has_suppression"] = df[[f"{c}_suppressed" for c in COUNT_COLS]].any(axis=1)
    expected_year_end = pd.to_datetime(
        df["year"].astype("Int64").astype("string") + "-12-31",
        format="%Y-%m-%d",
        errors="coerce",
    )
    df["is_partial_period"] = (
        df["group"].eq("By Month")
        & (df["end_date"] < (df["start_date"] + pd.offsets.MonthEnd(0)))
    ) | (
        df["group"].eq("By Year")
        & (df["end_date"] < expected_year_end)
    )
    df["covid_share_of_total"] = np.where(
        df["total_deaths"].fillna(0) > 0,
        df["covid_deaths"] / df["total_deaths"],
        np.nan,
    )
    return raw, df


raw, df = load_clean()
JURISDICTIONS = sorted(set(df["state"]) - {"United States", "New York City"})

summary = {
    "raw_rows": int(len(raw)),
    "raw_columns": int(raw.shape[1]),
    "clean_columns": int(df.shape[1]),
    "source_file_mb": round(DATA_PATH.stat().st_size / 1024**2, 2),
    "data_as_of": str(df["data_as_of"].dt.date.min()),
    "start_min": str(df["start_date"].dt.date.min()),
    "end_max": str(df["end_date"].dt.date.max()),
    "groups": {k: int(v) for k, v in df["group"].value_counts().items()},
    "states": int(df["state"].nunique()),
    "sexes": int(df["sex"].nunique()),
    "age_groups": int(df["age_group"].nunique()),
    "exact_duplicates": int(raw.duplicated().sum()),
    "key_duplicates": int(raw.duplicated(["Group", "Year", "Month", "State", "Sex", "Age Group"]).sum()),
    "footnoted_rows": int(raw["Footnote"].notna().sum()),
    "fully_observed_rows": int(raw["Footnote"].isna().sum()),
    "numeric_missing": {c: int(df[c].isna().sum()) for c in COUNT_COLS},
    "numeric_missing_pct": {c: round(float(df[c].isna().mean() * 100), 2) for c in COUNT_COLS},
    "negative_counts": int(sum((df[c] < 0).sum() for c in COUNT_COLS)),
    "covid_gt_total": int((df["covid_deaths"] > df["total_deaths"]).sum()),
    "pneumonia_gt_total": int((df["pneumonia_deaths"] > df["total_deaths"]).sum()),
    "influenza_gt_total": int((df["influenza_deaths"] > df["total_deaths"]).sum()),
    "overlap_gt_covid": int((df["pneumonia_covid_deaths"] > df["covid_deaths"]).sum()),
    "union_lt_components": int((df["pic_deaths"] < df[["covid_deaths", "pneumonia_deaths", "influenza_deaths"]].max(axis=1)).sum()),
    "invalid_dates": int((df["start_date"] > df["end_date"]).sum() + (df["end_date"] > df["data_as_of"]).sum()),
    "partial_period_rows": int(df["is_partial_period"].sum()),
}

# Missingness table and plot
miss = pd.DataFrame({
    "field": COUNT_COLS,
    "missing_count": [df[c].isna().sum() for c in COUNT_COLS],
    "missing_percent": [df[c].isna().mean() * 100 for c in COUNT_COLS],
})
miss.to_csv(TABLES / "missingness.csv", index=False)
plt.figure(figsize=(9, 4.8))
sns.barplot(data=miss, y="field", x="missing_percent", color="#377eb8")
plt.xlabel("Missing / suppressed values (%)")
plt.ylabel("")
plt.title("Missingness is concentrated in suppressed count fields")
for i, v in enumerate(miss["missing_percent"]):
    plt.text(v + 0.4, i, f"{v:.1f}%", va="center", fontsize=9)
plt.xlim(0, 38)
savefig("w1_missingness.png")

# Distribution / outlier comparison for coherent grain.
state_total = df[
    df["group"].eq("By Total")
    & df["sex"].eq("All Sexes")
    & df["age_group"].eq("All Ages")
    & df["state"].isin(JURISDICTIONS)
].copy()
q1, q3 = state_total["covid_deaths"].quantile([0.25, 0.75])
iqr = q3 - q1
outlier_cutoff = q3 + 1.5 * iqr
state_total["iqr_flag"] = state_total["covid_deaths"] > outlier_cutoff
summary["state_iqr_q1"] = float(q1)
summary["state_iqr_q3"] = float(q3)
summary["state_iqr_cutoff"] = float(outlier_cutoff)
summary["state_iqr_flag_count"] = int(state_total["iqr_flag"].sum())
summary["state_iqr_flag_names"] = state_total.loc[state_total["iqr_flag"], "state"].tolist()
plt.figure(figsize=(9, 4.8))
sns.histplot(state_total["covid_deaths"].astype(float), bins=18, color="#e41a1c")
plt.axvline(outlier_cutoff, color="black", ls="--", label=f"1.5×IQR cutoff: {outlier_cutoff:,.0f}")
plt.title("Cumulative COVID-19 deaths vary strongly across jurisdictions")
plt.xlabel("Reported COVID-19 deaths, 2020–23 September 2023")
plt.ylabel("Number of jurisdictions")
plt.legend()
savefig("w1_outlier_distribution.png")

# Week 2 EDA slices
national_monthly = df[
    df["group"].eq("By Month")
    & df["state"].eq("United States")
    & df["sex"].eq("All Sexes")
    & df["age_group"].eq("All Ages")
].sort_values("start_date").copy()
national_monthly["covid_pct_total"] = 100 * national_monthly["covid_deaths"] / national_monthly["total_deaths"]
national_complete = national_monthly[~national_monthly["is_partial_period"]].copy()
peak = national_complete.loc[national_complete["covid_deaths"].idxmax()]
low = national_complete.loc[national_complete["covid_deaths"].idxmin()]
summary["eda"] = {
    "national_cumulative_covid": int(state_total.loc[df.loc[state_total.index, "state"].eq("United States") if False else state_total.index[:0], "covid_deaths"].sum()) if False else int(df.loc[(df.group == "By Total") & (df.state == "United States") & (df.sex == "All Sexes") & (df.age_group == "All Ages"), "covid_deaths"].iloc[0]),
    "national_total_deaths": int(df.loc[(df.group == "By Total") & (df.state == "United States") & (df.sex == "All Sexes") & (df.age_group == "All Ages"), "total_deaths"].iloc[0]),
    "national_covid_share": round(float(df.loc[(df.group == "By Total") & (df.state == "United States") & (df.sex == "All Sexes") & (df.age_group == "All Ages"), "covid_share_of_total"].iloc[0] * 100), 2),
    "peak_month": peak["start_date"].strftime("%B %Y"),
    "peak_month_covid": int(peak["covid_deaths"]),
    "peak_month_share": round(float(peak["covid_pct_total"]), 2),
    "lowest_complete_month": low["start_date"].strftime("%B %Y"),
    "lowest_complete_month_covid": int(low["covid_deaths"]),
    "aug_2023_covid": int(national_monthly.loc[national_monthly.start_date.eq(pd.Timestamp("2023-08-01")), "covid_deaths"].iloc[0]),
}
plt.figure(figsize=(11, 5.5))
plt.plot(national_monthly["start_date"], national_monthly["covid_deaths"], marker="o", ms=3.5, lw=2, color="#d73027")
plt.scatter([peak["start_date"]], [peak["covid_deaths"]], s=70, color="black", zorder=5)
plt.annotate(f"Peak: {peak['start_date']:%b %Y}\n{peak['covid_deaths']:,.0f}",
             (peak["start_date"], peak["covid_deaths"]), xytext=(12, 8), textcoords="offset points")
plt.axvspan(pd.Timestamp("2023-09-01"), pd.Timestamp("2023-09-30"), color="gray", alpha=.2, label="Partial September 2023")
plt.title("United States monthly reported COVID-19 deaths")
plt.xlabel("Month")
plt.ylabel("COVID-19 deaths")
plt.legend(loc="upper right")
savefig("w2_national_monthly.png")

age = df[
    df["group"].eq("By Total") & df["state"].eq("United States")
    & df["sex"].eq("All Sexes") & df["age_group"].isin(EXCLUSIVE_AGE_GROUPS)
].copy()
age["age_group"] = pd.Categorical(age["age_group"], EXCLUSIVE_AGE_GROUPS, ordered=True)
age = age.sort_values("age_group")
age["share_covid"] = 100 * age["covid_deaths"] / age["covid_deaths"].sum()
age["covid_pct_total"] = 100 * age["covid_deaths"] / age["total_deaths"]
older = age[age["age_group"].isin(["65-74 years", "75-84 years", "85 years and over"])]
summary["eda"].update({
    "age_65plus_deaths": int(older["covid_deaths"].sum()),
    "age_65plus_share": round(float(older["covid_deaths"].sum() / age["covid_deaths"].sum() * 100), 2),
    "age_max_count_group": str(age.loc[age["covid_deaths"].idxmax(), "age_group"]),
    "age_max_count": int(age["covid_deaths"].max()),
    "age_highest_pct_group": str(age.loc[age["covid_pct_total"].idxmax(), "age_group"]),
    "age_highest_pct": round(float(age["covid_pct_total"].max()), 2),
})
fig, ax1 = plt.subplots(figsize=(11, 5.7))
x = np.arange(len(age))
ax1.bar(x, age["covid_deaths"].astype(float), color="#4575b4", alpha=.85)
ax1.set_ylabel("COVID-19 deaths")
ax1.set_xticks(x)
ax1.set_xticklabels(age["age_group"].astype(str), rotation=40, ha="right")
ax1.set_xlabel("Mutually exclusive age group")
ax1.set_title("COVID-19 death count and share of all deaths by age")
ax2 = ax1.twinx()
ax2.plot(x, age["covid_pct_total"].astype(float), color="#d73027", marker="o", lw=2)
ax2.set_ylabel("COVID-19 deaths as % of all deaths", color="#d73027")
savefig("w2_age_profile.png")

sex_age = df[
    df["group"].eq("By Total") & df["state"].eq("United States")
    & df["sex"].isin(["Female", "Male"]) & df["age_group"].isin(EXCLUSIVE_AGE_GROUPS)
].pivot(index="age_group", columns="sex", values="covid_deaths").reindex(EXCLUSIVE_AGE_GROUPS)
sex_age["male_to_female"] = sex_age["Male"] / sex_age["Female"]
summary["eda"].update({
    "male_total_selected_bins": int(sex_age["Male"].sum()),
    "female_total_selected_bins": int(sex_age["Female"].sum()),
    "male_female_ratio_selected_bins": round(float(sex_age["Male"].sum() / sex_age["Female"].sum()), 3),
    "max_sex_ratio_group": str(sex_age["male_to_female"].idxmax()),
    "max_sex_ratio": round(float(sex_age["male_to_female"].max()), 2),
})
sex_long = sex_age[["Female", "Male"]].reset_index().melt("age_group", var_name="Sex", value_name="COVID-19 deaths")
plt.figure(figsize=(11, 5.8))
sns.barplot(data=sex_long, x="age_group", y="COVID-19 deaths", hue="Sex", palette={"Female":"#984ea3", "Male":"#4daf4a"})
plt.xticks(rotation=40, ha="right")
plt.xlabel("Mutually exclusive age group")
plt.title("Male and female cumulative COVID-19 deaths by age")
savefig("w2_sex_age.png")

state_total["covid_pct_total"] = 100 * state_total["covid_deaths"] / state_total["total_deaths"]
top_state_share = state_total.nlargest(15, "covid_pct_total").sort_values("covid_pct_total")
summary["eda"].update({
    "highest_state_share_name": str(state_total.loc[state_total["covid_pct_total"].idxmax(), "state"]),
    "highest_state_share": round(float(state_total["covid_pct_total"].max()), 2),
    "lowest_state_share_name": str(state_total.loc[state_total["covid_pct_total"].idxmin(), "state"]),
    "lowest_state_share": round(float(state_total["covid_pct_total"].min()), 2),
    "state_share_median": round(float(state_total["covid_pct_total"].median()), 2),
})
plt.figure(figsize=(9, 6))
sns.barplot(data=top_state_share, y="state", x="covid_pct_total", color="#f46d43")
plt.xlabel("COVID-19 deaths as % of total reported deaths")
plt.ylabel("")
plt.title("Top 15 jurisdictions by COVID-19 share of all deaths")
savefig("w2_state_share.png")

corr_data = national_complete[["covid_deaths", "total_deaths", "pneumonia_deaths", "pneumonia_covid_deaths", "influenza_deaths", "pic_deaths"]].astype(float)
corr = corr_data.corr(method="spearman")
summary["eda"]["covid_pneumonia_spearman"] = round(float(corr.loc["covid_deaths", "pneumonia_deaths"]), 3)
summary["eda"]["covid_total_spearman"] = round(float(corr.loc["covid_deaths", "total_deaths"]), 3)
plt.figure(figsize=(8, 6.5))
labels = ["COVID", "All-cause", "Pneumonia", "Pneumonia+COVID", "Influenza", "P/I/C union"]
sns.heatmap(corr, annot=True, fmt=".2f", cmap="vlag", center=0, vmin=-1, vmax=1,
            xticklabels=labels, yticklabels=labels, square=True, cbar_kws={"label":"Spearman correlation"})
plt.title("Correlation among national monthly death-count series")
plt.xticks(rotation=40, ha="right")
savefig("w2_correlation.png")

# Tables for EDA
age[["age_group", "covid_deaths", "total_deaths", "share_covid", "covid_pct_total"]].to_csv(TABLES / "age_summary.csv", index=False)
state_total[["state", "covid_deaths", "total_deaths", "covid_pct_total"]].sort_values("covid_pct_total", ascending=False).to_csv(TABLES / "state_summary.csv", index=False)

# Week 3 clustering feature table
base = state_total.set_index("state").copy()
features = pd.DataFrame(index=JURISDICTIONS)
features["covid_pct_all_deaths"] = 100 * base["covid_deaths"] / base["total_deaths"]
features["pneumonia_pct_all_deaths"] = 100 * base["pneumonia_deaths"] / base["total_deaths"]
features["covid_with_pneumonia_pct"] = 100 * base["pneumonia_covid_deaths"] / base["covid_deaths"]
features["influenza_pct_all_deaths"] = 100 * base["influenza_deaths"] / base["total_deaths"]

age65 = df[
    df["group"].eq("By Total") & df["state"].isin(JURISDICTIONS)
    & df["sex"].eq("All Sexes")
    & df["age_group"].isin(["65-74 years", "75-84 years", "85 years and over"])
].groupby("state", observed=True)["covid_deaths"].sum(min_count=3)
features["age_65plus_covid_pct"] = 100 * age65 / base["covid_deaths"]
sex_cum = df[
    df["group"].eq("By Total") & df["state"].isin(JURISDICTIONS)
    & df["sex"].isin(["Female", "Male"]) & df["age_group"].eq("All Ages")
].pivot(index="state", columns="sex", values="covid_deaths")
features["male_covid_pct"] = 100 * sex_cum["Male"] / (sex_cum["Male"] + sex_cum["Female"])
yearly = df[
    df["group"].eq("By Year") & df["state"].isin(JURISDICTIONS)
    & df["sex"].eq("All Sexes") & df["age_group"].eq("All Ages")
    & df["year"].isin([2020, 2021, 2022])
].pivot(index="state", columns="year", values="covid_deaths")
features["covid_2021_pct_2020_22"] = 100 * yearly[2021] / yearly[[2020, 2021, 2022]].sum(axis=1)
features = features.astype(float)
assert features.notna().all().all(), features.isna().sum()

scaler = StandardScaler()
X = scaler.fit_transform(features)
cluster_metrics = []
for k in range(2, 9):
    km = KMeans(n_clusters=k, n_init=50, random_state=42)
    labels_k = km.fit_predict(X)
    cluster_metrics.append({"k": k, "inertia": km.inertia_, "silhouette": silhouette_score(X, labels_k)})
metrics_df = pd.DataFrame(cluster_metrics)
# Selection: highest silhouette; keep deterministic and data-led.
best_k = int(metrics_df.loc[metrics_df["silhouette"].idxmax(), "k"])
km = KMeans(n_clusters=best_k, n_init=100, random_state=42)
clusters = km.fit_predict(X)
features["cluster"] = clusters
# Re-label for stable semantic ordering by covid share then age profile.
old_order = features.groupby("cluster")["covid_pct_all_deaths"].mean().sort_values().index.tolist()
remap = {old: new for new, old in enumerate(old_order)}
features["cluster"] = features["cluster"].map(remap)
clusters = features["cluster"].to_numpy()

pca = PCA(n_components=2)
pcs = pca.fit_transform(X)
features["pc1"] = pcs[:, 0]
features["pc2"] = pcs[:, 1]
centroids = features.groupby("cluster").mean(numeric_only=True)
cluster_sizes = features["cluster"].value_counts().sort_index()
feature_means = features.drop(columns=["cluster", "pc1", "pc2"]).mean()
cluster_profile = centroids[feature_means.index]
cluster_profile_z = (cluster_profile - feature_means) / features[feature_means.index].std(ddof=0)

summary["clustering"] = {
    "n_jurisdictions": int(len(features)),
    "features": list(feature_means.index),
    "best_k": best_k,
    "best_silhouette": round(float(metrics_df.loc[metrics_df.k.eq(best_k), "silhouette"].iloc[0]), 3),
    "pca_variance_2d": round(float(pca.explained_variance_ratio_.sum() * 100), 2),
    "cluster_sizes": {str(int(k)): int(v) for k, v in cluster_sizes.items()},
    "cluster_members": {str(int(k)): features.index[features.cluster.eq(k)].tolist() for k in sorted(features.cluster.unique())},
    "cluster_means": {str(int(k)): {c: round(float(v), 3) for c, v in row.items()} for k, row in cluster_profile.iterrows()},
    "metrics": [{"k": int(r.k), "inertia": round(float(r.inertia), 3), "silhouette": round(float(r.silhouette), 3)} for _, r in metrics_df.iterrows()],
}
features.to_csv(TABLES / "cluster_features_assignments.csv")
metrics_df.to_csv(TABLES / "cluster_selection_metrics.csv", index=False)
cluster_profile.to_csv(TABLES / "cluster_profile_means.csv")

fig, ax1 = plt.subplots(figsize=(8.5, 5.2))
ax1.plot(metrics_df["k"], metrics_df["inertia"], marker="o", color="#377eb8")
ax1.set_xlabel("Number of clusters (k)")
ax1.set_ylabel("Within-cluster sum of squares (inertia)", color="#377eb8")
ax2 = ax1.twinx()
ax2.plot(metrics_df["k"], metrics_df["silhouette"], marker="s", color="#e41a1c")
ax2.set_ylabel("Silhouette score", color="#e41a1c")
ax1.axvline(best_k, color="black", ls="--", alpha=.7, label=f"Selected k={best_k}")
ax1.legend(loc="center right")
plt.title("Cluster-number selection balances compactness and separation")
savefig("w3_k_selection.png")

plt.figure(figsize=(9.5, 6.2))
palette = sns.color_palette("Set2", best_k)
for k in sorted(features.cluster.unique()):
    part = features[features.cluster.eq(k)]
    plt.scatter(part.pc1, part.pc2, s=70, color=palette[k], label=f"Cluster {k} (n={len(part)})", edgecolor="black", linewidth=.4)
    for state, row in part.iterrows():
        plt.text(row.pc1 + .04, row.pc2 + .04, state if len(state) <= 12 else "", fontsize=7, alpha=.8)
plt.xlabel(f"Principal component 1 ({pca.explained_variance_ratio_[0]*100:.1f}% variance)")
plt.ylabel(f"Principal component 2 ({pca.explained_variance_ratio_[1]*100:.1f}% variance)")
plt.title("Standardized jurisdiction mortality profiles in PCA space")
plt.legend()
savefig("w3_pca_clusters.png")

plt.figure(figsize=(10, max(4.5, best_k * .9 + 2)))
sns.heatmap(cluster_profile_z, cmap="vlag", center=0, annot=True, fmt=".2f", cbar_kws={"label":"Standard deviations from overall mean"})
plt.xlabel("Profile feature")
plt.ylabel("Cluster")
plt.title("Cluster characteristics relative to the jurisdiction average")
plt.xticks(rotation=35, ha="right")
savefig("w3_cluster_heatmap.png")

# Hierarchical dendrogram for robustness/context.
Z = linkage(X, method="ward")
plt.figure(figsize=(12, 5.5))
dendrogram(Z, labels=features.index.tolist(), leaf_rotation=90, leaf_font_size=6, color_threshold=None)
plt.ylabel("Ward linkage distance")
plt.title("Hierarchical view of jurisdiction profile similarity")
savefig("w3_dendrogram.png")

# Week 4 panel forecasting data and model
panel = df[
    df["group"].eq("By Month") & df["state"].isin(JURISDICTIONS)
    & df["sex"].eq("All Sexes") & df["age_group"].eq("All Ages")
].copy()
panel = panel[~panel["is_partial_period"]].sort_values(["state", "start_date"])
# Preserve suppression knowledge. Midpoint 5 is used only for lagged predictors, never as an observed evaluation target.
for c in ["covid_deaths", "total_deaths", "pneumonia_deaths"]:
    panel[f"{c}_for_lag"] = panel[c].fillna(5.0)
g = panel.groupby("state", observed=True)
panel["lag1_covid"] = g["covid_deaths_for_lag"].shift(1)
panel["lag2_covid"] = g["covid_deaths_for_lag"].shift(2)
panel["lag3_covid"] = g["covid_deaths_for_lag"].shift(3)
panel["rolling3_covid"] = g["covid_deaths_for_lag"].transform(
    lambda s: s.shift(1).rolling(3).mean()
)
panel["lag1_total"] = g["total_deaths_for_lag"].shift(1)
panel["lag1_pneumonia"] = g["pneumonia_deaths_for_lag"].shift(1)
panel["month_sin"] = np.sin(2 * np.pi * panel["month"].astype(float) / 12)
panel["month_cos"] = np.cos(2 * np.pi * panel["month"].astype(float) / 12)
panel["time_index"] = (panel["start_date"].dt.year - 2020) * 12 + panel["start_date"].dt.month - 1
model_data = panel.dropna(subset=["lag1_covid", "lag2_covid", "lag3_covid", "rolling3_covid", "lag1_total", "lag1_pneumonia", "covid_deaths"]).copy()

num_features = ["lag1_covid", "lag2_covid", "lag3_covid", "rolling3_covid", "lag1_total", "lag1_pneumonia", "month_sin", "month_cos", "time_index"]
cat_features = ["state"]
preprocessor = ColumnTransformer([
    ("numeric", StandardScaler(), num_features),
    ("state", OneHotEncoder(handle_unknown="ignore", sparse_output=False), cat_features),
], remainder="drop")
model = HistGradientBoostingRegressor(
    learning_rate=0.05, max_iter=350, max_leaf_nodes=15,
    min_samples_leaf=12, l2_regularization=1.0, random_state=42,
)
pipe = Pipeline([("preprocess", preprocessor), ("model", model)])


def metrics(y, pred):
    y = np.asarray(y, dtype=float)
    pred = np.clip(np.asarray(pred, dtype=float), 0, None)
    return {
        "MAE": mean_absolute_error(y, pred),
        "RMSE": mean_squared_error(y, pred) ** 0.5,
        "R2": r2_score(y, pred),
        "WAPE_pct": 100 * np.abs(y - pred).sum() / y.sum(),
    }

# Expanding-window folds chosen by calendar time.
fold_defs = [
    ("2021 H2", pd.Timestamp("2021-07-01"), pd.Timestamp("2021-12-01")),
    ("2022 H1", pd.Timestamp("2022-01-01"), pd.Timestamp("2022-06-01")),
    ("2022 H2", pd.Timestamp("2022-07-01"), pd.Timestamp("2022-12-01")),
]
cv_rows = []
for label, start, end in fold_defs:
    tr = model_data[model_data.start_date < start]
    va = model_data[model_data.start_date.between(start, end)]
    Xtr, ytr = tr[num_features + cat_features], np.log1p(tr["covid_deaths"].astype(float))
    Xva, yva = va[num_features + cat_features], va["covid_deaths"].astype(float)
    pipe.fit(Xtr, ytr)
    pred = np.expm1(pipe.predict(Xva))
    m = metrics(yva, pred)
    bm = metrics(yva, va["lag1_covid"].astype(float))
    cv_rows.append({"fold": label, "train_rows": len(tr), "validation_rows": len(va), **{f"model_{k}": v for k, v in m.items()}, **{f"baseline_{k}": v for k, v in bm.items()}})
cv_df = pd.DataFrame(cv_rows)

train = model_data[model_data.start_date < pd.Timestamp("2023-01-01")]
test = model_data[model_data.start_date.between(pd.Timestamp("2023-01-01"), pd.Timestamp("2023-08-01"))]
pipe.fit(train[num_features + cat_features], np.log1p(train["covid_deaths"].astype(float)))
test_pred = np.clip(np.expm1(pipe.predict(test[num_features + cat_features])), 0, None)
model_metrics = metrics(test["covid_deaths"].astype(float), test_pred)
baseline_metrics = metrics(test["covid_deaths"].astype(float), test["lag1_covid"].astype(float))
predictions = test[["state", "start_date", "covid_deaths", "lag1_covid"]].copy()
predictions["prediction"] = test_pred
predictions["absolute_error"] = np.abs(predictions["covid_deaths"].astype(float) - test_pred)
predictions["error"] = test_pred - predictions["covid_deaths"].astype(float)
worst = predictions.nlargest(10, "absolute_error")

summary["supervised"] = {
    "model_rows": int(len(model_data)),
    "train_rows": int(len(train)),
    "test_rows": int(len(test)),
    "train_period": f"{train.start_date.min():%b %Y} to {train.start_date.max():%b %Y}",
    "test_period": f"{test.start_date.min():%b %Y} to {test.start_date.max():%b %Y}",
    "model_metrics": {k: round(float(v), 3) for k, v in model_metrics.items()},
    "baseline_metrics": {k: round(float(v), 3) for k, v in baseline_metrics.items()},
    "mae_improvement_pct": round(float(100 * (baseline_metrics["MAE"] - model_metrics["MAE"]) / baseline_metrics["MAE"]), 2),
    "worst_cases": [{"state": r.state, "month": r.start_date.strftime("%b %Y"), "actual": round(float(r["covid_deaths"]), 1), "pred": round(float(r.prediction), 1), "abs_error": round(float(r.absolute_error), 1)} for _, r in worst.iterrows()],
    "cv": [{k: (round(float(v), 3) if isinstance(v, (int, float, np.integer, np.floating)) else v) for k, v in r.items()} for r in cv_rows],
}
cv_df.to_csv(TABLES / "supervised_cv_metrics.csv", index=False)
predictions.to_csv(TABLES / "supervised_test_predictions.csv", index=False)

# Actual vs predicted aggregated by month
plot_month = predictions.groupby("start_date").agg(actual=("covid_deaths", "sum"), predicted=("prediction", "sum"), baseline=("lag1_covid", "sum")).reset_index()
plt.figure(figsize=(10.5, 5.3))
plt.plot(plot_month.start_date, plot_month.actual, marker="o", lw=2.4, label="Actual")
plt.plot(plot_month.start_date, plot_month.predicted, marker="s", lw=2, label="Gradient-boosted model")
plt.plot(plot_month.start_date, plot_month.baseline, marker="^", lw=1.6, ls="--", label="Previous-month baseline")
plt.xlabel("Forecast month")
plt.ylabel("COVID-19 deaths across 52 jurisdictions")
plt.title("One-month-ahead test forecasts, January–August 2023")
plt.legend()
savefig("w4_forecast_monthly.png")

plt.figure(figsize=(7.2, 6.2))
plt.scatter(predictions["covid_deaths"].astype(float), predictions["prediction"], alpha=.55, s=30, color="#377eb8")
mx = max(predictions["covid_deaths"].max(), predictions["prediction"].max())
plt.plot([0, mx], [0, mx], "k--", lw=1.5, label="Perfect prediction")
plt.xscale("symlog", linthresh=10)
plt.yscale("symlog", linthresh=10)
plt.xlabel("Actual monthly COVID-19 deaths")
plt.ylabel("Predicted monthly COVID-19 deaths")
plt.title("Prediction errors grow during high-burden state-months")
plt.legend()
savefig("w4_actual_predicted.png")

plt.figure(figsize=(9.5, 5.2))
worst_plot = worst.copy()
worst_plot["label"] = worst_plot["state"] + "\n" + worst_plot["start_date"].dt.strftime("%b %Y")
sns.barplot(data=worst_plot, x="absolute_error", y="label", color="#e41a1c")
plt.xlabel("Absolute error (deaths)")
plt.ylabel("")
plt.title("Ten largest test-set forecast errors")
savefig("w4_worst_errors.png")

# Save preprocessed arrays for deep learning script.
# Fit preprocessing on training only; y transformed by log1p.
X_train = preprocessor.fit_transform(train[num_features + cat_features]).astype(np.float32)
X_test = preprocessor.transform(test[num_features + cat_features]).astype(np.float32)
y_train = np.log1p(train["covid_deaths"].astype(float).to_numpy()).astype(np.float32)
y_test = test["covid_deaths"].astype(float).to_numpy().astype(np.float32)
np.savez_compressed(ROOT / "model_arrays.npz", X_train=X_train, X_test=X_test, y_train=y_train, y_test=y_test)
test[["state", "start_date", "covid_deaths", "lag1_covid"]].to_csv(TABLES / "deep_test_meta.csv", index=False)
summary["supervised"]["n_features_after_encoding"] = int(X_train.shape[1])

# Output base summary for next stage.
with open(ROOT / "analysis_summary_pre_deep.json", "w", encoding="utf-8") as f:
    json.dump(summary, f, indent=2)
print(json.dumps(summary, indent=2))
