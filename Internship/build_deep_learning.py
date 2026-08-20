from __future__ import annotations

import json
import os
import random
from pathlib import Path

os.environ["TF_CPP_MIN_LOG_LEVEL"] = "2"
os.environ["TF_DETERMINISTIC_OPS"] = "1"
os.environ["OMP_NUM_THREADS"] = "2"
os.environ["TF_NUM_INTRAOP_THREADS"] = "2"
os.environ["TF_NUM_INTEROP_THREADS"] = "2"

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import tensorflow as tf
from sklearn.compose import ColumnTransformer
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.preprocessing import OneHotEncoder, StandardScaler

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
TABLES = ROOT / "tables"
SEED = 42
random.seed(SEED)
np.random.seed(SEED)
tf.keras.utils.set_random_seed(SEED)
try:
    tf.config.experimental.enable_op_determinism()
except Exception:
    pass

# Reuse the cleaned panel and feature logic from the executed analysis pipeline.
# Supplying __file__ lets the executed source resolve its own project-relative paths.
analysis_script = ROOT / "build_analysis.py"
namespace = {"__file__": str(analysis_script)}
source = analysis_script.read_text(encoding="utf-8")
exec(source.split("# Week 4 panel forecasting data and model")[0], namespace)
df = namespace["df"]
JURISDICTIONS = namespace["JURISDICTIONS"]

panel = df[
    df["group"].eq("By Month") & df["state"].isin(JURISDICTIONS)
    & df["sex"].eq("All Sexes") & df["age_group"].eq("All Ages")
].copy()
panel = panel[~panel["is_partial_period"]].sort_values(["state", "start_date"])
for c in ["covid_deaths", "total_deaths", "pneumonia_deaths"]:
    panel[f"{c}_for_lag"] = panel[c].fillna(5.0)
g = panel.groupby("state", observed=True)
panel["lag1_covid"] = g["covid_deaths_for_lag"].shift(1)
panel["lag2_covid"] = g["covid_deaths_for_lag"].shift(2)
panel["lag3_covid"] = g["covid_deaths_for_lag"].shift(3)
panel["rolling3_covid"] = g["covid_deaths_for_lag"].transform(lambda s: s.shift(1).rolling(3).mean())
panel["lag1_total"] = g["total_deaths_for_lag"].shift(1)
panel["lag1_pneumonia"] = g["pneumonia_deaths_for_lag"].shift(1)
panel["month_sin"] = np.sin(2 * np.pi * panel["month"].astype(float) / 12)
panel["month_cos"] = np.cos(2 * np.pi * panel["month"].astype(float) / 12)
panel["time_index"] = (panel["start_date"].dt.year - 2020) * 12 + panel["start_date"].dt.month - 1
num_features = ["lag1_covid", "lag2_covid", "lag3_covid", "rolling3_covid", "lag1_total", "lag1_pneumonia", "month_sin", "month_cos", "time_index"]
cat_features = ["state"]
model_data = panel.dropna(subset=num_features + ["covid_deaths"]).copy()

fit_data = model_data[model_data.start_date < pd.Timestamp("2022-07-01")]
val_data = model_data[model_data.start_date.between(pd.Timestamp("2022-07-01"), pd.Timestamp("2022-12-01"))]
test_data = model_data[model_data.start_date.between(pd.Timestamp("2023-01-01"), pd.Timestamp("2023-08-01"))]

preprocessor = ColumnTransformer([
    ("numeric", StandardScaler(), num_features),
    ("state", OneHotEncoder(handle_unknown="ignore", sparse_output=False), cat_features),
])
all_features = num_features + cat_features
X_fit = preprocessor.fit_transform(fit_data[all_features]).astype("float32")
X_val = preprocessor.transform(val_data[all_features]).astype("float32")
X_test = preprocessor.transform(test_data[all_features]).astype("float32")
# The network learns a correction to the previous-month forecast rather than
# relearning the entire count scale. This residual target is more stationary:
# zero means "use last month," negative means decline, and positive means growth.
y_fit = (
    np.log1p(fit_data["covid_deaths"].astype(float).to_numpy())
    - np.log1p(fit_data["lag1_covid"].astype(float).to_numpy())
).astype("float32")
y_val = (
    np.log1p(val_data["covid_deaths"].astype(float).to_numpy())
    - np.log1p(val_data["lag1_covid"].astype(float).to_numpy())
).astype("float32")
y_test = test_data["covid_deaths"].astype(float).to_numpy().astype("float32")


def make_model(regularized=True):
    reg = tf.keras.regularizers.l2(1e-3) if regularized else None
    layers = [tf.keras.layers.Input(shape=(X_fit.shape[1],))]
    for units, drop in [(32, 0.20), (16, 0.15)]:
        layers.append(tf.keras.layers.Dense(units, activation="relu", kernel_regularizer=reg))
        if regularized and drop:
            layers.append(tf.keras.layers.Dropout(drop))
    # A zero-initialized output starts at the strong previous-month baseline.
    layers.append(tf.keras.layers.Dense(
        1, activation="linear", kernel_initializer="zeros", bias_initializer="zeros"
    ))
    model = tf.keras.Sequential(layers)
    model.compile(
        optimizer=tf.keras.optimizers.Adam(learning_rate=1e-3),
        loss=tf.keras.losses.Huber(delta=0.3),
        metrics=[tf.keras.metrics.MeanAbsoluteError(name="mae_log_residual")],
    )
    return model

regularized_model = make_model(regularized=True)
callbacks = [
    tf.keras.callbacks.EarlyStopping(monitor="val_loss", patience=20, min_delta=1e-5, restore_best_weights=True),
    tf.keras.callbacks.ReduceLROnPlateau(monitor="val_loss", patience=10, factor=0.5, min_lr=1e-5),
]
history = regularized_model.fit(
    X_fit, y_fit,
    validation_data=(X_val, y_val),
    epochs=250,
    batch_size=64,
    verbose=0,
    shuffle=True,
    callbacks=callbacks,
)

# A deliberately unregularized reference makes the overfitting diagnosis observable,
# rather than asserting it from theory alone.
unregularized_model = make_model(regularized=False)
unreg_history = unregularized_model.fit(
    X_fit, y_fit,
    validation_data=(X_val, y_val),
    epochs=160,
    batch_size=64,
    verbose=0,
    shuffle=True,
)

predicted_log_residual = regularized_model.predict(X_test, verbose=0).ravel()
baseline = test_data["lag1_covid"].astype(float).to_numpy()
pred = np.clip(
    np.expm1(np.log1p(baseline) + predicted_log_residual),
    0,
    None,
)


def metrics(actual, prediction):
    return {
        "MAE": float(mean_absolute_error(actual, prediction)),
        "RMSE": float(mean_squared_error(actual, prediction) ** 0.5),
        "R2": float(r2_score(actual, prediction)),
        "WAPE_pct": float(100 * np.abs(actual - prediction).sum() / actual.sum()),
    }

model_metrics = metrics(y_test, pred)
baseline_metrics = metrics(y_test, baseline)

predictions = test_data[["state", "start_date", "covid_deaths", "lag1_covid"]].copy()
predictions["prediction"] = pred
predictions["absolute_error"] = np.abs(predictions["covid_deaths"].astype(float) - pred)
predictions.to_csv(TABLES / "deep_test_predictions.csv", index=False)

best_epoch = int(np.argmin(history.history["val_loss"]) + 1)
stop_epoch = len(history.history["loss"])
unreg_best = int(np.argmin(unreg_history.history["val_loss"]) + 1)
unreg_final_gap = float(unreg_history.history["val_loss"][-1] - unreg_history.history["loss"][-1])

# Training curves
plt.figure(figsize=(10, 5.4))
plt.plot(history.history["loss"], label="Training loss", lw=2)
plt.plot(history.history["val_loss"], label="Validation loss", lw=2)
plt.axvline(best_epoch - 1, color="black", ls="--", label=f"Best epoch: {best_epoch}")
plt.xlabel("Epoch")
plt.ylabel("Huber loss on log residual")
plt.title("Regularized network: training and validation learning curves")
plt.legend()
plt.tight_layout()
plt.savefig(ASSETS / "w5_training_curves.png", dpi=220, bbox_inches="tight", facecolor="white")
plt.close()

plt.figure(figsize=(10, 5.4))
plt.plot(unreg_history.history["loss"], label="Unregularized training loss", lw=2)
plt.plot(unreg_history.history["val_loss"], label="Unregularized validation loss", lw=2)
plt.axvline(unreg_best - 1, color="black", ls="--", label=f"Lowest validation loss: epoch {unreg_best}")
plt.xlabel("Epoch")
plt.ylabel("Huber loss on log residual")
plt.title("Unregularized reference: validation loss exposes overfitting")
plt.legend()
plt.tight_layout()
plt.savefig(ASSETS / "w5_overfitting_reference.png", dpi=220, bbox_inches="tight", facecolor="white")
plt.close()

# Architecture diagram
fig, ax = plt.subplots(figsize=(11, 3.7))
ax.axis("off")
labels = [f"Input\n{X_fit.shape[1]} features", "Dense 32\nReLU + L2", "Dropout\n20%", "Dense 16\nReLU + L2", "Dropout\n15%", "Output 1\nlog correction", "Forecast\nlog lag 1 + correction"]
xs = np.linspace(.07, .93, len(labels))
colors = ["#d9edf7", "#ccebc5", "#fddbc7", "#ccebc5", "#fddbc7", "#ccebc5", "#decbe4"]
for i, (x, label, color) in enumerate(zip(xs, labels, colors)):
    ax.text(x, .5, label, ha="center", va="center", fontsize=10,
            bbox=dict(boxstyle="round,pad=.55", fc=color, ec="#444444", lw=1.2))
    if i < len(labels) - 1:
        ax.annotate("", xy=(xs[i+1]-.055, .5), xytext=(x+.055, .5), arrowprops=dict(arrowstyle="->", lw=1.3))
ax.set_title("Feed-forward network used for one-month-ahead state-level forecasting", fontsize=14, pad=18)
plt.tight_layout()
plt.savefig(ASSETS / "w5_architecture.png", dpi=220, bbox_inches="tight", facecolor="white")
plt.close()

# Test forecast chart
monthly = predictions.groupby("start_date").agg(actual=("covid_deaths", "sum"), neural=("prediction", "sum"), baseline=("lag1_covid", "sum")).reset_index()
plt.figure(figsize=(10.5, 5.3))
plt.plot(monthly.start_date, monthly.actual, marker="o", lw=2.4, label="Actual")
plt.plot(monthly.start_date, monthly.neural, marker="s", lw=2, label="Neural network")
plt.plot(monthly.start_date, monthly.baseline, marker="^", ls="--", lw=1.6, label="Previous-month baseline")
plt.xlabel("Forecast month")
plt.ylabel("COVID-19 deaths across 52 jurisdictions")
plt.title("Neural-network test forecasts, January–August 2023")
plt.legend()
plt.tight_layout()
plt.savefig(ASSETS / "w5_forecast_monthly.png", dpi=220, bbox_inches="tight", facecolor="white")
plt.close()

with open(ROOT / "analysis_summary_pre_deep.json", "r", encoding="utf-8") as f:
    summary = json.load(f)
summary["deep_learning"] = {
    "framework": f"TensorFlow {tf.__version__}",
    "fit_rows": int(len(fit_data)),
    "validation_rows": int(len(val_data)),
    "test_rows": int(len(test_data)),
    "input_features": int(X_fit.shape[1]),
    "parameters": int(regularized_model.count_params()),
    "best_epoch": best_epoch,
    "stop_epoch": stop_epoch,
    "best_validation_loss": round(float(min(history.history["val_loss"])), 5),
    "unregularized_best_epoch": unreg_best,
    "unregularized_final_train_loss": round(float(unreg_history.history["loss"][-1]), 5),
    "unregularized_final_validation_loss": round(float(unreg_history.history["val_loss"][-1]), 5),
    "unregularized_final_gap": round(unreg_final_gap, 5),
    "model_metrics": {k: round(v, 3) for k, v in model_metrics.items()},
    "baseline_metrics": {k: round(v, 3) for k, v in baseline_metrics.items()},
    "worst_cases": [
        {"state": row.state, "month": row.start_date.strftime("%b %Y"), "actual": round(float(row["covid_deaths"]), 1), "pred": round(float(row.prediction), 1), "abs_error": round(float(row.absolute_error), 1)}
        for _, row in predictions.nlargest(8, "absolute_error").iterrows()
    ],
}
with open(ROOT / "analysis_summary.json", "w", encoding="utf-8") as f:
    json.dump(summary, f, indent=2)
print(json.dumps(summary["deep_learning"], indent=2))
