#!/usr/bin/env python3
"""
DAM405 - Machine Learning Operations | Practical 4
Automated Training & Tuning with an Open-Source AutoML Tool (FLAML)

This script performs every task of the practical, in order, and prints a clearly
labelled block for each one so that each block can be screenshotted:

    Task 1  Load data and split it (80/20, random_state=42)
    Task 2  Baseline: plain LogisticRegression
    Task 3  AutoML: FLAML search inside a fixed time budget
    Task 4  Inspect the best model (estimator, hyperparameters, CV score)
    Task 5  Evaluate AutoML vs baseline on the held-out test set
    Task 6  Save the AutoML model with joblib (and prove it reloads)
    Extras  (--extras) optional experiment: other time budgets + roc_auc metric

Everything that is printed is also written to outputs/results.json, which
build_report.py turns into the final Word report.

Usage (from the project root):
    python code/practical4_automl.py                     # required tasks, 60 s budget
    python code/practical4_automl.py --extras            # + optional experiments
    python code/practical4_automl.py --time-budget 120   # bigger search
    python code/practical4_automl.py --data my.csv --target label   # your own dataset
    python code/practical4_automl.py --baseline-only     # no FLAML needed
"""
import argparse
import json
import platform
import sys
import time
from datetime import datetime
from pathlib import Path

import joblib
import numpy as np
import sklearn
from sklearn.datasets import load_breast_cancer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, roc_auc_score
from sklearn.model_selection import train_test_split

ROOT = Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def banner(title):
    line = "=" * 72
    print(f"\n{line}\n{title}\n{line}", flush=True)


def to_jsonable(obj):
    """Make numpy scalars / arrays JSON serialisable."""
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        return float(obj)
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    return str(obj)


def auc_of(model, X, y, n_classes):
    """Test ROC-AUC for binary or multiclass (one-vs-rest)."""
    proba = model.predict_proba(X)
    if n_classes == 2:
        return float(roc_auc_score(y, proba[:, 1]))
    return float(roc_auc_score(y, proba, multi_class="ovr"))


def load_data(args):
    """Task 1 data loading. Default = the sheet's breast-cancer example."""
    if args.data:
        import pandas as pd

        df = pd.read_csv(args.data)
        if args.target not in df.columns:
            sys.exit(f"Target column '{args.target}' not found. Columns: {list(df.columns)}")
        y = df[args.target]
        X = df.drop(columns=[args.target])
        name = f"{Path(args.data).name} (target = {args.target})"
        return X, y, name
    X, y = load_breast_cancer(return_X_y=True)
    return X, y, "sklearn built-in breast_cancer"


def make_baseline(X_train):
    """Plain LogisticRegression exactly as in the sheet for numeric data.
    If a custom CSV has text / missing values, wrap it in a small preprocessing
    pipeline so the baseline can still be trained (still a 'plain' model)."""
    try:
        import pandas as pd
    except ImportError:  # numpy-only default path
        return LogisticRegression(max_iter=2000)
    if not isinstance(X_train, pd.DataFrame):
        return LogisticRegression(max_iter=2000)
    from sklearn.compose import ColumnTransformer
    from sklearn.impute import SimpleImputer
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import OneHotEncoder, StandardScaler

    cat = X_train.select_dtypes(exclude="number").columns.tolist()
    num = X_train.select_dtypes(include="number").columns.tolist()
    if not cat and not X_train.isna().any().any():
        return LogisticRegression(max_iter=2000)
    pre = ColumnTransformer(
        [
            ("num", Pipeline([("imp", SimpleImputer(strategy="median")), ("sc", StandardScaler())]), num),
            ("cat", Pipeline([("imp", SimpleImputer(strategy="most_frequent")),
                              ("oh", OneHotEncoder(handle_unknown="ignore"))]), cat),
        ]
    )
    return Pipeline([("pre", pre), ("clf", LogisticRegression(max_iter=2000))])


def run_automl(X_train, y_train, metric, budget, seed, log_file):
    from flaml import AutoML

    automl = AutoML()
    t0 = time.perf_counter()
    automl.fit(
        X_train,
        y_train,
        task="classification",
        metric=metric,
        time_budget=budget,  # seconds
        seed=seed,
        log_file_name=str(log_file),
    )
    return automl, time.perf_counter() - t0


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="DAM405 Practical 4 - FLAML AutoML vs baseline")
    ap.add_argument("--time-budget", type=int, default=60, help="FLAML time budget in seconds (sheet: 60)")
    ap.add_argument("--metric", default="accuracy", help="FLAML metric (sheet: accuracy)")
    ap.add_argument("--seed", type=int, default=42, help="seed passed to FLAML")
    ap.add_argument("--data", help="optional CSV file to use instead of breast_cancer")
    ap.add_argument("--target", help="target column name (required with --data)")
    ap.add_argument("--extras", action="store_true", help="also run the optional experiments")
    ap.add_argument("--baseline-only", action="store_true", help="run Tasks 1-2 only (no FLAML)")
    ap.add_argument("--outdir", default=str(ROOT / "outputs"), help="where results/model/log are written")
    args = ap.parse_args()
    if args.data and not args.target:
        sys.exit("--target is required when --data is used")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    results = {"started": datetime.now().isoformat(timespec="seconds")}

    # ---------------- environment ----------------
    banner("ENVIRONMENT")
    env = {"python": platform.python_version(), "platform": platform.platform(),
           "scikit_learn": sklearn.__version__, "joblib": joblib.__version__}
    try:
        import flaml
        env["flaml"] = flaml.__version__
    except ImportError:
        env["flaml"] = None
    for k, v in env.items():
        print(f"{k:13}: {v}")
    results["environment"] = env
    if env["flaml"] is None and not args.baseline_only:
        sys.exit("\nFLAML is not installed. Run:  pip install \"flaml[automl]\" scikit-learn\n"
                 "(or use --baseline-only to run Tasks 1-2 without it)")

    # ---------------- Task 1 ----------------
    banner("TASK 1 - Load data and split")
    X, y, ds_name = load_data(args)
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    classes, counts = np.unique(np.asarray(y), return_counts=True)
    n_classes = len(classes)
    print("Dataset          :", ds_name)
    print("Samples x features:", X.shape)
    print("Classes / counts :", dict(zip(map(str, classes), map(int, counts))))
    print("X_train / X_test :", X_train.shape, X_test.shape)
    results["dataset"] = {
        "name": ds_name, "n_samples": int(X.shape[0]), "n_features": int(X.shape[1]),
        "n_classes": int(n_classes), "class_counts": dict(zip(map(str, classes), map(int, counts))),
        "n_train": int(len(y_train)), "n_test": int(len(y_test)),
        "split": "train_test_split(test_size=0.2, random_state=42)",
    }

    # ---------------- Task 2 ----------------
    banner("TASK 2 - Baseline (LogisticRegression, max_iter=2000)")
    t0 = time.perf_counter()
    base = make_baseline(X_train).fit(X_train, y_train)
    base_time = time.perf_counter() - t0
    base_acc = accuracy_score(y_test, base.predict(X_test))
    print("Baseline accuracy:", round(base_acc, 4))
    print(f"Baseline training time: {base_time:.3f} s")
    results["baseline"] = {"model": "LogisticRegression(max_iter=2000)", "test_accuracy": float(base_acc),
                           "train_seconds": round(base_time, 4),
                           "test_correct": int(round(base_acc * len(y_test)))}

    if args.extras or args.baseline_only:
        # Supplementary (NOT part of the sheet): the plain baseline above does not converge
        # on unscaled data (see the ConvergenceWarning), so also report a scaled version.
        from sklearn.pipeline import make_pipeline
        from sklearn.preprocessing import StandardScaler
        scaled = make_pipeline(StandardScaler(), LogisticRegression(max_iter=2000)).fit(X_train, y_train)
        scaled_acc = accuracy_score(y_test, scaled.predict(X_test))
        print(f"[supplementary] Scaled baseline (StandardScaler + LogisticRegression): {scaled_acc:.4f}")
        results["scaled_baseline"] = {"model": "StandardScaler + LogisticRegression(max_iter=2000)",
                                      "test_accuracy": float(scaled_acc),
                                      "test_correct": int(round(scaled_acc * len(y_test)))}

    if args.baseline_only:
        (outdir / "results.json").write_text(json.dumps(results, indent=2, default=to_jsonable))
        print(f"\n--baseline-only: stopped after Task 2. Results -> {outdir / 'results.json'}")
        return

    # ---------------- Task 3 ----------------
    banner(f"TASK 3 - Run AutoML (FLAML, metric='{args.metric}', time_budget={args.time_budget}s)")
    automl, auto_time = run_automl(X_train, y_train, args.metric, args.time_budget, args.seed,
                                   outdir / "flaml_search.log")
    print(f"\nAutoML search finished in {auto_time:.1f} s (wall clock)")

    # ---------------- Task 4 ----------------
    banner("TASK 4 - Inspect the best model")
    best_cv = 1 - automl.best_loss
    print("Best estimator :", automl.best_estimator)
    print("Best config    :", automl.best_config)
    print("Best CV score  :", best_cv)
    extra_info = {}
    for attr in ("best_iteration", "best_config_train_time"):
        if hasattr(automl, attr):
            extra_info[attr] = getattr(automl, attr)
            print(f"{attr:15}:", extra_info[attr])
    print("Fitted model   :", automl.model)
    results["automl"] = {
        "metric": args.metric, "time_budget_s": args.time_budget, "seed": args.seed,
        "wall_clock_s": round(auto_time, 2), "best_estimator": automl.best_estimator,
        "best_config": automl.best_config, "best_cv_score": float(best_cv),
        **{k: v for k, v in extra_info.items()},
        "fitted_model": str(automl.model),
    }

    # ---------------- Task 5 ----------------
    banner("TASK 5 - Evaluate on the held-out test set")
    auto_acc = accuracy_score(y_test, automl.predict(X_test))
    print("AutoML test accuracy  :", round(auto_acc, 4))
    print("Baseline test accuracy:", round(base_acc, 4))
    n_test = len(y_test)
    diff = auto_acc - base_acc
    print(f"Difference (AutoML - baseline): {diff:+.4f}  "
          f"= {round(diff * n_test):+d} test sample(s) out of {n_test}")
    print(f"Extra compute: baseline {base_time:.2f} s vs AutoML {auto_time:.1f} s")
    results["comparison"] = {
        "baseline_test_accuracy": float(base_acc), "automl_test_accuracy": float(auto_acc),
        "difference": float(diff), "difference_in_samples": int(round(diff * n_test)),
        "n_test": int(n_test), "automl_test_correct": int(round(auto_acc * n_test)),
        "baseline_test_correct": int(round(base_acc * n_test)),
        "baseline_train_seconds": round(base_time, 4), "automl_wall_clock_s": round(auto_time, 2),
        "cv_score_vs_test_gap": float(best_cv - auto_acc),
    }

    # ---------------- Task 6 ----------------
    banner("TASK 6 - Save the best model")
    model_path = outdir / "automl_model.joblib"
    joblib.dump(automl, model_path)
    print("Saved", model_path.name)
    size_kb = model_path.stat().st_size / 1024
    print(f"File size: {size_kb:.1f} KB")
    reloaded = joblib.load(model_path)
    same = bool(np.array_equal(reloaded.predict(X_test), automl.predict(X_test)))
    print("Reload check - predictions identical after joblib.load():", same)
    results["saved_model"] = {"file": model_path.name, "size_kb": round(size_kb, 1),
                              "reload_predictions_identical": same}

    # ---------------- Extras ----------------
    if args.extras:
        banner("OPTIONAL EXPERIMENTS - other time budgets and the roc_auc metric")
        extras = {"budget_sweep": [], "roc_auc": None}

        budgets = sorted({10, 30, args.time_budget, 2 * args.time_budget})
        for b in budgets:
            if b == args.time_budget and args.metric == "accuracy":
                m, t = automl, auto_time  # reuse the main run
            else:
                print(f"\n--- accuracy metric, time_budget={b}s ---")
                m, t = run_automl(X_train, y_train, "accuracy", b, args.seed, outdir / f"flaml_search_{b}s.log")
            acc = accuracy_score(y_test, m.predict(X_test))
            row = {"time_budget_s": b, "wall_clock_s": round(t, 1), "best_estimator": m.best_estimator,
                   "best_cv_score": float(1 - m.best_loss), "test_accuracy": float(acc)}
            extras["budget_sweep"].append(row)
            print(f"budget={b:>4}s  estimator={m.best_estimator:<12} cv={row['best_cv_score']:.4f}  "
                  f"test_acc={acc:.4f}  wall={t:.1f}s")

        auc_metric = "roc_auc" if n_classes == 2 else "roc_auc_ovr"
        print(f"\n--- metric='{auc_metric}', time_budget={args.time_budget}s ---")
        m, t = run_automl(X_train, y_train, auc_metric, args.time_budget, args.seed, outdir / "flaml_search_roc_auc.log")
        extras["roc_auc"] = {
            "metric": auc_metric, "time_budget_s": args.time_budget, "wall_clock_s": round(t, 1),
            "best_estimator": m.best_estimator, "best_cv_score": float(1 - m.best_loss),
            "test_roc_auc": auc_of(m, X_test, y_test, n_classes),
            "test_accuracy": float(accuracy_score(y_test, m.predict(X_test))),
            "baseline_test_roc_auc": auc_of(base, X_test, y_test, n_classes),
        }
        print(json.dumps(extras["roc_auc"], indent=2, default=to_jsonable))
        results["extras"] = extras

    results["finished"] = datetime.now().isoformat(timespec="seconds")
    (outdir / "results.json").write_text(json.dumps(results, indent=2, default=to_jsonable))
    banner("DONE")
    print("Results written to", outdir / "results.json")
    print("Next: python code/build_report.py   (builds the Word report from these numbers)")


if __name__ == "__main__":
    main()
