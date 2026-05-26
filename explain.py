import os

import joblib
import matplotlib

os.environ.setdefault("MPLCONFIGDIR", ".mplconfig")
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from src.config import FIGURES_DIR, MODEL_PATH
from src.xai import random_forest_importance, lime_explanation, best_positive_sample_index


def main() -> None:
    FIGURES_DIR.mkdir(exist_ok=True)
    bundle = joblib.load(MODEL_PATH)
    pipeline = bundle["pipeline"]
    X_train = bundle["X_train"]
    X_test = bundle["X_test"]

    importance = random_forest_importance(pipeline)
    importance.rename(
        columns={"feature_tr": "Değişken", "importance": "Önem düzeyi"}
    )[["Değişken", "Önem düzeyi"]].to_csv(
        FIGURES_DIR / "rf_feature_importance_table.csv", index=False
    )

    plt.figure(figsize=(8, 5))
    plt.barh(importance["feature_tr"][::-1], importance["importance"][::-1], color="#2563eb")
    plt.xlabel("Önem düzeyi")
    plt.title("Modelin genel kararlarını etkileyen faktörler")
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "rf_feature_importance_from_explain.png", dpi=180)
    plt.close()

    sample_index = best_positive_sample_index(pipeline, X_test)
    lime_df = lime_explanation(pipeline, X_train, X_test, sample_index=sample_index)
    lime_df.to_csv(FIGURES_DIR / "lime_explanation_table.csv", index=False)

    colors = ["#dc2626" if value < 0 else "#16a34a" for value in lime_df["Katkı"]]
    ordered = lime_df.sort_values("Katkı")
    plt.figure(figsize=(8, 5))
    plt.barh(ordered["Değişken"], ordered["Katkı"], color=colors)
    plt.axvline(0, color="#111827", linewidth=0.8)
    plt.xlabel("Abone olma ihtimaline katkı")
    plt.title("Yerel karar açıklaması")
    plt.tight_layout()
    plt.savefig(FIGURES_DIR / "lime_explanation_from_explain.png", dpi=180)
    plt.close()

    print("Açıklanabilirlik çıktıları üretildi:")
    print(FIGURES_DIR / "rf_feature_importance_from_explain.png")
    print(FIGURES_DIR / "lime_explanation_from_explain.png")


if __name__ == "__main__":
    main()
