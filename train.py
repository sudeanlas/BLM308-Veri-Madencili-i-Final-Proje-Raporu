import joblib

from src.config import MODEL_PATH, METRICS_PATH, MODELS_DIR, OUTPUT_DIR
from src.data import load_dataset
from src.modeling import evaluate_all_models, train_one_model, evaluate_pipeline


def main() -> None:
    MODELS_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    df = load_dataset()

    print("Tüm modeller değerlendiriliyor...")
    metrics = evaluate_all_models(df)
    metrics.to_csv(METRICS_PATH, index=False)
    print(metrics.to_string(index=False))

    print("\nGerçekçi senaryo için Random Forest modeli kaydediliyor: duration hariç")
    pipeline, X_train, X_test, y_train, y_test = train_one_model(
        df,
        model_name="Random Forest",
        include_duration=False,
    )
    scores = evaluate_pipeline(pipeline, X_test, y_test)
    joblib.dump(
        {
            "pipeline": pipeline,
            "X_train": X_train,
            "X_test": X_test,
            "y_train": y_train,
            "y_test": y_test,
            "scores": scores,
        },
        MODEL_PATH,
    )
    print(f"\nModel kaydedildi: {MODEL_PATH}")
    print(f"Gerçekçi Random Forest ROC-AUC: {scores['ROC-AUC']:.3f}")


if __name__ == "__main__":
    main()
