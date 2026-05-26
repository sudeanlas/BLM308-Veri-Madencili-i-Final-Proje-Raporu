import numpy as np
import pandas as pd

from src.localization import feature_label
from src.modeling import feature_names


def random_forest_importance(pipeline, top_n: int = 12) -> pd.DataFrame:
    names = feature_names(pipeline.named_steps["prep"])
    values = pipeline.named_steps["model"].feature_importances_
    result = (
        pd.DataFrame({"feature": names, "importance": values})
        .sort_values("importance", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    result["feature_tr"] = result["feature"].map(feature_label)
    return result


def lime_explanation(pipeline, X_train, X_test, sample_index: int = 0, top_n: int = 8) -> pd.DataFrame:
    from lime.lime_tabular import LimeTabularExplainer

    preprocessor = pipeline.named_steps["prep"]
    model = pipeline.named_steps["model"]
    names = feature_names(preprocessor)
    translated_names = [feature_label(name) for name in names]
    X_train_transformed = preprocessor.transform(X_train)
    X_test_transformed = preprocessor.transform(X_test)

    explainer = LimeTabularExplainer(
        X_train_transformed,
        feature_names=translated_names,
        class_names=["abone olmaz", "abone olur"],
        discretize_continuous=True,
        random_state=42,
    )
    explanation = explainer.explain_instance(
        X_test_transformed[sample_index],
        model.predict_proba,
        num_features=top_n,
    )
    return pd.DataFrame(explanation.as_list(), columns=["Değişken", "Katkı"])


def best_positive_sample_index(pipeline, X_test) -> int:
    probabilities = pipeline.predict_proba(X_test)[:, 1]
    return int(np.argmax(probabilities))
