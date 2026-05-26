import pandas as pd

from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    roc_auc_score,
    confusion_matrix,
)
from sklearn.model_selection import train_test_split
from sklearn.neural_network import MLPClassifier
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.tree import DecisionTreeClassifier

from src.config import RANDOM_STATE
from src.data import split_features_target


def build_preprocessor(X: pd.DataFrame) -> ColumnTransformer:
    categorical_columns = X.select_dtypes(include=["object", "string"]).columns.tolist()
    numeric_columns = X.select_dtypes(exclude=["object", "string"]).columns.tolist()

    return ColumnTransformer(
        transformers=[
            (
                "cat",
                Pipeline(
                    steps=[
                        ("imputer", SimpleImputer(strategy="most_frequent")),
                        ("onehot", OneHotEncoder(handle_unknown="ignore", sparse_output=False)),
                    ]
                ),
                categorical_columns,
            ),
            (
                "num",
                Pipeline(
                    steps=[
                        ("imputer", SimpleImputer(strategy="median")),
                        ("scaler", StandardScaler()),
                    ]
                ),
                numeric_columns,
            ),
        ]
    )


def get_models() -> dict:
    return {
        "Logistic Regression": LogisticRegression(
            max_iter=1000,
            class_weight="balanced",
            random_state=RANDOM_STATE,
        ),
        "Decision Tree": DecisionTreeClassifier(
            max_depth=8,
            class_weight="balanced",
            random_state=RANDOM_STATE,
        ),
        "Random Forest": RandomForestClassifier(
            n_estimators=180,
            max_depth=12,
            min_samples_leaf=5,
            class_weight="balanced_subsample",
            n_jobs=-1,
            random_state=RANDOM_STATE,
        ),
        "MLP": MLPClassifier(
            hidden_layer_sizes=(48, 24),
            alpha=0.001,
            max_iter=90,
            early_stopping=True,
            random_state=RANDOM_STATE,
        ),
    }


def train_test_data(df: pd.DataFrame, include_duration: bool = False):
    X, y = split_features_target(df, include_duration=include_duration)
    return train_test_split(
        X,
        y,
        test_size=0.25,
        stratify=y,
        random_state=RANDOM_STATE,
    )


def train_one_model(df: pd.DataFrame, model_name: str = "Random Forest", include_duration: bool = False):
    X_train, X_test, y_train, y_test = train_test_data(df, include_duration=include_duration)
    preprocessor = build_preprocessor(X_train)
    model = get_models()[model_name]
    pipeline = Pipeline(steps=[("prep", preprocessor), ("model", model)])
    pipeline.fit(X_train, y_train)
    return pipeline, X_train, X_test, y_train, y_test


def evaluate_pipeline(pipeline, X_test, y_test) -> dict:
    y_pred = pipeline.predict(X_test)
    y_prob = pipeline.predict_proba(X_test)[:, 1]

    return {
        "Accuracy": accuracy_score(y_test, y_pred),
        "Precision": precision_score(y_test, y_pred, zero_division=0),
        "Recall": recall_score(y_test, y_pred, zero_division=0),
        "F1": f1_score(y_test, y_pred, zero_division=0),
        "ROC-AUC": roc_auc_score(y_test, y_prob),
        "Confusion Matrix": confusion_matrix(y_test, y_pred),
    }


def evaluate_all_models(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for include_duration, scenario in [(True, "duration dahil"), (False, "duration hariç")]:
        for model_name in get_models():
            pipeline, _, X_test, _, y_test = train_one_model(
                df,
                model_name=model_name,
                include_duration=include_duration,
            )
            scores = evaluate_pipeline(pipeline, X_test, y_test)
            rows.append(
                {
                    "Senaryo": scenario,
                    "Model": model_name,
                    "Accuracy": scores["Accuracy"],
                    "Precision": scores["Precision"],
                    "Recall": scores["Recall"],
                    "F1": scores["F1"],
                    "ROC-AUC": scores["ROC-AUC"],
                }
            )
    return pd.DataFrame(rows)


def feature_names(preprocessor) -> list[str]:
    names = []
    for transformer_name, transformer, columns in preprocessor.transformers_:
        if transformer_name == "cat":
            onehot = transformer.named_steps["onehot"]
            names.extend(onehot.get_feature_names_out(columns))
        elif transformer_name == "num":
            names.extend(columns)
    return list(names)
