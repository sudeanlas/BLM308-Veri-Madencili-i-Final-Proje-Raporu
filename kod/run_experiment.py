# BLM308 Final Projesi - Deney kodu
# Çalıştırma: python run_experiment.py
from pathlib import Path
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.ensemble import RandomForestClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.metrics import classification_report, roc_auc_score

df = pd.read_csv(Path('../veri/bank-additional-full.csv'), sep=';')
y = (df['y'] == 'yes').astype(int)
X = df.drop(columns=['y', 'duration'])  # kampanya öncesi gerçekçi senaryo
cat_cols = X.select_dtypes(include=['object']).columns.tolist()
num_cols = X.select_dtypes(exclude=['object']).columns.tolist()

prep = ColumnTransformer([
    ('cat', Pipeline([('imp', SimpleImputer(strategy='most_frequent')),
                      ('ohe', OneHotEncoder(handle_unknown='ignore'))]), cat_cols),
    ('num', Pipeline([('imp', SimpleImputer(strategy='median')),
                      ('scaler', StandardScaler())]), num_cols),
])

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=.25, stratify=y, random_state=42)
models = {
    'Random Forest': RandomForestClassifier(n_estimators=180, max_depth=12, class_weight='balanced_subsample', random_state=42, n_jobs=-1),
    'MLP': MLPClassifier(hidden_layer_sizes=(48, 24), early_stopping=True, random_state=42, max_iter=90),
}

for name, model in models.items():
    pipe = Pipeline([('prep', prep), ('model', model)])
    pipe.fit(X_train, y_train)
    pred = pipe.predict(X_test)
    prob = pipe.predict_proba(X_test)[:, 1]
    print('\n', name)
    print(classification_report(y_test, pred, target_names=['no', 'yes']))
    print('ROC-AUC:', roc_auc_score(y_test, prob))
