import joblib
import pandas as pd
import streamlit as st

from src.config import MODEL_PATH
from src.data import load_dataset
from src.localization import column_label, value_label
from src.xai import random_forest_importance


st.set_page_config(
    page_title="Banka Pazarlama Tahmin Sistemi",
    page_icon="📈",
    layout="wide",
)


@st.cache_resource
def load_model_bundle():
    return joblib.load(MODEL_PATH)


@st.cache_data
def load_reference_data():
    df = load_dataset()
    return df.drop(columns=["y", "duration"])


st.title("Banka Pazarlama Kampanyalarında Müşteri Tahmin ve Açıklanabilirlik Sistemi")
st.caption("BLM308 Veri Madenciliği Final Projesi - Sude ANLAŞ, Doğan Fatih OĞUR ve Merve Naz ORAN")

if not MODEL_PATH.exists():
    st.error("Model dosyası bulunamadı. Önce terminalde `python train.py` komutunu çalıştırın.")
    st.stop()

bundle = load_model_bundle()
pipeline = bundle["pipeline"]
reference = load_reference_data()

st.write(
    "Bu sistem, banka pazarlama kampanyalarında bir müşterinin vadeli mevduata abone olma "
    "ihtimalini tahmin eder ve modelin kararlarını etkileyen temel faktörleri anlaşılır biçimde gösterir."
)

left, right = st.columns([1, 1], gap="large")

with left:
    st.subheader("Müşteri bilgileri")
    form_values = {}
    for column in reference.columns:
        if reference[column].dtype == "object" or str(reference[column].dtype).startswith("str"):
            options = sorted(reference[column].dropna().unique().tolist())
            mode_value = reference[column].mode()[0]
            default_index = options.index(mode_value) if mode_value in options else 0
            form_values[column] = st.selectbox(
                column_label(column),
                options,
                index=default_index,
                format_func=value_label,
            )
        else:
            min_value = float(reference[column].min())
            max_value = float(reference[column].max())
            mean_value = float(reference[column].mean())
            form_values[column] = st.number_input(
                column_label(column),
                value=mean_value,
                min_value=min_value,
                max_value=max_value,
            )

    predict_clicked = st.button("Tahmin oluştur", type="primary")

with right:
    st.subheader("Tahmin sonucu")
    if predict_clicked:
        sample = pd.DataFrame([form_values])
        probability = float(pipeline.predict_proba(sample)[0, 1])
        prediction = (
            "Vadeli mevduata abone olma ihtimali yüksek"
            if probability >= 0.5
            else "Vadeli mevduata abone olma ihtimali düşük"
        )
        st.metric("Model kararı", prediction)
        st.metric("Abone olma olasılığı", f"%{probability * 100:.1f}")
        st.write("Gerçekçi kampanya öncesi kullanım için görüşme süresi modele verilmez.")
    else:
        st.info("Sol taraftan müşteri bilgilerini seçip tahmin oluşturun.")

st.divider()
st.subheader("Modelin Genel Kararlarını Etkileyen Faktörler")
st.write(
    "Aşağıdaki tablo, eğitilen Random Forest modelinin kararlarında en fazla etkili olan "
    "değişkenleri gösterir."
)
importance = random_forest_importance(pipeline, top_n=10).rename(
    columns={"feature_tr": "Değişken", "importance": "Önem düzeyi"}
)
importance = importance[["Değişken", "Önem düzeyi"]]
st.dataframe(importance, use_container_width=True, hide_index=True)
st.bar_chart(importance.set_index("Değişken"))
