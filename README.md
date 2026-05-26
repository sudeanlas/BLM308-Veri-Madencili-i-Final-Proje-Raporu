# BLM308 Veri Madenciliği Final Projesi

## Proje Başlığı

Banka Pazarlama Kampanyalarında Müşteri Tahmin ve Açıklanabilirlik Sistemi

## Öğrenciler

- Sude ANLAŞ - 231041038
- Doğan Fatih OĞUR - 231041048
- Merve Naz ORAN - 231041058

## Proje Türü

Interpretable ML / XAI

## Proje Amacı

Bu proje, banka pazarlama kampanyalarında bir müşterinin vadeli mevduata abone olma ihtimalini tahmin etmeyi amaçlar. Çalışmada yalnızca tahmin skoru üretilmez; modelin kararlarında hangi müşteri, kampanya ve ekonomik göstergelerin etkili olduğu da açıklanır.

Bu yönüyle proje, ders yönergesindeki açıklanabilir yapay zeka başlığına uygundur. Random Forest ve MLP gibi kara kutu modeller kurulmuş, Random Forest modeli için kararları etkileyen faktörler açıklanabilir hale getirilmiştir.

## Kullanılan Veri Seti

Veri seti UCI Bank Marketing veri setidir. Veri seti, bir bankanın telefonla pazarlama kampanyalarına ait müşteri ve kampanya bilgilerini içerir.

Hedef değişken müşterinin vadeli mevduata abone olup olmadığını gösterir. Arayüzde bu hedef, kullanıcıya İngilizce sınıf adlarıyla değil, şu şekilde gösterilir:

- Vadeli mevduata abone olma ihtimali düşük
- Vadeli mevduata abone olma ihtimali yüksek

## Kullanılan Yöntem

Proje CRISP-DM sürecine göre hazırlanmıştır:

1. İş anlayışı
2. Veri anlayışı
3. Veri hazırlığı
4. Modelleme
5. Değerlendirme
6. İş değeri ve açıklanabilirlik yorumu

## Kullanılan Modeller

- Lojistik Regresyon
- Karar Ağacı
- Random Forest
- Çok Katmanlı Algılayıcı

Random Forest modeli, performans ve açıklanabilirlik dengesi nedeniyle uygulama arayüzünde kullanılan ana modeldir. Üç kişilik proje kapsamına uygun olarak değerlendirme kısmı istatistiksel model karşılaştırmasıyla genişletilmiştir.

## Açıklanabilirlik Yaklaşımı

Modelin genel kararlarını etkileyen faktörler, Random Forest değişken önem düzeyleriyle gösterilir. Ayrıca `explain.py` dosyası çalıştırıldığında yerel karar açıklaması da üretilir.

Rapor bölümünde SHAP ve LIME yaklaşımları açıklanmıştır. Bu yaklaşımlar, modelin kararlarının hangi değişkenlerden etkilendiğini yorumlamak için kullanılır.

## Önemli Akademik Not

`duration` değişkeni telefon görüşmesi bittikten sonra bilinen görüşme süresidir. Müşteri aranmadan önce bu bilgi bilinmez. Bu nedenle gerçek kampanya öncesi kullanımda bu değişken modele verilmez.

Projede iki senaryo karşılaştırılmıştır:

- `duration` dahil: modelin üst sınır performansını gösterir.
- `duration` hariç: gerçek uygulama için daha doğru senaryodur.

Bu ayrım, veri sızıntısı farkındalığı gösterdiği için projenin en önemli teknik yorumlarından biridir.

## VS Code ile Çalıştırma

1. Bu klasörü Visual Studio Code ile açın.
2. Terminali açın.
3. Gerekli paketleri kurun:

```powershell
python -m pip install -r requirements.txt
```

4. Modelleri eğitin ve metrikleri üretin:

```powershell
python train.py
```

5. Açıklanabilirlik grafiklerini üretin:

```powershell
python explain.py
```

6. Tahmin arayüzünü açın:

```powershell
python -m streamlit run app.py
```

Komut çalıştıktan sonra terminalde verilen yerel adres tarayıcıda açılır. Açılan sayfada müşteri bilgileri seçilerek model kararı ve abone olma olasılığı görüntülenebilir.

## Klasör Yapısı

```text
.
+-- app.py
+-- train.py
+-- explain.py
+-- requirements.txt
+-- src/
|   +-- config.py
|   +-- data.py
|   +-- localization.py
|   +-- modeling.py
|   +-- xai.py
+-- veri/
|   +-- bank-additional-full.csv
+-- figures/
|   +-- rf_feature_importance_from_explain.png
|   +-- lime_explanation_from_explain.png
+-- models/
|   +-- random_forest_duration_haric.joblib
+-- teslim/
    +-- rapor.docx
    +-- rapor.tex
    +-- sunum_beamer.tex
    +-- metrics.csv
    +-- README.txt
```

## Hocaya Gösterilecek Uygulamalı Akış

1. `python train.py` komutu çalıştırılır.
2. Terminalde model performansları gösterilir.
3. `python explain.py` komutu çalıştırılır.
4. Açıklanabilirlik grafikleri `figures/` klasöründe oluşur.
5. `python -m streamlit run app.py` komutu ile tahmin arayüzü açılır.
6. Arayüzde müşteri bilgileri seçilir, sistemin karar sonucu ve etkili faktörler gösterilir.
