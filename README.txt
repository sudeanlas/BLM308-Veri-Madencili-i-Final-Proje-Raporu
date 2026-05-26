BLM308 Veri Madenciliği Final Projesi

Proje başlığı:
Banka Pazarlama Kampanyalarında Müşteri Tahmin ve Açıklanabilirlik Sistemi

Öğrenciler:
- Sude ANLAŞ - 231041038
- Doğan Fatih OĞUR - 231041048
- Merve Naz ORAN - 231041058

Proje türü:
Interpretable ML / XAI

Teslim içeriği:
- rapor.docx: Ana Word raporu
- rapor.tex: Bonus için LaTeX rapor kaynağı
- sunum_beamer.tex: Gönüllü sunum için Beamer sunum taslağı
- metrics.csv: Deney metrikleri
- app.py: Türkçe tahmin arayüzü
- train.py: Model eğitimi ve metrik üretimi
- explain.py: Açıklanabilirlik grafikleri
- src/: Veri yükleme, modelleme ve Türkçe etiketleme kodları
- veri/: Bank Marketing veri seti
- figures/: Grafik çıktıları
- models/: Eğitilmiş Random Forest modeli

VS Code ile çalıştırma:
1. Proje klasörünü VS Code ile açın.
2. Terminalden paketleri kurun:
   python -m pip install -r requirements.txt
3. Modelleri eğitin:
   python train.py
4. Açıklanabilirlik çıktılarını üretin:
   python explain.py
5. Tahmin arayüzünü açın:
   python -m streamlit run app.py

Proje özeti:
UCI Bank Marketing veri seti kullanılarak müşterinin vadeli mevduata abone olup olmayacağı tahmin edilmiştir. Lojistik Regresyon, Karar Ağacı, Random Forest ve Çok Katmanlı Algılayıcı modelleri karşılaştırılmıştır. Random Forest ve Çok Katmanlı Algılayıcı kara kutu model olarak ele alınmış, Random Forest kararları açıklanabilirlik yaklaşımıyla yorumlanmıştır. Üç kişilik ekip kapsamı için McNemar istatistiksel karşılaştırması ve daha geniş değerlendirme bölümü eklenmiştir.

En önemli teknik not:
duration değişkeni telefon görüşmesi bittikten sonra bilinen bir bilgi olduğu için kampanya öncesi kullanımda veri sızıntısı riski taşır. Bu nedenle raporda duration dahil ve duration hariç iki senaryo karşılaştırılmıştır.
