Banka Pazarlama Kampanyalarında Müşteri Tahmin ve Açıklanabilirlik Sistemi
Bu proje, BLM308 Veri Madenciliği dersi final projesi kapsamında hazırlanmıştır. Çalışmanın amacı, banka pazarlama kampanyalarında bir müşterinin vadeli mevduata abone olma ihtimalini tahmin etmek ve modelin verdiği kararları açıklanabilir hale getirmektir.

Proje kapsamında UCI Bank Marketing veri seti kullanılmıştır. Veri setinde müşterilere ait profil bilgileri, kampanya geçmişi, iletişim bilgileri ve ekonomik göstergeler yer almaktadır. Hedef değişken, müşterinin kampanya sonucunda vadeli mevduata abone olup olmamasıdır.

Proje Konusu
Seçilen proje başlığı:

Interpretable Machine Learning / Explainable AI

Bu kapsamda kara kutu olarak değerlendirilebilecek makine öğrenmesi modelleri kurulmuş ve model kararları açıklanabilir yapay zeka yöntemleriyle yorumlanmıştır. Özellikle Random Forest modeli ana model olarak kullanılmış, modelin genel kararlarını etkileyen değişkenler ve tek müşteri özelindeki tahmin gerekçeleri incelenmiştir.

Kullanılan Veri Seti
Projede kullanılan veri seti UCI Bank Marketing Dataset veri setidir.

Veri seti, banka müşterilerine yapılan pazarlama kampanyalarına ait kayıtları içermektedir. Amaç, müşterinin kampanya sonunda vadeli mevduata abone olup olmayacağını tahmin etmektir.

Veri setinde örnek olarak şu değişkenler bulunmaktadır:

Yaş
Meslek
Medeni durum
Eğitim durumu
Konut kredisi bilgisi
İhtiyaç kredisi bilgisi
İletişim kanalı
Kampanya kapsamında arama sayısı
Önceki kampanya sonucu
Ekonomik göstergeler
Kullanılan Yöntemler
Projede veri madenciliği süreci CRISP-DM yaklaşımına uygun şekilde ele alınmıştır.

Genel süreç şu adımlardan oluşmaktadır:

İş probleminin tanımlanması
Veri setinin incelenmesi
Veri ön işleme adımlarının uygulanması
Makine öğrenmesi modellerinin eğitilmesi
Model performanslarının karşılaştırılması
Açıklanabilirlik çıktılarının üretilmesi
Türkçe arayüz ile tahmin sisteminin hazırlanması
Bu yapı sayesinde proje yalnızca model eğitimiyle sınırlı kalmamış, veri hazırlığından uygulama arayüzüne kadar uçtan uca tamamlanmıştır.

Kullanılan Modeller
Projede farklı sınıflandırma modelleri denenmiştir:

Lojistik Regresyon
Karar Ağacı
Random Forest
Çok Katmanlı Algılayıcı
Random Forest modeli, performans ve açıklanabilirlik dengesi nedeniyle ana model olarak seçilmiştir. Model karşılaştırması yapılırken doğruluk, kesinlik, duyarlılık, F1 skoru ve ROC-AUC değerleri birlikte değerlendirilmiştir.

Açıklanabilir Yapay Zeka
Modelin sadece tahmin üretmesi yeterli görülmemiştir. Bu nedenle modelin hangi değişkenlerden etkilendiğini göstermek için açıklanabilir yapay zeka yaklaşımı kullanılmıştır.

Bu bölümde amaçlananlar:

Modelin genel kararlarını etkileyen değişkenleri görmek
Tek bir müşteri için tahmin sonucunu etkileyen faktörleri yorumlamak
Kara kutu model kararlarını daha anlaşılır hale getirmek
Bankacılık gibi kararların açıklanabilir olması gereken bir alanda model güvenilirliğini artırmak
Önemli Teknik Not
Veri setinde yer alan görüşme süresi değişkeni, telefon görüşmesi bittikten sonra bilinen bir bilgidir. Müşteri aranmadan önce bu bilgi bilinmediği için gerçek kampanya öncesi tahmin senaryosunda bu değişkenin kullanılması veri sızıntısı riski oluşturur.

Bu nedenle projede iki farklı senaryo değerlendirilmiştir:

Görüşme süresi dahil senaryo
Görüşme süresi hariç gerçekçi senaryo
Ana yorumlar, gerçek kullanım koşullarına daha uygun olduğu için görüşme süresi hariç senaryo üzerinden yapılmıştır.

Proje Dosyaları
Dosya	Açıklama
train.py	Veri ön işleme, model eğitimi ve performans karşılaştırması
explain.py	Model açıklanabilirliği için grafik ve çıktı üretimi
app.py	Streamlit tabanlı Türkçe tahmin arayüzü
metrics.csv	Model performans sonuçları
requirements.txt	Projede kullanılan Python kütüphaneleri
Kurulum
Projeyi çalıştırmak için önce gerekli kütüphaneler kurulmalıdır.

pip install -r requirements.txt
Modeli Eğitme
Model eğitimi ve performans çıktıları için:

python train.py
Açıklanabilirlik Çıktılarını Üretme
Model açıklamalarını ve ilgili grafikleri oluşturmak için:

python explain.py
Uygulamayı Çalıştırma
Türkçe arayüzü başlatmak için:

python -m streamlit run app.py
Uygulama açıldıktan sonra kullanıcı, müşteri bilgilerini girerek vadeli mevduata abone olma ihtimalini görebilir. Sonuçlar Türkçe olarak gösterilir ve modelin genel kararlarını etkileyen faktörler arayüzde sunulur.

Projenin Amacı
Bu proje ile yalnızca yüksek doğruluk değerine sahip bir model kurmak değil, aynı zamanda modelin kararlarını yorumlanabilir hale getirmek hedeflenmiştir. Bankacılık alanında karar destek sistemlerinin güvenilir olması için tahminlerin nedenleriyle birlikte sunulması önemlidir.

Bu nedenle proje, makine öğrenmesi modeli, açıklanabilirlik analizi ve kullanıcı arayüzünü bir araya getiren bütüncül bir çalışma olarak hazırlanmıştır.

Ekip
Sude ANLAŞ - 231041038
Doğan Fatih OĞUR - 231041048
Merve Naz ORAN - 231041058
Ders Bilgisi
BLM308 Veri Madenciliği
Final Projesi
Gedik Üniversitesi - Bilgisayar Mühendisliği
