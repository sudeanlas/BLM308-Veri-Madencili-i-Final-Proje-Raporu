"""Türkçe arayüz ve model açıklama etiketleri."""

COLUMN_LABELS = {
    "age": "Yaş",
    "job": "Meslek",
    "marital": "Medeni durum",
    "education": "Eğitim durumu",
    "default": "Kredi temerrüt durumu",
    "housing": "Konut kredisi",
    "loan": "İhtiyaç kredisi",
    "contact": "İletişim kanalı",
    "month": "Son iletişim ayı",
    "day_of_week": "Son iletişim günü",
    "duration": "Görüşme süresi",
    "campaign": "Kampanya kapsamında arama sayısı",
    "pdays": "Önceki kampanyadan geçen gün sayısı",
    "previous": "Önceki kampanya temas sayısı",
    "poutcome": "Önceki kampanya sonucu",
    "emp.var.rate": "İstihdam değişim oranı",
    "cons.price.idx": "Tüketici fiyat endeksi",
    "cons.conf.idx": "Tüketici güven endeksi",
    "euribor3m": "Piyasa faiz göstergesi",
    "nr.employed": "İstihdam göstergesi",
}

VALUE_LABELS = {
    "admin.": "yönetici/asistan",
    "blue-collar": "işçi",
    "entrepreneur": "girişimci",
    "housemaid": "ev hizmetleri",
    "management": "yönetici",
    "retired": "emekli",
    "self-employed": "serbest meslek",
    "services": "hizmet sektörü",
    "student": "öğrenci",
    "technician": "teknisyen",
    "unemployed": "işsiz",
    "unknown": "bilinmiyor",
    "divorced": "boşanmış",
    "married": "evli",
    "single": "bekar",
    "basic.4y": "ilkokul 4 yıl",
    "basic.6y": "ilkokul 6 yıl",
    "basic.9y": "ortaokul",
    "high.school": "lise",
    "illiterate": "okur yazar değil",
    "professional.course": "mesleki eğitim",
    "university.degree": "üniversite mezunu",
    "no": "hayır",
    "yes": "evet",
    "cellular": "cep telefonu",
    "telephone": "sabit telefon",
    "jan": "ocak",
    "feb": "şubat",
    "mar": "mart",
    "apr": "nisan",
    "may": "mayıs",
    "jun": "haziran",
    "jul": "temmuz",
    "aug": "ağustos",
    "sep": "eylül",
    "oct": "ekim",
    "nov": "kasım",
    "dec": "aralık",
    "mon": "pazartesi",
    "tue": "salı",
    "wed": "çarşamba",
    "thu": "perşembe",
    "fri": "cuma",
    "failure": "başarısız",
    "nonexistent": "önceki temas yok",
    "success": "başarılı",
}


def column_label(column: str) -> str:
    return COLUMN_LABELS.get(column, column)


def value_label(value) -> str:
    return VALUE_LABELS.get(str(value), str(value))


def feature_label(feature_name: str) -> str:
    """One-hot dönüştürülmüş model değişkenlerini Türkçe gösterir."""
    for column in sorted(COLUMN_LABELS, key=len, reverse=True):
        prefix = f"{column}_"
        if feature_name.startswith(prefix):
            raw_value = feature_name[len(prefix):]
            return f"{column_label(column)}: {value_label(raw_value)}"
    return column_label(feature_name)
