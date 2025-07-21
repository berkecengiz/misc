"""
sendika.py - Temizlenmiş veri oluşturmak için script

Kullanım:
Terminalde aşağıdaki komutu çalıştırın:

    python sendika.py

Gerekli dosyalar:
- files/4.csv : Kaynak CSV dosyası
- files/unvan_il_whitelist.xlsx : Ünvan-İl eşleşme dosyası

Çıktılar:
- files/temizlenmis_veri.xlsx : Temizlenmiş ve eşleşmiş kayıtlar
- files/eslesemeyen_kayitlar.xlsx : Eşleşemeyen kayıtlar

Not: Pandas ve openpyxl kütüphaneleri kurulu olmalıdır.
"""


import os
import sys
import re
import pandas as pd


CSV_PATH = 'files/4.csv'
WHITELIST_PATH = 'files/unvan_il_whitelist.xlsx'
OUTPUT_PATH = 'files/temizlenmis_veri.xlsx'
UNMATCHED_PATH = 'files/eslesemeyen_kayitlar.xlsx'


ILLER_LISTESI = [
    "Adana", "Adıyaman", "Afyonkarahisar", "Ağrı", "Aksaray", "Amasya", "Ankara", "Antalya", "Artvin", "Aydın",
    "Balıkesir", "Bartın", "Batman", "Bayburt", "Bilecik", "Bingöl", "Bitlis", "Bolu", "Burdur", "Bursa",
    "Çanakkale", "Çankırı", "Çorum", "Denizli", "Diyarbakır", "Düzce", "Edirne", "Elazığ", "Erzincan", "Erzurum",
    "Eskişehir", "Gaziantep", "Giresun", "Gümüşhane", "Hakkari", "Hatay", "Iğdır", "Isparta", "İstanbul", "İzmir",
    "Kahramanmaraş", "Karabük", "Karaman", "Kars", "Kastamonu", "Kayseri", "Kilis", "Kırıkkale", "Kırklareli",
    "Kırşehir", "Kocaeli", "Konya", "Kütahya", "Malatya", "Manisa", "Mardin", "Mersin", "Muğla", "Muş", "Nevşehir",
    "Niğde", "Ordu", "Osmaniye", "Rize", "Sakarya", "Samsun", "Siirt", "Sinop", "Sivas", "Şanlıurfa", "Şırnak",
    "Tekirdağ", "Tokat", "Trabzon", "Tunceli", "Uşak", "Van", "Yalova", "Yozgat", "Zonguldak", "Ardahan"
]

KISA_IL = {
    "C.KALE": "Çanakkale", "Ç.KALE": "Çanakkale", "ARD.": "Ardahan", "ANK.": "Ankara", "ANT.": "Antalya",
    "GAZ.": "Gaziantep", "IST.": "İstanbul", "IZM.": "İzmir", "T.DAĞ": "Tekirdağ", "Ş.URFA": "Şanlıurfa", "URFA": "Şanlıurfa"
}

ILCE_IL_HARITA = {
    "SANDIKLI": "Afyonkarahisar", "SERİK": "Antalya", "ESPİYE": "Giresun", "ÜMRANİYE": "İstanbul",
    "SARIYER": "İstanbul", "ÇANKAYA": "Ankara", "BODRUM": "Muğla", "FETHİYE": "Muğla", "GEBZE": "Kocaeli",
    "TORBALI": "İzmir", "CİZRE": "Şırnak", "SALİHLİ": "Manisa", "GÜLNAR": "Mersin", "YATAĞAN": "Muğla",
    "KIZILTEPE": "Mardin", "EYYUBİYE": "Şanlıurfa", "GAZİEMİR": "İzmir", "ADAPAZARI": "Kocaeli", "ULA": "İzmir",
    "ICEL": "Mersin", "MANAVGAT": "Antalya", "AFYON": "Afyonkarahisar", "SIVEREK": "Şanlıurfa", "SOMA": "Manisa",
    "AKSEHIR": "Konya", "TURGUTLU": "Manisa", "SURUC": "Şanlıurfa", "ORTACA": "Muğla", "KALKAN": "Antalya",
    "MILAS": "Muğla", "TARSUS": "Mersin", "ALANYA": "Antalya", "TAVSANLI": "Kütahya", "GEDIZ": "Kütahya",
    "KOYCEGIZ": "Muğla", "FOCA": "İzmir", "PASAKOY": "Kırklareli", "IPSALA": "Edirne", "BEYTUSSEBAP": "Şırnak",
    "OSMANIY": "Osmaniye", "AKHISAR": "Manisa", "BABAESKI": "Kırklareli", "MARAS": "Kahramanmaraş",
    "KORKUTELI": "Antalya", "KARGI": "Çorum", "NIKSAR": "Tokat", "MENDERES": "İzmir", "HAVSA": "Edirne",
    "BORCKA": "Artvin", "ESKIPAZAR": "Karabük", "EZINE": "Çanakkale", "URGUP": "Nevşehir", "ELAZIG": "Elazığ",
    "BISMIL": "Diyarbakır", "DATCA": "Muğla", "SILIFKE": "Mersin", "KARABURUN": "İzmir",
}


def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = text.translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU"))
    return re.sub(r"[^A-Z0-9 ]+", " ", text).strip()


def load_whitelist(path):
    df = pd.read_excel(path, engine='openpyxl')
    return {normalize(u): i for u, i in zip(df["Unvan"], df["İl"])}


def find_city(row, whitelist):
    unvan = normalize(row.get("Ünvan", ""))
    adres = normalize(row.get("Adres", ""))

    # Check whitelist first
    for k, il in whitelist.items():
        if k in unvan:
            return il

    # Create normalized lookup dictionaries
    normalized_kisa_il = {normalize(k): v for k, v in KISA_IL.items()}
    normalized_ilce_il = {normalize(k): v for k, v in ILCE_IL_HARITA.items()}
    normalized_iller = {normalize(il): il for il in ILLER_LISTESI}
    
    # Combine all normalized keys for checking
    all_normalized_keys = set(normalized_kisa_il.keys()) | set(normalized_ilce_il.keys()) | set(normalized_iller.keys())

    words = adres.split()
    for word in reversed(words):
        if word in all_normalized_keys:
            # Check in order of priority: abbreviations, districts, then cities
            if word in normalized_kisa_il:
                return normalized_kisa_il[word]
            elif word in normalized_ilce_il:
                return normalized_ilce_il[word]
            elif word in normalized_iller:
                return normalized_iller[word]

    # Check if any district or abbreviation keys are contained in the address
    for original_key, il in ILCE_IL_HARITA.items():
        if normalize(original_key) in adres:
            return il
    
    for original_key, il in KISA_IL.items():
        if normalize(original_key) in adres:
            return il

    # Check if any city name is contained in the address
    for il in ILLER_LISTESI:
        if normalize(il) in adres:
            return il

    return pd.NA

def main():
    for path in [CSV_PATH, WHITELIST_PATH]:
        if not os.path.exists(path):
            print(f"Hata: {path} dosyası bulunamadı!", file=sys.stderr)
            sys.exit(1)

    df = pd.read_csv(CSV_PATH, encoding='windows-1254', sep=';')
    df["Çalışan Sayısı"] = pd.to_numeric(df["Çalışan Sayısı"], errors="coerce")
    df = df[df["Çalışan Sayısı"] >= 5].copy()

    whitelist = load_whitelist(WHITELIST_PATH)
    df["İl"] = df.apply(lambda row: find_city(row, whitelist), axis=1)

    df.drop(columns=["Faks"], inplace=True, errors="ignore")
    df.sort_values(by=["İl", "Çalışan Sayısı"], ascending=[True, False], inplace=True)
    df = df[["İl"] + [c for c in df.columns if c != "İl"]]

    df.to_excel(OUTPUT_PATH, index=False)
    df[df["İl"].isna()].to_excel(UNMATCHED_PATH, index=False)
    print(f"Çıktılar oluşturuldu:\n - {OUTPUT_PATH}\n - {UNMATCHED_PATH}")


if __name__ == "__main__":
    main()
