#!/usr/bin/env python3

import os
import sys
import re
import pandas as pd
import numpy as np

# --- Dosya ve çıktı yolları ---
CSV_PATH = 'files/4.csv'
WHITELIST_PATH = 'files/unvan_il_whitelist.xlsx'
OUTPUT_PATH = 'files/temizlenmis_veri.xlsx'
UNMATCHED_PATH = 'files/eslesemeyen_kayitlar.xlsx'

# --- Türkiye şehir listesi vb. ---
iller_listesi = [ "Adana", "Adıyaman", "Afyonkarahisar", "Ağrı", "Aksaray", "Amasya",
    "Ankara", "Antalya", "Artvin", "Aydın", "Balıkesir", "Bartın", "Batman", "Bayburt",
    "Bilecik", "Bingöl", "Bitlis", "Bolu", "Burdur", "Bursa", "Çanakkale", "Çankırı",
    "Çorum", "Denizli", "Diyarbakır", "Düzce", "Edirne", "Elazığ", "Erzincan", "Erzurum",
    "Eskişehir", "Gaziantep", "Giresun", "Gümüşhane", "Hakkari", "Hatay", "Iğdır",
    "Isparta", "İstanbul", "İzmir", "Kahramanmaraş", "Karabük", "Karaman", "Kars",
    "Kastamonu", "Kayseri", "Kilis", "Kırıkkale", "Kırklareli", "Kırşehir", "Kocaeli",
    "Konya", "Kütahya", "Malatya", "Manisa", "Mardin", "Mersin", "Muğla", "Muş",
    "Nevşehir", "Niğde", "Ordu", "Osmaniye", "Rize", "Sakarya", "Samsun", "Siirt",
    "Sinop", "Sivas", "Şanlıurfa", "Şırnak", "Tekirdağ", "Tokat", "Trabzon", "Tunceli",
    "Uşak", "Van", "Yalova", "Yozgat", "Zonguldak", "Ardahan"
]

kisa_il = {
    "C.KALE": "Çanakkale", "Ç.KALE": "Çanakkale", "ARD.": "Ardahan",
    "ANK.": "Ankara", "ANT.": "Antalya", "GAZ.": "Gaziantep",
    "IST.": "İstanbul", "IZM.": "İzmir", "T.DAĞ": "Tekirdağ", "Ş.URFA": "Şanlıurfa", "URFA": "Şanlıurfa"
}

ilce_il_harita = {
    "SANDIKLI": "Afyonkarahisar", "SERİK": "Antalya", "ESPİYE": "Giresun",
    "ÜMRANİYE": "İstanbul", "SARIYER": "İstanbul", "ÇANKAYA": "Ankara",
    "BODRUM": "Muğla", "FETHİYE": "Muğla", "GEBZE": "Kocaeli",
    "TORBALI": "İzmir", "CİZRE": "Şırnak", "SALİHLİ": "Manisa",
    "GÜLNAR": "Mersin", "YATAĞAN": "Muğla", "KIZILTEPE": "Mardin",
    "EYYUBİYE": "Şanlıurfa", "GAZİEMİR": "İzmir", "ADAPAZARI": "Kocaeli",
    "ULA": "İzmir", "ICEL": "Mersin", "MANAVGAT": "Antalya",
    "AFYON": "Afyonkarahisar",
    "SIVEREK": "Şanlıurfa",
    "SOMA": "Manisa",
    "AKSEHIR": "Konya",
    "TURGUTLU": "Manisa",
    "SURUC": "Şanlıurfa",
    "ORTACA": "Muğla",
    "KALKAN": "Antalya",
    "MILAS": "Muğla",
    "TARSUS": "Mersin",
    "ALANYA": "Antalya",
    "TAVSANLI": "Kütahya",
    "GEDIZ": "Kütahya",
    "KOYCEGIZ": "Muğla",
    "FOCA": "İzmir",
    "PASAKOY": "Kırklareli",       # Paşaköy ilçesi Babaeski'ye bağlı
    "IPSALA": "Edirne",
    "BEYTUSSEBAP": "Şırnak",
    "OSMANIY": "Osmaniye",         # Yazım hatalıydı, düzeltildi
    "AKHISAR": "Manisa",
    "BABAESKI": "Kırklareli",
    "MARAS": "Kahramanmaraş",       # Kısaltma gibi, Kahramanmaraş için geçerli
    "KORKUTELI": "Antalya",
    "KARGI": "Çorum",
    "NIKSAR": "Tokat",
    "MENDERES": "İzmir",
    "HAVSA": "Edirne",
    "BORCKA": "Artvin",
    "ESKIPAZAR": "Karabük",
    "EZINE": "Çanakkale",
    "URGUP": "Nevşehir",
    "ELAZIG": "Elazığ",       # Yazım düzeltilmiş hali
    "BISMIL": "Diyarbakır",
}

def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = text.translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU"))
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def find_city_row(row, unvan_whitelist):
    unvan_norm = normalize(row.get("Ünvan", ""))
    adres_norm = normalize(row.get("Adres", ""))

    # 1. Ünvan whitelist eşleşmesi
    for key, il in unvan_whitelist.items():
        if key in unvan_norm:
            return il

    # 2. Adresin son kelimesinden eşleşme
    words = adres_norm.split()
    for word in reversed(words):
        for il in iller_listesi:
            if normalize(il) == word:
                return il
        for kisa, il in kisa_il.items():
            if normalize(kisa) == word:
                return il
        for ilce, il in ilce_il_harita.items():
            if normalize(ilce) == word:
                return il

    # 3. İlçe adının adres içinde geçmesi
    for ilce, il in ilce_il_harita.items():
        if normalize(ilce) in adres_norm:
            return il

    # 4. Kısaltma adres içinde
    for kisa, il in kisa_il.items():
        if normalize(kisa) in adres_norm:
            return il

    # 5. Şehir adının adres içinde yer alması
    for il in iller_listesi:
        if normalize(il) in adres_norm:
            return il

    return np.nan

def main():
    # Dosya kontrolü
    for path in [CSV_PATH, WHITELIST_PATH]:
        if not os.path.exists(path):
            print(f"Hata: '{path}' dosyası bulunamadı!", file=sys.stderr)
            sys.exit(1)

    # CSV oku
    df = pd.read_csv(CSV_PATH, encoding='windows-1254', sep=';')
    df["Çalışan Sayısı"] = pd.to_numeric(df["Çalışan Sayısı"], errors="coerce")
    df_filtered = df[df["Çalışan Sayısı"] >= 5].copy()

    # Whitelist oku
    whitelist_df = pd.read_excel(WHITELIST_PATH, engine='openpyxl')
    unvan_il_whitelist = {
        normalize(unvan): il for unvan, il in zip(whitelist_df["Unvan"], whitelist_df["İl"])
    }

    # İl eşleştir
    df_filtered["İl"] = df_filtered.apply(
        lambda row: find_city_row(row, unvan_il_whitelist),
        axis=1
    )

    # Faks sütununu kaldır
    if "Faks" in df_filtered.columns:
        df_filtered = df_filtered.drop(columns=["Faks"])
    
    # --- Sıralama ve görsel düzenleme ---
    df_filtered = df_filtered.sort_values(by=["İl", "Çalışan Sayısı"], ascending=[True, False])
    columns_ordered = ["İl"] + [col for col in df_filtered.columns if col != "İl"]
    df_filtered = df_filtered[columns_ordered]
    
    # Çıktılar
    df_filtered.to_excel(OUTPUT_PATH, index=False)
    df_filtered[df_filtered["İl"].isna()].to_excel(UNMATCHED_PATH, index=False)
    print(f"✔️ '{OUTPUT_PATH}' ve eşleşmeyen kayıtlar '{UNMATCHED_PATH}' olarak kaydedildi.")

if __name__ == "__main__":
    main()
