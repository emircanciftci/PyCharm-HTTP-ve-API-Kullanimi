import easyocr
import pandas as pd
import os
from PIL import Image
import re
import time

print("Merhaba")
time.sleep(2)
print("İşlem Bir Kaç Dakika Sürebilir,Lütfen Bekleyiniz...")

reader = easyocr.Reader(['tr', 'en'])
folder_path = r"resimler"
all_data = []

for filename in os.listdir(folder_path):
    if filename.lower().endswith((".png", ".jpg", ".jpeg")):
        image_path = os.path.join(folder_path, filename)

        try:
            im = Image.open(image_path)
            max_size = 3000
            if max(im.size) > max_size:
                im.thumbnail((max_size, max_size))
        except Exception as e:
            print(f"{filename} açılamadı: {e}")
            continue

        try:
            results = reader.readtext(im, detail=0)
        except Exception as e:
            print(f"{filename} okunamadı: {e}")
            continue

        iller = (
            "Adana|Adıyaman|Afyonkarahisar|Ağrı|Amasya|Ankara|Antalya|Artvin|Aydın|Balıkesir|"
            "Bilecik|Bingöl|Bitlis|Bolu|Burdur|Bursa|Çanakkale|Çankırı|Çorum|Denizli|Diyarbakır|"
            "Edirne|Elazığ|Erzincan|Erzurum|Eskişehir|Gaziantep|Giresun|Gümüşhane|Hakkari|Hatay|"
            "Isparta|Mersin|İstanbul|İzmir|Kars|Kastamonu|Kayseri|Kırklareli|Kırşehir|Kocaeli|Konya|"
            "Kütahya|Malatya|Manisa|Kahramanmaraş|Mardin|Muğla|Muş|Nevşehir|Niğde|Ordu|Rize|Sakarya|"
            "Samsun|Siirt|Sinop|Sivas|Tekirdağ|Tokat|Trabzon|Tunceli|Şanlıurfa|Uşak|Van|Yozgat|"
            "Zonguldak|Aksaray|Bayburt|Karaman|Kırıkkale|Batman|Şırnak|Bartın|Ardahan|Iğdır|Yalova|"
            "Karabük|Kilis|Osmaniye|Düzce"
        )

        full_text = " ".join(results)
        full_yazi = full_text

        ad_soyad = ""
        adres = ""
        barkod_no = ""
        poset_no = ""
        siparis_no = ""
        il = ""
        ilce = ""

        if re.search(r"n11|nl1", full_text, re.IGNORECASE):

            barkod_match = re.search(r"\b\d{15}\b", full_text)
            barkod_no = barkod_match.group(0) if barkod_match else ""

            poset_match = re.search(r"\b\d{8}\b", full_text)
            if poset_match:
                poset_no = poset_match.group(0)

            siparis_no_match = re.search(r"Sipariş\s+Numarası:\s+(\d+)", full_text, re.IGNORECASE)
            siparis_no = siparis_no_match.group(1) if siparis_no_match else ""

            ad_soyad_match = re.search(r"AdISoyad:\s+(.*?)\s+Adres", full_text, re.IGNORECASE)
            if ad_soyad_match:
                ad_soyad = ad_soyad_match.group(1).strip()

            adres_match = re.search(rf"Adres:\s+(.*?)(?:\s+(\S+))\s+({iller})", full_text, re.IGNORECASE)
            if adres_match:
                adres = adres_match.group(1).strip()
                ilce = adres_match.group(2).strip().capitalize()
                il = adres_match.group(3).strip().capitalize()
        else:
            barkod_match = re.search(r"\b\d{16}\b", full_text)
            barkod_no = barkod_match.group(0) if barkod_match else ""

            poset_match = re.search(r"\b\d{8}\b", full_text)
            if poset_match:
                poset_no = poset_match.group(0)

            siparis_no_match = re.search(r"Sipariş\s+No\s+(\d+)", full_text, re.IGNORECASE)
            siparis_no = siparis_no_match.group(1) if siparis_no_match else ""


            ad_soyad_match = re.search(r"Ad-?Soyad\s+(.*?)\s+Adres", full_text, re.IGNORECASE)
            if ad_soyad_match:
                ad_soyad = ad_soyad_match.group(1).strip()


            adres_match = re.search(rf"Adres\s+(.*?)(?:\s+(\S+))\s+({iller})", full_text, re.IGNORECASE)
            if adres_match:
                adres = adres_match.group(1).strip()
                ilce = adres_match.group(2).strip().capitalize()
                il = adres_match.group(3).strip().capitalize()

        all_data.append({
            "Poşet No": poset_no,
            "Barkod No": barkod_no,
            "Sipariş No": siparis_no,
            "Ad-Soyad": ad_soyad,
            "Adres": adres,
            "İlçe": ilce,
            "İl": il,
        })

output_excel = r"Kargo_Detay_Listesi.xlsx"
df = pd.DataFrame(all_data)
df.to_excel(output_excel, index=False)
print(f"Excel oluşturuldu: {output_excel}")
