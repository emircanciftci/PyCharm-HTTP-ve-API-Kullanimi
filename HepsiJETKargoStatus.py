import pandas as pd
import requests
import os
import glob

# 1️⃣ EXE'nin bulunduğu klasör
folder = os.getcwd()

# 2️⃣ Klasördeki Excel dosyalarını al, geçici dosyaları atla
excel_files = [f for f in glob.glob(os.path.join(folder, "*.xls*")) if not os.path.basename(f).startswith("~$")]

if not excel_files:
    print("❌ Klasörde hiç Excel dosyası bulunamadı!")
    input("Çıkmak için Enter'a basın...")
    exit()

# 3️⃣ En yeni Excel dosyasını seç
excel_path = max(excel_files, key=os.path.getctime)
print(f"✅ Excel bulundu: {excel_path}")

# 4️⃣ Excel'i oku
try:
    df = pd.read_excel(excel_path, dtype=str)
except Exception as e:
    print(f"❌ Excel okunamadı: {e}")
    exit()

# 5️⃣ "Kargo Takip Kodu" sütununu kontrol et
if "Kargo Takip Kodu" not in df.columns:
    print("❌ Excel'de 'Kargo Takip Kodu' sütunu bulunamadı!")
    input("Çıkmak için Enter'a basın...")
    exit()

# 6️⃣ Her takip kodunu sorgula
for index, row in df.iterrows():
    code = row["Kargo Takip Kodu"]
    url = f"https://website-backend.hepsijet.com/server/tms/webDeliveryTracking?customerDeliveryNo={code}&captchaKey=&captchaValue="

    try:
        response = requests.get(url, timeout=10)
        data = response.json()
        # Her işlemde durum kontrolü
        if data.get("status") != "OK":
            print(f"{code} → Hata: Status OK değil")
            continue

        shipment = data.get("data", [{}])[0]  # Eğer liste boşsa boş dict
        current_location = shipment.get("currentLocation", "Bilgi yok")
        last_transaction = shipment.get("lastTransaction", "Bilgi yok")

        print(f"{code} → currentLocation: {current_location}, lastTransaction: {last_transaction}")

    except Exception as e:
        print(f"{code} → Hata: {e}")
