from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests
import os, glob, shutil
import time

folder = os.getcwd()
excel_files = glob.glob(os.path.join(folder, "*.xls*"))
if not excel_files:
    print("Klasörde hiç Excel dosyası bulunamadı!")
    input("Çıkmak için Enter'a basın...")
    exit()
excel_path = max(excel_files, key=os.path.getctime)
backup_path = os.path.splitext(excel_path)[0] + "_sonuclar.xlsx"
shutil.copy(excel_path, backup_path)
df = pd.read_excel(backup_path)
if "Kargo Takip Kodu" not in df.columns:
    print("Excel'de 'Kargo Takip Kodu' sütunu bulunamadı!")
    input("Çıkmak için Enter'a basın...")
    exit()




def login(filepath):
    deger = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for val in f:
            val = val.strip()
            if "=" in val:
                key, value = val.split("=", 1)
                deger[key.strip()] = value.strip()
    return deger
giris = login("credentials.txt")
username = giris.get("username")
password = giris.get("password")

log_url = "https://jetpartner-backend.hepsijet.com/api/auth/login"
log_payload = {"email": username, "password": password}
log_headers = {"Content-Type": "application/json"}
loginn = requests.post(log_url, json=log_payload, headers=log_headers)

if loginn.status_code != 200:
    print("Login başarısız:", loginn.text)
    exit()
loginn_data = loginn.json()
token = loginn_data.get("token")
if not token:
    print("Token Alınmadı!")
    exit()
print("Token alındı!")
time.sleep(2)

# statuses = []
# for index,row in df.iterrows():
#     code = row["Kargo Takip Kodu"]
#     cargoname = row["Kargo Şirketi"]
#     namess = row["Alıcı"]
#     if cargoname == "HepsiJet":
#         url = f"https://jetpartner-backend.hepsijet.com/api/reports/delivery-search?barcode={code}"
#         headers = {
#             "Authorization": f"Bearer {token}",
#             "Accept": "application/json",
#             "Content-Type": "application/json"
#         }
#         try:
#             data = requests.get(url,headers=headers, timeout=10).json()
#             rows = data.get("data", {}).get("rows", [])  # rows listesini al
#             status = rows[-1].get("translation_tr", "Bulunamadı") if rows else "Bulunamadı"
#         except Exception as e:
#             status = f"Hata: {e}"
#         print(f"{code} - {namess} → {status}")
#         statuses.append(status)
#
#     elif cargoname == "YurtiçiKargo" or cargoname == "Geliver":
#         url = f"https://www.yurticikargo.com/service/shipmentstracking?id={code}&language=tr"
#         try:
#             data = requests.get(url, timeout=10).json()
#             status = data.get("ShipmentStatus", "Bulunamadı")
#         except Exception as e:
#             status = f"Hata: {e}"
#         print(f"{code} - {namess} → {status}")
#         statuses.append(status)
#     else:
#         status = "Bilinmeyen Kargo"
#     statuses.append(status)



statuses = []
def check_status(row):
    code = row["Kargo Takip Kodu"]
    cargoname = str(row["Kargo Şirketi"]).strip().lower()

    if cargoname == "hepsijet":
        url = f"https://jetpartner-backend.hepsijet.com/api/reports/delivery-search?barcode={code}"
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        data = requests.get(url, headers=headers, timeout=10).json()
        rows = data.get("data", {}).get("rows", [])
        if rows:
            translation = rows[-1].get("translation_tr") or ""
            location = rows[-1].get("location_tr") or ""
        else:
            translation = "Bulunamadı"
            location = "Bulunamadı"
        return " / ".join(filter(None, [translation, location]))


    elif "yurtiçi" in cargoname or "geliver" in cargoname:
        url = f"https://www.yurticikargo.com/service/shipmentstracking?id={code}&language=tr"
        payload = {f"id={code}&language=tr"}
        data = requests.get(url, timeout=10).json()
        return data.get("ShipmentStatus", "Bulunamadı")

    else:
        return "Bilinmeyen Kargo"


with ThreadPoolExecutor(max_workers=200) as executor:
    statuses = list(executor.map(check_status, df.to_dict('records')))

df["KargoDurumu"] = statuses
df.to_excel(backup_path, index=False)
print(f"Sonuçlar {os.path.basename(backup_path)} dosyasına yazıldı.")
