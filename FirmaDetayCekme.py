import requests
import pandas as pd

def read_credentials(filepath):
    creds = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for val in f:
            val = val.strip()
            if "=" in val:
                key, value = val.split("=", 1)
                creds[key.strip()] = value.strip()
    return creds
credentials = read_credentials("credentials.txt")
username = credentials.get("username")
password = credentials.get("password")

login_url = "https://mikroapi.loomis.com.tr/api/account/login"
login_payload = {"Email": username, "Password": password}
login_headers = {"Content-Type": "application/json"}

login_response = requests.post(login_url, json=login_payload, headers=login_headers)
if login_response.status_code != 200:
    print("Login başarısız:", login_response.text)
    exit()

login_data = login_response.json()
token = login_data.get("Data", {}).get("Token")
if not token:
    print("Token alınamadı!")
    exit()
print("Token alındı!")

url = "https://mikroapi.loomis.com.tr/api/firma/tum-firma-listesi"
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}

response = requests.post(url, headers=headers, json={})
if response.status_code != 200:
    print("Firma listesi alınamadı:", response.status_code, response.text)
    exit()

data = response.json().get("Data", [])

rows = []
for firma in data:
    row = {
    "ID": firma.get("Id"),
    "ŞubeliFirma": "EVET" if firma.get("SubeliFirma") else "HAYIR",
    "AnaFirma": firma.get("AnaFirma", {}).get("KisaAdi") if firma.get("AnaFirma") else None,
    "Aktif": "AKTİF" if firma.get("Aktif") else "PASİF",
    "Firma Adı": firma.get("KisaAdi"),
    "Ünvanı": firma.get("Unvani"),
    "Muhasebe Kodu": firma.get("MuhasebeKodu"),
    "Mail": firma.get("MuhasebeMail"),
    "Vergi Dairesi": firma.get("VergiDairesi"),
    "Vergi No": firma.get("VergiNo"),
    "Yetkili Adı": firma.get("YetkiliAdi"),
    "Telefon": firma.get("Telefon"),
    "İl": firma.get("Il", {}).get("Adi"),
    "İlçe": firma.get("Ilce", {}).get("Adi"),
    "Adres": firma.get("Adres"),
    "Sigorta Oranı": firma.get("SigortaOrani"),
    "YK": firma.get("YK"),
    "UPS": firma.get("UPS"),
    "MNG": firma.get("MNG"),
    "HJ": firma.get("HJ"),
    "KG": firma.get("KG"),
    "GV": firma.get("GV"),
    "Sözleşmeyi Yapan": firma.get("SozlesmeyiYapan"),
    "Müşteri Türü": firma.get("MusteriTuru"),
    "Sözleşme Başlangıcı": firma.get("SozlesmeBaslangici"),
    }
    rows.append(row)

df = pd.DataFrame(rows)
df.to_excel("Firma Listesi.xlsx", index=False)
print("Excel dosyası oluşturuldu: Firma Listesi.xlsx")
