import time
import requests
import pandas
from sympy import false
from concurrent.futures import ThreadPoolExecutor


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

log_url = "https://mikroapi.loomis.com.tr/api/account/login"
log_payload = {"Email": username, "Password": password}
log_headers = {"Content-Type": "application/json"}
loginn = requests.post(log_url, json=log_payload, headers=log_headers)

if loginn.status_code != 200:
    print("Login başarısız:", loginn.text)
    exit()

loginn_data = loginn.json()
token = loginn_data.get("Data", {}).get("Token")
if not token:
    print("Token Alınmadı!")
    exit()
print("Token alındı!")
time.sleep(2)

while True:
    startdate = input("Başlangıç Tarihi Giriniz Format(YYYY-MM-DD): ")
    enddate = input("Bitiş Tarihi Giriniz Format(YYYY-MM-DD): ")
    karar = input(f"Doğru mu? {startdate} - {enddate} (D/Y): ")

    if karar == "d" or karar == "D":
        break
    else:
        print("Tekrar tarih seçiniz.\n")

print("Lütfen Bekleyiniz...")

url = "https://mikroapi.loomis.com.tr/api/mikrokargo/fatura-liste"
payload = {"FromDate":startdate+"T21:00:00.000Z","ToDate":enddate+"T21:00:00.000Z"}
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}
response = requests.post(url, headers=headers, json=payload, timeout=120)

if response.status_code != 200:
    print("Liste alınamadı:", response.status_code, response.text)
    exit()

data = response.json().get("Data", [])

rows = []
for alici in data:
    if alici.get("MuhasebeyeGonderildi") == false:
        row = {
        "FaturaFirma": alici.get("FaturaFirma"),
        "FaturaFirmaMuhasebeKodu": alici.get("FaturaFirmaMuhasebeKodu"),
        "FaturaSeriNo": alici.get("FaturaSeriNo"),
        "FaturaTarihi": alici.get("FaturaTarihi"),
        "ToplamTutar": alici.get("ToplamTutar"),
        "GonderiAdet": alici.get("GonderiAdet"),
        "BaslangicTarihi": alici.get("BaslangicTarihi"),
        "BitisTarihi": alici.get("BitisTarihi"),
        "Muhasebeye Gönderildi Mi":"EVET" if (alici.get("MuhasebeyeGonderildi")) else "HAYIR",
        "MuhasebeGonderimTarihi": alici.get("MuhasebeGonderimTarihi"),
        }
        print(row.get("FaturaFirma", '')[:30].ljust(30, '-'),row.get("ToplamTutar"))
        rows.append(row)
    else:
        row = {}

df = pandas.DataFrame(rows)
df.to_excel("Fatura Listesi.xlsx", index=False)
print("Excel Dosyası Oluşturuldu: Fatura Listesi.xlsx")

print("Faturası Gönderilecek Firma Sayısı = ", len(rows))
while True:
    karar = input(f"Doğru mu? {startdate} - {enddate} (D/Y): ")

    if karar == "d" or karar == "D":
        break
    else:
        exit()
print("İşleme Devam Ediliyor, Lütfen Bekleyiniz...")

def gonder(row):
    url = "https://mikroapi.loomis.com.tr/api/mikrokargo/logo-fatura-muhasebelestir"
    payload = [row.get("FaturaSeriNo")]
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=payload, timeout=200)
    print(row.get("FaturaFirma", '')[:30].ljust(30, '-'), (response.json() or {}).get("ExceptionString"))
with ThreadPoolExecutor(max_workers=200) as executor:
    executor.map(gonder, rows)

print("İşlem Tamamlandı.")