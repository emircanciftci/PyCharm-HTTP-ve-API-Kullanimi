import requests
import pandas

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

url = "https://mikroapi.loomis.com.tr/api/mikrokargo/alici-listesi"
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}
response = requests.post(url, headers=headers, json={}, timeout=120)

if response.status_code != 200:
    print("Alıcı listesi alınamadı:", response.status_code, response.text)
    exit()

data = response.json().get("Data", [])

rows = []
for alici in data:
    row = {
    "ID": alici.get("Id"),
    "Gönderici Firma": (alici.get("GondericiFirma") or {}).get("Unvani", ""),
    "Alıcı Adı": alici.get("AliciAdi"),
    "Alıcı Tel": alici.get("AliciTelefon"),
    "Alıcı Firma": (alici.get("AliciFirma") or {}).get("Unvani", ""),
    "Durumu": "AKTİF" if alici.get("Aktif") else "PASİF",
    "Adresleri": alici.get("Adresler"),
    }
    rows.append(row)

df = pandas.DataFrame(rows)
df.to_excel("Alıcı Listesi.xlsx", index=False)
print("Excel dosyası oluşturuldu: Alıcı Listesi.xlsx")