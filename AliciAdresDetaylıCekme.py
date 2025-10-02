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
sayi = 0
for alici in data:
    sayi += 1
    print(f"\rVerisi Alınan Alıcı Sayısı: {sayi}", end="", flush=True)

    url = "https://mikroapi.loomis.com.tr/api/mikrokargo/alici-adresleri"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = {"Id": alici.get("Id")}

    responses = requests.post(url, headers=headers, json=payload, timeout=120)
    if responses.status_code != 200:
        print("Alıcı listesi alınamadı:", responses.status_code, responses.text)
        exit()

    adres_listesi = responses.json().get("Data", [])
    adresler = adres_listesi.get("Adresler", [])
    reis = []
    for adres in adresler:
        if adres.get("MikroKargoAliciId") not in reis:
            row = {
                "ID": adres_listesi.get("Id"),
                "Alıcı Adı": adres_listesi.get("AliciAdi"),
                "Alıcı Telefon": adres_listesi.get("AliciTelefon"),
                "Alıcı Firma": (adres_listesi.get("AliciFirma") or {}).get("Unvani", ""),
                "Durumu": "AKTİF" if (adres_listesi.get("Aktif")) else "",
                "Adres ID": adres.get("Id", ""),
                "Adres Aktif": adres.get("IsEnable", True),
                "Adres Başlık": adres.get("Baslik", ""),
                "Adres Ad": adres.get("Ad", ""),
                "Adres Soyad": adres.get("Soyad", ""),
                "Adres Telefon": adres.get("Telefon", ""),
                "İl": (adres.get("Il") or {}).get("Adi", ""),
                "İlçe": (adres.get("Ilce") or {}).get("Adi", ""),
                "Mahalle": (adres.get("Mahalle") or {}).get("Adi", ""),
                "Adres": adres.get("Adres", ""),
            }
            reis.append(adres.get("MikroKargoAliciId"))
            rows.append(row)
        else:
            row = {
                "ID": adres_listesi.get("Id"),
                "Alıcı Adı": "",
                "Alıcı Telefon": "",
                "Alıcı Firma": "",
                "Durumu": "",
                "Adres ID": (adres.get("Adresler") or {}).get("Id", ""),
                "Adres Aktif": "AKTİF" if (adres.get("IsEnable")) else "",
                "Adres Başlık": adres.get("Baslik"),
                "Adres Ad": adres.get("Ad"),
                "Adres Soyad": adres.get("Soyad"),
                "Adres Telefon": adres.get("Telefon"),
                "İl": (adres.get("Il") or {}).get("Adi", ""),
                "İlçe": (adres.get("Ilce") or {}).get("Adi", ""),
                "Mahalle": (adres.get("Mahalle") or {}).get("Adi", ""),
                "Adres": adres.get("Adres"),

            }
            rows.append(row)

df = pandas.DataFrame(rows)
df.to_excel("Alıcı Listessi.xlsx", index=False)
print("Excel dosyası oluşturuldu: Alıcı Listessi.xlsx")

