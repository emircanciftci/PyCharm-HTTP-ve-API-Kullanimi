import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
from selenium import webdriver


options = Options()
options.add_experimental_option("detach", True)

service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)

driver.get("https://mikrokargo.loomis.com.tr/login")
driver.maximize_window()
wait.until(EC.url_contains("/login"))

def read_credentials(filepath):
    creds = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if "=" in line:
                key, value1 = line.split("=", 1)
                creds[key.strip()] = value1.strip()
    return creds
credentials = read_credentials("credentials.txt")
username = credentials.get("username")
password = credentials.get("password")

username_input = wait.until(EC.element_to_be_clickable((By.ID, "mat-input-0")))
username_input.send_keys(username)

password_input = wait.until(EC.element_to_be_clickable((By.ID, "mat-input-1")))
password_input.send_keys(password)
password_input.send_keys(Keys.ENTER)

WebDriverWait(driver, 30).until(EC.url_contains("/mikrokargo/dashboard"))
driver.get("https://mikrokargo.loomis.com.tr/mikrokargo/gonderi/listesi")
WebDriverWait(driver, 30).until(EC.url_contains("/mikrokargo/gonderi/listesi"))

search_input = driver.find_element(By.CLASS_NAME, "dx-texteditor-input")
sayi = 1
previous_value = ""

while True:
    try:
        search_input = driver.find_element(By.CLASS_NAME, "dx-texteditor-input")
        value = search_input.get_attribute("value")
        if len(value) == 15 and previous_value != value:
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.dx-checkbox-icon"))).click()

            original_window_count = len(driver.window_handles)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()=' Hepsijet Barkod ']]"))).click()

            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > original_window_count)

            while len(driver.window_handles) > original_window_count:
                pyautogui.press('enter')
                time.sleep(0.2)

            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.dx-checkbox-icon"))).click()
            search_input.clear()
            search_input.click()
            print(str(sayi) + "- Barkod Yazdırıldı: " + value)
            sayi = sayi+1
            previous_value = value

    except Exception as e:
        print(f"Hata: {e}")
        time.sleep(0.2)
