from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# install require package python (optional, jika blm punya)
# pip3 install selenium pandas openpyxl

try:
    # Setup
    service = Service("/Users/erikcahya/Documents/Project/chromedriver/chromedriver") # chromedriver location path
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()

    # Login
    driver.get("http://localhost:8000/login") # url website
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "email"))).send_keys("admin@gmail.com")
    driver.find_element(By.NAME, "password").send_keys("admin123lsp")
    
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-primary")))
    login_button.click()

    # Isi Form
    df = pd.read_excel("data.xlsx") # excel file location path
    for index, row in df.iterrows():
        driver.get("http://localhost:8000/manajemen/create") # url website
        
        # Tunggu form ready
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "nama_manajemen")))
        
        # Isi form
        fields = {
            "nama_manajemen": row["Nama"],
            "no_telp": row["No Telp"],
            "jabatan": row["Jabatan"],
            "alamat": row["Alamat"]
        }
        
        for name, value in fields.items():
            element = driver.find_element(By.NAME, name)
            element.clear()
            element.send_keys(value)
            time.sleep(0.5)  # Jeda antar field

        # Submit dengan cara paling robust
        submit_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.btn-info")))
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", submit_button)
        
        time.sleep(3)  # Tunggu proses submit

    # REDIRECT KE HALAMAN INDEX SETELAH SELESAI
    print("Proses selesai! Mengarahkan ke halaman index...")
    driver.get("http://localhost:8000/manajemen")  # Ganti dengan URL index
    time.sleep(3)

    # print("Proses berhasil!")

except Exception as e:
    print("Error:", e)
    driver.save_screenshot("error.png")  # Debug screenshot
finally:
    input("Tekan enter untuk exit")
