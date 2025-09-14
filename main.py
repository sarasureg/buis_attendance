from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
import os
import numpy as np
import pandas as pd
from datetime import datetime

# ----------------------------
# CONFIG
# ----------------------------
now = datetime.now()
target_year = [2021, 2022, 2023, 2024, 2025]
today = datetime.today().strftime('%Y-%m-%d')

# Folders
base_dir = os.getcwd()
temp_folder = os.path.join(base_dir, "temp")
final_folder = os.path.join(base_dir, "final")
os.makedirs(temp_folder, exist_ok=True)
os.makedirs(final_folder, exist_ok=True)

# Credentials from GitHub Secrets (or local env vars)
USERNAME = os.environ.get("USERNAME", "soumen_saha")
PASSWORD = os.environ.get("PASSWORD", "newpass")

# ----------------------------
# Chrome Setup (Headless for GitHub Actions)
# ----------------------------
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--enable-popup-blocking")
chrome_options.add_argument("--enable-notifications")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": temp_folder,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})

driver = webdriver.Chrome(options=chrome_options)

# ----------------------------
# Utility Functions
# ----------------------------
def clean_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

def forcefully_enter_val(xpath, value):
    element = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    element.click()
    actions = ActionChains(driver)
    actions.move_to_element(element).click().send_keys(value).send_keys(Keys.ENTER).perform()

def forcefully_click_val(xpath):
    element = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()

# ----------------------------
# Script Starts
# ----------------------------
clean_folder(temp_folder)

try:
    driver.get("https://buis.brainwareuniversity.org.in/")
    time.sleep(2)

    # Close popup
    driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/button").click()
    time.sleep(1)

    # Select role
    forcefully_enter_val("//*[@id='root']/div/div/div[2]/div/div[3]/form/div/div[1]/div/div/div/div[1]/div[2]", "staff")

    # Enter username
    forcefully_enter_val('//*[@id="root"]/div/div/div[2]/div/div[3]/form/div/div[2]/div/input', USERNAME)

    # Enter password
    forcefully_enter_val('//*[@id="root"]/div/div/div[2]/div/div[3]/form/div/div[3]/div/span/input', PASSWORD)

    # Captcha (your script extracts text, assuming it’s not an image)
    captcha_text = driver.find_element(By.XPATH, "//*[@id='capcode']/div").text
    forcefully_enter_val('//*[@id="capture_code"]', captcha_text)

    # Login
    login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    login_button.click()
    time.sleep(3)

except Exception as e:
    print(f"Login failed: {e}")

# ----------------------------
# Navigate to Attendance
# ----------------------------
try:
    forcefully_click_val('//*[@id="root"]/div/div[1]/main/div/div[2]/div/div[1]/div/div/div/h3')
    time.sleep(1)
    forcefully_click_val('//*[@id="details.report"]')
    time.sleep(1)
    forcefully_click_val('//*[@id="student.report.semester-attendance"]')
    time.sleep(1)
except Exception as e:
    print(f"Navigation failed: {e}")

# ----------------------------
# Download Attendance Files
# ----------------------------
for item in target_year:
    try:
        # Select year
        forcefully_enter_val("//*[@id='root']/div/div/div[2]/div/main/div/div/div[2]/div[1]/div/div/div[1]/div/div/div[1]/div[2]", f"{item}")
        time.sleep(1)

        # Program
        forcefully_enter_val("//*[@id='root']/div/div/div[2]/div/main/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]", "ALL")
        time.sleep(1)

        # Running batch
        forcefully_enter_val("//*[@id='root']/div/div/div[2]/div/main/div/div/div[2]/div[1]/div/div/div[3]/div/div/div[1]/div[2]", "Running")
        time.sleep(1)

        # Submit
        driver.find_element(By.XPATH, "//*[@id='root']/div/div/div[2]/div/main/div/div/div[2]/div[1]/div/div/div[5]/button").click()
        time.sleep(1)

        # Download Excel
        forcefully_click_val("//*[@id='root']/div/div/div[2]/div/main/div/div/div[2]/div[2]/div/div[1]/div[1]/span/button[2]/span")
        time.sleep(5)

        # Wait for file
        timeout = 120
        start_time = time.time()
        downloaded_file = None
        while True:
            files = os.listdir(temp_folder)
            downloaded_file = [f for f in files if not f.endswith(".crdownload")]
            if downloaded_file:
                downloaded_file = downloaded_file[0]
                break
            if time.time() - start_time > timeout:
                print("Download timed out.")
                break
            time.sleep(1)

        if downloaded_file:
            current_download_path = os.path.join(temp_folder, downloaded_file)
            desired_name = f"{item}_{today}.xlsx"
            renamed_path = os.path.join(temp_folder, desired_name)
            os.rename(current_download_path, renamed_path)
            shutil.move(renamed_path, os.path.join(final_folder, desired_name))
            print(f"Downloaded & moved: {desired_name}")

    except Exception as e:
        print(f"Error processing year {item}: {e}")

# ----------------------------
# Merge into Single Excel
# ----------------------------
df_final = pd.DataFrame()
for item in target_year:
    file_path = os.path.join(final_folder, f"{item}_{today}.xlsx")
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, "Sheet1")
        df_final = pd.concat([df_final, df])

attendance_columns = [
    'Sem1 Attendance %','Sem2 Attendance %','Sem3 Attendance %','Sem4 Attendance %',
    'Sem5 Attendance %','Sem6 Attendance %','Sem7 Attendance %','Sem8 Attendance %',
    'Sem9 Attendance %','Sem10 Attendance %','Year1 Attendance','Year2 Attendance',
    'Year3 Attendance','Year4 Attendance'
]

df_final['Current Sem'] = df_final[attendance_columns].apply(
    lambda row: row.dropna().iloc[-1] if not row.dropna().empty else "-", axis=1
)

output_file = now.strftime("%d-%m-%Y") + ".xlsx"
df_final.to_excel(output_file, sheet_name="Sheet1", index=False)
print(f"Final merged file saved as {output_file}")

# Clean up
clean_folder(final_folder)
driver.quit()

