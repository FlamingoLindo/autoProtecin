import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import random
import time
from dotenv import load_dotenv
import os
from pynput.keyboard import Key, Controller

load_dotenv()

# Load the Excel file
df = pd.read_excel('./protecin.xlsx')
print(df)

# Path to your ChromeDriver
driver_path = './chromedriver.exe'
s = Service(driver_path)
driver = webdriver.Chrome(service=s)  

# Open the web page
driver.get(os.getenv('PROTECIN_URL'))

# Initialize WebDriverWait
wait = WebDriverWait(driver, 5)

# Wait for email input to be clickable
email_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="email"]')))
email_input.send_keys(os.getenv('PROTECIN_LOGIN'))

# Wait for password input to be clickable
password_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]')))
password_input.send_keys(os.getenv('PROTECIN_PASSWORD'))

# Click on login button
login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/main/section/form/button')))
login_btn.click()

# Click on the extinguisher page
ext_page = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/nav/a[5]')))
ext_page.click()

print("🤖SCRPIT STARTING🤖")
for index, row in df.iterrows():
    # Click on the extinguisher register page
    ext_register = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/section/a')))
    ext_register.click()

    # Show what index its currently inputing (Just for debuging)
    print(index)
    
    # Selects the company in the drop-down
    company_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[1]/div[1]/div/div/div[1]/div[2]')))
    company_dropdown.click()
    pyautogui.write("saint")
    company_input = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), 'SAINT GOBAIN ABRASIVOS CUMBICA')]")))
    company_input.click()

    # Inputs control number
    control_num_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#controlNumber")))
    control_num_input.send_keys(row['numero'])

    # Inputs seal number
    seal_num_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#seal")))
    seal_num_input.send_keys(row['numeroLacre'])

    # Inputs the manufacturer
    # Doesnt need it 

    # Inputs the serial number
    serial_num_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#serialNumber")))
    serial_num_input.send_keys(row['numeroFabricacao'])

    # Inputs the agent drop-down
    agent_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[1]/div[6]/div/div/div[1]')))
    agent_dropdown.click()
    pyautogui.write({row['agente']}) 
    agent_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{row['agente']}')]")))
    driver.execute_script("arguments[0].click();",agent_option)

    # Inputs the extinguisher capacity drop-down
    ext_cap_dropdown = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[1]/div[7]/div/div/div[1]')))
    ext_cap_dropdown.click()
    pyautogui.write({row['capacidade']})
    ext_cap_option = wait.until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(text(), '{row['capacidade']}')]")))
    driver.execute_script("arguments[0].click();", ext_cap_option)

    # Inputs the extinguisher load drop-down
    ext_load_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[1]/div[8]/div/div/div[1]')))
    ext_load_dropdown.click()
    pyautogui.write({row['cap']})
    ext_load_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{row['cap']}')]")))
    driver.execute_script("arguments[0].click();", ext_load_option)

    # Inputs last TH date
    last_TH_date_str = row['validadeTeste'].strftime('%d/%m/%Y')
    last_TH_date_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dateOfLastTH")))
    last_TH_date_input.send_keys(last_TH_date_str)

    # Inputs last recharge date
    last_reload_date_str = row['validadeCarga'].strftime('%d/%m/%Y')
    last_reload_date_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#lastRechargeDate")))
    last_reload_date_input.send_keys(last_reload_date_str)

    # Assuming row is a pandas Series and wait is an instance of WebDriverWait
    data_semestral = row['dataSemestral']

    # Check if dataSemestral is NaT (Not a Time)
    if pd.isna(data_semestral):
        print("dataSemestral is NaT. Skipping input.")
        
    else:
        # Convert the date to string in the desired format
        weighing_date_str = data_semestral.strftime('%d/%m/%Y')
        # Locate the input element
        last_weighing_date_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dateOfLastWeighing")))
        # Send the date string to the input element
        last_weighing_date_input.send_keys(weighing_date_str)

    # Inputs the acess type name drop-down
    acess_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[1]/div[12]/div/div')))
    acess_dropdown.click()
    acess_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), 'Área Comum')]")))
    acess_option.click()

    # Inputs location
    location_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="subLocation"]')))
    location_input.send_keys(row['localizacao'])

    # Inputs sector
    setor = row['setor']
    if pd.isna(setor):
        print("Setor is NaN. Skipping input.")
        
    else:
        sector_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sector"]')))
        sector_input.send_keys(row['setor'])

    # Inputs inmetro seal number
    inmetro_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#inmetroSealNumber")))
    inmetro_input.send_keys(row['numeroImetro'])

    # Inputs full weight
    weight_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fullWeight"]')))
    weight_input.send_keys(row['pesoSemestral'])

    # Submit the form 
    send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/fieldset[2]/button')))
    send_btn.click()

    done_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/form/div/div/button')))
    done_btn.click()
    
# Close the browser
driver.quit()
print("SCRPIT ENDED SUCCESSFULLY 😲!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
