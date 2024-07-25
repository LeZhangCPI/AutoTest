import random
import time
from datetime import datetime

from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
#login


driver = webdriver.Chrome()


driver.maximize_window()
driver.get("https://r10v2test.cpisystems.com/")

#agree
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.modal-close.btn-primary"))).click()

#login page
driver.find_element(By.CSS_SELECTOR, "input[id='Email']").send_keys("lzhang@computerpackages.com")
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "input[id='Password']").send_keys("Password1")
time.sleep(1)
#click login button
xpath = "//button[@type='submit' and not(contains(@class, 'btn-Microsoft')) and contains(@class, 'btn-primary')]"
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

#click patent
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a#menu-Patent"))).click()
#click duedate
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link[href*='/Patent/Report/DueDate']"))).click()
time.sleep(5)

#click checkbox
#left most checkbox
checkboxes_includes1 = driver.find_elements(By.CLASS_NAME, "disable-ams-only")
for checkbox in checkboxes_includes1:
    if not checkbox.is_selected():
        checkbox.click()
        time.sleep(1)

checkboxes_includes2 = driver.find_elements(By.CSS_SELECTOR, "div.col-md-auto input[type='checkbox']")
for checkbox in checkboxes_includes2:
    if not checkbox.is_selected():
        checkbox.click()
        time.sleep(1)

checkboxes_systems = driver.find_elements(By.CLASS_NAME, "PrintSystemCheckBox")
for checkbox in checkboxes_systems:
    if not checkbox.is_selected():
        checkbox.click()
        time.sleep(1)

#select date
#random ymd
year = 2023
month = random.randint(1, 12)
day = random.randint(1, 28)  # avoid Feb
random_date = datetime(year, month, day).strftime('%m/%d/%Y')  # format date
time.sleep(1)
#start date
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "FromDate_DueDateList")))
date_input = driver.find_element(By.ID, "FromDate_DueDateList")
date_input.clear()  # clear content
date_input.send_keys(random_date)
time.sleep(1)
#to date
current_date_str = datetime.now().strftime('%d-%b-%Y').upper() #select current date
date_input = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.ID, "ToDate_DueDateList"))
)
date_input.clear()
date_input.send_keys(current_date_str)
time.sleep(1)
#print button
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link.page-nav.print-report"))).click()

time.sleep(15)