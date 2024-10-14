import random
import time
from datetime import datetime

from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

driver = webdriver.Chrome()
driver.maximize_window()

driver.get("https://corp.cpisystems.com")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.modal-close.btn-primary"))).click()

# log in 
driver.find_element(By.CSS_SELECTOR, "input[id = 'Email']").send_keys("vzheng@computerpackages.com")
driver.find_element(By.CSS_SELECTOR, "input[id='Password']").send_keys("Keepcalmandlove7%!")

xpath = "//button[@type='submit' and not(contains(@class, 'btn-Microsoft')) and contains(@class, 'btn-primary')]"
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

#click Patent & Master List
driver.find_element(By.CSS_SELECTOR, "a#menu-Patent").click()
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link[href*='/Patent/Report/PatMasterList?name=Master%20List']"))).click()

# click checkboxes
includesList1 = driver.find_elements(By.CLASS_NAME, "CountryDependence")
for checkbox in includesList1:
    if not checkbox.is_selected():
        checkbox.click()

#enter date
year, month, day = 2023, random.randint(1,12), random.randint(1,28)
fromDate = datetime(year, month, day).strftime('%m-%d-Y')

WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "FromDate_PatMasterList")))

enterDate = driver.find_element(By.ID, "FromDate_PatMasterList")
enterDate.send_keys(fromDate)

toDate = datetime.now().strftime('%m-%d-%y')

enterDate = driver.find_element(By.ID, "ToDate_PatMasterList")
enterDate.send_keys(toDate)

# print report
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link.page-nav.print-report"))).click()
time.sleep(5)




