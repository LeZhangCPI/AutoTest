import time
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
from datetime import datetime, timedelta

from excelGenerate import append_result, setup_workbook, save_workbook

workbook, sheet = setup_workbook()
calibri_font = Font(name='Arial', size=11)

def get_current_date():
    return datetime.now().strftime("%Y-%m-%d")
current_date = get_current_date()

def clickable_try_action_and_log_result(driver, locator_type, locator_string, action_type, expected_results, failure_message, sheet, item, test_date, tested_by, font):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((locator_type, locator_string))
        )
        element.click()

        append_result(sheet=sheet, item=item, expected_results=expected_results, test_date=test_date,
                      pass_fail="Y", tested_by=tested_by, comments="Test passed successfully.", font=font)
    except Exception as e:
        append_result(sheet=sheet, item=item, expected_results=failure_message, test_date=test_date,
                      pass_fail="N", tested_by=tested_by, comments=f"Test Failed: {str(e)}", font=font)

def interact_with_input_and_log_result(driver, locator_type, locator_string, keys_to_send, success_message, failure_message, sheet, item, test_date, tested_by, font):
    try:
        input_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((locator_type, locator_string))
        )
        input_element.send_keys(keys_to_send)

        dropdown_locator = "#Country_countryApplicationDetailsView_listbox li:nth-child(1)"
        dropdown_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((locator_type, dropdown_locator))
        )
        dropdown_item.click()

        # Log success after entire interaction
        append_result(sheet=sheet, item=item, expected_results=success_message, test_date=test_date,
                      pass_fail="Y", tested_by=tested_by, comments="Country entered and first dropdown selected successfully.", font=font)

    except Exception as e:
        append_result(sheet=sheet, item=item, expected_results=failure_message, test_date=test_date,
                      pass_fail="N", tested_by=tested_by, comments=f"Interaction Failed: {str(e)}", font=font)

def interact_with_date_input_and_log_result(driver, locator_type, locator_string, date_to_enter, success_message, failure_message, sheet, item, test_date, tested_by, font):
    try:
        # Locate the date input element and clear any existing content
        date_input_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((locator_type, locator_string))
        )
        date_input_element.clear()  # Clear any existing date
        date_input_element.send_keys(date_to_enter)  # Enter the new date
        print("Debug: Date entered successfully.")  # Debug print

        # Log success after entering the date
        append_result(sheet=sheet, item=item, expected_results=success_message, test_date=test_date,
                      pass_fail="Y", tested_by=tested_by, comments="Date entered successfully.", font=font)

    except Exception as e:
        print(f"Debug: Error encountered - {str(e)}")  # Exception detail print
        append_result(sheet=sheet, item=item, expected_results=failure_message, test_date=test_date,
                      pass_fail="N", tested_by=tested_by, comments=f"Date Entry Failed: {str(e)}", font=font)

def check_table_rows_and_log(driver, sheet, random_date_forCheck, font):
    print("Debug: Starting check_table_rows_and_log.")
    time.sleep(2)
    try:
        # Ensure the table is loaded
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#actionsGrid_countryApplicationDetailsView table"))
        )
        print("Debug: Table loaded successfully.")

        # Find all rows in the table
        rows = driver.find_elements(By.CSS_SELECTOR, "div#actionsGrid_countryApplicationDetailsView tbody tr")
        print(f"Debug: Found {len(rows)} rows.")

        # Iterate through each row and check conditions
        for idx, row in enumerate(rows):
            action_type = row.find_element(By.CSS_SELECTOR, 'td[role="gridcell"]:nth-child(1)').text
            action_due = row.find_element(By.CSS_SELECTOR, 'td[role="gridcell"]:nth-child(2)').text
            due_date = row.find_element(By.CSS_SELECTOR, 'td[role="gridcell"]:nth-child(3)').text
            print(f"Debug: Row {idx + 1} - ActionType: {action_type}, Action Due: {action_due}, Due Date: {due_date}")

            if action_type and action_due and due_date:
                try:
                    # Parse due_date assuming it is in the format 'DD-Mon-YYYY'
                    due_date_dt = datetime.strptime(due_date, '%d-%b-%Y')
                    print(f"Debug: Parsed Date {due_date_dt.day}")
                    print(f"Debug: random_date_forCheck {random_date_forCheck.day}")

                    # Check conditions
                    if (due_date_dt.day == random_date_forCheck.day and
                            (due_date_dt.month > random_date_forCheck.month or
                             (due_date_dt.month == random_date_forCheck.month and due_date_dt.year > random_date_forCheck.year))):
                        print(f"Debug: Row {idx + 1} meets criteria. Logging PASS.")
                        append_result(sheet=sheet, item=f"Action Type, Action Due, and Due Date of Row {idx + 1}", expected_results="Action Type and Action Due has correct content, and the Due date is generated based on Filing Date",
                                      test_date=datetime.today().strftime('%Y-%m-%d'),
                                      pass_fail="Y", tested_by="Le Zhang",
                                      comments=f"ActionType: {action_type}, Action Due: {action_due}, Due Date: {due_date}",
                                      font=font)
                    else:
                        print(f"Debug: Row {idx + 1} does not meet criteria. Logging FAIL.")
                        append_result(sheet=sheet, item=f"Due Date of Row {idx + 1}", expected_results="FAIL - Date does not meet criteria",
                                      test_date=datetime.today().strftime('%Y-%m-%d'),
                                      pass_fail="N", tested_by="Le Zhang",
                                      comments=f"Due Date: {due_date}", font=font)
                except ValueError as e:
                    print(f"Debug: Date parsing error in Row {idx + 1} - {str(e)}")
                    append_result(sheet=sheet, item=f"Date format of Row {idx + 1}", expected_results="FAIL - Incorrect date format",
                                  test_date=datetime.today().strftime('%Y-%m-%d'),
                                  pass_fail="N", tested_by="Le Zhang",
                                  comments=f"Due Date: {due_date}, Error: {str(e)}", font=font)
            else:
                print(f"Debug: Missing data in Row {idx + 1}. Logging FAIL.")
                append_result(sheet=sheet, item=f"Action Type, Action Due, or Due Date of Row {idx + 1}", expected_results="FAIL - Missing data",
                              test_date=datetime.today().strftime('%Y-%m-%d'),
                              pass_fail="N", tested_by="Le Zhang",
                              comments=f"ActionType: {action_type}, Action Due: {action_due}, Due Date: {due_date}", font=font)

    except Exception as e:
        print(f"Debug: Unexpected error - {str(e)}")

    print("Debug: Completed check_table_rows_and_log.")

def log_datePicker_and_result(driver, date_input_id, date_format, sheet, item, test_date, tested_by, font):
    try:
        date_input_element = WebDriverWait(driver, 10).until(
            lambda driver: driver.find_element(By.ID, date_input_id) if driver.find_element(By.ID, date_input_id).get_attribute('value') else False
        )
        picked_date_str = date_input_element.get_attribute('value')
        picked_date = datetime.strptime(picked_date_str, date_format)
        append_result(sheet, item, "Picked date is present and correct", test_date, "Y", tested_by, f"Picked Date: {picked_date_str}", font)
    except Exception as e:
        print(f"Debug: Exception - {str(e)}")
        append_result(sheet, item, "Failed to pick date", test_date, "N", tested_by, f"Error: {str(e)}", font)

def verify_and_log_date(driver, date_input_id, expected_date_function, sheet, item, font, test_date=None, tested_by="Automation"):
    test_date = test_date or datetime.today().strftime('%Y-%m-%d')
    try:
        exp_date_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, date_input_id))
        )
        time.sleep(2)
        expiration_date_str = exp_date_input.get_attribute("value")
        expiration_date = datetime.strptime(expiration_date_str, '%d-%b-%Y')
        parent_date = WebDriverWait(driver, 10).until(
            lambda driver: driver.find_element(By.ID, "ParentFilDate_countryApplicationDetailsView").get_attribute(
                'value')
        )
        parent_date_forCheck = datetime.strptime(parent_date, '%d-%b-%Y')
        twenty_years_later = expected_date_function(parent_date_forCheck)

        if expiration_date == twenty_years_later:
            append_result(sheet, item, "Date is correct", test_date, "Y", tested_by, f"Found: {expiration_date_str}", font)
            print("Date is correct: 20 years after the specified date.")
        else:
            append_result(sheet, item, "Date is incorrect", test_date, "N", tested_by, f"Found: {expiration_date_str}, Expected: {twenty_years_later.strftime('%d-%b-%Y')}", font)
            print("Date is incorrect. Found:", expiration_date_str, "Expected:", twenty_years_later.strftime('%d-%b-%Y'))

    except Exception as e:
        append_result(sheet, item, "Failed to verify date", test_date, "N", tested_by, f"Error: {str(e)}", font)
        print(f"Error verifying date: {str(e)}")

def calculate_twenty_years_later(date):
    return date + timedelta(days=365 * 20 + 5)

#login
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://corp.cpisystems.com/Login?ReturnUrl=%2FPatent")

WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.modal-close.btn-primary"))).click()

driver.find_element(By.CSS_SELECTOR, "input[id='Email']").send_keys("autotest@admin.com")

driver.find_element(By.ID, "Password").send_keys("Password1")
xpath = "//button[@type='submit' and not(contains(@class, 'btn-Microsoft')) and contains(@class, 'btn-primary')]"
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.XPATH,
    locator_string=xpath,
    action_type='xpath',
    expected_results="Successful login",
    failure_message="Fail to login",
    sheet=sheet,
    item="Login Functionality",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#click patent
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="a#menu-Patent",
    action_type='css',
    expected_results="Patent menu should be accessible",
    failure_message="Fail to access Patent menu",
    sheet=sheet,
    item="Access Patent Menu Test",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
#click country application
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="a.nav-link[href*='/Patent/CountryApplication']",
    action_type='css',
    expected_results="Country Application should be accessible",
    failure_message="Fail to access Country Application",
    sheet=sheet,
    item="Access Country Application Test",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#Invention
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="a.k-button.page-nav",
    action_type='click',
    expected_results="Navigation to the country input clicked",
    failure_message="Failed to click navigation to country input",
    sheet=sheet,
    item="Click Navigation Button",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

# click country 'USA'
interact_with_input_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="input[name='Country_input']",
    keys_to_send='US',
    success_message="Country 'US' selected",
    failure_message="Failed to enter country 'US' or select option",
    sheet=sheet,
    item="Enter Country and Select",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
#click case type
interact_with_input_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="input[name='CaseType_input']",
    keys_to_send='DIV',
    success_message="Case Type 'DIV' selected",
    failure_message="Failed to enter Case Type 'DIV' or select option",
    sheet=sheet,
    item="Enter Case Type and Select",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="#CaseType_countryApplicationDetailsView_listbox li:nth-child(1)",
    action_type='click',
    expected_results="First item in 'Case Type' selected successfully",
    failure_message="Failed to select first item in 'Case Type'",
    sheet=sheet,
    item="Select First Item in Case Type",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#click Status
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="button[aria-label='expand combobox'][aria-controls='ApplicationStatus_countryApplicationDetailsView_listbox']",
    action_type='click',
    expected_results="Combobox expanded successfully",
    failure_message="Failed to expand combobox",
    sheet=sheet,
    item="Expand Application Status Combobox",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="#ApplicationStatus_countryApplicationDetailsView_listbox li:nth-child(2)",
    action_type='click',
    expected_results="Second item in combobox selected successfully",
    failure_message="Failed to select second item in combobox",
    sheet=sheet,
    item="Select Second Item in Application Status",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
#click case number
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="button[aria-label='expand combobox'][aria-controls='CaseNumber_countryApplicationDetailsView_listbox']",
    action_type='click',
    expected_results="Combobox for 'Case Number' expanded successfully",
    failure_message="Failed to expand combobox for 'Case Number'",
    sheet=sheet,
    item="Expand Case Number Combobox",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

# Wait and click the fifth item in the "Case Number" list
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="#CaseNumber_countryApplicationDetailsView_listbox li:nth-child(9)",
    action_type='click',
    expected_results="Seventh item in 'Case Number' selected successfully",
    failure_message="Failed to select seventh item in 'Case Number'",
    sheet=sheet,
    item="Select Seventh Item in Case Number",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#pick up Filling date
year = 2023
month = random.randint(1, 12)
day = random.randint(1, 28)
random_date = datetime(year, month, day).strftime('%m/%d/%Y')
random_date_forCheck = datetime(year, month, day)
interact_with_date_input_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="FilDate_countryApplicationDetailsView",
    date_to_enter=random_date,
    success_message="Filing date entered correctly",
    failure_message="Failed to enter filing date",
    sheet=sheet,
    item="Filing Date Entry",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#pick up Parent Filing Date
current_date_str = datetime.now().strftime('%d-%b-%Y').upper()
interact_with_date_input_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="ParentFilDate_countryApplicationDetailsView",
    date_to_enter=current_date_str,  # Assuming current_date_str is properly formatted
    success_message="Parent Filing Date entered correctly",
    failure_message="Failed to enter Parent Filing Date",
    sheet=sheet,
    item="Parent Filing Date Entry",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
#input application number
interact_with_input_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="AppNumber",
    keys_to_send='_12345678',
    success_message="Application number entered",
    failure_message="Failed to enter application number",
    sheet=sheet,  # assuming 'sheet' is defined
    item="Enter Application Number",
    test_date="current_date",  # assuming 'current_date' is defined
    tested_by="Your Name",
    font="calibri_font"  # assuming 'calibri_font' is defined
)
#Click Save button
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="a.save-changes[title='Save']",
    action_type='click',
    expected_results="Save button clicked successfully",
    failure_message="Failed to click Save button",
    sheet=sheet,
    item="Click Save Button",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#Click Action Tab
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="countryAppActionsTab",
    action_type='click',
    expected_results="Actions tab clicked successfully",
    failure_message="Failed to click Actions tab",
    sheet=sheet,
    item="Click Actions Tab",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

time.sleep(5)
check_table_rows_and_log(driver, sheet, random_date_forCheck, calibri_font)

#click case info to check exp date
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="countryAppMainInfoTab",
    action_type='click',
    expected_results="Case Info tab clicked successfully",
    failure_message="Failed to click Case Info tab",
    sheet=sheet,
    item="Click Case Info Tab",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

#input issue date
year = 2022
month = random.randint(1, 12)
day = random.randint(1, 28)
issue_random_date = datetime(year, month, day).strftime('%m/%d/%Y')
issue_random_date_forCheck = datetime(year, month, day)
interact_with_date_input_and_log_result(
    driver=driver,
    locator_type=By.ID,
    locator_string="IssDate_countryApplicationDetailsView",
    date_to_enter=issue_random_date,  # Assuming issue_random_date is properly formatted
    success_message="Issue Date entered correctly",
    failure_message="Failed to enter Issue Date",
    sheet=sheet,
    item="Issue Date Entry",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

# Click on the 'Save' button
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="a.nav-link.save-changes",
    action_type='click',
    expected_results="Second Save button clicked successfully",
    failure_message="Failed to click second Save button",
    sheet=sheet,
    item="Click Second Save Button",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)

# Click on the 'Accept' button
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.modal-body")))
clickable_try_action_and_log_result(
    driver=driver,
    locator_type=By.CSS_SELECTOR,
    locator_string="button.btn.btn-primary.save-exp-close",
    action_type='click',
    expected_results="Accept button in modal clicked successfully",
    failure_message="Failed to click accept button in modal",
    sheet=sheet,
    item="Click Accept Button in Modal",
    test_date=current_date,
    tested_by="Le Zhang",
    font=calibri_font
)
# Verify the date in the date picker
verify_and_log_date(
    driver=driver,
    date_input_id="ExpDate_countryApplicationDetailsView",
    expected_date_function=calculate_twenty_years_later,
    sheet=sheet,
    item="Expiration Date Verification",
    font=calibri_font,
    tested_by="Le Zhang"
)

save_workbook(workbook)
time.sleep(5)