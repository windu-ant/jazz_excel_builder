from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import os
import time
import pickle


# URL of Jazz website
url = input('Please enter the URL of the "send application documents" results for a job:\n')

# Load the excel data
jnd = pd.read_excel('jazz_no_docs.xlsx')
jndframe = pd.DataFrame()

# Define the webdriver options
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
"download.default_directory": os.getcwd(),  # Define default directory
"download.prompt_for_download": False,  # To auto download the file
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True  # To download PDF files
})

# Set the driver for selenium
driver = webdriver.Chrome(options=options)

# Go to website so we can load the appropriate cookies
driver.get(url)

# COOKIES!
cookies = pickle.load(open("cookies_JAZZ.pkl", "rb"))
for cookie in cookies:
    driver.add_cookie(cookie)
driver.refresh()

# Go to website we actually want to go to
driver.get(url)

# Give the page time to load
time.sleep(5)

# Set the first candidate to be clickable
first_candidate = search_box = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/ui-view/ui-view/ui-view/ui-view/div/div/div[1]/div[3]/table/tbody/tr[2]/td[1]/a')))

# Click it
first_candidate.click()

# Initialize our variables
LOI_location = None
LOI_signed = None
app_signed = None

# Wait for page load
time.sleep(3)

while True:

    # Find the element for LOI then try to check if it is signed or incomplete
    try:
        title_element = driver.find_element(By.XPATH, '//a[@title="Eagle Pass LOI - Updated.docx eSignature Template"]')
        LOI_location = "Eagle Pass"
        
        # Set the parent div
        parent_div = title_element.find_element(By.XPATH, './..')

        # Navigate to its parent div and search for the span with the desired classes
        span_element = parent_div.find_element(By.XPATH, './/span[contains(@class, "label-success") or contains(@class, "label-yellow ng-binding")]')
        
        # Set complete or incomplete if document is signed
        if "label-success" in span_element.get_attribute("class"):
            LOI_signed = "complete"
        elif "label-yellow ng-binding" in span_element.get_attribute("class"):
            LOI_signed = "incomplete"
            
    except NoSuchElementException:
        print("Element not found!")

    # Print the results
    if LOI_location:
        print(f"Document Name: {LOI_location}")
    if LOI_signed:
        print(f"Status: {LOI_signed}")

    # Find the element for the full app then try to check if it is signed or incomplete
    try:
        title_element = driver.find_element(By.XPATH, '//a[@title="CBP Full Application - 2023.pdf eSignature Template"]')
        
        # Set the parent div
        parent_div = title_element.find_element(By.XPATH, './..')

        # Navigate to its parent div and search for the span with the desired classes
        span_element = parent_div.find_element(By.XPATH, './/span[contains(@class, "label-success") or contains(@class, "label-yellow ng-binding")]')
        
        # Set complete or incomplete if document is signed
        if "label-success" in span_element.get_attribute("class"):
            app_signed = "complete"
        elif "label-yellow ng-binding" in span_element.get_attribute("class"):
            app_signed = "incomplete"

        # Get the date of application 
        date_signed_element = parent_div.find_element(By.XPATH, ".//span[@class='jz-utl-text-nowrap ng-binding' and contains(@ng-bind, 'dateTime') and contains(@ng-bind, 'numericShortDate')]")
        date_signed = date_signed_element.text

    except NoSuchElementException:
        print("Element not found!")

    # Print the results
    if app_signed:
        print(f"App Signed: {app_signed}")
    if date_signed:
        print(f"Date of Application: {date_signed}")

    # Now get more info for spreadsheet
    # Name
    get_name_element = driver.find_element(By.XPATH, "//h1[@class='candidate-name fs-data-mask ng-binding']")
    applicant_name = get_name_element.text
    print(applicant_name)
    split_name_at_space = applicant_name.split(" ", 1)
    applicant_first_name = split_name_at_space[0]
    applicant_last_name = split_name_at_space[1] if len(split_name_at_space) > 1 else ""

    # Phone
    get_phone_element = driver.find_element(By.XPATH, "//a[@class='ng-binding' and starts-with(@href, 'tel:')]")
    applicant_phone = get_phone_element.text
    applicant_phone = applicant_phone.replace(" ","").replace("-","").replace("(","").replace(")","").replace("+","")
    print(applicant_phone)

    # Location
    get_location_element = driver.find_element(By.XPATH, "//span[@class='jz-utl-color-black ng-binding' and contains(@ng-bind, 'location')]")
    applicant_loc = get_location_element.text
    print(applicant_loc)
    if "McAllen," in applicant_loc:
        applicant_loc = 'RGV'
    elif "Eagle Pass," in applicant_loc:
        applicant_loc = 'Eagle Pass'

    # Email
    get_email_element = driver.find_element(By.XPATH, "//a[@class='ng-binding' and @ng-click='$ctrl.sendEmail()']")
    applicant_email = get_email_element.text
    print(applicant_email)

    # Write to Excel spreadsheet if email does not exist in column G
    if applicant_email not in jnd['Email'].values:
        # Set all above variables to columns
        new_row = {'Recruiter': "Tony", 'Sector': applicant_loc, 'Location': applicant_loc, 'Last Name': applicant_last_name, 'First Name': applicant_first_name, 'Phone': applicant_phone, 'Email': applicant_email, 'Date Applied': date_signed, 'LOI': LOI_signed, 'APP': app_signed} 
        # Write to sheet
        jnd.loc[len(jnd)] = new_row

        # Save dataframe to excel sheet
        jnd.to_excel('jazz_no_docs.xlsx', index=False)

    next_candidate_element = driver.find_element(By.XPATH, "//button[@class='jz-btn-secondary is-icon-right' and @ng-click='$ctrl.goToAdjacentProfile($ctrl.candidates.next)' and contains(text(), 'Next Candidate')]")
    if next_candidate_element.get_attribute("disabled"):
            break
    next_candidate_element.click()

    time.sleep(3)


# Input wait for testing
input('Press enter to exit...')
driver.quit()