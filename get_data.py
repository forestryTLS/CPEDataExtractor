from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import re
from bs4 import BeautifulSoup
import pandas as pd
import argparse
import os
from dotenv import load_dotenv

import distribute

load_dotenv()

# get selected browser from environment variables, default to Chrome
browser = os.environ.get("BROWSER")

if browser == "Edge":
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
    from selenium.webdriver.edge.service import Service as EdgeService

    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
elif browser == "Firefox":
    from webdriver_manager.firefox import GeckoDriverManager
    from selenium.webdriver.firefox.service import Service as FirefoxService

    driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
elif browser == "Chromium":
    from webdriver_manager.core.utils import ChromeType
    from selenium.webdriver.chrome.service import Service as ChromiumService
    
    driver = webdriver.Chrome(service=ChromiumService(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()))
else:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service as ChromeService

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# NOTE: KEEP THIS VALID_COURSES AND FULL_OPTION_NAME UP TO DATE
VALID_COURSES = ["CBBD", "CACE", "CNR", "CSRP", "CVA", "EFO", "FCM", "HTC", "FHM", "FSTB", "SMS", "TWS", "ZCBS"] 
FULL_OPTION_NAME = {
                    "CBBD - ": "CBBD - Online Micro-Certificate: Circular Bioeconomy Business Development",
                    "CACE - ": "CACE - Online Micro-Certificate: Climate Action and Community Engagement",
                    "CNR - ": 'CNR - Online Micro-Certificate: Co-Management of Natural Resources',
                    "CSRP - ": "CSRP - Online Micro-Certificate: Communication Strategies for Resource Practitioners",
                    "CVA - ": "CVA - Online Micro-Certificate: Climate Vulnerability & Adaptation",
                    "EFO - ": "EFO - Online Micro-Certificate: Environmental Footprints of Organizations",
                    "FCM - ": "FCM - Online Micro-Certificate: Forest Carbon Management",
                    "HTC - ": "HTC - Online Micro-Certificate: Hybrid Timber Construction",
                    "FHM - ": "FHM - Online Micro-Certificate: Forest Health Management",
                    "FSTB - ": "FSTB - Online Micro-Certificate: Fire Safety for Timber Buildings",
                    "SMS - ": "SMS - Online Micro-Certificate: Strategic Management for Sustainability",
                    "TWS - ": "TWS - Online Micro-Certificate: Tall Wood Structures",
                    "ZCBS - ": "ZCBS - Online Micro-Certificate: Zero Carbon Building Solutions"
                    }

ENROLLMENT_STATUSES = ['Active', 'Completed', 'Concluded', 'Dropped']

def print_decorator(func):
    # This just prints the function name before and after, useful for debugging
    def wrapper(*args, **kwargs):
        print(f"{'-'*15}STARTING {func.__name__}{'-'*15}")
        result = func(*args, **kwargs)
        print(f"{'+'*15}FINISHED {func.__name__}{'+'*15}")
        return result
    return wrapper

def append_data_to_excel(filename, df_new_data):
    """ Take in an excel and combine the data, then remove duplicates prioriziting keeping the existing data """
    if os.path.isfile(filename):
        try:
            df_old = pd.read_excel(filename)
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            df_combined = df_new_data
        else:
            df_combined = (pd.concat([df_old, df_new_data], ignore_index=True, sort=False)
                .drop_duplicates(keep='first'))

    else:
        df_combined = df_new_data

    df_combined.to_excel(filename, index=False)
    return df_new_data


@print_decorator
def login():
    """ Open the url which will prompt a login """
    driver.get("https://courses.cpe.ubc.ca/new_analytics/enrollments")
    # Click the login button
    link = driver.find_element(By.XPATH, '//a[@href="http://ubccpe.instructure.com/login/saml"]')
    link.click()

    SECONDS_TO_LOGIN = 90
    wait = WebDriverWait(driver, SECONDS_TO_LOGIN)
    wait.until(EC.url_contains('enrollments'))

def check_page_source(driver, option):
    """ This returns true if the enrollment filtering option has shown up on the page """
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    # Find all <div> elements with the "title" attribute
    divs_with_title = soup.find_all('div', {'title': True})
    
    # Iterate over the matching <div> elements
    for div in divs_with_title:
        text = div.get_text()
        
        # Search for the pattern in the text
        pattern = re.escape(option)
        found = re.search(pattern, text)
        
        if found:
            print("FOUND:", text)
            return True
    
    return False

@print_decorator
def filtering(courses):
    """ This clicks the filter button on the enrollments page and searches for all the courses then selects them """
    wait = WebDriverWait(driver, 10)
    button = wait.until(EC.visibility_of_element_located((By.XPATH,  "//button[@data-automation='Filter__Show__Filters__Button']")))

    button.click()
    
    # Wait until the dropdown menu is visible
    dropdown_menu = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')))
    dropdown_menu.click()
    
    # Add " - " to courses since that differentiates a program from a course
    options_to_select = [course + " - " for course in courses]

    # Iterate over the options you want to select
    for option in options_to_select:
        # Type the name of the option to filter the dropdown menu
        dropdown_menu = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
        dropdown_menu.clear()
        dropdown_menu.send_keys(option)

        try:
            wait.until(lambda driver: check_page_source(driver, FULL_OPTION_NAME[option]))
            # Then send the ENTER key
            catalog_filter = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
            catalog_filter.send_keys(Keys.ARROW_DOWN)
            catalog_filter.send_keys(Keys.ENTER)

        except (TimeoutException, KeyboardInterrupt):
            print("OPTION NOT FOUND IN TIME", option)

@print_decorator
def filter_enrollment_status(status_list):
    """ Apply the specified status filters when the ```--status``` argument is used. """
    wait = WebDriverWait(driver, 10)

    # expand the "Enrollments" accordion
    enrollments_accordion_button = driver.find_element(By.CSS_SELECTOR, 'button[data-automation="FilterPanel__Toggle__Details"]')
    enrollments_accordion_button.click()

    #wait until the "Status" dropdown is visible
    status_dropdown = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Enrollment__Status"]')))
    status_dropdown.click()

    # iterate over status list and select each status for filtering
    for status in status_list:
        status_dropdown = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Enrollment__Status"]')
        status_dropdown.clear()
        status_dropdown.send_keys(status)

        # check if selected status exists in page and select if so
        try:
            wait.until(lambda driver: check_page_source(driver, status))
            status_filter = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Enrollment__Status"]')
            status_filter.send_keys(Keys.ARROW_DOWN)
            status_filter.send_keys(Keys.ENTER)
        except (TimeoutException, KeyboardInterrupt):
            print("COULDN'T FIND STATUS:", status)

@print_decorator
def filter_enrollment_date(manually_filter):
    """ If manually_filter, this clicks the date filter button and waits for user input before continuing """
    if not manually_filter:
        apply = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[form="filter-panel-form"]'))
        )
        apply.click()
        return

    # Toggle the tab to filter enrollment dates
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "css-10dz9lm-toggleDetails__summary"))
    )
    
    element.click()
    input("Please apply any additional filters and hit apply. Once you see the table loaded, please hit enter in this terminal")

def check_and_click_next_button():
    """ If the next button exists, click it and return True, else return False"""
    # Find the span element containing the buttons
    try:
        pagination_span = driver.find_element(By.CLASS_NAME, "css-ighgvd-view--inlineBlock-pagination__pages")

        # Find the button with aria-current="page"
        current_button = pagination_span.find_element(By.CSS_SELECTOR,"button[aria-current='page']")

        # Find the sibling button if it exists
        next_button = current_button.find_element(By.XPATH, "following-sibling::button")
        if next_button:
            print("CLICKED NEXT")
            next_button.click()
            return True
        else:
            return False

    except NoSuchElementException:
        return False
        
def convert_numeric_columns(df):
    """
    An issue with excel is that it will automatically convert numeric data so the raw data and excel data will be considered different,
    Convert the numeric columns so the values compare correctly.
    """
    for column in df.columns:
        try:
            df[column] = pd.to_numeric(df[column], errors='raise')
        except (ValueError, TypeError):
            pass  # Ignore columns that cannot be converted to numeric

    return df

def extract_table_data(table_data):
    """ Extract aria-labels or text, assumes page has a table, uses the data-testid property as the column header """
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
    except TimeoutException:
        print("NO DATA FOUND.")
        exit()
    
    soup = BeautifulSoup(driver.page_source.encode("utf-8"), 'html.parser')

    table = soup.find('table') 
    tbody = table.find('tbody')
    for row in tbody.find_all('tr'):

        row_data = {}
        # iterate over each column in the row
        for td in row.find_all('td'):
            # getting the column's label from data-testid attribute
            if 'data-testid' in td.attrs:
                label = td['data-testid']
                span_with_aria_labels = td.find_all('span', attrs={'aria-label': True})
                # if there are multiple spans with aria-label
                if len(span_with_aria_labels) > 1:
                    for i, span in enumerate(span_with_aria_labels):
                        value = span['aria-label']
                        row_data[label + '_' + str(i)] = value  # unique column name with index
                else:
                    value = td.text
                    row_data[label] = value

        table_data.append(row_data)
    return table_data

@print_decorator
def extract_enrollment_table():
    """ This accumulates the data on each page """
    table_data = []
    
    while True:
        table_data = extract_table_data(table_data)
        if check_and_click_next_button() == False:
            break

    # Create a DataFrame from your data
    df = pd.DataFrame(table_data)
    
    df = convert_numeric_columns(df)
    return append_data_to_excel(os.environ.get("RAW_DATA_PATH_ENROLLMENTS"), df)


@print_decorator
def extract_users(manually_filter_users):
    """ 
    This will go to the users page, and optionally pause to allow for manual user filtering
    It will then go through each page and put the data in the excel
    """
    driver.get('https://courses.cpe.ubc.ca/new_analytics/users')

    if manually_filter_users:
        button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-automation="Filter__Show__Filters__Button"]')))
        button.click()
        input("Please apply any additional filters and hit apply. Once you see the table loaded, please hit enter in this terminal")

    table_data = []
    
    while True:
        table_data = extract_table_data(table_data)
        if check_and_click_next_button() == False:
            break

    # Create a DataFrame from your data
    df = pd.DataFrame(table_data)
    df = convert_numeric_columns(df)

    # Append data to Excel file
    append_data_to_excel(os.environ.get("RAW_DATA_PATH_USERS"), df)
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='This Script uses Selenium to login to Canvas Catalog and extracts enrollments + users')
    
    # Optional command line arguments
    parser.add_argument('--mfe', action='store_true', help='Manually Filter Enrollments. Include this argument if you want the bot to pause when filtering enrollments')
    parser.add_argument('--mfu', action='store_true', help='Manually Filter Users. Include this argument if you want the bot to pause when filtering users')
    parser.add_argument('--courses', nargs='+', choices=VALID_COURSES, default=VALID_COURSES, help='Include courses that you want selected. Example: --courses CACE CNR CVA. Defaults to all courses')
    parser.add_argument('--status', nargs='+', choices=ENROLLMENT_STATUSES, default=ENROLLMENT_STATUSES, help='Indicate which enrollment statuses you wish to filter for. Example: --status Active Completed. Defaults to any status.')

    # Parse the command line arguments
    args = parser.parse_args()
    
    login()
    filtering(args.courses)

    # skip enrollment status filtering if all statuses are selected (redundant)
    if(set(args.status) != set(ENROLLMENT_STATUSES)):
        filter_enrollment_status(args.status)
    
    filter_enrollment_date(args.mfe)
    enrollment_df = extract_enrollment_table()
    extract_users(args.mfu)

    distribute.distribute_enrollment_data(enrollment_df, os.environ.get("RAW_DATA_PATH_USERS"), os.environ.get("PROCESSED_DATA_PATH"))