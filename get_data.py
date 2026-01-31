from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import re
from bs4 import BeautifulSoup
from bs4.element import Tag
import pandas as pd
import argparse
import os
from dotenv import load_dotenv
from datetime import date, timedelta
from utils.logging import bcolors
import distribute
import json

from typing import TypedDict

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
VALID_COURSES = [
    "CBBD", 
    "CACE", 
    "CNR", 
    "CSRP", 
    "CVA", 
    "EFO", 
    "FCM", 
    "HTC", 
    "FHM", 
    "FSTB", 
    "SMS", 
    "TWS", 
    "ZCBS", 
    "FMP", 
    "EBSC", 
    "LCACF",
    "LLFM",
    "FAS",
    "CLF",
    "CGF",
    "FCMo"
]
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
    "ZCBS - ": "ZCBS - Online Micro-Certificate: Zero Carbon Building Solutions",
    "FMP - ": "FMP - Online Micro-Certificate: Forest Management Planning",
    "EBSC - ": "EBSC - Online Micro-Certificate: Engineered Bamboo for Sustainable Construction",
    "LCACF - ": "LCACF - Online Micro-Certificate: Life Cycle Assessment in Clean Fuels",
    "LLFM - ": "LLFM - Online Micro-Certificate: Landscape Level Forest Modeling",
    "FAS - ": "FAS - Online Micro-Certificate: Foundations of Advanced Silviculture",
    "CLF - ": "CLF - Online Micro-Certificate: Advanced Life Cycle Assessment of Clean Liquid Fuels",
    "CGF - ": "CGF - Online Micro-Certificate: Advanced Life Cycle Assessment of Clean Gaseous Fuels",
    "FCMo - ": "FCMo - Online Micro-Certificate: Forest Carbon Modelling"
}

ENROLLMENT_STATUSES = ['Active', 'Completed', 'Concluded', 'Dropped']

# Configuration to quickly set filters for analytics search
SESSION_STORAGE_ENROLMENTS_SETTINGS_KEY = "analytics-settings-enrollments"
SESSION_STORAGE_USERS_SETTINGS_KEY = "analytics-settings-users"

DEFAULT_CREATION_DATE_FROM = (date.today() - timedelta(days=7)).strftime("%Y-%m-%d")
DEFAULT_CREATION_DATE_TO = date.today().strftime("%Y-%m-%d")

class AccountFilter(TypedDict):
    id: int
    name: str

DEFAULT_ENROLMENTS_SETTINGS_OBJECT = {
    "filter": {
        "account_ids": [],
        "product_ids": [],
        "product_statuses": [],
        "student_ids": [],
        "enrollment_date_preset": "past_week",
        "enrollment_date_from": DEFAULT_CREATION_DATE_FROM,
        "enrollment_date_to": DEFAULT_CREATION_DATE_TO,
        "enrollment_statuses": [],
        "completion_date_preset": "all_time",
        "completion_date_from": "",
        "completion_date_to": "",
        "enrollment_completion_percentage_min": "",
        "enrollment_completion_percentage_max": ""
    },
    "page": 0,
    "pageSize": 10,
    "search": "",
    "sortBy": "enrollment_date",
    "sortDirection": "desc",
    "wideContent": True,
    "showChart": False,
    "baseAccountId": 512
}

DEFAULT_USERS_SETTINGS_OBJECT = {
    "filter": {
        "account_ids": [],
        "student_ids": [],
        "enrollment_count_min": "",
        "enrollment_count_max": "",
        "last_enrollment_date_preset": "all_time",
        "last_enrollment_date_from": "",
        "last_enrollment_date_to": "",
        "registration_date_preset": "past_week",
        "registration_date_from": DEFAULT_CREATION_DATE_FROM,
        "registration_date_to": DEFAULT_CREATION_DATE_TO
    },
    "page": 0,
    "pageSize": 10,
    "search": "",
    "sortBy": "registration_date",
    "sortDirection": "desc",
    "wideContent": False,
    "showChart": True,
    "baseAccountId": 512
}

CATALOG_PROGRAM_IDS = {
    "CACE - Online Micro-Certificate: Climate Action and Community Engagement": 526,
    "CBBD - Online Micro-Certificate: Circular Bioeconomy Business Development": 1581,
    "CGF - Online Micro-Certificate: Advanced Life Cycle Assessment of Clean Gaseous Fuels": 2693,
    "CLF - Online Micro-Certificate: Advanced Life Cycle Assessment of Clean Liquid Fuels": 2569,
    "CNR - Online Micro-Certificate: Co-Management of Natural Resources": 521,
    "CSRP - Online Micro-Certificate: Communication Strategies for Resource Practitioners": 785,
    "CVA - Online Micro-Certificate: Climate Vulnerability & Adaptation": 515,
    "EBSC - Online Micro-Certificate: Engineered Bamboo for Sustainable Construction": 1958,
    "EFO - Online Micro-Certificate: Environmental Footprints of Organizations": 1112,
    "FAS - Online Micro-Certificate: Foundations of Advanced Silviculture": 2654,
    "FCM - Online Micro-Certificate: Forest Carbon Management": 520,
    "FHM - Online Micro-Certificate: Forest Health Management": 946,
    "FMP - Online Micro-Certificate: Forest Management Planning": 1723,
    "FSTB - Online Micro-Certificate: Fire Safety for Timber Buildings": 951,
    "HTC - Online Micro-Certificate: Hybrid Timber Construction": 956,
    "LCACF - Online Micro-Certificate: Life Cycle Assessment in Clean Fuels": 1944,
    "LLFM - Online Micro-Certificate: Landscape Level Forest Modeling": 1953,
    "SMS - Online Micro-Certificate: Strategic Management for Sustainability": 780,
    "TWS - Online Micro-Certificate: Tall Wood Structures": 941,
    "ZCBS - Online Micro-Certificate: Zero Carbon Building Solutions": 1116,
    "FCMo - Online Micro-Certificate: Forest Carbon Modelling": 2991
}


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
    driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")
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
def filter_records_through_session_storage(courses: list[str], statuses: list[str]):
    """ Adds the selected program filters to the session storage and navigates to enrolments page. """

    if not courses or len(courses) == 0:
        courses = VALID_COURSES
    
    num_filter_courses = len(courses)

    if num_filter_courses > 20:
        print(bcolors.WARNING + f"WARNING: More than 20 programs are selected. Since Catalog cannot filter for more than 20 programs, the last {num_filter_courses - 20} courses will be excluded." + bcolors.ENDC)
        print(bcolors.WARNING + "Excluded courses: " + ", ".join(courses[20:]) + bcolors.ENDC)
        courses = courses[:20]

    if not statuses or len(statuses) == 0:
        statuses = ENROLLMENT_STATUSES

    enrolments_settings_dict = dict(DEFAULT_ENROLMENTS_SETTINGS_OBJECT)

    selected_course_abbreviations = [course + " - " for course in courses]

    # add the course catalog filters
    for course_abbrev in selected_course_abbreviations:
        full_course_catalog_name = FULL_OPTION_NAME[course_abbrev]
        course_catalog_id = CATALOG_PROGRAM_IDS[full_course_catalog_name]

        # add the course to the account filters
        enrolments_settings_dict["filter"]["account_ids"].append(
            AccountFilter(
                id=course_catalog_id,
                name=full_course_catalog_name
            )
        )
    
    # add the status filters
    for status in statuses:
        enrolments_settings_dict["filter"]["enrollment_statuses"].append(
            {
                "id": status.upper(),
                "label": status
            }
        )
    
    # update the corresponding session storage key
    driver.execute_script(f'sessionStorage.setItem("{SESSION_STORAGE_ENROLMENTS_SETTINGS_KEY}", JSON.stringify({json.dumps(enrolments_settings_dict)}))')

    # TODO: Add support for modifying "Users" filters
    driver.execute_script(f'sessionStorage.setItem("{SESSION_STORAGE_USERS_SETTINGS_KEY}", JSON.stringify({json.dumps(DEFAULT_USERS_SETTINGS_OBJECT)}))')

    # navigate to the enrolments page and wait for table navigation element to load
    driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")

    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, "[data-sortable-table]")
    ))
    

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
    
def find_and_click_next_page():
    """If there is more than one page of results, find the next page <button> and click it."""

    try:
        wait = WebDriverWait(driver, 10)
        div_pagination = driver.find_element(By.CSS_SELECTOR, "[data-automation='Pagination']")

        button_next_page = div_pagination.find_element(By.CSS_SELECTOR, "li:has(button[aria-current='page']) + li > button") 
            
        if button_next_page:
            print(f"Navigating to page {button_next_page.text}...")
            driver.execute_script("arguments[0].click();", button_next_page)
            return True
        
        return False
    except NoSuchElementException:
        print("No additional pages found. Proceeding...")
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
        td: Tag
        for td in row.find_all(['td', 'th']):
            # getting the column's label from data-testid attribute
            if 'data-testid' in td.attrs:
                
                label = td['data-testid']

                if label == 'student_name':
                    search_string = td.text

                    # span_name: Tag = td.find_all(lambda tag: tag.name == 'span' and 'truncateText' in tag['class'])
                    # span_email: Tag = td.find_all(lambda tag: tag.name == 'span' and 'screenReaderContent' in tag['class'])

                    span_name: Tag = td.select_one('div[class*=view] span[class*=truncateText]')
                    span_email: Tag = td.select_one('span > span[class*=screenReaderContent]')

                    if span_name or span_email:
                        row_data[f'{label}_0'] = str(span_name.text).strip() if span_name else ''
                        row_data[f'{label}_1'] = re.sub(r"\,\s*Email:\s*", " | ", re.sub(r"ID:\s*", "#", str(span_email.text))).strip() if span_email else ''
                        
                    # if no regex match for td innerText, insert full innerText into first column
                    else:
                        column_text_preview = str(td.text).replace("\n", " ")[:10]
                        print(bcolors.WARNING + f"WARNING: Could not process data for row \"{column_text_preview}...\". Skipping." + bcolors.ENDC)
                        continue
                elif label == 'product_name':
                    # try to find the <span> tag that contains the full name (only inserted when text istruncated)
                    # if not found, there is no truncation, thus <a> tag text should have full name
                    span_listing_name = td.select_one("a span[class*=screenReaderContent]")
                    
                    if span_listing_name:
                        listing_name = str(span_listing_name.text).strip()
                    else:
                        anchor_listing_name = td.select_one('a')
                        listing_name = str(anchor_listing_name.text).split("\n")[0] if anchor_listing_name else None
                    
                    # listing ID row always has the screenReaderContent <span> (I can't figure out why this is different for the life of me)
                    span_listing_id = td.select_one('div + span > span[class*=screenReaderContent]')
                    listing_id = str(span_listing_id.text).replace("ID:", "").strip() if span_listing_id else None

                    if listing_name or listing_id:
                        row_data[f'{label}_0'] = listing_name if listing_name else ""
                        row_data[f'{label}_1'] = listing_id if listing_id else ""
                    else:
                        column_text_preview = str(td.text).replace("\n", " ")[:10]
                        print(bcolors.WARNING + f"WARNING: Could not process data for row \"{column_text_preview}...\". Skipping." + bcolors.ENDC)
                        continue
                else:
                    row_data[label] = td.text

        table_data.append(row_data)
    return table_data

@print_decorator
def extract_enrollment_table():
    """ This accumulates the data on each page """
    table_data = []
    
    while True:
        table_data = extract_table_data(table_data)
        if find_and_click_next_page() == False:
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
    driver.get('https://courses.cpe.ubc.ca/analytics/users')

    if manually_filter_users:
        button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-automation="Filter__Show__Filters__Button"]')))
        button.click()
        input("Please apply any additional filters and hit apply. Once you see the table loaded, please hit enter in this terminal")

    table_data = []
    
    while True:
        table_data = extract_table_data(table_data)
        if find_and_click_next_page() == False:
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
    parser.add_argument('--courses', nargs='+', choices=VALID_COURSES, default=VALID_COURSES, type=str, help='Include courses that you want selected. Example: --courses CACE CNR CVA. Defaults to all courses')
    parser.add_argument('--status', nargs='+', choices=ENROLLMENT_STATUSES, default=ENROLLMENT_STATUSES, type=str.capitalize, help='Indicate which enrollment statuses you wish to filter for. Example: --status Active Completed. Defaults to any status.')

    # Parse the command line arguments
    args = parser.parse_args()
    
    login()
    #filtering(args.courses)

    # skip enrollment status filtering if all statuses are selected (redundant)
    # if(set(args.status) != set(ENROLLMENT_STATUSES)):
    #     filter_enrollment_status(args.status)

    filter_records_through_session_storage(args.courses, args.status)
    
    #filter_enrollment_date(args.mfe)
    enrollment_df = extract_enrollment_table()
    extract_users(args.mfu)

    distribute.distribute_enrollment_data(enrollment_df, os.environ.get("RAW_DATA_PATH_USERS"), os.environ.get("PROCESSED_DATA_PATH"))