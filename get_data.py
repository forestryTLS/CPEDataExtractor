from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
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

load_dotenv()

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

def print_decorator(func):
    # This just prints the function name before and after
    def wrapper(*args, **kwargs):
        print(f"{'-'*15}STARTING {func.__name__}{'-'*15}")
        result = func(*args, **kwargs)
        print(f"{'+'*15}FINISHED {func.__name__}{'+'*15}")
        return result
    return wrapper

def append_data_to_excel(filename, df_new_data):
    if os.path.isfile(filename):
        try:
            df_old = pd.read_excel(filename)
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            df_combined = df_new_data
        else:
            df_combined = (pd.concat([df_old, df_new_data], ignore_index=True, sort =False)
                .drop_duplicates(keep='first'))

    else:
        df_combined = df_new_data

    df_combined.to_excel(filename, index=False)


@print_decorator
def login():
    driver.get("https://courses.cpe.ubc.ca/new_analytics/enrollments")
    wait = WebDriverWait(driver, 30)
    wait.until(EC.url_contains('enrollments'))

def check_page_source(driver, option):
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
    button = driver.find_element(By.XPATH, "//button[@data-automation='AnalyticsPage__Show__Filters__Button']")
    button.click()
    
    # Wait until the dropdown menu is visible
    wait = WebDriverWait(driver, 10)
    dropdown_menu = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')))
    dropdown_menu.click()
    
    # Add " - " to courses since that differentiates a program from a course
    options_to_select = [course + " - " for course in courses]

    full_option_name = {
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

    # Iterate over the options you want to select
    for option in options_to_select:
        # Type the name of the option to filter the dropdown menu
        dropdown_menu = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
        dropdown_menu.clear()
        dropdown_menu.send_keys(option)

        try:
            wait.until(lambda driver: check_page_source(driver, full_option_name[option]))
            # Then send the ENTER key
            catalog_filter = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
            catalog_filter.send_keys(Keys.ARROW_DOWN)
            catalog_filter.send_keys(Keys.ENTER)

        except (TimeoutException, KeyboardInterrupt):
            print("OPTION NOT FOUND IN TIME", option)

@print_decorator
def filter_enrollment_date(manually_filter):
    if not manually_filter:
        apply = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[form="analytics-filter-form"]'))
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
    for column in df.columns:
        try:
            df[column] = pd.to_numeric(df[column], errors='raise')
        except (ValueError, TypeError):
            pass  # Ignore columns that cannot be converted to numeric

    return df

def extract_table_data(table_data):
    """ Extract aria-labels or text, assumes page has a table """
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
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
    table_data = []
    
    while True:
        table_data = extract_table_data(table_data)
        if check_and_click_next_button() == False:
            break

    # Create a DataFrame from your data
    df = pd.DataFrame(table_data)
    
    df = convert_numeric_columns(df)
    append_data_to_excel(os.environ.get("RAW_DATA_PATH") + "enrollment.xlsx", df)


@print_decorator
def extract_users(manually_filter_users):
    driver.get('https://courses.cpe.ubc.ca/new_analytics/users')

    if manually_filter_users:
        button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-automation="AnalyticsPage__Show__Filters__Button"]')))
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
    append_data_to_excel(os.environ.get("RAW_DATA_PATH") + 'user_data.xlsx', df)
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='This Script uses Selenium to login to Canvas Catalog and extracts enrollments + users')
    
    # Optional command line arguments
    parser.add_argument('--mfe', action='store_true', help='Manually Filter Enrollments. Include this argument if you want the bot to pause when filtering enrollments')
    parser.add_argument('--mfu', action='store_true', help='Manually Filter Users. Include this argument if you want the bot to pause when filtering users')
    valid_courses = ["CACE", "CNR", "CSRP", "CVA", "EFO", "FCM", "HTC", "FHM", "FSTB", "SMS", "TWS", "ZCBS"] 
    parser.add_argument('--courses', nargs='+', choices=valid_courses, default=valid_courses, help='Include courses that you want selected. Example: --courses CACE CNR CVA. Defaults to all courses')
    
    # Parse the command line arguments
    args = parser.parse_args()
    
    login()
    filtering(args.courses)
    filter_enrollment_date(args.mfe)
    extract_enrollment_table()
    extract_users(args.mfu)