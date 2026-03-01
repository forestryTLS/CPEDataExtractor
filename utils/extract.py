import json
from enum import StrEnum
import re
from typing import TypedDict

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.webdriver import WebDriver as EdgeWebDriver
from selenium.webdriver.firefox.webdriver import WebDriver as FirefoxWebDriver
from selenium.webdriver.chromium.webdriver import ChromiumDriver
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeWebDriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from bs4 import BeautifulSoup
from bs4.element import Tag

import pandas as pd

from utils.logging import bcolors
from utils.common import (
    ENROLLMENT_STATUSES,
    CERTIFICATE_PROGRAMS
)

RE_ID_EMAIL_SPLIT = re.compile(r'\s*(?:Email:|\|)\s*', re.I)
RE_CATALOG_ID_SEARCH = re.compile(r'(#|ID:\s)(?P<id>\d+)')
RE_CATALOG_PROG_NAME_SPLIT = re.compile(r'(?:\s*\-*\s)')
RE_SESSION_NAME = re.compile(r'\d{4}\s*[A-Z]+', re.I)
RE_CONSECUTIVE_SPACES = re.compile(r'\s{2,}')

type SeleniumDriver = EdgeWebDriver | FirefoxWebDriver | ChromiumDriver | ChromeWebDriver



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

# Configuration to quickly set filters for analytics search
SESSION_STORAGE_ENROLMENTS_SETTINGS_KEY = "analytics-settings-enrollments"
SESSION_STORAGE_USERS_SETTINGS_KEY = "analytics-settings-users"

class AccountFilter(TypedDict):
    id: int
    name: str

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

CATALOG_COL_ID_NAME_MAP = {
    "student_name": "Full Name",
    "student_id": "Student Catalog ID",
    "student_email": "Student Email",
    "account_name": "Catalog",
    "program_name": "Program",
    "product_name": "Listing",
    "listing_id": "Listing ID",
    "session": "Session",
    "product_status": "Listing Status",
    "canvas_course_id": "Canvas Course ID",
    "canvas_section_id": "Canvas Section ID",
    "enrolment_id": "Enrollment ID",
    "enrolment_status": "Enrollment Status",
    "custom_fields_relevant-degree-or-experience": "Relevant Degrees or Experience",
    "certificate_offered": "Certificate",
    "requirement_details": "Completion Percentage",
    "registration_date": "Registration Date",
    "enrollment_count": "Enrollment Count",
    "last_enrollment_date": "Last Enrolment Date",
    "transcript": "Transcript",
    "custom_fields_home-address": "Home Address",
    "custom_fields_indigenous-self-declaration": "Self-Identify as Indigenous?",
    "custom_fields_is-fof-alum": "Is Forestry Alum?",
    "custom_fields_mailing-address": "Mailing Address",
    "custom_fields_organization": "Organization",
    "custom_fields_phone-number": "Phone Number",
    "custom_fields_title": "Title",
}

def initialize_selenium_driver(browser: str):
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

    return driver

def login_to_canvas_catalog(driver: SeleniumDriver):
    """ Open Catalog Analytics URL for enrolments and wait for user to authenticate. By default, will wait for 90 seconds. """
    driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")
    # Click the login button
    link = driver.find_element(By.XPATH, '//a[@href="http://ubccpe.instructure.com/login/saml"]')
    link.click()

    login_timeout_seconds = 90

    wait = WebDriverWait(driver, login_timeout_seconds)
    wait.until(EC.url_contains('enrollments'))

def set_enrollments_program_filters(
    enrollments_settings: dict,
    programs: list[str]
):
    """ Format the selected programs as required by the Catalog Analytics session storage filters. """

    if len(programs) == 0:
        programs = CERTIFICATE_PROGRAMS

    # reset the account filter array in case script must do multiple runs
    enrollments_settings["filter"]["account_ids"] = []

    selected_program_abbreviations = [program + " - " for program in programs]

    for course_abbrev in selected_program_abbreviations:
        full_course_catalog_name = FULL_OPTION_NAME[course_abbrev]
        course_catalog_id = CATALOG_PROGRAM_IDS[full_course_catalog_name]

        # add the course to the account filters
        enrollments_settings["filter"]["account_ids"].append(
            AccountFilter(
                id=course_catalog_id,
                name=full_course_catalog_name
            )
        )
    
    return enrollments_settings

def set_enrollment_status_filters(
    enrollments_settings: dict,
    statuses: list[str]
):
    """ Format the selected statuses as required by the Catalog Analytics session storage filters. """
    if len(statuses) == 0:
        statuses = ENROLLMENT_STATUSES

    for status in statuses:
        enrollments_settings["filter"]["enrollment_statuses"].append(
            {
                "id": status.upper(),
                "label": status
            }
        )
    
    return enrollments_settings

def set_record_filters_via_session_storage(
    driver: SeleniumDriver,
    enrollments_settings: dict,
    users_settings: dict
):
    """ Sets the Catalog Analytics filters using session storage. """

    driver.execute_script(f'sessionStorage.setItem("{SESSION_STORAGE_ENROLMENTS_SETTINGS_KEY}", JSON.stringify({json.dumps(enrollments_settings)}))')

    driver.execute_script(f'sessionStorage.setItem("{SESSION_STORAGE_USERS_SETTINGS_KEY}", JSON.stringify({json.dumps(users_settings)}))')

def find_and_click_pagination_next_button(driver: SeleniumDriver):
    """ Attempt to find the 'next page' button in Catalog Analytics pagination and click it if it exists. If button is found, returns True. Otherwise, returns False. """

    try:
        WebDriverWait(driver, 10)
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
    
def convert_df_columns_to_numeric(df: pd.DataFrame):
    """ Attempt to convert the provided DataFrame's columns to numeric. This step aims to mitigate comparison errors caused by Excel's automatic conversion of numeric data. """
    for column in df.columns:
        try:
            df[column] = pd.to_numeric(df[column], errors='raise')
        except (ValueError, TypeError):
            pass  # Ignore columns that cannot be converted to numeric

    return df

class AnalyticsTables(StrEnum):
    ENROLLMENTS = 'enrolments'
    USERS = 'users'

def extract_table_data_to_df(driver: SeleniumDriver):
    """ Attempt to find a table on the current page and extract its data to a DataFrame. """
    table_data = []
    
    while True:
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        except TimeoutException:
            return None
        
        soup = BeautifulSoup(driver.page_source.encode("utf-8"), 'html.parser')

        table = soup.find('table')
        tbody = table.find('tbody')
        for row in tbody.find_all('tr'):
            row_data = {}

            # iterate over each column in the row
            td: Tag
            for td in row.find_all(['td', 'th']):
                # attempt to identify column only using data-testid attribute
                if 'data-testid' in td.attrs:
                    label = td['data-testid']

                    if label == 'student_name':

                        span_name: Tag = td.select_one('div[class*=view] span[class*=truncateText]')
                        span_id_email: Tag = td.select_one('span > span[class*=screenReaderContent]')

                        if span_name or span_id_email:
                            row_data[CATALOG_COL_ID_NAME_MAP[label]] = str(span_name.text).strip() if span_name else ''
                            
                            # attempt to split the ID and email address (put in same line by Catalog)
                            if span_id_email:
                                #raw_student_id, raw_email_address = RE_ID_EMAIL_SPLIT.split(str(span_id_email.text))
                                id_email_split = RE_ID_EMAIL_SPLIT.split(str(span_id_email.text))

                                if len(id_email_split) > 1:
                                    raw_student_id, raw_email_address = id_email_split[:2]
                                elif len(id_email_split) == 1:
                                    raw_student_id, raw_email_address = id_email_split[0]
                                else:
                                    raw_student_id, raw_email_address = str(span_id_email.text).strip()

                                row_data[CATALOG_COL_ID_NAME_MAP['student_email']] = raw_email_address.strip()

                                student_id_match = RE_CATALOG_ID_SEARCH.search(raw_student_id)

                                if student_id_match:
                                    row_data[CATALOG_COL_ID_NAME_MAP['student_id']] = student_id_match.group('id')
                                else:
                                    row_data[CATALOG_COL_ID_NAME_MAP['student_id']] = raw_student_id
                        else:
                            column_text_preview = str(td.text).replace("\n", " ")[:10]
                            print(bcolors.WARNING + f"WARNING: Could not process data for row \"{column_text_preview}...\". Skipping." + bcolors.ENDC)
                            continue
                    elif label == 'product_name':
                        # try to find the <span> tag that contains the full name (only inserted when text is truncated)
                        # if not found, there is no truncation, thus <a> tag text should have full name
                        span_listing_name = td.select_one("a span[class*=screenReaderContent]")

                        if span_listing_name:
                            listing_name = str(span_listing_name.text).strip()
                        else:
                            anchor_listing_name = td.select_one('a')
                            listing_name = str(anchor_listing_name.text).split("\n")[0] if anchor_listing_name else ''

                        row_data[CATALOG_COL_ID_NAME_MAP[label]] = listing_name

                        # attempt to get the session from the listing name
                        listing_name_search = RE_SESSION_NAME.search(listing_name)

                        if listing_name_search:
                            row_data[CATALOG_COL_ID_NAME_MAP['session']] = RE_CONSECUTIVE_SPACES.sub(' ', listing_name_search.group(0))

                        listing_id_search = RE_CATALOG_ID_SEARCH.search(td.text)

                        if listing_id_search:
                            row_data[CATALOG_COL_ID_NAME_MAP['listing_id']] = listing_id_search.group('id')
                    elif label == 'account_name':
                        # try to find the <span> tag that contains the full name (only inserted when text is truncated)
                        # if not found, there is no truncation, thus td tag text should have full name
                        span_catalog_name = td.select_one("span > span[class*=screenReaderContent]")

                        if span_catalog_name:
                            full_catalog_name = str(span_catalog_name.text).strip()
                        else:
                            full_catalog_name = str(td.text).strip()

                        # save the full Catalog name
                        row_data[CATALOG_COL_ID_NAME_MAP[label]] = full_catalog_name
                        
                        # attempt to get the program name (abbreviated) from the catalog name and save it
                        catalog_name_split = RE_CATALOG_PROG_NAME_SPLIT.split(full_catalog_name)

                        if len(catalog_name_split) > 0:
                            row_data[CATALOG_COL_ID_NAME_MAP['program_name']] = catalog_name_split[0]
                    else:
                        if label in CATALOG_COL_ID_NAME_MAP:
                            row_data[CATALOG_COL_ID_NAME_MAP[label]] = str(td.text).strip()                     
            
            if row_data:
                table_data.append(row_data)
    
        if not find_and_click_pagination_next_button(driver):
            break

    return pd.DataFrame(table_data)
