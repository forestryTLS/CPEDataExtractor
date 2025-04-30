import argparse
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
import os
from dotenv import load_dotenv
from orders_data_processor import ordersDataProcessor
import time
from selenium.common.exceptions import ElementClickInterceptedException




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

# NOTE: KEEP THIS VALID_COURSES AND NUMBER_OF_COURSES_IN_PROGRAM UP TO DATE
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
    "LCACF",
    "LLFM",
    "EBSC",
    "Foundations of Silviculture"
]


NUMBER_OF_COURSES_IN_PROGRAM = {
    "CVA": 5,
    "CSRP": 4,
    "CNR": 3,
    "CBBD": 4,
    "CACE": 3,
    "FCM": 4,
    "FHM": 4,
    "SMS": 4,
    "LCACF": 3,
    "ZCBS": 4,
    "EFO": 3,
    "HTC": 4,
    "TWS": 4,
    "FMP": 4,
    "EBSC": 4,
    "LLFM": 4,
    "FSTB": 4,
    "Foundations of Silviculture": 4
}

SESSION = "2025 Spring"

ENROLLMENT_STATUSES = ['Active', 'Completed', 'Concluded', 'Dropped']

def print_decorator(func):
    # This just prints the function name before and after, useful for debugging
    def wrapper(*args, **kwargs):
        print(f"{'-' * 15}STARTING {func.__name__}{'-' * 15}")
        result = func(*args, **kwargs)
        print(f"{'+' * 15}FINISHED {func.__name__}{'+' * 15}")
        return result

    return wrapper

@print_decorator
def login():
    """ Open the url which will prompt a login """
    driver.get("https://courses.cpe.ubc.ca/new_analytics/orders")
    # Click the login button
    link = driver.find_element(By.XPATH, '//a[@href="http://ubccpe.instructure.com/login/saml"]')
    link.click()

    SECONDS_TO_LOGIN = 90
    wait = WebDriverWait(driver, SECONDS_TO_LOGIN)
    wait.until(EC.url_contains('orders'))

def check_page_source(driver, option):
    """ This returns true if the enrollment filtering option has shown up on the page """
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all <div> elements with the "title" attribute
    divs_with_title = soup.find_all('div', {'title': True})

    # Iterate over the matching <div> elements
    for div in divs_with_title:
        text = div.get_text()
        print(text)

        # Search for the pattern in the text
        pattern = re.escape(option)
        print(pattern)
        found = re.search(pattern, text)

        if found:
            print("FOUND:", text)
            return True

    return False

def extract_table_data(table_data):
    """ Extract aria-labels or text, assumes page has a table, uses the data-testid property as the column header """
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
    except TimeoutException:
        print("NO DATA FOUND.")
        # exit()
        return

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

                # spans_with_aria_labels = td.find_all('span', attrs={'aria-label': True})

                if label == 'student_name':
                    spans_with_aria_labels = td.find_all(lambda tag: tag.name == 'span' and tag.has_attr('aria-label'))

                    search_string = td.text
                    name_regex = '(^[0-9A-Za-z\\u0100-\\u017FÀ-ÖØ-öø-ÿ\\s\\-\\(\\)\'\\.]+)'
                    email_regex = '([A-z0-9\\.\\#\\-\\_\\|]+@[A-z0-9\\.\\-]{4,})'
                    full_regex = f'{name_regex}(#[0-9]+)(\\s\\|\\s)?{email_regex}?'

                    full_match = re.search(full_regex, search_string)

                    if full_match:
                        name_found = False
                        email_found = False

                        # if <span> with aria-label exists, get name and/or email from that
                        if len(spans_with_aria_labels) > 0:
                            for span in spans_with_aria_labels:
                                name_match = re.search(name_regex, span['aria-label'], re.I)
                                email_match = re.search(email_regex, span['aria-label'], re.I)
                                if name_match:
                                    row_data[f'{label}_0'] = name_match.group(1)
                                    name_found = True
                                if email_match:
                                    row_data[f'{label}_1'] = email_match.string
                                    email_found = True

                                    # check if the full name and/or email were found in a span's aria-label property
                        # if not, get from innerText match
                        if name_found is False:
                            row_data[f'{label}_0'] = full_match.group(1)
                        if email_found is False:
                            if (len(full_match.groups()) > 1):
                                row_data[f'{label}_1'] = ''.join(map(str, full_match.groups()[1:]))
                            else:
                                row_data[f'{label}_1'] = '—'
                        # otherwise, try to get name and/or email from td contents
                        # else:
                        #     if len(full_match.groups()) > 1:
                        #         row_data[f'{label}_0'] = full_match.group(1)
                        #         row_data[f'{label}_1'] = ''.join(map(str,full_match.groups())[1:])
                        #     else:
                        #         row_data[f'{label}_1'] = '—'
                    # if no regex match for td innerText, insert full innerText into first column
                    else:
                        row_data[f'{label}_0'] = search_string
                elif label == 'product_name':
                    span_with_aria_label = td.find(lambda tag: tag.name == 'span' and tag.has_attr('aria-label'))

                    # if truncated text, get full listing name from aria-label and id from innerText
                    if span_with_aria_label:
                        row_data[f'{label}_0'] = span_with_aria_label['aria-label']

                        id_pattern = re.compile('[0-9]{4,}$')

                        row_data[f'{label}_1'] = id_pattern.search(td.text).group(0)

                    # if no truncated text, get listing name and id from innerText
                    else:
                        id_pattern = re.compile('[0-9]{4,}$')

                        match = id_pattern.search(td.text)

                        if match:
                            listing_id = match.group(0)

                            screen_reader_span = td.find_all("span",
                                                             class_=re.compile("screenReaderContent", re.IGNORECASE),
                                                             limit=1)

                            # only listing names that overflow the cell contain a <span> element with the ...-screenReaderContent class
                            if len(screen_reader_span) > 0:
                                listing_name = screen_reader_span[0].text
                            else:
                                listing_name = td.text.replace(listing_id, "")

                            row_data[f'{label}_0'] = listing_name
                            row_data[f'{label}_1'] = listing_id
                        else:
                            row_data[f'{label}_0'] = td.text

                else:
                    row_data[label] = td.text

        table_data.append(row_data)
    return table_data

def find_and_click_next_page():
    """If there is more than one page of results, find the next page <button> and click it."""

    try:
        wait = WebDriverWait(driver, 10)
        div_pagination = driver.find_element(By.CSS_SELECTOR, "[data-automation='Pagination']")

        button_next_page = div_pagination.find_element(By.CSS_SELECTOR,
                                                       "li:has(button[aria-current='page']) + li > button")

        if button_next_page:
            print(f"Navigating to page {button_next_page.text}...")
            driver.execute_script("arguments[0].click();", button_next_page)
            return True

        return False
    except NoSuchElementException:
        print("No additional pages found. Proceeding...")
        return False

@print_decorator
def extract_enrollment_table():
    """ This accumulates the data on each page """
    table_data = []

    while True:
        table_data = extract_table_data(table_data)
        if find_and_click_next_page() == False:
            break
    if table_data is None:
        return
    # Create a DataFrame from your data
    df = pd.DataFrame(table_data)

    df = convert_numeric_columns(df)
    return append_data_to_excel(os.environ.get("RAW_DATA_PATH_ENROLLMENTS"), df)

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

def append_data_to_excel(filename, df_new_data):
    # """ Take in an excel and combine the data, then remove duplicates prioriziting keeping the existing data """

    df_combined = df_new_data

    df_combined[['Name', 'Email']] = df_combined['purchaser_name'].str.extract(r'^([\w\s\'\.,-]+)#\d+ \| ([^#]+)')
    df_combined["bulk_seats"] = df_combined["bulk_purchase"].apply(lambda x: int(x[3:].split()[0]) if x.startswith("Yes") else 0)
    df_combined.to_excel(filename, index=False)
    return df_new_data

@print_decorator
def filtering(courses, session):
    """ This removes all the applied filters, clicks the filter button on the orders page, searches for catalog
    and searches for all the courses the selected program for the given session and then selects them """

    read_excel_if_exists()
    for program in courses:

        wait = WebDriverWait(driver, 10)

        # Remove Existing filters from previous search
        try:
            active_filters = wait.until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'css-1rh6zyi-view-tag'))
            )
            while active_filters:
                try:
                    active_filters[0].click()
                except ElementClickInterceptedException:
                    driver.execute_script("arguments[0].click();", active_filters[0])
                wait.until(EC.staleness_of(active_filters[0]))
                active_filters = wait.until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, 'css-1rh6zyi-view-tag'))
                )

        except TimeoutException:
            print("No more existing filters")

        button = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//button[@data-automation='Filter__Show__Filters__Button']")))

        button.click()

        # Wait until the dropdown menu is visible
        dropdown_menu = wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')))
        dropdown_menu.click()

        notFound = 0
        option = program + " - "

        dropdown_menu = driver.find_element(By.CSS_SELECTOR, 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
        dropdown_menu.clear()
        dropdown_menu.send_keys(option)

        try:
            wait.until(lambda driver: check_page_source(driver, option))
            # Then send the ENTER key
            catalog_filter = driver.find_element(By.CSS_SELECTOR,
                                                 'input[data-automation="AnalyticsPage__Filter__Catalog"]')
            catalog_filter.send_keys(Keys.ARROW_DOWN)
            catalog_filter.send_keys(Keys.ENTER)

        except (TimeoutException, KeyboardInterrupt):
            print("OPTION NOT FOUND IN TIME", option)
            order_data_processor = ordersDataProcessor(session, program)
            order_data_processor.save_all_data()
            continue

        print("___Reached Selecting Listing___")

        if program == "CVA":
            numListings = NUMBER_OF_COURSES_IN_PROGRAM[program] + 2
        else:
            numListings = NUMBER_OF_COURSES_IN_PROGRAM[program] + 1

        for listingNumber in range(numListings):
            dropdown_menu = driver.find_element(By.CSS_SELECTOR,
                                                'input[data-automation="AnalyticsPage__Filter__Listing"]')
            dropdown_menu.clear()
            dropdown_menu.send_keys(session)

            try:
                wait.until(lambda driver: check_page_source(driver, session))
                # Then send the ENTER key
                catalog_filter = driver.find_element(By.CSS_SELECTOR,
                                                     'input[data-automation="AnalyticsPage__Filter__Listing"]')
                catalog_filter.send_keys(Keys.ARROW_DOWN)
                catalog_filter.send_keys(Keys.ENTER)

            except (TimeoutException, KeyboardInterrupt):
                notFound = notFound + 1
                print("OPTION NOT FOUND IN TIME", session)

        if notFound == numListings:
            apply = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[form="filter-panel-form"]'))
            )
            apply.click()
            order_data_processor = ordersDataProcessor(session, program)
            order_data_processor.save_all_data()
            continue

        apply = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[form="filter-panel-form"]'))
        )
        apply.click()
        enrollment_df = extract_enrollment_table()
        if enrollment_df is None:
            order_data_processor = ordersDataProcessor(session, program)
            order_data_processor.save_all_data()
            continue
        order_data_processor = ordersDataProcessor(session, program)
        order_data_processor.read_orders_table('enrollment.xlsx')

    df = pd.read_excel('financeTable.xlsx', sheet_name=session)
    df = calculate_finance_subtotal(df)
    df.to_excel('financeTable.xlsx', sheet_name=session, index=False)
    finance_file_name = os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH")+"FinanceSheet.xlsx"
    df.to_excel(finance_file_name, sheet_name=session, index=False)

    df = pd.read_excel('enrolmentTable.xlsx', sheet_name=session)
    df = calculate_total_enrolment_by_column(df)
    df.to_excel('enrolmentTable.xlsx', sheet_name=session, index=False)
    enrolment_file_name = os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH") + "EnrolmentSheet.xlsx"
    df.to_excel(enrolment_file_name, sheet_name=session, index=False)
    return

def read_excel_if_exists():
    """Read the financial and enrolment sheet for the session if already exists"""

    if os.path.isfile(os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH") + "EnrolmentSheet.xlsx"):
        try:
            file = pd.ExcelFile(os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH") + "EnrolmentSheet.xlsx")
            if SESSION in file.sheet_names:
                df = file.parse(SESSION)
                df.to_excel('enrolmentTable.xlsx', sheet_name = SESSION, index=False)
        except Exception as e:
            print(f"Error reading the Excel file: {e}")

    if os.path.isfile(os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH")+"FinanceSheet.xlsx"):
        try:
            file = pd.ExcelFile(os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH")+"FinanceSheet.xlsx")
            if SESSION in file.sheet_names:
                df = file.parse(SESSION)
                df.to_excel('financeTable.xlsx', sheet_name = SESSION, index=False)
        except Exception as e:
            print(f"Error reading the Excel file: {e}")



def calculate_finance_subtotal(df):
    """ Calculates the subtotal row of the finance table """

    if "Subtotal" in df['Program'].values:
        df.drop(df[df['Program'] == "Subtotal"].index, inplace=True)
    individual_course_total = calculate_individual_course_total(df)
    total = df['Total Amount'].sum()
    FSG_total_amount = df['Total Claim Amount'].sum()
    partial_FSG_claim_amount = df['Partial FSG & Paid FSG Claim Amount'].sum()
    partial_FSG_amount_forestry = df['Partial FSG & Paid Amount to Forestry'].sum()
    dropped_admin_amount = df['Dropped Admin Amount'].sum()
    full_program_discount = df['Full Program Discount Amount'].sum()
    bulk_full_program = df['Bulk Registration in Catalog'].sum()
    non_FSG_full_program = df['Non FSG Full Program (including bank transfer)'].sum() * ordersDataProcessor.FULL_PROGRAM_COST
    total_row = {"Program": "Subtotal", 'Full FSG': FSG_total_amount,
                 'Partial FSG & Paid FSG Claim Amount': partial_FSG_claim_amount,
                 'Partial FSG & Paid Amount to Forestry': partial_FSG_amount_forestry,
                 'Non FSG Full Program (including bank transfer)': non_FSG_full_program,
                 'Non FSG Individual': individual_course_total,
                 'Dropped Admin Amount': dropped_admin_amount,
                 'Full Program Discount Amount': full_program_discount,
                 'Bulk Registration in Catalog': bulk_full_program, "Total Amount": total}
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    df.drop('Total Claim Amount', axis=1, inplace=True)
    return df

def calculate_individual_course_total(df):
    """ Calculates the total amount received through individual course registrations in a program"""
    total = 0
    for program in df["Program"].values:
        p = program.split(" ")
        if p[0] in ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM and \
                ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[p[0]] == 3:
            individual_course_cost = 850
        else:
            individual_course_cost = 650
        total = total + (individual_course_cost*(df.loc[df["Program"] == program, "Non FSG Individual"].iloc[0]))
    return total

def calculate_total_enrolment_by_column(df):
    """ Calculates the total enrolment number for each column in the enrolment table """

    if "Subtotal" in df['Program'].values:
        df.drop(df[df['Program'] == "Subtotal"].index, inplace=True)
    total_row = {"Program": "Subtotal"}
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    for cols in df.columns:
        if cols in ["FSG", "Non FSG", "Returning Non FSG", "Free Slots", "FSG & Paid", "Total Enrolment"] or \
                bool(re.fullmatch(r"\d+\s+returning FSG", cols)):
            total = df[cols].sum()
            df.loc[df["Program"] == "Subtotal", cols] = total

    return df

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='This Script uses Selenium to login to Canvas Catalog and extracts enrollments + users')

    # Optional command line arguments
    parser.add_argument('--mfe', action='store_true',
                        help='Manually Filter Enrollments. Include this argument if you want the bot to pause when filtering enrollments')
    parser.add_argument('--mfu', action='store_true',
                        help='Manually Filter Users. Include this argument if you want the bot to pause when filtering users')
    parser.add_argument('--courses', nargs='+', choices=VALID_COURSES, default=VALID_COURSES,
                        help='Include courses that you want selected. Example: --courses CACE CNR CVA. Defaults to all courses')
    parser.add_argument('--status', nargs='+', choices=ENROLLMENT_STATUSES, default=ENROLLMENT_STATUSES,
                        help='Indicate which enrollment statuses you wish to filter for. Example: --status Active Completed. Defaults to any status.')
    parser.add_argument('--session', default=SESSION,
                        help='Indicate which enrollment session you wish to filter for. Example: --session "2025 Spring". Defaults to 2025 Spring.')

    # Parse the command line arguments
    args = parser.parse_args()

    login()
    filtering(args.courses, args.session)

    # enrollment_df = extract_enrollment_table()