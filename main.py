import argparse
import os
from pathlib import Path

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

import pandas as pd 
from dotenv import load_dotenv

load_dotenv()

from utils.common import (
    bcolors,

    CERTIFICATE_PROGRAMS, 
    ENROLLMENT_STATUSES,
    DEFAULT_ENROLMENTS_SETTINGS_OBJECT,
    DEFAULT_USERS_SETTINGS_OBJECT
)
from utils.gui import display_filter_config_gui
from utils.extract import (
    initialize_selenium_driver, 
    login_to_canvas_catalog,
    set_enrollments_program_filters,
    set_enrollment_status_filters,
    extract_table_data_to_df,
    set_record_filters_via_session_storage
)

if __name__ == "__main__":
    # ===========================
    # REGISTRATION FOLDER PARSING
    # ===========================
    try:
        registration_folder_path = Path(os.environ.get("REGISTRATION_DATA_FOLDER_PATH"))
    except TypeError:
        raise Exception("The required environment variable REGISTRATION_DATA_FOLDER_PATH must be set to a valid string path.")

    if not registration_folder_path.is_dir():
        raise Exception(f"The path {registration_folder_path.absolute()} does not point to a valid directory.")

    # =====================
    # CLI ARGUMENT PARSING
    # =====================
    parser = argparse.ArgumentParser(description='This Script uses Selenium to login to Canvas Catalog and extracts enrollments + users')
    
    # Optional command line arguments
    parser.add_argument('--mfe', action='store_true', help='Manually Filter Enrollments. Include this argument if you want the bot to pause when filtering enrollments')
    parser.add_argument('--mfu', action='store_true', help='Manually Filter Users. Include this argument if you want the bot to pause when filtering users')
    parser.add_argument('--program', nargs='+', choices=CERTIFICATE_PROGRAMS, default=CERTIFICATE_PROGRAMS, type=str, help='Include courses that you want selected. Example: --courses CACE CNR CVA. Defaults to all courses')
    parser.add_argument('--status', nargs='+', choices=ENROLLMENT_STATUSES, default=ENROLLMENT_STATUSES, type=str.capitalize, help='Indicate which enrollment statuses you wish to filter for. Example: --status Active Completed. Defaults to any status.')
    parser.add_argument('--filter', action='store_true', help='Display a window to customize the default filters.')

    args = parser.parse_args()

    # ==============================
    # ANALYTICS CONFIG DEFINITION
    # ==============================
    enrollments_settings = dict(DEFAULT_ENROLMENTS_SETTINGS_OBJECT)
    users_settings = dict(DEFAULT_USERS_SETTINGS_OBJECT)

    if args.filter is True:
        enrollments_settings, users_settings = display_filter_config_gui(
            enrollments_settings,
            users_settings
        )

    # ===============
    # WEB DRIVER INIT
    # ===============
    browser = os.environ.get("BROWSER")
    driver = initialize_selenium_driver(browser)

    # ===============
    # DATA EXTRACTION
    # ===============
    login_to_canvas_catalog(driver)

    # divide the program list to account for the 20 filter limit
    total_runs = (len(args.program) // 20) + 1

    df_all_enrollments = pd.DataFrame()

    # ENROLLMENTS
    # ++++++++++++
    if total_runs > 1:
        print(bcolors.OKCYAN + f"INFO: More than 20 programs selected. The script will perform data extraction and distribution in {total_runs} batches." + bcolors.ENDC)
    
    print("\n")
    print("-" * 30)
    print("RETRIEVING ENROLLMENTS DATA")
    print("-" * 30)

    for i in range(total_runs):
        programs = args.program[20*i:20*(i+1)]

        print(bcolors.OKCYAN + f"INFO: Extracting data for programs: {', '.join(programs)}" + bcolors.ENDC)

        # @@@ SET FILTERS @@@@
        enrollments_settings = set_enrollment_status_filters(enrollments_settings, args.status)
        enrollments_settings = set_enrollments_program_filters(enrollments_settings, programs)

        set_record_filters_via_session_storage(driver, enrollments_settings, users_settings)

        # @@@ NAVIGATE TO ANALYTICS > ENROLLMENTS @@@
        driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")

        wait = WebDriverWait(driver, 10)

        try:
            wait.until(EC.visibility_of_element_located(
                (By.CSS_SELECTOR, "[data-sortable-table]")
            ))
        except TimeoutException:
            print(bcolors.WARNING + "WARNING: No enrollment data found." + bcolors.ENDC)
            continue

        # @@@ EXTRACT ENROLLMENTS DATA @@@
        df_all_enrollments = pd.concat([df_all_enrollments, extract_table_data_to_df(driver)])

    if len(df_all_enrollments) > 0:
        df_all_enrollments = df_all_enrollments.replace("No Data—", pd.NA)
        df_all_enrollments = df_all_enrollments.fillna("")

        df_all_enrollments.to_excel("all-enrollments.xlsx", index=False)
    else:
        print(bcolors.WARNING + "WARNING: No enrollment data to save." + bcolors.ENDC)

    print("\n")
    print("-" * 30)
    print("RETRIEVING USER DATA")
    print("-" * 30)

    # USERS
    # +++++++
    df_all_users = pd.DataFrame()
    
    driver.get("https://courses.cpe.ubc.ca/analytics/users")

    wait = WebDriverWait(driver, 10)

    try:
        wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "[data-sortable-table]")
        ))

        df_all_users = extract_table_data_to_df(driver)

        df_all_users.to_excel("all-users.xlsx", index=False)
    except TimeoutException:
        print(bcolors.WARNING + "WARNING: No user data found." + bcolors.ENDC)


    print("\n")
    print("-" * 30)
    print("DISTRIBUTING DATA")
    print("-" * 30)

    # TODO: Write the actual code lmao

    