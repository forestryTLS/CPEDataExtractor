import argparse
import os
from pathlib import Path
import logging

# silence lower level log records (which the two libraries below emit A LOT of)
logging.getLogger("selenium").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.CRITICAL)

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

import pandas as pd 
from dotenv import load_dotenv

from utils.common import (
    ReportColumn,
    PROGRAM_TO_EXCEL_MAP,
    HEADER_ROW_IDX,
    REPORT_COLUMNS_TO_POPULATE,
    CERTIFICATE_PROGRAMS
)
from utils import distribute

from utils.extract import (
    initialize_selenium_driver, 
    login_to_canvas_catalog,
    set_catalog_filters_via_ui,
    extract_table_data_to_df,
    click_apply_filters,
    open_filters_modal
)

from utils.logging import CustomStreamHandler

load_dotenv()

ENROLLMENT_STATUS_KEY = "Enrollment Status"

BASE_DIR = Path(__file__).parent
    
logger = logging.getLogger(__name__)

formatter = logging.Formatter(
    fmt='[%(asctime)s] %(levelname)s: (%(module)s) %(message)s',
    datefmt='%Y-%m-%d %I:%M:%S %p'
)

# +++++++++++++
# FILE HANDLER
# +++++++++++++
file_handler = logging.FileHandler(
    filename=BASE_DIR / "catalog-data-extractor.log",
    encoding='utf-8',
    mode='a'
)

file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

# +++++++++++++++
# STREAM HANDLER
# +++++++++++++++
stream_handler = CustomStreamHandler()

stream_handler.setLevel(logging.DEBUG)
stream_handler.setFormatter(formatter)

# don't send log records of CRITICAL level to stdout (code already raises exceptions)
stream_handler.addFilter(lambda log_record: log_record.levelno != logging.CRITICAL)

# configure the root logger (`stream` and `filename` not needed as we're defining handlers)
logging.basicConfig(
    level=logging.DEBUG,
    handlers=[file_handler, stream_handler]
)

# ===========================
# REGISTRATION FOLDER PARSING
# ===========================
try:
    registration_folder_path = Path(os.environ.get("REGISTRATION_DATA_FOLDER_PATH"))
except TypeError:
    logger.critical("The environment variable REREGISTRATION_DATA_FOLDER_PATH is not set to a valid string path.")
    raise Exception("The required environment variable REGISTRATION_DATA_FOLDER_PATH must be set to a valid string path.")

if not registration_folder_path.is_dir():
    logger.critical(f"The environment variable REREGISTRATION_DATA_FOLDER_PATH does not point to a valid directory: \"{registration_folder_path.absolute()}\".")
    raise Exception(f"The path \"{registration_folder_path.absolute()}\" does not point to a valid directory.")

raw_data_dir = registration_folder_path / "0RawData"

if not raw_data_dir.is_dir():
    logger.warning(f"Could not find an existing raw data directory. Creating directory at \"{raw_data_dir.absolute()}\".")
    raw_data_dir.mkdir(exist_ok=True)

enrollments_file_path = raw_data_dir / "enrollment.xlsx"
users_file_path = raw_data_dir / "user_data.xlsx"

if not enrollments_file_path.is_file():
    logger.warning(f"No enrollments Excel file found. Creating file at \"{enrollments_file_path.absolute()}\".")
    pd.DataFrame().to_excel(enrollments_file_path) # quick way to create a valid Excel file

if not users_file_path.is_file():
    logger.warning(f"No users Excel file found. Creating file at \"{users_file_path.absolute()}\".")
    pd.DataFrame().to_excel(users_file_path)

# =====================
# CLI ARGUMENT PARSING
# =====================
parser = argparse.ArgumentParser(description='This Script uses Selenium to login to Canvas Catalog and extracts enrollments + users')

# Optional command line arguments
parser.add_argument('--mfe', action='store_true', help='Manually Filter Enrollments. Include this argument if you want the script to pause when filtering enrollments')
parser.add_argument('--mfu', action='store_true', help='Manually Filter Users. Include this argument if you want the script to pause when filtering users')
parser.add_argument('--program', nargs='+', choices=CERTIFICATE_PROGRAMS, default=CERTIFICATE_PROGRAMS, type=str, help='Include programs that you want selected. Example: --courses CACE CNR CVA. Defaults to all programs')

args = parser.parse_args()

logger.info("====== STARTING CATALOG DATA EXTRACTION ======")

optional_arguments = []

if args.mfe:
    optional_arguments.append("--mfe")

if args.mfu:
    optional_arguments.append("--mfu")

if len(optional_arguments) > 0:
    logger.info(f"Running script with optional arguments: {' '.join(optional_arguments)}.")

logger.info(f"Extracting data for {len(args.program)} program(s): {', '.join(args.program)}.")

# ===============
# WEB DRIVER INIT
# ===============
browser = os.environ.get("BROWSER")

with initialize_selenium_driver(browser) as driver:
    # ===============
    # DATA EXTRACTION
    # ===============
    logger.info("Requesting manual input from user for Canvas Catalog login.")

    login_to_canvas_catalog(
        driver,
        next_url="https://courses.cpe.ubc.ca/analytics/users"
    )

    # ++++++
    # USERS
    # ++++++
    if args.mfu:
        open_filters_modal(driver)

        logger.info("Waiting for user to manually define Catalog user record filters via UI.")

        input("Please apply any additional filters (e.g. custom date range) and press enter in the terminal to continue...")
        
        click_apply_filters(driver)

    logger.info("Scraping data from Catalog Analytics \"Users\" table.")

    df_all_users = pd.DataFrame()

    wait = WebDriverWait(driver, 10)

    try:
        wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "[data-sortable-table]")
        ))

        df_all_users = extract_table_data_to_df(driver)
        df_all_users = df_all_users.replace(r"(?i)No Data.*", pd.NA, regex=True)
        df_all_users = df_all_users.fillna("")

        logger.info(f"Found {df_all_users.shape[0]} user records. Saving data to \"{users_file_path.absolute()}\".")

        distribute.append_df_data_to_excel(
            df_all_users,
            users_file_path
        )
    except TimeoutException:
        logger.warning("No data found in \"Users\" Catalog Analytics tab. No data will be appended to the users Excel file.")

    # +++++++++++
    # ENROLLMENTS
    # +++++++++++

    logger.info("Scraping data from Catalog Analytics \"Enrollments\" table.")

    driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")

    # divide the program list to account for the 20 filter limit
    total_runs = (len(args.program) // 20) + 1

    df_all_enrollments = pd.DataFrame()

    if total_runs > 1:
        logger.info(f"More than 20 programs selected. Script will scrape data in {total_runs} batches.")

    for i in range(total_runs):
        programs = args.program[20*i:20*(i+1)]

        logger.info(f"Extracting data for batch {i+1} programs: {', '.join(programs)}.")

        # @@@ NAVIGATE TO ANALYTICS > ENROLLMENTS @@@
        driver.get("https://courses.cpe.ubc.ca/analytics/enrollments")

        wait = WebDriverWait(driver, 10)

        try:
            wait.until(EC.visibility_of_element_located(
                (By.CSS_SELECTOR, "[data-sortable-table]")
            ))
        except TimeoutException:
            logger.warning("No erollment data found.")
            continue
        
        set_catalog_filters_via_ui(driver, programs)

        if args.mfe:
            logger.info("Waiting for user to manually define Catalog enrollment record filters via UI.")
            input("Please apply any additional filters (e.g. custom date range) and press enter in the terminal to continue...")

        click_apply_filters(driver)

        # @@@ EXTRACT ENROLLMENTS DATA @@@
        df_all_enrollments = pd.concat(
            [df_all_enrollments, extract_table_data_to_df(driver)],
            ignore_index=True
        )

    # +++++++++++++++++++++++++++++++++++
    # DISTRIBUTE DATA TO INTERNAL RECORDS
    # +++++++++++++++++++++++++++++++++++
    if df_all_enrollments.shape[0] > 0:
        df_all_enrollments = df_all_enrollments.replace(r"(?i)No Data.*", pd.NA, regex=True)
        df_all_enrollments = df_all_enrollments.fillna("")

        # account for cases where student drops, then re-registers
        df_all_enrollments = df_all_enrollments.drop_duplicates(
            [str(ReportColumn.USER_ID), "Listing ID"],
            keep='first' # Catalog always renders records from most to least recent
        )

        logger.info(f"Saving {df_all_enrollments.shape[0]} enrollment records to \"{enrollments_file_path.absolute()}\".")

        distribute.append_df_data_to_excel(
            df_all_enrollments,
            enrollments_file_path
        )
    else:
        logger.warning("No data found in \"Enrollments\" Catalog Analytics tab. No data will be appended to the enrollment Excel file.")

    logger.info("Distributing data to corresponding Excel files.")

    if ENROLLMENT_STATUS_KEY not in df_all_enrollments.columns:
        logger.critical("Column \"Enrollment Status\" could not be extracted from enrollment data. Cannot proceed with distribution.")
        raise Exception("\"Enrollment Status\" not extracted. Cannot proceed with distribution.")
    
    unique_programs = df_all_enrollments["Program"].unique()
    unique_sessions = df_all_enrollments["Session"].unique()

    for program in unique_programs:
        excel_path = registration_folder_path / PROGRAM_TO_EXCEL_MAP[program]

        if not excel_path.is_file():
            logger.error(f"The file \"{excel_path.absolute()}\" does not exist. Skipping distribution of data for program {program}.")
            continue

        for session in unique_sessions:
            logger.info(f"Distributing enrollments for {session} {program}.")

            df_program_single_session = df_all_enrollments[
                (df_all_enrollments["Program"] == program)
                & (df_all_enrollments["Session"] == session)
            ].copy()

            df_program_single_session: pd.DataFrame
            if not all(
                common_column in df_program_single_session.columns
                for common_column in [ReportColumn.FULL_NAME, ReportColumn.EMAIL_ADDRESS, ReportColumn.ORGANIZATION]
            ):
                logger.warning(f"Could not find required columns for {program} ({session}). Skipping.")
                continue

            # ===========================================================================
            # Attempt to fill empty columns in enrollments table using users table data
            # ===========================================================================
            df_program_single_session = df_program_single_session.apply(
                lambda series: distribute.fill_null_columns_from_first(
                    series,
                    df_all_users,
                    match_key=str(ReportColumn.USER_ID)
                ),
                axis=1,
                result_type='broadcast'
            )

            # =============================================
            # Separate dropped from non-dropped enrollments
            # =============================================
            df_new_dropped = df_program_single_session[
                df_program_single_session[ENROLLMENT_STATUS_KEY].astype(str, errors='ignore').str.lower().str.strip() == "dropped"
            ]

            df_new_valid = df_program_single_session[
                df_program_single_session[ENROLLMENT_STATUS_KEY].astype(str, errors='ignore').str.lower().str.strip() != "dropped"
            ]

            logger.info(f"New records to append - VALID: {df_new_valid.shape[0]}, DROPPED: {df_new_dropped.shape[0]}.")

            # ================================================
            # Read and combine existing dropped data (if any)
            # ================================================
            session_year, session_season = str(session).split(" ")
            shortened_session_name = f"{session_year}{session_season[0]}"

            dropped_sheet = f"Dropped ({shortened_session_name})"

            if df_new_dropped.shape[0] > 0:
                ordered_columns = df_new_dropped.columns.to_list()

                target_columns = {
                    str(ReportColumn.FULL_NAME),
                    str(ReportColumn.EMAIL_ADDRESS),
                    str(ReportColumn.ORGANIZATION),
                    str(ReportColumn.TITLE),
                    str(ReportColumn.USER_ID),
                    str(ReportColumn.SINGLE_LISTING_ID),
                    str(ReportColumn.CATALOG_NAME)
                }

                # filter for the target columns while ensuring order is kept
                for column in ordered_columns:
                    if column in target_columns:
                        ordered_columns.append(column)

                #df_new_dropped = df_new_dropped[list(filter_for_columns)]
                df_new_dropped = df_new_dropped[filtered_columns]

                distribute.append_df_data_to_excel(
                    df_new_dropped,
                    excel_path,
                    dropped_sheet,
                    header_row_idx=HEADER_ROW_IDX,
                    drop_duplicates_indices=[
                        str(ReportColumn.FULL_NAME),
                        str(ReportColumn.EMAIL_ADDRESS),
                        str(ReportColumn.SINGLE_LISTING_ID)
                    ]
                )

            # =============================================
            # Combine new valid enrollments with existing
            # =============================================
            try:
                df_existing_enrollments = pd.read_excel(
                    excel_path,
                    sheet_name=session,
                    dtype={
                        str(ReportColumn.USER_ID): str,
                        str(ReportColumn.LISTING_IDS): str,
                        str(ReportColumn.PHONE_NUMBER): str,
                        str(ReportColumn.PROGRAM_START): object,
                        str(ReportColumn.PROGRAM_EXPIRY): object
                    },
                    keep_default_na=False,
                    header=HEADER_ROW_IDX
                )

                if df_all_enrollments.shape[0] == 0 \
                or not all(
                    common_column in df_existing_enrollments.columns
                    for common_column in [ReportColumn.FULL_NAME, ReportColumn.EMAIL_ADDRESS, ReportColumn.ORGANIZATION]
                ):
                    logger.warning(f"Sheet {session} was found in \"{excel_path.absolute()}\", however, it is either empty or doesn't contain the required columns. Its data will be fully replaced.")
                    df_existing_enrollments = None
            except ValueError:
                df_existing_enrollments = None

            if df_existing_enrollments is not None:

                df_existing_enrollments: pd.DataFrame

                # CHECK: Ensure "Student Catalog ID" and "Listing IDs" are valid columns in existing data
                if ReportColumn.USER_ID not in df_existing_enrollments.columns:
                    logger.warning(f"Column \"{ReportColumn.USER_ID}\" not found in sheet \"{session}\" ({program}). A blank column will be inserted.")

                    df_existing_enrollments.insert(
                        len(df_existing_enrollments.columns),
                        column=str(ReportColumn.USER_ID),
                        value=""
                    )

                if ReportColumn.LISTING_IDS not in df_existing_enrollments.columns:
                    logger.warning(f"Column \"{ReportColumn.LISTING_IDS}\" not found in sheet \"{session}\" ({program}). A blank column will be inserted.")
                    
                    df_existing_enrollments.insert(
                        len(df_existing_enrollments.columns),
                        column=str(ReportColumn.LISTING_IDS),
                        value=""
                    )

                added_rows = []

                # Merge new enrollments into existing data
                for _, new_data in df_new_valid.iterrows():
                    df_found_record = df_existing_enrollments[
                        df_existing_enrollments[str(ReportColumn.USER_ID)] == new_data[str(ReportColumn.USER_ID)]
                    ]

                    if df_found_record.shape[0] == 0:
                        df_found_record = df_existing_enrollments[
                            df_existing_enrollments[str(ReportColumn.EMAIL_ADDRESS)].str.strip().str.lower()
                            == str(new_data[str(ReportColumn.EMAIL_ADDRESS)]).strip().lower()
                        ]

                    if df_found_record.shape[0] > 0:
                        # if matching record exists, update it
                        s_found_record = df_found_record.iloc[0, :].copy()

                        record_listing_ids = (
                            str(s_found_record.at[str(ReportColumn.LISTING_IDS)]).strip().split(";")
                            if s_found_record.at[str(ReportColumn.LISTING_IDS)]
                            else []
                        )

                        new_listing_id = str(int(new_data["Listing ID"]))

                        if new_listing_id not in record_listing_ids:
                            record_listing_ids.append(new_listing_id)

                        s_found_record.at[str(ReportColumn.LISTING_IDS)] = ";".join(record_listing_ids)

                        # Check that:
                        # 1. The column exists in the corresponding registration sheet table
                        # 2. The column exists in the incoming data. If not, set to empty string
                        for column in REPORT_COLUMNS_TO_POPULATE:
                            if column in df_existing_enrollments.columns:
                                if column in new_data.index:
                                    s_found_record.at[column] = new_data.at[column]
                                else:
                                    s_found_record.at[column] = ""
                        
                        df_existing_enrollments.iloc[s_found_record.name] = s_found_record
                    else:
                        new_row = new_data.copy()

                        for column in new_row.index:
                            if column not in df_existing_enrollments.columns:
                                new_row = new_row.drop(column)

                        # add the "Listing IDs" column data (not present in Catalog enrollment data)
                        new_row.at[str(ReportColumn.LISTING_IDS)] = new_data[str(ReportColumn.SINGLE_LISTING_ID)]

                        df_existing_enrollments.loc[df_existing_enrollments.shape[0]] = new_row
                
                # Update records if corresponding dropped enrollments exist
                for _, drop_data in df_new_dropped.iterrows():
                    df_found_record = df_existing_enrollments[
                        df_existing_enrollments[str(ReportColumn.USER_ID)] == drop_data[str(ReportColumn.USER_ID)]
                    ]

                    if df_found_record.shape[0] == 0:
                        df_found_record = df_existing_enrollments[
                            df_existing_enrollments[str(ReportColumn.EMAIL_ADDRESS)].str.strip().str.lower() 
                            == str(drop_data[str(ReportColumn.EMAIL_ADDRESS)]).strip().lower()
                        ]
                    
                    if df_found_record.shape[0] > 0:
                        s_found_record: pd.Series = df_found_record.iloc[0, :]

                        record_listing_ids = str(s_found_record.at[str(ReportColumn.LISTING_IDS)]).strip().split(";")

                        try:
                            drop_listing_id = str(int(drop_data["Listing ID"])).strip()
                        except ValueError:
                            drop_listing_id = drop_data["Listing ID"]

                        # remove the dropped entry's listing ID from the list, and if the existing record's listing IDs are empty
                        # remove the row (all courses/programs dropped)
                        if drop_listing_id in record_listing_ids:
                            logger.debug(f"Dropped record found for {session} {program}. Student: {drop_data['Full Name']}, Listing ID: {drop_listing_id}.")

                            record_listing_ids.pop(
                                record_listing_ids.index(drop_listing_id)
                            )

                            if len(record_listing_ids) == 0:
                                logger.debug(f"Student {drop_data['Full Name']} dropped all courses in {session} {program}. Their enrollment record will be removed.")
                                df_existing_enrollments = df_existing_enrollments.drop(s_found_record.name)
            else:
                # ensure data for the current program-session combination exists
                if df_new_valid.shape[0] > 0:
                    logger.warning(f"Could not find sheet \"{session}\" for program {program} in \"{excel_path.absolute()}\". A new sheet will be created.")

                    # use enrollment data directly to create new table
                    df_existing_enrollments = df_new_valid

                    
                    unique_student_catalog_ids = list(df_existing_enrollments[str(ReportColumn.USER_ID)].unique())
                    student_id_to_listing_ids_map = {}

                    df_existing_enrollments.insert(len(df_existing_enrollments.columns), str(ReportColumn.LISTING_IDS), "")

                    # group each student's listing IDs into a list mapped to their student ID
                    for student_id in unique_student_catalog_ids:
                        current_student_lisitng_ids = list(df_existing_enrollments[
                            df_existing_enrollments[str(ReportColumn.USER_ID)] == student_id
                        ][str(ReportColumn.SINGLE_LISTING_ID)])

                        student_id_to_listing_ids_map[student_id] = current_student_lisitng_ids

                    def populate_listing_ids_column(series: pd.Series):
                        current_user_id = series.at[str(ReportColumn.USER_ID)]

                        if current_user_id in student_id_to_listing_ids_map:
                            series.at[str(ReportColumn.LISTING_IDS)] = ";".join(student_id_to_listing_ids_map[current_user_id])

                        return series
                    
                    df_existing_enrollments = df_existing_enrollments.apply(
                       populate_listing_ids_column,
                       axis=1,
                       result_type='broadcast'
                    )

                    ordered_columns = df_existing_enrollments.columns.to_list()

                    program_lowered = str(program).lower().strip()

                    # keep only necessary columns
                    target_columns = {
                        str(ReportColumn.FULL_NAME),
                        str(ReportColumn.ORGANIZATION),
                        str(ReportColumn.TITLE),
                        str(ReportColumn.EMAIL_ADDRESS),
                        str(ReportColumn.PHONE_NUMBER),
                        str(ReportColumn.MAILING_ADDRESS) if program_lowered == 'cnr' else str(ReportColumn.HOME_ADDRESS),
                        str(ReportColumn.IS_ALUM),
                        str(ReportColumn.USER_ID),
                        str(ReportColumn.LISTING_IDS)
                    }

                    # conditionally include program-specific columns
                    if program_lowered == 'cnr':
                        target_columns.update((
                            str(ReportColumn.INDIGENOUS_IDENTITY)
                        ))
                    elif program_lowered == 'fmp':
                        target_columns.update((
                            str(ReportColumn.DEGREES_EXPERIENCE)
                        ))

                    filtered_columns = []

                    for column in ordered_columns:
                        if column in target_columns:
                            filtered_columns.append(column)

                    df_existing_enrollments = df_existing_enrollments[filtered_columns]
                    df_existing_enrollments = df_existing_enrollments.drop_duplicates()

                    # insert manually populated columns with empty cells
                    for col in [
                        ReportColumn.PROGRAM_START,
                        ReportColumn.PROGRAM_EXPIRY,
                        ReportColumn.COMPLETION_STATUS,
                        ReportColumn.RECEIVED_FSG,
                        ReportColumn.PROMO_CODE,
                        ReportColumn.GRANT_RECEIVED,
                        ReportColumn.AMOUNT_PAID,
                        ReportColumn.CERTIFICATE_STATUS,
                        ReportColumn.NOTES
                    ]:
                        df_existing_enrollments.insert(len(df_existing_enrollments.columns), col, "")
                    
            
            logger.info(f"Saving enrollment data for {session} {program} to file.")

            distribute.append_df_data_to_excel(
                df_existing_enrollments,
                excel_path,
                sheet_name=session,
                header_row_idx=HEADER_ROW_IDX,
                merge_into_existing=False
            )