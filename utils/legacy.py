# NOTE: The functions in this file no longer affect Catalog filtering due to vendor changes.

from typing import TypedDict
import json
from datetime import date, timedelta

from utils.common import CERTIFICATE_PROGRAMS, ENROLLMENT_STATUSES
from utils.extract import (
    FULL_PROGRAM_CATALOG_NAMES,
    CATALOG_PROGRAM_IDS,
    SeleniumDriver
)

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

ENROLLMENT_STATUSES = ['Active', 'Completed', 'Concluded', 'Dropped']

ENROLMENTS_DATE_PRESET_KEY = 'enrollment_date_preset'
ENROLMENTS_DATE_FROM_KEY = 'enrollment_date_from'
ENROLMENTS_DATE_TO_KEY = 'enrollment_date_to'

USERS_DATE_PRESET_KEY = 'registration_date_preset'
USERS_DATE_FROM_KEY = 'registration_date_from'
USERS_DATE_TO_KEY = 'registration_date_to'

DEFAULT_CREATION_DATE_FROM = (date.today() - timedelta(days=7)).strftime("%Y-%m-%d")
DEFAULT_CREATION_DATE_TO = date.today().strftime("%Y-%m-%d")

DEFAULT_ENROLMENTS_SETTINGS_OBJECT = {
    "filter": {
        "account_ids": [],
        "product_ids": [],
        "product_statuses": [],
        "student_ids": [],
        ENROLMENTS_DATE_PRESET_KEY: "past_week",
        ENROLMENTS_DATE_FROM_KEY: DEFAULT_CREATION_DATE_FROM,
        ENROLMENTS_DATE_TO_KEY: DEFAULT_CREATION_DATE_TO,
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
        USERS_DATE_PRESET_KEY: "past_week",
        USERS_DATE_FROM_KEY: DEFAULT_CREATION_DATE_FROM,
        USERS_DATE_TO_KEY: DEFAULT_CREATION_DATE_TO
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

# Configuration to quickly set filters for analytics search
SESSION_STORAGE_ENROLMENTS_SETTINGS_KEY = "analytics-settings-enrollments"
SESSION_STORAGE_USERS_SETTINGS_KEY = "analytics-settings-users"

class AccountFilter(TypedDict):
    id: int
    name: str

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
        full_course_catalog_name = FULL_PROGRAM_CATALOG_NAMES[course_abbrev]
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

# from utils.common import (
#     ENROLMENTS_DATE_PRESET_KEY,
#     ENROLMENTS_DATE_FROM_KEY,
#     ENROLMENTS_DATE_TO_KEY,

#     USERS_DATE_PRESET_KEY,
#     USERS_DATE_FROM_KEY,
#     USERS_DATE_TO_KEY
# )

# def display_filter_config_gui(
#     enrollments_settings: dict,
#     users_settings: dict
# ):
#     import tkinter as tk
#     from tkinter import ttk
#     from tkcalendar import DateEntry

#     enrollments_settings['filter'][ENROLMENTS_DATE_PRESET_KEY] = 'custom'
#     users_settings['filter'][USERS_DATE_PRESET_KEY] = 'custom'

#     enrolments_initial_from = enrollments_settings['filter']['enrollment_date_from']
#     enrolments_initial_to = enrollments_settings['filter']['enrollment_date_to']

#     # initialize the app and create a frame to hold widgets
#     root = tk.Tk()
#     frm = ttk.Frame(root, padding=10)
#     frm.grid()

#     # define labels and date widgets to filter enrolments and users by
#     ttk.Label(frm, text='FILTER RECORDS').grid(row=0, column=0, columnspan=5)
#     ttk.Label(frm, text="From:").grid(row=1, column=0)

#     sv_date_from = tk.StringVar()
#     sv_date_to = tk.StringVar()

#     start_date_entry = DateEntry(
#         frm, 
#         date_pattern='yyyy-MM-dd',
#         textvariable=sv_date_from
#     )
    
#     start_date_entry.grid(row=1, column=1)

#     sv_date_from.set(enrolments_initial_from)

#     ttk.Label(frm, text='-').grid(row=1, column=2)

#     ttk.Label(frm, text="To:").grid(row=1, column=3)

#     end_date_entry = DateEntry(
#         frm,
#         date_pattern='yyyy-MM-dd',
#         textvariable=sv_date_to
#     )
    
#     end_date_entry.grid(row=1, column=4)

#     sv_date_to.set(enrolments_initial_to)

#     def set_enrolments_date_from(sv, index, mode):
#         from_date = sv_date_from.get()

#         enrollments_settings['filter'][ENROLMENTS_DATE_FROM_KEY] = from_date 
#         users_settings['filter'][USERS_DATE_FROM_KEY] = from_date

#     def set_enrolments_date_to(sv, index, mode):
#         to_date = sv_date_to.get()

#         enrollments_settings['filter'][ENROLMENTS_DATE_TO_KEY] = to_date
#         users_settings['filter'][USERS_DATE_TO_KEY] = to_date

#     sv_date_from.trace_add('write', set_enrolments_date_from)
#     sv_date_to.trace_add('write', set_enrolments_date_to)

#     ttk.Button(frm, text="Save", command=root.destroy).grid(row=2, column=0, columnspan=5)
#     root.mainloop()

#     return enrollments_settings, users_settings