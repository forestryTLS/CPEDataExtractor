# NOTE: The functions in this file no longer affect Catalog filtering due to vendor changes.

from typing import TypedDict
import json

from utils.common import CERTIFICATE_PROGRAMS, ENROLLMENT_STATUSES
from utils.extract import (
    FULL_PROGRAM_CATALOG_NAMES,
    CATALOG_PROGRAM_IDS,
    SeleniumDriver
)

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