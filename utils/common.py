from datetime import date, timedelta
from enum import StrEnum

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

ENROLMENTS_DATE_PRESET_KEY = 'enrollment_date_preset'
ENROLMENTS_DATE_FROM_KEY = 'enrollment_date_from'
ENROLMENTS_DATE_TO_KEY = 'enrollment_date_to'

USERS_DATE_PRESET_KEY = 'registration_date_preset'
USERS_DATE_FROM_KEY = 'registration_date_from'
USERS_DATE_TO_KEY = 'registration_date_to'

ENROLLMENT_STATUSES = ['Active', 'Completed', 'Concluded', 'Dropped']

CERTIFICATE_PROGRAMS = [
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

EXCELS = {
    "CBBD": ("Circular Bioeconomy Business Development - Registrations.xlsx"),
    "CACE": ("Climate Action and Community Engagement - Registrations.xlsx"),
    "CVA": ("Climate Vulnerability and Adaptation - Registrations.xlsx"),
    "CNR": ("Co-Management of Natural Resources - Registrations.xlsx"),
    "CSRP": ("Communication Strategies for Resource Practitioners - Registrations.xlsx"),
    "EFO": ("Environmental Footprints of Organizations - Registrations.xlsx"),
    "FSTB": ("Fire Safety for Timber Buildings - Registrations.xlsx"),
    "FCM": ("Forest Carbon Management - Registrations.xlsx"),
    "FHM": ("Forest Health Management - Registrations.xlsx"),
    "HTC": ("Hybrid Timber Construction - Registrations.xlsx"),
    "SMS": ("Strategic Management for Sustainability - Registrations.xlsx"),
    "TWS": ("Tall Wood Structures - Registrations.xlsx"),
    "ZCBS": ("Zero Carbon Building Solutions - Registrations.xlsx"),
    "FMP": ("Forest Management Planning - Registrations.xlsx"),
    "EBSC": ("Engineered Bamboo for Sustainable Construction - Registrations.xlsx"),
    "LCACF": ("Life Cycle Assessment of Clean Fuels - Registrations.xlsx"),
    "LLFM": ("Landscape Level Forest Modeling - Registrations.xlsx"),
    "FAS": ("Foundations of Advanced Silviculture - Registration.xlsx"),
    "CLF": ("Advanced Life Cycle Assessment of Clean Liquid Fuels - Registration.xlsx"),
    "CGF": ("Advanced Life Cycle Assessment of Clean Gaseous Fuels - Registration.xlsx"),
    "FCMo": ("Forest Carbon Modeling - Registrations.xlsx"),
}

HEADER_ROW_IDX = 1

class ReportColumn(StrEnum):
    FULL_NAME = "Full Name"
    EMAIL_ADDRESS = "Email Address"
    ORGANIZATION = "Organization"
    TITLE = "Title"
    USER_ID = "Student Catalog ID"
    LISTING_IDS = "Listing IDs"
    PHONE_NUMBER = "Phone Number"
    HOME_ADDRESS = "Home Address"
    IS_ALUM = "Is Forestry Alum?",
    INDIGENOUS_IDENTITY = "Self-Identify as Indigenous?"
    MAILING_ADDRESS = "Mailing Address"
    DEGREES_EXPERIENCE = "Relevant Degrees or Experience"
    PROGRAM_START = "Program Start Date"
    PROGRAM_EXPIRY = "Program Expiry Date"
    SINGLE_LISTING_ID = "Listing ID"
    CATALOG_NAME = "Catalog"

REPORT_COLUMNS_TO_POPULATE = [
    ReportColumn.FULL_NAME,
    ReportColumn.ORGANIZATION,
    ReportColumn.TITLE,
    ReportColumn.USER_ID,
    ReportColumn.EMAIL_ADDRESS,
    ReportColumn.PHONE_NUMBER,
    ReportColumn.HOME_ADDRESS,
    ReportColumn.IS_ALUM,
    ReportColumn.INDIGENOUS_IDENTITY,
    ReportColumn.MAILING_ADDRESS,
    ReportColumn.DEGREES_EXPERIENCE
]