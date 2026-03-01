from datetime import date, timedelta

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