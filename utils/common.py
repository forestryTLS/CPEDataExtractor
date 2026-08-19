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
    "FCMo",
    "FSD"
]

PROGRAM_TO_EXCEL_MAP = {
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
    "FAS": ("Foundations of Advanced Silviculture - Registrations.xlsx"),
    "CLF": ("Advanced Life Cycle Assessment of Clean Liquid Fuels - Registrations.xlsx"),
    "CGF": ("Advanced Life Cycle Assessment of Clean Gaseous Fuels - Registrations.xlsx"),
    "FCMo": ("Forest Carbon Modeling - Registrations.xlsx"),
    "FSD": ("Forest Stand and Development - Registrations.xlsx"),
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
    COMPLETION_STATUS = "Completion Status (if full certificate)"
    RECEIVED_FSG = "Received FSG?"
    PROMO_CODE = "Promotion Code (if applicable)"
    GRANT_RECEIVED = "Grant Amount Received"
    AMOUNT_PAID = "Amount Paid to Forestry"
    CERTIFICATE_STATUS = "Certificate Status"
    NOTES = "Notes"

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