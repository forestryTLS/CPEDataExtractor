from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
import pandas as pd

from utils.logging import bcolors

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

def append_df_data_to_file(
    df: pd.DataFrame,
    data_file: Path
):
    """ Appends the data in the provided DataFrame to the specified file. This function assumes both store information using the same format and columns. """
    if not data_file.is_file() or data_file.suffix != '.xlsx':
        print(bcolors.WARNING + f'Could not append user data. The file {data_file.absolute()} is not a valid Excel (.xlsx) file.' + bcolors.ENDC)
        return False
    
    if len(df) == 0:
        print(bcolors.WARNING + 'No user data to append to existing records.' + bcolors.ENDC)
        return False
    
    # read data from the existing file (assume file only has one sheet containing the data)
    df_existing_users = pd.read_excel(data_file)

    df_existing_users = pd.concat([df, df_existing_users], ignore_index=True)
    df_existing_users.drop_duplicates()

    df_existing_users.to_excel(
        data_file,
        index=False
    )

 