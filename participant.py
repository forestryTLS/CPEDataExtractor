import os

import pandas as pd
from dotenv import load_dotenv


class Participant:

    # NOTE: KEEP THIS ClaimsSheetProgramName UP TO DATE

    ClaimsSheetProgramName = {
        "CACE": "Climate Action and Community Engagement (UBCV)",
        "CBBD": "Circular Bioeconomy Development Micro-Certificate",
        "CSRP": "Communication Strategies for Resource Practitioners Micro-Certificate",
        "CNR": "Co-Management of Natural Resources (UBCV)",
        "CVA": "Climate Vulnerability and Adaptation Micro-certificate (UBCV)",
        "FCM": "Forest Carbon Management Micro-certificate (UBCV)",
        "FHM": "Forest Health Management Micro-certificate (UBCV)",
        "LCACF": "Life Cycle Assessment of Clean Fuels Micro-Certificate",
        "SMS": "Strategic Management for Sustainability (UBCV)",
        "EFO": "Environmental Footprints of Organizations Micro-Certificate",
        "HTC": "Hybrid Timber Construction Micro-Certificate",
        "FSTB": "Fire Safety for Timber Buildings (UBCV)",
        "ZCBS": "Zero Carbon Building Solutions Micro-Certificate",
        "FMP": "Forest Management Planning Micro-Certificate",
        "TWS": "Tall Wood Structures Micro-certificate (UBCV)",
        "LLFM": "Landscape Level Forest Modeling Micro-Certificate"

    }

    def __init__(self, row, course, claim_data):
        self.course = course
        self.name = row['Name']
        self.email = row["Email"].lower()
        self.earliestFSGYear = -1
        self.earliestReturningYear = -1
        self.claimAmount = -1
        self.revenue = row["revenue"]
        self.discount = row["discount"]
        self.courses_enrolled = row['product_name_0']
        self.promo_codes = row['promo_codes']
        self.bulk_seats = row['bulk_seats']
        self.FSG_CLAIM_DATA = claim_data
        load_dotenv()

    def set_earliest_FSG_year(self, session):
        self.earliestFSGYear = session

    def set_earliest_returning_year(self, session):
        self.earliestReturningYear = session

    def set_claim_amount(self, amount):
        self.claimAmount = amount

    def get_email(self):
        return self.email

    def get_earliest_FSG_year(self):
        return self.earliestFSGYear

    def get_earliest_returning_year(self):
        return self.earliestReturningYear

    def get_revenue(self):
        return self.revenue

    def get_discount(self):
        return self.discount

    def get_claim_amount(self):
        return self.claimAmount

    def get_courses_participant(self):
        return self.courses_enrolled

    def get_bulk_seats(self):
        return self.bulk_seats

    def find_ealier_session(self, assignedSession, sheetName):
        if (assignedSession == -1):
            return sheetName
        curr_assigned_year = assignedSession.split()[0]
        curr_assigned_sess = assignedSession.split()[1]
        possible_year = sheetName.split()[0]
        possible_sess = sheetName.split()[1]

        if (int(curr_assigned_year) > int(possible_year)):
            return sheetName
        if (int(possible_year) > int(curr_assigned_year)):
            return assignedSession
        if (curr_assigned_year == possible_year):
            if (curr_assigned_sess == "Fall" and possible_sess == "Spring"):
                return sheetName
            else:
                return assignedSession

    def find_earlier_session_for_year(self, sheetName, year):
        sheet_year = sheetName.split()[0]
        sheet_sess = sheetName.split()[1]

        if sheet_year == str(year) and sheet_sess == "Fall":
            return sheetName
        elif sheet_year == str(int(year) + 1) and sheet_sess == "Spring" and self.earliestFSGYear != (str(year) + " Fall"):
            return sheetName
        else:
            return self.earliestFSGYear


    def received_FSG(self, sheet):
        claims_data = self.FSG_CLAIM_DATA
        for year in claims_data:
            if claims_data[year].loc[
                ((claims_data[year]['EMAIL'] == self.email) | (claims_data[year]['EMAIL'] == self.email.lower())) &
                (claims_data[year]['PROGRAM NAME'] == Participant.ClaimsSheetProgramName[self.course]) &
                (claims_data[year]['STATUS'] == "Claimed"), 'STATUS'].any():
                earliestFSGYear = self.find_earlier_session_for_year(sheet, year)
                self.set_earliest_FSG_year(earliestFSGYear)
                claimAmount = claims_data[year].loc[
                    ((claims_data[year]['EMAIL'] == self.email) | (claims_data[year]['EMAIL'] == self.email.lower())) &
                    (claims_data[year]['STATUS'] == 'Claimed') &
                    (claims_data[year]['PROGRAM NAME'] == Participant.ClaimsSheetProgramName[self.course]),
                    'TOTAL CLAIM AMOUNT'].iloc[0]
                self.set_claim_amount(claimAmount)

    def is_FSG_recipient(self, sheet_name, sheet_data):

        if 'Received FSG?' in sheet_data.columns:
            if sheet_data.loc[sheet_data['Email Address'] == self.email, 'Received FSG?'].iloc[0] == 'yes' or (
                    sheet_data.loc[sheet_data['Email Address'] == self.email, 'Received FSG?'].iloc[0] == 'Yes'):
                earliestFSGYear = self.find_ealier_session(self.earliestFSGYear, sheet_name)
                self.set_earliest_FSG_year(earliestFSGYear)
            else:
                earliestReturningYear = self.find_ealier_session(self.earliestReturningYear, sheet_name)
                self.set_earliest_returning_year(earliestReturningYear)
        elif "FSG?" in sheet_data.columns:
            if sheet_data.loc[sheet_data['Email Address'] == self.email, 'FSG?'].iloc[0] == 'yes' or (
                    sheet_data.loc[sheet_data['Email Address'] == self.email, 'FSG?'].iloc[0] == 'Yes'):
                earliestFSGYear = self.find_ealier_session(self.earliestFSGYear, sheet_name)
                self.set_earliest_FSG_year(earliestFSGYear)
            else:
                earliestReturningYear = self.find_ealier_session(self.earliestReturningYear, sheet_name)
                self.set_earliest_returning_year(earliestReturningYear)
        else:
            earliestReturningYear = self.find_ealier_session(self.earliestReturningYear, sheet_name)
            self.set_earliest_returning_year(earliestReturningYear)

    def process_participant(self, program_excel_data):
        for sheet in program_excel_data:
            if len(sheet.split()) == 2:
                sheet_data = program_excel_data[sheet]
                sheet_data['Email Address'] = sheet_data['Email Address'].str.strip()
                sheet_data['Email Address'] = sheet_data['Email Address'].str.lower()
                if self.email in sheet_data['Email Address'].values:
                    self.received_FSG(sheet)
                    if self.earliestFSGYear == -1:
                        self.is_FSG_recipient(sheet, sheet_data)

        print("Email ", self.email)
        print("Name ", self.name)
        print("Revenue ", self.revenue)
        print("earliestFSGYear ", self.earliestFSGYear)
        print("earliestReturningYear ", self.earliestReturningYear)
        print("discount ", self.discount)
        print("Claim Amount ", self.claimAmount)
        print("Courses: ", self.courses_enrolled)



if __name__ == "__main__":
    participant = Participant({
        "discount": 0,
        "Email": "Example@ubc.ca",
        "Name": "Example Name",
        "revenue": 0,
        "discount": 2400,
        "promo_codes": ["Example Code"],
        "bulk_seats": 0,
        "product_name_0": ["Example",
                           "Example",
                           "Example",
                           "Example",
                           "Example"]

    }, "CSRP")
    participant.received_FSG("Example")
