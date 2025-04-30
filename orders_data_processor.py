import math
import os

import pandas as pd
import numpy as np

from participant import Participant
from dotenv import load_dotenv
import numbers


class ordersDataProcessor:

    # NOTE: KEEP THIS EXCEL_SHEETS, ClaimsSheetProgramName AND NUMBER_OF_COURSES_IN_PROGRAM UP TO DATE
    EXCEL_SHEETS = {
        "CBBD": "Circular Bioeconomy Business Development - Registrations.xlsx",
        "CSRP": "Communication Strategies for Resource Practitioners - Registrations.xlsx",
        "FCM": "Forest Carbon Management - Registrations.xlsx",
        "CACE": "Climate Action and Community Engagement - Registrations.xlsx",
        "FHM": "Forest Health Management - Registrations.xlsx",
        "CNR": "Co-Management of Natural Resources - Registrations.xlsx",
        "SMS": "Strategic Management for Sustainability - Registrations.xlsx",
        "LCACF": "Life Cycle Assessment of Clean Fuels - Registrations.xlsx",
        "CVA": "Climate Vulnerability and Adaptation - Registrations.xlsx",
        "EFO": "Environmental Footprints of Organizations - Registrations.xlsx",
        "HTC": "Hybrid Timber Construction - Registrations.xlsx",
        "FSTB": "Fire Safety for Timber Buildings - Registrations.xlsx",
        "ZCBS": "Zero Carbon Building Solutions - Registrations.xlsx",
        "FMP": "Forest Management Planning - Registrations.xlsx",
        "TWS": "Tall Wood Structures - Registrations.xlsx",
        "LLFM": "Landscape Level Forest Modeling - Registrations.xlsx",
        "EBSC": "Engineered Bamboo for Sustainable Construction - Registrations.xlsx"
    }

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

    FULL_PROGRAM_COST = 2400

    def __init__(self, session, course):
        self.session = session
        self.course = course
        self.course_excel_data = {}
        self.summary_enrolment = {"Program": course, "FSG": 0, "Non FSG": 0, "Returning Non FSG": 0, "Free Slots": 0,
                                  "FSG & Paid": 0}
        self.finance_table = {"Program": course, "Full FSG": 0,
                              "Partial FSG & Paid FSG Claim Amount": 0,
                              "Partial FSG & Paid Amount to Forestry": 0,
                              "Non FSG Full Program (including bank transfer)": 0,
                              "Non FSG Individual": 0, "Dropped Admin Amount": 0,
                              "Full Program Discount Amount": 0,
                              "Bulk Registration in Catalog": 0, "Total Amount": 0,
                              "Total Claim Amount": 0}
        load_dotenv()
        #NOTE: Keep this self.FSG_CLAIM_DATA UP TO DATE
        self.FSG_CLAIM_DATA = {
            "2023": pd.read_excel(os.environ.get("CLAIMS_2023")),
            "2024": pd.read_excel(os.environ.get("CLAIMS_2024"))
        }

    def is_session_sheet(self, sheetName, session):
        """ Checks if the sheet name follows the pattern '[year] Fall' or '[year] Spring' """

        session_components = session.split()
        sheet_components = sheetName.split()
        if len(sheet_components) == 2 and (sheet_components[1] == "Spring" or sheet_components[1] == "Fall"):
            if int(session_components[0]) > int(sheet_components[0]) and int(sheet_components[0]) >= 2023:
                enrolment_category = sheet_components[0] + " returning FSG"
                self.summary_enrolment[enrolment_category] = 0
                return True
            elif int(session_components[0]) == int(sheet_components[0]) and session_components[1] == "Fall"\
                    and int(sheet_components[0]) >= 2023:
                enrolment_category = sheet_components[0] + " returning FSG"
                self.summary_enrolment[enrolment_category] = 0
                return True
            elif int(session_components[0]) == int(sheet_components[0]) and (
                    session_components[1] == sheet_components[1]):
                return True
        else:
            return False

    def is_sheet_valid(self, sheetName, session):
        """Checks if the sheet name follows the pattern '[year] Bank Transfer' or '[year] All Refunds'."""
        session_components = session.split()
        sheet_components = sheetName.split()

        if len(sheet_components) == 3 and sheet_components[0] == session_components[0] and (
                sheet_components[1].lower() == "payment" and sheet_components[2].lower() == "transactions"):
            return True
        elif len(sheet_components) == 3 and sheet_components[0] == session_components[0] and (
                sheet_components[1].lower() == "all" and sheet_components[2].lower() == "refunds"):
            return True
        else:
            return False

    def process_refund_sheet_data(self):
        """Preprocesses the 'Email' and 'Case' columns of the all refund dataframe to ensure:
            - All email addresses are converted to lowercase.
            - Any leading or trailing spaces around the email addresses are removed.
            - The 'Case' column is also processed to be in lowercase."""

        refund_sheet_name = self.session.split()[0] + " All Refunds"
        if refund_sheet_name in self.course_excel_data:
            # Removing space in front of emails
            self.course_excel_data[refund_sheet_name]['Email'] = self.course_excel_data[refund_sheet_name][
                'Email'].str.strip()
            self.course_excel_data[refund_sheet_name]['Email'] = self.course_excel_data[refund_sheet_name][
                'Email'].str.lower()
            self.course_excel_data[refund_sheet_name]['Case'] = self.course_excel_data[refund_sheet_name][
                'Case'].str.lower()

    def process_bank_sheet_data(self):
        """Preprocesses the 'Email', 'Registration Type', 'Is Bulk?' and 'Is Active?' columns of the payment
            transactions dataframe to ensure:
            - All email addresses are converted to lowercase.
            - Any leading or trailing spaces around the email addresses are removed.
            - The 'Is Active?', 'Registration Type', 'Is Bulk?' column is also processed to be in lowercase.
            - Any leading or trailing spaces around the values of 'Is Active?' and 'Is Bulk?' column is removed"""

        bank_pay_sheet_name = self.session.split()[0] + " Payment Transactions"
        if bank_pay_sheet_name in self.course_excel_data:
            bank_sheet_data = self.course_excel_data[bank_pay_sheet_name]
            # Removing space in front of emails
            self.course_excel_data[bank_pay_sheet_name]['Email'] = self.course_excel_data[bank_pay_sheet_name][
                'Email'].str.strip()
            self.course_excel_data[bank_pay_sheet_name]['Email'] = self.course_excel_data[bank_pay_sheet_name][
                'Email'].str.lower()
            self.course_excel_data[bank_pay_sheet_name]['Is Active?'] = self.course_excel_data[bank_pay_sheet_name][
                'Is Active?'].str.lower()
            self.course_excel_data[bank_pay_sheet_name]['Is Active?'] = self.course_excel_data[bank_pay_sheet_name][
                'Is Active?'].str.strip()
            self.course_excel_data[bank_pay_sheet_name]['Registration Type'] = \
                self.course_excel_data[bank_pay_sheet_name]['Registration Type'].str.lower()
            self.course_excel_data[bank_pay_sheet_name]['Is Bulk?'] = self.course_excel_data[bank_pay_sheet_name][
                'Is Bulk?'].str.lower()
            self.course_excel_data[bank_pay_sheet_name]['Is Bulk?'] = self.course_excel_data[bank_pay_sheet_name][
                'Is Bulk?'].str.strip()

    def get_program_sheet_data(self):
        """Reads the program excel sheet and stores the relevant sheet data """

        print(os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH"))
        path_excel_file = os.environ.get("ORDERS_REGISTRATIONS_FOLDER_PATH") + ordersDataProcessor.EXCEL_SHEETS[
            self.course]
        excel_file_info = pd.ExcelFile(path_excel_file)

        # get data of sheets in program excel file
        for sheet in excel_file_info.sheet_names:
            if self.is_session_sheet(sheet, self.session):
                self.course_excel_data[sheet] = excel_file_info.parse(sheet, header=1)
            elif self.is_sheet_valid(sheet, self.session):
                self.course_excel_data[sheet] = excel_file_info.parse(sheet, header=0)
                self.process_refund_sheet_data()
                self.process_bank_sheet_data()

    def combine_rows(self, df):
        """ Combines rows of the dataframe that have the same email and aggregates the data """

        combined_df = (
            df.groupby('Email', as_index=False)
                .agg({
                'Name': "first",
                'promo_codes': lambda x: list(filter(pd.notna, x)) if any(pd.notna(x)) else None,
                'product_name_0': lambda x: list(filter(pd.notna, x)) if any(pd.notna(x)) else None,
                # Collect promotions into a list or None
                'discount': lambda x: x.dropna().sum() if x.notna().any() else None,  # Sum discounts or None
                'revenue': 'sum',  # Always sum revenue
                'bulk_seats': 'sum'
            })
        )
        return combined_df

    def read_orders_table(self, path):
        """ Takes path of enrolment excel file as argument, reads pre-processes the data and
                processes record for each participant"""
        df = pd.read_excel(path)
        df.replace('—', np.nan, inplace=True)

        # Coverting discount and revenue to numeric
        df['discount'] = df['discount'].replace({'\$': '', ',': ''}, regex=True)
        df['discount'] = df['discount'].apply(lambda x: pd.to_numeric(x, errors='coerce') if pd.notna(x) else x)
        df['revenue'] = df['revenue'].replace({'\$': '', ',': ''}, regex=True)
        df['revenue'] = df['revenue'].apply(lambda x: pd.to_numeric(x, errors='coerce') if pd.notna(x) else x)

        combined_df = self.combine_rows(df)

        print(combined_df)

        # Get program excel file from teams
        self.get_program_sheet_data()

        self.process_participants(combined_df)

    def count_individual_course_enrolment(self, courses_enrolled):
        """ Returns the number of distinct individual courses in the courses_enrolled list """

        # print("---- Printing from count_individual_course_enrolment -----")
        # print("courses_enrolled", courses_enrolled)
        distinct_courses = list(set(courses_enrolled))
        # print("distinct_courses", distinct_courses)
        if len(distinct_courses) == ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course] or (
                len(distinct_courses) > ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course]
        ):
            return ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course]
        else:
            return len(distinct_courses)

    def process_full_FSG(self, participant):
        """ Updates the financial and enrolment table for Full FSG category either in current or previous session"""

        if participant.get_earliest_FSG_year() == self.session:
            if participant.get_claim_amount() == -1 or math.isnan(participant.get_claim_amount()):
                if self.check_all_refund_sheet_for_case(participant, 'fsg'):
                    refund_sheet_name = self.session.split()[0] + " All Refunds"
                    refund_sheet_data = self.course_excel_data[refund_sheet_name]
                    claim_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                         (refund_sheet_data[
                                                              'Case'] == 'fsg'), 'Claim Amount'].iloc[0]
                    amount_paid = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                         (refund_sheet_data[
                                                              'Case'] == 'fsg'), 'Amount Paid (FSG)'].iloc[0]
                    print("Claim Amount: ", claim_amount)
                    print("Amount Paid: ", amount_paid)
                    if int(claim_amount) == ordersDataProcessor.FULL_PROGRAM_COST or amount_paid == 0 or \
                            math.isnan(amount_paid):
                        self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                        self.finance_table["Full FSG"] = self.finance_table["Full FSG"] + 1
                        self.finance_table["Total Claim Amount"] = self.finance_table[
                                                                       "Total Claim Amount"] + claim_amount

                    else:
                        self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                        self.finance_table["Partial FSG & Paid Amount to Forestry"] = self.finance_table[
                                                                                          "Partial FSG & Paid Amount to Forestry"] + \
                                                                                      amount_paid
                        self.finance_table["Partial FSG & Paid FSG Claim Amount"] = self.finance_table[
                                                                                        "Partial FSG & Paid FSG Claim Amount"] + \
                                                                                    claim_amount
            else:
                print("Added to FSG")
                self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                self.finance_table["Full FSG"] = self.finance_table["Full FSG"] + 1
                self.finance_table["Total Claim Amount"] = self.finance_table["Total Claim Amount"] \
                                                           + participant.get_claim_amount()
        else:
            enrolment_category = participant.get_earliest_FSG_year().split()[0] + " returning FSG"
            print("Added to ", enrolment_category)
            self.summary_enrolment[enrolment_category] = self.summary_enrolment[enrolment_category] + 1

    def check_all_refund_sheet_for_case(self, participant, case):
        """ Checks if participant received refund fr given case """

        refund_sheet_name = self.session.split()[0] + " All Refunds"
        if refund_sheet_name in self.course_excel_data:
            refund_sheet_data = self.course_excel_data[refund_sheet_name]
            if participant.get_email() in refund_sheet_data['Email'].values and (
                    refund_sheet_data.loc[refund_sheet_data['Email'] == participant.get_email(), 'Case'].iloc[
                        0] == case):
                return True
        return False

    def process_partial_FSG_N_paid(self, participant):
        """ Updates the financial and enrolment table for participants who made payment on catalog
            and also claimed FSG"""

        if participant.get_revenue() == participant.get_claim_amount() or (participant.get_claim_amount() ==
                                                                           ordersDataProcessor.FULL_PROGRAM_COST):
            print("refunded for FSG")
            if participant.get_earliest_FSG_year() == self.session:
                self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                self.finance_table["Full FSG"] = self.finance_table["Full FSG"] + 1
                self.finance_table["Total Claim Amount"] = self.finance_table["Total Claim Amount"] \
                                                           + participant.get_claim_amount()
            else:
                enrolment_category = participant.get_earliest_FSG_year().split()[0] + " returning FSG"
                print("Added to ", enrolment_category)
                self.summary_enrolment[enrolment_category] = self.summary_enrolment[enrolment_category] + 1
        elif participant.get_earliest_FSG_year() == self.session:
            print("Added to FSG and Paid")
            if self.check_all_refund_sheet_for_case(participant, 'fsg'):
                refund_sheet_name = self.session.split()[0] + " All Refunds"
                refund_sheet_data = self.course_excel_data[refund_sheet_name]
                refund_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                      (refund_sheet_data[
                                                           'Case'] == 'fsg'), 'Balance Refunded'].iloc[0]
                claim_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                      (refund_sheet_data[
                                                           'Case'] == 'fsg'), 'Claim Amount'].iloc[0]
                print("Refund Amount: ", refund_amount)
                if int(refund_amount) == ordersDataProcessor.FULL_PROGRAM_COST or (
                        int(refund_amount) >= participant.get_revenue()):
                    self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                    self.finance_table["Full FSG"] = self.finance_table["Full FSG"] + 1
                    self.finance_table["Total Claim Amount"] = self.finance_table["Total Claim Amount"]+ claim_amount

                else:
                    self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                    self.finance_table["Partial FSG & Paid Amount to Forestry"] = self.finance_table[
                                                                          "Partial FSG & Paid Amount to Forestry"] + \
                                                                      (participant.get_revenue() - int(
                                                                          refund_amount))
                    self.finance_table["Partial FSG & Paid FSG Claim Amount"] = self.finance_table["Partial FSG & Paid FSG Claim Amount"] +\
                                                                                claim_amount
            else:
                self.summary_enrolment["FSG"] = self.summary_enrolment["FSG"] + 1
                if participant.get_claim_amount() != 0 and participant.get_claim_amount() != -1 and (
                        isinstance(participant.get_claim_amount(), numbers.Number)):
                    self.finance_table["Partial FSG & Paid Amount to Forestry"] = self.finance_table["Partial FSG & Paid Amount to Forestry"] + \
                                                                      (
                                                                                  participant.get_revenue() - participant.get_claim_amount())
                    self.finance_table["Partial FSG & Paid FSG Claim Amount"] = self.finance_table[
                                                                                    "Partial FSG & Paid FSG Claim Amount"] + \
                                                                                participant.get_claim_amount()
                else:
                    self.finance_table["Partial FSG & Paid Amount to Forestry"] = self.finance_table["Partial FSG & Paid Amount to Forestry"] + \
                                                                      participant.get_revenue()
                    self.finance_table["Partial FSG & Paid FSG Claim Amount"] = self.finance_table[
                                                                                    "Partial FSG & Paid FSG Claim Amount"] + \
                                                                                participant.get_claim_amount()
        else:
            self.summary_enrolment["FSG & Paid"] = self.summary_enrolment["FSG & Paid"] + 1

    def process_free_slot_case(self, participant):
        """ Updates the enrolment table for Free slot category"""

        print("Added to Free Column in Enrolment Table Only")
        self.summary_enrolment["Free Slots"] = self.summary_enrolment["Free Slots"] + 1

    def process_dropped(self, participant):
        refund_sheet_name = self.session.split()[0] + " All Refunds"
        refund_sheet_data = self.course_excel_data[refund_sheet_name]
        refund_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                              (refund_sheet_data['Case'] == 'dropped'),
                                              'Non-Refunded Balance'].iloc[0]
        print("Refund Amount: ", refund_amount)
        print("Added to Dropped")
        self.finance_table["Dropped Admin Amount"] = self.finance_table["Dropped Admin Amount"] + refund_amount

    def process_bank_or_free_case(self, participant):
        """ Updates the financial and enrolment table for participants who either made payment through bank or
            received the free slot"""

        if self.check_all_refund_sheet_for_case(participant, 'dropped'):
            self.process_dropped(participant)
            return
        bank_pay_sheet_name = self.session.split()[0] + " Payment Transactions"
        if bank_pay_sheet_name in self.course_excel_data:
            bank_sheet_data = self.course_excel_data[bank_pay_sheet_name]
            if participant.get_email() in bank_sheet_data['Email'].values and (
                    bank_sheet_data.loc[bank_sheet_data['Email'] == participant.get_email(), 'Is Active?'].iloc[
                        0] == 'active'):
                if (bank_sheet_data.loc[(bank_sheet_data['Email'] == participant.get_email()) &
                                        (bank_sheet_data['Is Active?'] == 'active'),
                                        'Is Bulk?'].iloc[0] == 'yes'):
                    return
                if (bank_sheet_data.loc[(bank_sheet_data['Email'] == participant.get_email()) &
                                        (bank_sheet_data['Is Active?'] == 'active'),
                                        'Registration Type'].iloc[0] == 'full program'):
                    self.summary_enrolment["Non FSG"] = self.summary_enrolment["Non FSG"] + 1
                    self.finance_table["Non FSG Full Program (including bank transfer)"] = \
                        self.finance_table["Non FSG Full Program (including bank transfer)"] + 1
            else:
                # Add to free column
                self.process_free_slot_case(participant)
        else:
            # Add to free column
            self.process_free_slot_case(participant)

    def process_Non_FSG(self, participant):
        """ Updates the financial and enrolment table for participants either dropped the course
            or are Non FSG registrant"""

        if self.check_all_refund_sheet_for_case(participant, 'dropped'):
            if self.check_all_refund_sheet_for_case(participant, 'dropped'):
                self.process_dropped(participant)
        else:
            print("Added to Non FSG")
            self.summary_enrolment["Non FSG"] = self.summary_enrolment["Non FSG"] + 1
            if (participant.get_revenue() >= self.FULL_PROGRAM_COST):
                self.finance_table["Non FSG Full Program (including bank transfer)"] = \
                    self.finance_table["Non FSG Full Program (including bank transfer)"] + 1
            else:
                if self.check_all_refund_sheet_for_case(participant, 'program discount'):
                    refund_sheet_name = self.session.split()[0] + " All Refunds"
                    refund_sheet_data = self.course_excel_data[refund_sheet_name]
                    refund_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                          (refund_sheet_data['Case'] == 'program discount'),
                                                          'Balance Refunded'].iloc[0]
                    print("Refund Amount: ", refund_amount)
                    self.finance_table["Full Program Discount Amount"] = \
                        self.finance_table["Full Program Discount Amount"] + refund_amount
                self.increment_non_FSG_individual_course(participant)

    def process_Non_FSG_returning(self, participant):
        """ Updates the financial and enrolment table for participants who were previously registered as Non FSG either
            in full program or individual course and have made a payment this session"""

        if self.check_all_refund_sheet_for_case(participant, 'already paid'):
            if participant.get_earliest_returning_year() != self.session:
                self.summary_enrolment["Returning Non FSG"] = self.summary_enrolment["Returning Non FSG"] + 1
                return

        print("Added to Non FSG Individual")
        self.summary_enrolment["Non FSG"] = self.summary_enrolment["Non FSG"] + 1
        if self.course == "CVA":
            num_courses = ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course] - 1
        else:
            num_courses = ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course]
        individual_course_amount = (ordersDataProcessor.FULL_PROGRAM_COST / num_courses) + 50
        print("Individual Course Amount: ", individual_course_amount)
        if (participant.get_discount()):
            remainder_discount = participant.get_discount() % individual_course_amount
            print("Remainder Discount: ", remainder_discount)
        else:
            remainder_discount = None
        refund_sheet_name = self.session.split()[0] + " All Refunds"
        if participant.get_revenue() < ordersDataProcessor.FULL_PROGRAM_COST and remainder_discount == (
                (individual_course_amount * num_courses) - ordersDataProcessor.FULL_PROGRAM_COST):
            self.finance_table["Full Program Discount Amount"] = self.finance_table[
                                                                     "Full Program Discount Amount"] + remainder_discount
            self.finance_table["Non FSG Individual"] = self.finance_table["Non FSG Individual"] + int(
                (participant.get_revenue() + remainder_discount) / individual_course_amount)
        elif self.check_all_refund_sheet_for_case(participant, 'program discount'):
            refund_sheet_data = self.course_excel_data[refund_sheet_name]
            refund_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                  (refund_sheet_data['Case'] == 'program discount'),
                                                  'Balance Refunded'].iloc[0]
            print("Refund Amount: ", refund_amount)
            self.finance_table["Full Program Discount Amount"] = \
                self.finance_table["Full Program Discount Amount"] + refund_amount
            self.increment_non_FSG_individual_course(participant)
        else:
            # Individual Course
            self.increment_non_FSG_individual_course(participant)

    def increment_non_FSG_individual_course(self, participant):
        """ Updates the financial table for number of individual Non FSG courses registered by the participant"""

        if self.course == "CVA":
            individual_course_cost = (ordersDataProcessor.FULL_PROGRAM_COST /
                                      (ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course] - 1)) + 50
        else:
            individual_course_cost = (ordersDataProcessor.FULL_PROGRAM_COST /
                                      ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course]) + 50
        print(individual_course_cost)

        if (participant.get_revenue() == (len(participant.get_courses_participant()) * individual_course_cost)):
            num_individual_courses = self.count_individual_course_enrolment(participant.get_courses_participant())
            print("Num Individual Courses ", num_individual_courses)
            self.finance_table["Non FSG Individual"] = self.finance_table["Non FSG Individual"] + num_individual_courses
        elif participant.get_discount() != 0 and ((participant.get_revenue() + participant.get_discount())
                                                  == (
                                                          len(participant.get_courses_participant()) * individual_course_cost)):
            num_individual_courses = self.count_individual_course_enrolment(participant.get_courses_participant())
            print("Num Individual Courses ", num_individual_courses)
            self.finance_table["Non FSG Individual"] = self.finance_table["Non FSG Individual"] + num_individual_courses
        else:
            number_of_courses_paid = int(participant.get_revenue() / individual_course_cost)
            self.finance_table["Non FSG Individual"] = self.finance_table["Non FSG Individual"] + number_of_courses_paid

    def process_bulk_or_dropped_case(self, participant):
        """ Updates the financial and enrolment table for participants who made bulk payments
            or dropped course/program"""

        if (participant.get_bulk_seats() > 0):
            if (participant.get_revenue() == (participant.get_bulk_seats() * ordersDataProcessor.FULL_PROGRAM_COST)):
                print("Added Bulk Seats")
                self.summary_enrolment["Non FSG"] = self.summary_enrolment["Non FSG"] + participant.get_bulk_seats()
                self.finance_table["Bulk Registration in Catalog"] = self.finance_table[
                                                              "Bulk Registration in Catalog"] + participant.get_revenue()
        elif (participant.get_bulk_seats() == 0):
            if self.check_all_refund_sheet_for_case(participant, 'dropped'):
                refund_sheet_name = self.session.split()[0] + " All Refunds"
                refund_sheet_data = self.course_excel_data[refund_sheet_name]
                refund_amount = refund_sheet_data.loc[(refund_sheet_data['Email'] == participant.get_email()) &
                                                      (refund_sheet_data['Case'] == 'dropped'),
                                                      'Non-Refunded Balance'].iloc[0]
                print("Refund Amount: ", refund_amount)
                print("Added to Dropped")
                self.finance_table["Dropped Admin Amount"] = self.finance_table["Dropped Admin Amount"] + refund_amount
        else:
            print("-------Participant not added to either Tables-----")

    def process_participants(self, combined_df):
        """ Updates the enrollment and financial for all the participants registered in the session"""

        for index, row in combined_df.iterrows():

            participant = Participant(row, self.course, self.FSG_CLAIM_DATA)
            participant.process_participant(self.course_excel_data)

            if (participant.get_revenue() == 0 and participant.get_earliest_FSG_year() != -1):
                self.process_full_FSG(participant)
            elif (participant.get_revenue() != 0 and participant.get_earliest_FSG_year() != -1):
                # Add to FSG&Paid column of the session
                self.process_partial_FSG_N_paid(participant)
            elif (participant.get_revenue() == 0 and participant.get_earliest_FSG_year() == -1 and
                  participant.get_earliest_returning_year() != -1):
                # Add to returning NonFSG column
                if participant.get_earliest_returning_year() != self.session:
                    print("Added to Returning Non FSG")
                    self.summary_enrolment["Returning Non FSG"] = self.summary_enrolment["Returning Non FSG"] + 1
                elif participant.get_earliest_returning_year() == self.session:
                    self.process_bank_or_free_case(participant)
            elif (participant.get_revenue() != 0 and participant.get_earliest_FSG_year() == -1 and
                  participant.get_earliest_returning_year() == self.session):
                # Add to Non FSG
                self.process_Non_FSG(participant)
            elif (participant.get_revenue() != 0 and participant.get_earliest_FSG_year() == -1 and
                  participant.get_earliest_returning_year() != -1):
                self.process_Non_FSG_returning(participant)
            elif (participant.get_revenue() != 0 and participant.get_earliest_FSG_year() == -1 and
                  participant.get_earliest_returning_year() == -1):
                self.process_bulk_or_dropped_case(participant)

            print("Enrollment Table")
            print(self.summary_enrolment)
            print("Finance")
            print(self.finance_table)

        self.save_all_data()

        return

    def save_all_data(self):
        """ Processes the enrolment and finance table and saves them"""

        self.calculate_total_enrolments()
        self.process_finance_table()
        self.save_table(self.finance_table, 'financeTable.xlsx')
        self.save_table(self.summary_enrolment, 'enrolmentTable.xlsx')

    def calculate_total_enrolments(self):
        """ Calculate total enrolment for course in the given session and update the enrolment table"""
        self.summary_enrolment["Total Enrolment"] = 0
        for category in self.summary_enrolment:
            if category != "Total Enrolment" and category != "Program":
                self.summary_enrolment["Total Enrolment"] = self.summary_enrolment["Total Enrolment"] + \
                                                            self.summary_enrolment[category]

    def process_finance_table(self):
        """ Calculates the total amount in the finance table and add negative sign in front of Full Program Discount"""
        if ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course] == 3:
            individual_course_cost = 850
        else:
            individual_course_cost = 650

        self.finance_table["Full Program Discount Amount"] = -abs(self.finance_table["Full Program Discount Amount"])
        self.finance_table["Total Amount"] = (self.finance_table[
                                                  "Non FSG Full Program (including bank transfer)"] * ordersDataProcessor.FULL_PROGRAM_COST) + \
                                             (self.finance_table["Non FSG Individual"] * individual_course_cost) + \
                                             self.finance_table["Dropped Admin Amount"] + self.finance_table[
                                                 "Full Program Discount Amount"] + \
                                             (self.finance_table["Bulk Registration in Catalog"]) + \
                                             self.finance_table["Partial FSG & Paid Amount to Forestry"]

    def save_table(self, courseTable, filePath):
        """ Updates course table with enrollment/financial records for the given course"""

        if (ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[self.course] == 3):
            courseTable["Program"] = self.course + " (3 courses)"
        if os.path.isfile(filePath):
            try:
                df = pd.read_excel(filePath)
            except Exception as e:
                print(f"Error reading the Excel file: {e}")
                df = pd.DataFrame(columns=courseTable.keys())
        else:
            df = pd.DataFrame(columns=courseTable.keys())

        if courseTable["Program"] in df["Program"].values:
            df.loc[df["Program"] == courseTable["Program"], courseTable.keys()] = courseTable.values()
            # df = pd.concat([df, pd.DataFrame([courseTable])], ignore_index=True)
        else:
            df = pd.concat([df, pd.DataFrame([courseTable])], ignore_index=True)

        if "Total Enrolment" in df.columns:
            temp_cols = df.columns.tolist()
            index = df.columns.get_loc("Total Enrolment")
            new_cols = temp_cols[0:index] + temp_cols[index + 1:] + temp_cols[index:index + 1]
            df = df[new_cols]

        df.to_excel(filePath, sheet_name = self.session, index=False)

    def calculate_total_enrolment_by_column(self, df):
        for cols in df.columns:
            if cols != "Program":
                total = df[cols].sum()
                df.loc[df["Program"] == "Subtotal", cols] = total

        return df

    def calculate_individual_course_total(self, df):
        total = 0
        for program in df["Program"].values:
            if ordersDataProcessor.NUMBER_OF_COURSES_IN_PROGRAM[program] == 3:
                individual_course_cost = 850
            else:
                individual_course_cost = 650
            total = total + (individual_course_cost*(df.loc[df["Program"] == program, "Non FSG Individual"].iloc[0]))
        return total


if __name__ == "__main__":
    data_processor = ordersDataProcessor("2025 Spring", "LCACF")
    data_processor.read_orders_table('enrollment.xlsx')
