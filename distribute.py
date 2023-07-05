from openpyxl import load_workbook
import pandas as pd

REGISTRATIONS_FOLDER = "Registrations/"
CURRENT_SHEETS = ['2023 FALL']
# Sheets are {program code: (File name, Sheet name)} No more sheet name, use current sheets
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
}

"""
Steps to distribute data for each enrollment entry
1. Iterate through new enrollment data (must be a pandas dataframe)
Enrollment data must contain the following columns:
student_name_0 -> Full name, student_name_1 -> Email address, product_name_0 (contains program code and session)
Then using their email address retrieve the following data from user_data (take the last match, potentially duplicate emails if data changed):
custom_fields_organization -> Organization
custom_fields_title -> Title
custom_fields_phone-number -> Phone Number
custom_fields_mailing-address -> Mailing Address (Only check for CNR)
custom_fields_indigenous-self-declaration -> Self-Identify as Indigenous? (Only check for CNR)

Also pull the following data from processed_data by searching for student's email in processed_data
Grant amount to give -> Received FSG?, Grant Amount

2. Use account_name + product_name_0 to find the correct sheet to use (excel name and sheet session)
3. Check sheet if user (based on email) already exists save the row index if yes 
4. Format data into a dictionary dynamically based on all the sheet headers (on row 2). Columns not specified will be either None or existing data
5. Append the data to the end of the sheet, or write to the exiting row

"""
def extract_user_data(row, user_data_row, user_grant_row):
    email = row['student_name_1'].split(' ')[2]

    data = {
        'Full Name': row['student_name_0'],
        'Email Address': email,
    }

    if not user_data_row.empty:
        extra = {
            'Organization': user_data_row['custom_fields_organization'].values[0],
            'Title': user_data_row['custom_fields_title'].values[0],
            'Phone Number': user_data_row['custom_fields_phone-number'].values[0],
            'Mailing Address': user_data_row['custom_fields_mailing-address'].values[0],
            'Self-Identify as Indigenous?': 'Yes' if user_data_row['custom_fields_indigenous-self-declaration'].values[0].lower().strip() == '1' else 'No',
        }
        data.update(extra)

    if not user_grant_row.empty:
        extra = {
            'Received FSG?': 'Yes',
            'Grant Amount Received': user_grant_row['Grant amount to give'].values[0]
        }
        data.update(extra)

    return data

def find_sheet(row):
    course_code = row['account_name'].split(" ")[0]
    course_info = row['product_name_0'].split(" ")
    course_session = " ".join(course_info[-2:])

    print("COURES CODE IS", course_code)

    excel_path = REGISTRATIONS_FOLDER + EXCELS[course_code]
    workbook = load_workbook(filename=excel_path)
    return (workbook[course_session], excel_path, workbook)

def search_email_in_sheet(sheet, email):
    """ Return the row index where email is found, -1 if not found """
    header_row = 2 # Headers for excel must be on row 2
    email_column = None

    for cell in sheet[header_row]:
        if cell.value == 'Email Address':
            email_column = cell.column_letter
            break

    if email_column is None:
        return -1

    for row_idx, cell in enumerate(sheet[email_column], start=1):
        if cell.value and cell.value.lower().strip() == email:
            return row_idx

    return -1

def find_empty_row(sheet):
    """ Starting form row 3 of the sheet, find a row where the first column is empty"""
    for row_index, row in enumerate(sheet.iter_rows(min_row=3), start=3):
        if row[0].value is None:
            return row_index
    return -1

def insert_or_append_row(sheet, data, existing_row):
    row = existing_row if existing_row != -1 else None
    header_row = 2
    if row is None:
        row = find_empty_row(sheet)
        
    for col_header, value in data.items():
        col_index = None
        for cell in sheet[header_row]: # Find the column where the current header is
            if cell.value == col_header:
                col_index = cell.column
                break

        if col_index is not None:
            target_cell = sheet.cell(row=row, column=col_index)
            if target_cell.value is None:
                target_cell.value = value
            
def distribute_enrollment_data(df_enrollment, path_to_user_data, path_to_grant_data):
    df_user_data = pd.read_excel(path_to_user_data)
    df_grant_data = pd.read_excel(path_to_grant_data)

    for _, row in df_enrollment.iterrows():
        # Search for the row in user_data based on student_name_1 and email inside student_name_1
        user_data_row = df_user_data[df_user_data['student_name_1'].str.lower().str.strip() == row['student_name_1'].lower().strip()].tail(1)
        user_email = row['student_name_1'].split(' ')[2].lower().strip() 
        user_grant_row = df_grant_data[df_grant_data['Email'].str.lower().str.strip() == user_email].tail(1)
        data = extract_user_data(row, user_data_row, user_grant_row)
        print("UESR EMAIL IS", user_email, data)
        # 2: find the correct sheet to use
        (sheet, excel_path, workbook) = find_sheet(row)

        # 3: Check if email already in sheet
        existing_row = search_email_in_sheet(sheet, user_email)

        # 4: Format row into correct format to match sheet
        # row_values = [data.get(column_name, '') for column_name in sheet.columns]
        # print("ROW VALUES ARE", row_values)
        # 5: Insert data at the end or write to the existing row
        insert_or_append_row(sheet, data, existing_row)
        workbook.save(excel_path)

            
if __name__ == '__main__':
    df = pd.read_excel("Raw_data/enrollment.xlsx")
    distribute_enrollment_data(df, 'Raw_data/user_data.xlsx', 'Raw_data/processed_data.xlsx')
    
workbooks = []
sheets = []

# https://chat.openai.com/share/3fd5f762-1473-4db1-86b8-9d79b57de137
# # Create a dictionary representing the new row of data
# Create a dictionary representing the new row of data
# new_row = {'Full name': 'Bob', 'Phone Number': 'Joe'}

# # Get the column names from the first row
# column_names = [cell.value for cell in sheet[1]]

# # Create a list representing the new row with values in the appropriate order
# ordered_row = [new_row.get(column_name, None) for column_name in column_names]

# # Flag to check if row already exists
# row_exists = False

# # Iterate over existing rows and compare values
# for row in sheet.iter_rows(min_row=2, values_only=True):
#     if all(cell_value == ordered_value or cell_value is None for cell_value, ordered_value in zip(row, ordered_row)):
#         # Row already exists, set flag and break the loop
#         row_exists = True
#         break

# # Append the new row if it doesn't already exist
# if not row_exists:
#     sheet.append(ordered_row)

# # Save the updated workbook
# workbook.save('your_excel_file.xlsx')

# for sheet, value in excels.items():
#     name = root_folder + value[0]
#     print("LOADING", name)
#     workbook = load_workbook(filename=name)
#     workbooks.append(workbook)

#     # print(workbook.sheetnames)
#     temp = workbook[value[1]]
#     sheets.append(temp)
#     print(list(temp.values))
