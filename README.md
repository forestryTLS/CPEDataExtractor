# Setup
These instructions below are using Unix commands.
1. Install with ```git clone https://github.com/WillKang01/CPEDataExtractor.git```.
2. Go to the new folder ```cd CPEDataExtractor``` then install a virtual environment such as with ```python3 -m venv env``` and ```source env/bin/activate```
3. Install the requirements with ```pip install -r requirements.txt```
4. Copy .env.example into .env with ```cp .env.example .env```
5. Update .env with the correct paths
6. Run the program with ```python get_data.py```

# Notes
- You must keep the following constants updated in the code: inside get_data.py: ```VALID_COURSES, FULL_OPTION_NAME```. Inside distribute.py ```EXCELS```
- You can pass in the following arguments into get_data.py: ```--mfe, --mfu, --courses```. Example: ```python get_data.py --mfe --mfu --courses CVA CNR```.
That command will pause at the filtering stage for enrollments and users so you can customize it. It also only searches for the courses CVA and CNR. Use ```python get_data.py --help```
for more information.
- Each excel sheet must have the right sheet names such as 2023 Fall. If a user registers for a program that doesn't have a sheet created for it yet, the program will fail.

# Data Created
- Inside 0RawData, it creates a running list of all the enrollments, and users so far. Duplicate rows are avoided by checking if every column entry is the same.
- Inside 0EnrollmentHistory, all the data that was distributed to the various sheets is saved as a history.
- Users are identified by their email. Emails are used to cross check the user_data.xlsx sheet and processed_data.xlsx. If the emails do not match, they are not considered the same user
and a new row will be created in the sheet. Else the program will write data to empty columns in the existing row.

# Debugging Tips
- Since Canvas Catlog's page is entirely dynamic, you may run into issues when trying to inspect the page and the element disappears. To get around this you can use this command in the inspect terminal ```setTimeout(function(){debugger;}, 5000)``` which will pause the screen after 5 seconds.

  




