# Setup
## Unix
1. Install with ```git clone https://github.com/WillKang01/CPEDataExtractor.git```.
2. Go to the new folder ```cd CPEDataExtractor``` then install a virtual environment such as with ```python3 -m venv env``` and ```source env/bin/activate```
3. Install the requirements with ```pip install -r requirements.txt```
4. Copy .env.example into .env with ```cp .env.example .env```
5. Update .env with the correct paths
6. Run the program with ```python get_data.py```

## Windows
1. Install with `git clone https://github.com/WillKang01/CPEDataExtractor.git`.
   - Open the Windows Terminal or Command Prompt.
   - Install Git if it is not already installed.
   - Run the command `git clone https://github.com/WillKang01/CPEDataExtractor.git` to clone the repository.
2. Go to the new folder `cd CPEDataExtractor` then install a virtual environment such as with `python3 -m venv env` and `source env/bin/activate`.
   - Change the directory to the `CPEDataExtractor` folder using the command `cd CPEDataExtractor`.
   - Create a virtual environment using the command `python -m venv env`.
   - Activate the virtual environment using the command `.\env\Scripts\activate`.
3. Install the requirements with `pip install -r requirements.txt`.
   - Make sure you have Python and pip installed.
   - Run the command `pip install -r requirements.txt` to install the required packages.
4. Copy `.env.example` into `.env` with `cp .env.example .env`.
   - In Windows, you can use the `copy` command instead of `cp`.
   - Run the command `copy .env.example .env` to copy the file.
5. Update `.env` with the correct paths.
   - Open the `.env` file using a text editor and update the paths as required.
6. Run the program with `python get_data.py`.
   - Run the command `python get_data.py` to execute the program.

# Notes
- Set the browser variable to a browser that you don't use as your default. Such as if you use Chrome as your default, set browser to Edge.
- The program should open a new browser everytime forcing you to log in if everything is set up correctly.
If that doesn't work, try adding the following (with the correct browser):
```
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument("--remote-debugging-port=9222")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()),
    options=options
)
```
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
