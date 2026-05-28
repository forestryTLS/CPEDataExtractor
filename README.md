# Canvas Catalog Scraping Script

Script that scrapes data from the Canvas Catalog "Analytics" page and distributes it to various internal records.

## Table of Contents

- [Important Information](#important-information)
  - [Limitations](#limitations)
  - [Rules for Canvas Catalog listing and catalog names](#rules-for-canvas-catalog-listing-and-catalog-names)
  - [Storage and distribution of scraped data](#storage-and-distribution-of-scraped-data)
  - [Adding new micro-certificate programs](#adding-new-micro-certificate-programs)
- [Setup](#setup)

## Important Information

### Limitations

This script scrapes information directly from the Canvas Catalog Analytics webpage. Thus, some manual input is required for the script to fully run. Mainly, manual input is required when:

- Entering login credentials
- Completing MFA
- Applying custom filters

### Storage and distribution of scraped data

The script assumes the data should be distributed to Excel workbook in the folder pointed to by the `REGISTRATION_DATA_FOLDER_PATH` environment variable.

The folder should contain a workbook for each micro-certificate program, and each workbook should contain a sheet for every offering of the program.

### Rules for Canvas Catalog listing and catalog names

This script depends on specific formatting of listing and catalog names on Canvas Catalog to extract data correctly. Below are the most important rules to follow.

#### All programs must have an abbreviation

**e.g.** Co-Management of Natural Resources > <ins>CNR</ins>

#### Listings must always end with the program offering

- **Programs:** CNR - Online Micro-Certificate: Co-Management of Natural Resources <ins>2026 Spring</ins>
- **Courses:** Co-Management <ins>2026 Spring</ins>

#### Program listings must begin with the program abbreviation

- <ins>CNR</ins> - Online Micro-Certificate: Co-Management of Natural Resources 2026 Spring

#### Catalogs must begin with the program abbreviation and a spaced dash character

- <ins>CNR - </ins>Online Micro-Certificate: Co-Management of Natural Resources

**NOTE:** These rules are based on how UBC FES sets up their online programs on Canvas Catalog. The script and the modules it depends on will require significant modification if you set up your programs differently.

#### Course catalogs must be children of program catalogs

For the script to properly filter for course listings, the parent catalog of course catalogs must be set to the corresponding program catalog.

Additionally, course listings should be associated with course catalogs, and program listings with program catalogs.

### Adding new micro-certificate programs

To add a new micro-certificate program, you will need to follow the steps below:

1. Create the corresponding listings, catalogs, etc. on Canvas Catalog keeping the naming rules above in mind
2. Create a new Excel workbook for the program in the folder at `REGISTRATION_DATA_FOLDER_PATH`
3. Add the program abbreviation to `CERTIFICATE_PROGRAMS` in [utils/common.py][common.py]
4. Add an entry mapping the program abbreviation to its Excel workbook name to `PROGRAM_TO_EXCEL_MAP` in [utils/common.py][common.py]
5. If the program introduces any new custom fields, add an entry to `CATALOG_COL_ID_NAME_MAP` in [utils/extract.py](./utils/extract.py)


## Setup

1\. Clone the repo on your local machine

```bash
git clone https://github.com/UBCForestryTLS/catalog-data-extractor.git
```

2\. Change into the repo directory, create a virtual environment and activate it
```bash
cd catalog-data-extractor/
python -m venv venv/
source venv/bin/activate
# [Windows] source venv/Scripts/activate 
```

3\. Install the project dependendcies
```bash
pip install -r requirements.txt
```

4\. Create a `.env` file by copying `.env.example` and replace the placeholder values accordingly
```bash
cp .env.example .env
```

| Environment Variable | Description | Possible Values |
| -------- | ----------- | --------------- |
| `REGISTRATION_DATA_FOLDER_PATH` | The path to the folder containing each micro-certificate program's enrollment Excel worksheets. | |
| `BROWSER` | The web browser to run the script on. If an invalid browser is specified, the script will default to using the Selenium Chrome driver. | `Edge`, `Firefox`, `Chromium`, `Chrome` |

5\. Run the script. With no options specified, the script will attempt to scrape data for all programs in the `CERTIFICATE_PROGRAMS` list (see [utils/common.py][common.py])
```bash
python main.py
```

For a complete list of options, run `python main.py -h`.

[common.py]: ./utils/common.py