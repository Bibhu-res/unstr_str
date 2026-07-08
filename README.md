# Import Python's date/time library and call it dt for short.
import datetime as dt

# Import Path so we can work with folder and file paths easily.
from pathlib import Path

# Import pandas, which helps us read Excel/CSV files and append data.
import pandas as pd


# This is the config file where entity serial number and entity code are kept.
CONFIG_FILE = "config/entities.csv"

# This is the folder where all input Excel files will be placed.
INPUT_FOLDER = "input"

# This is the folder where the final appended Excel file will be created.
OUTPUT_FOLDER = "output"

# This tells Python which Excel row contains the real column headers.
# In your screenshot, the real header row is row 3.
HEADER_ROW_NUMBER = 3

# This is the column where red risk section names are appearing.
# Example red section names: Treasury Risk, People Risk, Financial Crime Compliance Risk.
RISK_TYPE_SOURCE_COLUMN = "Metrics"


# Keep this as None if you want the script to use the current system month.
# Example: if today's month is July 2026, the script will use JUL2026.
PERIOD_OVERRIDE = None

# If you want to force a specific month for testing, use this instead:
# PERIOD_OVERRIDE = "APR2026"


# This function decides which month/year period the script should use.
def get_period():
    # If PERIOD_OVERRIDE has a value like "APR2026", use that value.
    if PERIOD_OVERRIDE:
        # Return the manually entered period.
        return PERIOD_OVERRIDE

    # Get today's date from the computer system.
    today = dt.date.today()

    # Convert today's month into a 3-letter month name like JAN, FEB, MAR.
    month_name = today.strftime("%b").upper()

    # Join month name and year together, for example JUL2026.
    return f"{month_name}{today.year}"


# This function checks whether a cell is blank.
def is_blank(value):
    # pandas uses NaN for blank Excel cells, so pd.isna checks those blanks.
    if pd.isna(value):
        # Return True because this value is blank.
        return True

    # Convert the value to text, remove spaces, and check if nothing is left.
    return str(value).strip() == ""


# This function checks whether a row is a red risk section row.
def is_risk_type_row(row):
    # If the Metrics column is missing, we cannot find risk type rows.
    if RISK_TYPE_SOURCE_COLUMN not in row.index:
        # Return False because there is no Metrics column to check.
        return False

    # Read the text from the Metrics column.
    metric_text = row[RISK_TYPE_SOURCE_COLUMN]

    # If the Metrics cell is blank, this is not a risk type row.
    if is_blank(metric_text):
        # Return False because blank text cannot be a risk type.
        return False

    # Convert the Metrics text to normal text and remove extra spaces.
    metric_text = str(metric_text).strip()

    # Risk type rows are headings like "Treasury Risk" or "People Risk".
    if not metric_text.lower().endswith("risk"):
        # Return False because this row does not look like a risk heading.
        return False

    # On red section rows, Frequency is normally blank.
    frequency_is_blank = "Frequency" not in row.index or is_blank(row["Frequency"])

    # On red section rows, Metric Category is normally blank.
    category_is_blank = "Metric Category" not in row.index or is_blank(row["Metric Category"])

    # Return True only when it looks like a section heading, not a real metric row.
    return frequency_is_blank and category_is_blank


# This function checks whether a row is a repeated header row.
def is_repeated_header_row(row):
    # These are the main column names we expect in the input file.
