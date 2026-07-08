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
    header_columns = ["Frequency", "Metric Category", "Metrics"]

    # Start a counter for how many cells look like repeated headers.
    matched_headers = 0

    # Check each important header column.
    for column in header_columns:
        # Only check the column if it exists in the file.
        if column in row.index:
            # Compare the cell value with the column name.
            if str(row[column]).strip() == column:
                # Increase the counter when the cell repeats the header name.
                matched_headers = matched_headers + 1

    # If two or more header names repeat in the row, treat it as a header row.
    return matched_headers >= 2


# This function adds risk_type to normal data rows and removes red heading rows.
def add_risk_type_and_remove_heading_rows(data):
    # Start with no risk type until we find the first red heading row.
    current_risk_type = ""

    # This list will store the risk type for each row.
    risk_types = []

    # This list will store True for rows we want to keep.
    rows_to_keep = []

    # Go through each row in the Excel data.
    for _, row in data.iterrows():
        # Check if the row is a repeated header row from another file or section.
        if is_repeated_header_row(row):
            # Add blank risk type for this header row because we will remove it.
            risk_types.append("")

            # Mark this repeated header row to be removed.
            rows_to_keep.append(False)

        # Check if the row is a red risk heading row.
        elif is_risk_type_row(row):
            # Save this heading as the current risk type.
            current_risk_type = str(row[RISK_TYPE_SOURCE_COLUMN]).strip()

            # Add blank risk type for this heading row because we will remove it.
            risk_types.append("")

            # Mark this heading row to be removed.
            rows_to_keep.append(False)

        # If this is a normal data row, keep it.
        else:
            # Add the latest risk type to this normal data row.
            risk_types.append(current_risk_type)

            # Mark this normal data row to be kept.
            rows_to_keep.append(True)

    # Add the risk_type column as the first column in the file data.
    data.insert(0, "risk_type", risk_types)

    # Keep only normal data rows and remove the red heading rows.
    data = data[rows_to_keep]

    # Reset row numbering after removing heading rows.
    data = data.reset_index(drop=True)

    # Return the cleaned data.
    return data


# Get the final period that will be used to search files.
period = get_period()

# Read the entity config CSV file into a pandas table.
config = pd.read_csv(CONFIG_FILE)

# Sort the config by serial number so files are appended in correct order.
config = config.sort_values("serial")

# Create an empty list where each entity file's data will be stored.
all_data = []

# Go through each row in the config file one by one.
for _, row in config.iterrows():
    # Read the entity code from the current config row, for example PINB.
    entity_code = row["entity_code"]

    # Build the expected file name pattern for this entity and period.
    # Example: PINB_Risk_Appetite_Statement_Actual_APR2026*.xlsx
    file_pattern = f"{entity_code}_Risk_Appetite_Statement_Actual_{period}*.xlsx"

    # Search the input folder for files matching the expected file pattern.
    matching_files = list(Path(INPUT_FOLDER).glob(file_pattern))

    # Check if no file was found for this entity.
    if len(matching_files) == 0:
        # Print a message so we know this entity file is missing.
        print(f"No file found for {entity_code}")

        # Skip this entity and move to the next row in the config.
        continue

    # Check if more than one matching file was found for the same entity.
    if len(matching_files) > 1:
        # Print a warning. The script will use the first file from the list.
        print(f"More than one file found for {entity_code}. Using first file.")

    # Pick the first matching file for this entity.
    file_path = matching_files[0]

    # Print which file is being read.
    print(f"Reading {file_path.name}")

    # Read the Excel file into a pandas table.
    # header=HEADER_ROW_NUMBER - 1 is used because Python starts counting from 0.
    # Example: Excel row 3 becomes Python header row 2.
    data = pd.read_excel(file_path, header=HEADER_ROW_NUMBER - 1)

    # Remove any completely blank rows from the file.
    data = data.dropna(how="all")

    # Add risk_type to each row and remove red risk heading rows.
    data = add_risk_type_and_remove_heading_rows(data)

    # Add entity_code as the first column, so we know which entity each row belongs to.
    data.insert(0, "entity_code", entity_code)

    # Add source_file as the second column, so we know which file each row came from.
    data.insert(1, "source_file", file_path.name)

    # Add this entity's data table into the all_data list.
    all_data.append(data)


# Check if no data was collected from any file.
if len(all_data) == 0:
    # Print a message and stop because there is nothing to append.
    print("No files found. Nothing to append.")

# If at least one file was read, create the final appended output.
else:
    # Append all entity tables one below another into one final table.
    final_data = pd.concat(all_data, ignore_index=True)

    # Create the output folder if it does not already exist.
    Path(OUTPUT_FOLDER).mkdir(exist_ok=True)

    # Create the final output file name, for example RAS_Appended_APR2026.xlsx.
    output_file = Path(OUTPUT_FOLDER) / f"RAS_Appended_{period}.xlsx"

    # Write the final appended table into an Excel file.
    final_data.to_excel(output_file, index=False)

    # Print a success message with the output file path.
    print(f"Done. Created {output_file}")
