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
# In your screenshot, the yellow header row is row 2.
HEADER_ROW_NUMBER = 2


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
    # Example: Excel row 2 becomes Python header row 1.
    data = pd.read_excel(file_path, header=HEADER_ROW_NUMBER - 1)

    # Remove any completely blank rows from the file.
    data = data.dropna(how="all")

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
