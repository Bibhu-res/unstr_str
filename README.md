""
DATA TYPE CLASSIFIER FOR MESSY SPREADSHEET VALUES
===================================================
Goal: look at ONE cell value at a time (as it comes out of Excel) and decide
which "bucket" it belongs to:
 
    1. rag         -> Red / Green / Amber status  (R, G, A, Red, Green, Amber)
    2. percentage  -> A number that represents a percentage   (e.g. "12.05%")
    3. numeric     -> A pure number, even if units/arrows are stuck to it
                       (e.g. "2654.70m", "-414m", "USD (-222,811)", "↑ 200k")
    4. text        -> Anything else that doesn't cleanly fit the above
                       (labels, mixed sentences, notes like "Nil", "0 VH, 0H")
 
We ALSO return a "cleaned_value" - the value after stripping the noise
(arrows, currency codes, commas, brackets, unit letters like m/k/b) so it's
ready to use in a real calculation if you ever need it.
 
WHY THIS APPROACH?
Real spreadsheets mix many "dialects" of data in one column (status codes,
percentages, raw numbers, notes). Instead of one giant regex, we run the
value through a small PIPELINE of checks, in order, and stop at the first
one that matches. This makes the logic easy to read and easy to extend
later (just add a new check function).
"""
 
import re                     # 're' = regular expressions. Used to test if a string
                               # "looks like" a number or percentage (pattern matching).
import pandas as pd           # 'pandas' lets us read Excel files and store results
                               # in a table (DataFrame) so we can view/save them easily.
 
 
# ---------------------------------------------------------------------------
# STEP 1: Define the "vocabulary" of known noise/keywords, in one place.
# Keeping these as constants at the top means if your sheet uses a slightly
# different word or unit later, you only change it here - not deep in the code.
# ---------------------------------------------------------------------------
 
# Words that mean a RAG (Red/Amber/Green) status. Dictionary maps the
# lowercase word -> the standardised single-letter code we want to output.
RAG_WORDS = {
    "r": "R", "red": "R",
    "g": "G", "green": "G",
    "a": "A", "amber": "A",
}
 
# Up/down trend arrow characters seen in the sheet. These carry meaning
# (direction of change) but the USER asked us to ignore/strip them for
# the purpose of typing the value, so we just delete them.
ARROW_CHARS = ["↑", "↓", "⬆", "⬇", "→", "←"]
 
# Unit letters that can be glued onto the END of a number, e.g. "200k",
# "2654.70m". Longer ones are listed too (e.g. "bn", "mn") so we can match
# them before the shorter single-letter versions.
UNIT_SUFFIXES = ["bn", "mn", "m", "k", "b"]
 
# Currency-style prefixes that can appear BEFORE a number, e.g. "USD 1200".
CURRENCY_PREFIXES = ["usd", "inr", "rs", "$", "₹"]
 
 
# ---------------------------------------------------------------------------
# STEP 2: Small "helper" functions - each one does ONE job.
# Breaking the problem into small functions makes it easier to test and
# easier to explain (this is called the "Single Responsibility Principle").
# ---------------------------------------------------------------------------
 
def clean_noise(value: str) -> str:
    """
    Removes arrows, currency prefixes, commas, and accounting-style
    brackets (negative numbers) from a string. Does NOT touch a unit
    suffix like 'm'/'k' yet - that is handled by a separate function
    because we need to check the REST of the string is a clean number
    before we decide the suffix really is a unit and not part of a word.
    """
    cleaned = value.strip()                       # remove leading/trailing spaces
 
    # Remove every arrow character, wherever it sits in the string
    for arrow in ARROW_CHARS:
        cleaned = cleaned.replace(arrow, "")
 
    cleaned = cleaned.strip()                      # a space may be left behind after removing the arrow
 
    # Remove a currency prefix, but ONLY if it's at the very start of the string
    lowered = cleaned.lower()
    for prefix in CURRENCY_PREFIXES:
        if lowered.startswith(prefix):
            cleaned = cleaned[len(prefix):].strip()
            lowered = cleaned.lower()               # refresh lowered copy after trimming
            break                                    # stop after the first prefix match
 
    # Accounting-style negative numbers look like "(222,811)" meaning -222811.
    # If the string is wrapped in brackets, remember that it's negative and
    # strip the brackets off.
    is_negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        is_negative = True
        cleaned = cleaned[1:-1]                     # drop the outer "(" and ")"
 
    cleaned = cleaned.replace(",", "")               # remove thousand-separator commas, e.g. "222,811" -> "222811"
 
    if is_negative and not cleaned.startswith("-"):
        cleaned = "-" + cleaned                      # re-apply the negative sign we detected above
 
    return cleaned.strip()
 
 
def strip_unit_suffix(value: str):
    """
    If the string ENDS with a known unit letter (m, k, b, bn, mn), remove
    it and tell the caller which suffix was found.
    Returns a tuple: (value_without_suffix, suffix_found_or_None)
    """
    lowered = value.lower()
    # Check longer suffixes ("bn", "mn") before shorter ones ("b", "m") so we
    # don't accidentally chop only half of "bn" off and leave a stray "n".
    for suffix in sorted(UNIT_SUFFIXES, key=len, reverse=True):
        if lowered.endswith(suffix) and len(value) > len(suffix):
            return value[: -len(suffix)], suffix
    return value, None                                # no suffix found - return unchanged
 
 
def is_number(text: str) -> bool:
    """
    Checks whether a string looks EXACTLY like a plain number,
    e.g. "123", "-45.6", "0.53". Returns True or False.
    """
    # ^        -> start of string
    # -?       -> an optional minus sign
    # \d+      -> one or more digits
    # (\.\d+)? -> an optional decimal point followed by more digits
    # $        -> end of string  (this forces the WHOLE string to match, not just part of it)
    pattern = r"^-?\d+(\.\d+)?$"
    return bool(re.match(pattern, text))
 
 
# ---------------------------------------------------------------------------
# STEP 3: The main function that ties everything together.
# This is the one function you actually call from the outside.
# ---------------------------------------------------------------------------
 
def classify_value(raw_value):
    """
    Takes ONE cell value (could be text, int, float, or None/NaN coming
    from Excel) and returns a dictionary with:
        - original_value : exactly what was passed in
        - data_type      : 'empty', 'rag', 'percentage', 'numeric', or 'text'
        - cleaned_value  : the noise-free version (float for numbers/percentages,
                            standardised letter for rag, original text for text)
    """
 
    # ---- Handle empty / missing cells first (Excel often gives blanks as None or NaN) ----
    if raw_value is None or (isinstance(raw_value, str) and raw_value.strip() == ""):
        return {"original_value": raw_value, "data_type": "empty", "cleaned_value": None}
 
    # If pandas/Excel already gave us a real number (not text), there's nothing to clean.
    if isinstance(raw_value, (int, float)):
        return {"original_value": raw_value, "data_type": "numeric", "cleaned_value": float(raw_value)}
 
    text_value = str(raw_value).strip()               # force everything else to a plain string to test against
 
    # ---- CHECK 1: is the WHOLE cell just a RAG word? ----
    if text_value.lower() in RAG_WORDS:
        return {
            "original_value": raw_value,
            "data_type": "rag",
            "cleaned_value": RAG_WORDS[text_value.lower()],
        }
 
    # ---- Clean obvious noise before testing for percentage/numeric ----
    cleaned = clean_noise(text_value)
 
    # ---- CHECK 2: does it end in "%" and is everything before it a clean number? ----
    if cleaned.endswith("%"):
        number_part = cleaned[:-1].strip()             # drop the "%" character itself
        if is_number(number_part):
            return {
                "original_value": raw_value,
                "data_type": "percentage",
                "cleaned_value": float(number_part),   # e.g. 12.05 means "12.05%"
            }
 
    # ---- CHECK 3: after removing a unit suffix (m/k/b/bn/mn), is it a clean number? ----
    number_candidate, suffix_found = strip_unit_suffix(cleaned)
    if is_number(number_candidate):
        return {
            "original_value": raw_value,
            "data_type": "numeric",
            "cleaned_value": float(number_candidate),  # unit is DROPPED as per your instructions,
        }                                               # i.e. "2654.70m" -> 2654.70, not 2,654,700,000
 
    # ---- CHECK 4: nothing matched -> treat it as free text ----
    return {"original_value": raw_value, "data_type": "text", "cleaned_value": text_value}
 
 
# ---------------------------------------------------------------------------
# STEP 4: A convenience function to run this over an ENTIRE Excel sheet.
# ---------------------------------------------------------------------------
 
def classify_excel_sheet(file_path: str, sheet_name=0) -> pd.DataFrame:
    """
    Reads every cell of an Excel sheet and classifies it.
    header=None tells pandas "don't treat row 1 as column titles" - we want
    the raw grid exactly as it appears, since your sheet has data starting
    right from the first cell.
    Returns a tidy DataFrame with one row PER CELL: row, column, original_value,
    data_type, cleaned_value.
    """
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
 
    results = []                                       # empty list we will fill up, one dict per cell
    for row_idx, row in df_raw.iterrows():              # loop over every row
        for col_idx, cell_value in row.items():         # loop over every column in that row
            result = classify_value(cell_value)         # run our classifier on this one cell
            result["row"] = row_idx                     # remember which row it came from
            result["column"] = col_idx                  # remember which column it came from
            results.append(result)
 
    return pd.DataFrame(results)                        # turn the list of dicts into a table
 
 
# ---------------------------------------------------------------------------
# STEP 5: DEMO - run the classifier on sample values similar to your sheet.
# This block only runs when you execute this file directly
# (python classify_data_types.py), not when you import it elsewhere.
# ---------------------------------------------------------------------------
 
if __name__ == "__main__":
    sample_values = [
        "0%", "6.4", "2654.70m", "13.8%", "122.7%", "Nil", "0.08%",
        "104%", "100", "2%", "USD (-222,811)", "9%", "G", "R", "Green",
        "Amber", "FRP", "↓ 51.7%", "↑ 200k", "0.53%", "5726m", "-414m",
        "0.984m", "0 VH, 0H", "H=0,M=0", "100%", "0.984",
        "CD1 : 94.6%CD2: 91.8%CD3: 95.9%",
    ]
 
    print(f"{'ORIGINAL':35} {'TYPE':12} {'CLEANED'}")
    print("-" * 65)
    for value in sample_values:
        result = classify_value(value)
        print(f"{str(result['original_value']):35} {result['data_type']:12} {result['cleaned_value']}")
 
    # Uncomment the two lines below and put your real file path to run this
    # on your actual Excel file, then save the results to a new CSV file.
    #
    # results_df = classify_excel_sheet("your_file.xlsx", sheet_name="Sheet2")
    # results_df.to_csv("classified_results.csv", index=False)
