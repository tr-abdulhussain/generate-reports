import pandas as pd
import datetime
import asyncio
import sys
from openpyxl.utils import get_column_letter
from gspread_formatting import DataValidationRule, BooleanCondition, set_data_validation_for_cell_range
from snipeit_functions import correct_asset_status
from google_functions import authenticate_google_services, create_google_sheet, write_to_google_sheet, download_google_sheet, delete_google_sheet

CURRENT_MONTH = datetime.datetime.now().strftime("%B")
CURRENT_YEAR = datetime.datetime.now().strftime("%Y")

SHARE_WITH_EMAIL = "infra@tooltwist.com"

MAPPING_SHEET_ID = "1iiwUvw5QaRM6Hvq7uY8B5UnRvg45EICZDnQWITRxvrc"
INITIAL_SHEET_NAME = f"Initial MACHINE REPORTS {CURRENT_MONTH} {CURRENT_YEAR}"
XLSX_FILE = f"MACHINE REPORTS {CURRENT_MONTH} {CURRENT_YEAR}.xlsx"

RAW_SHEET_NAME = "Raw"
DUPLICATE_SHEET_NAME = "Duplicate Machines"
WRONG_STATUS_SHEET_NAME = "Wrong Statuses"

ASSET_TAG_COLUMN = "Asset Tag"
ASSET_NAME_COLUMN = "Asset Name"
CATEGORY_COLUMN = "Category"
MODEL_COLUMN = "Model"
MACHINE_COLUMN = "Machine"
MACHINE_STATUS_COLUMN = "Machine Status"
STATUS_COLUMN = "Status"
ACTIONS_COLUMN = "Actions to be taken"

EXCLUDED_EMPLOYEES = ["Brent Kearney"]
EXCULDED_ASSET_KEYWORDS = ["Personal", "Client"]

INCLUDED_COLUMNS = [
    "Company", "Asset Name", "Asset Tag", "Model", "Category", "Cost", "Order Number", "Supplier", 
    "Status", "Year Model", "Type.1", "Storage Size", "Storage Type", "RAM", "Processor", 
    "Processor Speed", "PBI Number", "PBO Number"
]

def format_raw_sheet(csv_file):
    df = pd.read_csv(csv_file)
    df = df[df[CATEGORY_COLUMN] == MACHINE_COLUMN]
    client, _, _ = authenticate_google_services()
    mapping_spreadsheet = client.open_by_key(MAPPING_SHEET_ID)
    machine_status_worksheet = mapping_spreadsheet.worksheet(MACHINE_STATUS_COLUMN)
    list_rows_machine_status_worksheet = machine_status_worksheet.get_all_values()
    machine_status_df = pd.DataFrame(
        list_rows_machine_status_worksheet[1:], columns=list_rows_machine_status_worksheet[0]
    )
    df = df[INCLUDED_COLUMNS]
    df = df.rename(columns={"Type.1": "Type"})

    df[STATUS_COLUMN] = df[STATUS_COLUMN].astype(str).str.replace(r'\s+', '', regex=True)
    machine_status_df["Raw Data"] = machine_status_df["Raw Data"].astype(str).str.replace(r'\s+', '', regex=True)

    df = df[~df[ASSET_TAG_COLUMN].str.contains("|".join(EXCULDED_ASSET_KEYWORDS), na=False)]
    df = df[~df[ASSET_NAME_COLUMN].isin(EXCLUDED_EMPLOYEES)]
    
    df = df.merge(machine_status_df.rename(columns={"Raw Data": STATUS_COLUMN}), on=STATUS_COLUMN, how="left")

    df[ASSET_TAG_COLUMN] = df[ASSET_TAG_COLUMN].fillna("")
    duplicate_df = df[df[ASSET_NAME_COLUMN].notna() & (df[ASSET_NAME_COLUMN] != "") & (df[STATUS_COLUMN].str.contains("Assigned"))]
    duplicates = duplicate_df[duplicate_df[ASSET_NAME_COLUMN].duplicated()][ASSET_NAME_COLUMN].to_frame()

    wrong_status_df = df[df[MACHINE_STATUS_COLUMN] == "Wrong Status"]
    wrong_statuses = wrong_status_df[[ASSET_TAG_COLUMN, STATUS_COLUMN]].copy()
    wrong_statuses[ACTIONS_COLUMN] = ""
    snipeit_has_errors = not wrong_statuses.empty

    spreadsheet = create_google_sheet(INITIAL_SHEET_NAME, snipeit_has_errors, SHARE_WITH_EMAIL)

    if snipeit_has_errors:
        write_to_google_sheet(spreadsheet, WRONG_STATUS_SHEET_NAME, wrong_statuses)
        worksheet = spreadsheet.worksheet(WRONG_STATUS_SHEET_NAME)

        actions = ["Set as Assigned", "Set as Spare"]
        col_letter = get_column_letter(len(wrong_statuses.columns))
        cell_range = f"{col_letter}2:{col_letter}{len(wrong_statuses)+1}"

        rule = DataValidationRule(
            condition=BooleanCondition("ONE_OF_LIST", actions),
            showCustomUi=True
        )
        set_data_validation_for_cell_range(worksheet, cell_range, rule)
        print(f"Check {WRONG_STATUS_SHEET_NAME} Sheet for assets with wrong status.")

        patched_asset_tags = fix_status_snipeit(spreadsheet)
        df_filtered = df[df[ASSET_TAG_COLUMN].isin(patched_asset_tags)]
        df_filtered = df_filtered.merge(wrong_statuses[[ASSET_TAG_COLUMN, ACTIONS_COLUMN]], on=ASSET_TAG_COLUMN, how='left')
        status_mapping = {
            "Set as Assigned": "Assigned",
            "Set as Spare": "Spare"
        }
        df_filtered[MACHINE_STATUS_COLUMN] = df_filtered[ACTIONS_COLUMN].map(status_mapping).fillna(df_filtered[MACHINE_STATUS_COLUMN])
        df.update(df_filtered[[ASSET_TAG_COLUMN, MACHINE_STATUS_COLUMN]])
    if not duplicates.empty:
        write_to_google_sheet(spreadsheet, DUPLICATE_SHEET_NAME, duplicates)
    write_to_google_sheet(spreadsheet, RAW_SHEET_NAME, df)
    spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    excel_path = download_google_sheet(spreadsheet, XLSX_FILE)
    delete_google_sheet(spreadsheet, INITIAL_SHEET_NAME)
    return RAW_SHEET_NAME, excel_path, df, SHARE_WITH_EMAIL

def fix_status_snipeit(spreadsheet):
    take_actions = input("Have you assigned the actions to be taken? Y/n: ")
    if take_actions == "Y":
        worksheet = spreadsheet.worksheet(WRONG_STATUS_SHEET_NAME)

        data = worksheet.get_all_values()
        headers = data[0]
        rows = data[1:]
        wrong_statuses = pd.DataFrame(rows, columns=headers)

        return asyncio.run(correct_asset_status(wrong_statuses[[ASSET_TAG_COLUMN, ACTIONS_COLUMN]]))
    if take_actions == "n":
        print("Wrong statuses have to be corrected before proceeding to generate reports.")
        delete_google_sheet(spreadsheet, INITIAL_SHEET_NAME)
        sys.exit()
    print("Y or n are the only allowed values")
    fix_status_snipeit(spreadsheet)
