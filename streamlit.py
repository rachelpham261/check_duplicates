import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import os

def process_excel(file_path):
    # Load the Excel workbook
    wb = load_workbook(file_path)
    all_sheets = wb.sheetnames

    # Filter and sort sheets based on date format MM.DD.YY
    date_sheets = []
    for sheet in all_sheets:
        try:
            # Try to parse the sheet name as a date
            parsed_date = datetime.strptime(sheet, '%m.%d.%y')
            date_sheets.append((sheet, parsed_date))
        except ValueError:
            # If the sheet name doesn't match the date format, ignore it
            continue

    # Sort sheets by the parsed date
    date_sheets.sort(key=lambda x: x[1])
    sorted_sheets = [sheet[0] for sheet in date_sheets]

    # Initialize an empty dataframe for the aggregation
    aggregate_df = pd.DataFrame(columns=['Full Site Address', 'Date'])

    # Iterate over each sorted sheet
    for sheet in sorted_sheets:
        df = pd.read_excel(file_path, sheet_name=sheet)

        # Convert all relevant address columns to string and strip them
        address_columns = [
            'Site Address House Number', 
            'Site Address Street Prefix', 
            'Site Address Street Name', 
            'Site Address Unit Number', 
            'Site Address City', 
            'Site Address State', 
            'Site Address Zip+4'
        ]
        df[address_columns] = df[address_columns].fillna('').astype(str).apply(lambda x: x.str.strip())

        # Concatenate the address components in the desired format
        df['Full Site Address'] = (
            df['Site Address House Number'] + ' ' +
            df['Site Address Street Prefix'] + ' ' + df['Site Address Street Name'] + ', ' +
            df['Site Address Unit Number'].apply(lambda x: f"Unit {x}, " if x != '' else '') +
            df['Site Address City'] + ', ' +
            df['Site Address State'] + ' ' + df['Site Address Zip+4']
        ).str.replace(r'\s+', ' ', regex=True).str.strip().replace(', ,', ',', regex=False)

        # Select only the necessary columns and add the sheet (date) as a column
        sheet_df = df[['Full Site Address']].copy()
        sheet_df['Date'] = sheet

        # Append to the summary dataframe using pd.concat
        aggregate_df = pd.concat([aggregate_df, sheet_df], ignore_index=True)

    # Identify duplicates and list all dates where duplicates exist, excluding the current date
    aggregate_df['Duplicate Dates'] = aggregate_df.apply(
        lambda row: ', '.join(sorted(set(aggregate_df.loc[aggregate_df['Full Site Address'] == row['Full Site Address'], 'Date']) - {row['Date']})),
        axis=1
    )

    # Save the processed file
    processed_file_path = os.path.splitext(file_path)[0] + '_Processed.xlsx'
    with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
        aggregate_df.to_excel(writer, sheet_name='Aggregation', index=False)

    # Load the workbook to apply conditional formatting
    wb = load_workbook(processed_file_path)
    ws = wb['Aggregation']

    # Apply conditional formatting to highlight duplicates in the 'Full Site Address' column
    duplicate_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    formula = f'COUNTIF($A$2:$A${len(aggregate_df) + 1}, A2)>1'
    ws.conditional_formatting.add(f'A2:A{len(aggregate_df) + 1}', FormulaRule(formula=[formula], fill=duplicate_fill))

    # Save the workbook with the formatting applied
    wb.save(processed_file_path)

    return processed_file_path

def main():
    st.title("Check for Duplicate Addresses")

    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Save the uploaded file temporarily
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.write("Processing the file...")
        processed_file = process_excel(uploaded_file.name)

        st.write("Download the processed file:")
        with open(processed_file, "rb") as f:
            st.download_button(label="Download", data=f, file_name=os.path.basename(processed_file))

if __name__ == "__main__":
    main()
