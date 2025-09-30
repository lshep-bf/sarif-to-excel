import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
import os
import sys
import argparse

def process_sarif(input_file_path):
    # Load SARIF file with utf-8 encoding
    with open(input_file_path, 'r', encoding='utf-8') as file:
        sarif = json.load(file)

    # Extract the results section
    results = sarif.get('runs', [])[0].get('results', [])
    
    # Process results into a DataFrame
    rows = []
    for result in results:
        file_path = result.get('locations', [{}])[0].get('physicalLocation', {}).get('artifactLocation', {}).get('uri', 'N/A')
        file_name = os.path.basename(file_path)  # Extract the file name from the path
        rows.append({
            "Severity": result.get('level', 'N/A'),
            "Message": result.get('ruleId', 'N/A'),  # Renamed from Rule ID
            "Details": result.get('message', {}).get('text', 'N/A'),  # Renamed from Message
            "Path": file_path,  # Renamed from File
            "Page": file_name,  # Use the file name as the page value
            "Line": result.get('locations', [{}])[0].get('physicalLocation', {}).get('region', {}).get('startLine', 'N/A')
        })
    
    # Convert to DataFrame
    df = pd.DataFrame(rows, columns=["Severity", "Message", "Details", "Path", "Page", "Line"])
    
    # Generate output file path in same directory as input, with .xlsx extension
    input_dir = os.path.dirname(input_file_path)
    input_filename = os.path.basename(input_file_path)
    output_filename = os.path.splitext(input_filename)[0] + '.xlsx'
    output_file = os.path.join(input_dir, output_filename)
    
    # Save to Excel
    df.to_excel(output_file, index=False, sheet_name="Results")

    # Add table formatting and adjust column widths
    add_excel_table_and_adjust_columns(output_file, "Results", wrap_columns=["Message", "Details"], auto_fit_columns=["Path", "Page", "Line"])
    print(f"Processed Report saved to {output_file}")


def add_excel_table_and_adjust_columns(file_path, sheet_name, wrap_columns, auto_fit_columns):
    # Load the workbook and worksheet
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    
    # Determine the range of the table (e.g., A1:F20)
    start_cell = ws.cell(row=1, column=1)
    end_cell = ws.cell(row=ws.max_row, column=ws.max_column)
    table_range = f"{start_cell.coordinate}:{end_cell.coordinate}"
    
    # Create a table
    table = Table(displayName="SARIFTable", ref=table_range)
    
    # Add a table style (banded rows but no banded columns)
    style = TableStyleInfo(
        name="TableStyleMedium9",  # Choose a predefined Excel table style
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,  # Enable banded rows
        showColumnStripes=False  # Disable banded columns
    )
    table.tableStyleInfo = style
    
    # Add the table to the worksheet
    ws.add_table(table)

    # Adjust column widths and wrap text
    for col_name in ws[1]:
        if col_name.value in auto_fit_columns:
            # Auto-fit column widths based on content length
            col_index = col_name.column
            column_letter = ws.cell(row=1, column=col_index).column_letter
            max_length = max(len(str(ws.cell(row=row, column=col_index).value or "")) for row in range(2, ws.max_row + 1))  # Exclude header
            ws.column_dimensions[column_letter].width = max_length + 2  # Add padding
        elif col_name.value in wrap_columns:
            # Adjust width to 3/4 of content and enable text wrapping
            col_index = col_name.column
            column_letter = ws.cell(row=1, column=col_index).column_letter
            max_length = max(len(str(ws.cell(row=row, column=col_index).value or "")) for row in range(2, ws.max_row + 1))  # Exclude header
            ws.column_dimensions[column_letter].width = (max_length * 0.75)  # Set width to 3/4 of content length
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_index)
                cell.alignment = Alignment(wrap_text=True)  # Enable text wrapping
    
    # Save the workbook
    wb.save(file_path)
    print(f"Table formatting and column adjustments added to {file_path}")

def main():
    parser = argparse.ArgumentParser(description='Convert SARIF reports to Excel format')
    parser.add_argument('sarif_file', help='Path to the SARIF file to process')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.sarif_file):
        print(f"Error: File '{args.sarif_file}' not found.")
        sys.exit(1)
    
    process_sarif(args.sarif_file)

if __name__ == '__main__':
    main()
