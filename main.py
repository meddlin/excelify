import csv
from string import ascii_uppercase
import argparse, textwrap
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.filters import (
    FilterColumn,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
    Filters,
)

# Note: This looks helpful.
# Openpyxl - How to apply autofilter to columns
# https://openpyxl.readthedocs.io/en/latest/filters.html
def configure_filters(worksheet, data_col_height, cols_amount):
    """ configure filters for the spreadsheet """
    ## NOTE: worksheet -> the worksheet we're working on
    ## NOTE: data_col_height -> number of rows (height) of the particular column
    ## NOTE: cols_amount -> Amount of columns on the worksheet

    filters = worksheet.auto_filter
    # filters.ref = f"A1:C{data_col_height}"
    filters.ref = f"A1:{ascii_uppercase[cols_amount - 1]}{data_col_height}"
    col = FilterColumn(colId = 0)
    filters.filterColumn.append(col)

def set_zoom_scale(workbook):
    for ws in workbook.worksheets:
        ws.sheet_view.zoomScale = 110

def autofit_columns(ws):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

def bold_header(ws):
    """ Bold every cell in the header row. (Iterating over uppercase letters.) """
    for c in ascii_uppercase:
        ws[f'{c}1'].font = Font(bold=True)

# This was helpful: https://stackoverflow.com/questions/37182528/how-to-append-data-using-openpyxl-python-to-excel-file-from-a-specified-row
def read_csv(filename: str) -> dict[str, dict[str, str]]:
    """ Read CSV into two dictionaries. A full and abridged version. """
    full = []
    abridged = []

    with open(filename, 'r', encoding = 'utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            full.append(row)
            abridged.append({
                                'Posting Date': row['Posting Date'], 
                                'Amount': row['Amount'], 
                                'Description': row['Description']
                            })
    
    return { 'full': full, 'abridged': abridged }

def create_workbooks(full: dict[str, str], abridged: dict[str, str]):
    """ Create workbooks and control what changes are made to them. (This is where the magic happens)"""

    wb = Workbook()
    del wb['Sheet'] # Delete the default sheet

    ws1 = wb.create_sheet('Transactions')
    ws2 = wb.create_sheet('raw_data')
    
    wb.active = ws1
    ws1.append(['Posting Date', 'Amount', 'Description', 'Transaction Category', 'Extended Description'])
    bold_header(ws1) # NOTE: Must execute AFTER data has been written to the header row.

    row_counter = 0
    num_of_cols = len(abridged[0].keys())
    for row in abridged:
        # ws1.append([row['Posting Date'], row['Amount'], row['Description'], row['Transaction Category'], row['Extended Description']])
        ws1.append([row['Posting Date'], row['Amount'], row['Description']])
        row_counter += 1
    
    autofit_columns(ws1)
    configure_filters(ws1, row_counter, num_of_cols)

    
    wb.active = ws2
    ws2.append([
        'Transaction ID', 
        'Posting Date', 
        'Effective Date', 
        'Transaction Type', 
        'Amount', 
        'Check Number', 
        'Reference Number',
        'Description',
        'Transaction Category',
        'Type',
        'Balance',
        'Memo',
        'Extended Description'
    ])
    bold_header(ws2) # NOTE: Must execute AFTER data has been written to the header row.
    full_sheet_row_counter = 0
    num_of_cols = len(full[0].keys())
    for row in full:
        ws2.append([
            row['Transaction ID'], 
            row['Posting Date'], 
            row['Effective Date'],
            row['Transaction Type'],
            row['Amount'], 
            row['Check Number'],
            row['Reference Number'],
            row['Description'],
            row['Transaction Category'],
            row['Type'],
            row['Balance'],
            row['Memo'],
            row['Extended Description']
        ])
        full_sheet_row_counter += 1
    
    autofit_columns(ws2)
    configure_filters(ws2, full_sheet_row_counter, num_of_cols)

    set_zoom_scale(wb)
    wb.save('transactions-report.xlsx')

def main():
    print("Excelify...")
    parser = argparse.ArgumentParser(prog='excelify',
                                     formatter_class=argparse.RawDescriptionHelpFormatter,
                                     epilog=textwrap.dedent('''
                                                            Examples:
                                                                excelify.py --csv [file]
                                                                excelify.py --csv [file] --options bold_header,filter,af_cols
                                                        '''))
    parser.add_argument('--csv', type=str, required=True, dest='cli_args', help="Path to .csv file to process")
    parser.add_argument('--output', type=str, required=True, dest='cli_args', help="Output path for resulting .xlsx file")
    args = parser.parse_args()
    cli_args = args.cli_args

    print(f'cli_args: {cli_args}')

if __name__ == "__main__":
    main()