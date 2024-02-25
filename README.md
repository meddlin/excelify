# excelify

A utility to convert CSV files to Excel workbooks. 

For the rather specific situation, when you have a CSV file you need to open
frequently in Excel, but it doesn't make sense to simply convert it (like say
a CSV report that falls out of an automated tool). And it's getting tiresome 
to run move through the tedium of: (1) open file (2) adjust columns (3) bold 
the header (4) turn on filters, and *finally*... (5) get to work.

This utility can convert those .csv files to .xlsx files with all of those
settings already *set*.

## How to Use

- Your CSV file at: `./report.csv`
- Run: `python main.py --csv ./report.csv --output report.xlsx --sheet new_filter --filter-cols Col1,Col2`
- Open your new `report.xlsx`

### Typical Use-case

This tool was inspired by incredibly wide reports, i.e. 20-30 columns wide. With such wide reports,
it can be helpful to have the most useful columns in their own worksheet within a larger workbook.

So, taken from the command above, this example would work like so:

```bash
$> python main.py --csv ./report.csv --output report.xlsx --sheet your_view --filter-cols "Date,Name,Some Field"
```

**Original `.csv`**

| Date | Name | Some Field | Another | Why | More Data | Moar Data | Yes... | Even More | Really Why?... |
| ---- | ---- | ---------- | ------- | --- | --------- | --------- | ------ | --------- | -------------- |
| 01/01/2024 | John S. | content | content | content | content | content | content | content | content | content |
| 01/02/2024 | Jane S. | content | content | content | content | content | content | content | content | content |

**New Excel workbook --> `.xlsx`**

**Sheet: `new_filter`**

| Date | Name | Some Field |
| ---- | ---- | ---------- |
| 01/01/2024 | John S. | content |
| 01/02/2024 | Jane S. | content |

**Sheet: `your_view`**

| Date | Name | Some Field | Another | Why | More Data | Moar Data | Yes... | Even More | Really Why?... |
| ---- | ---- | ---------- | ------- | --- | --------- | --------- | ------ | --------- | -------------- |
| 01/01/2024 | John S. | content | content | content | content | content | content | content | content | content |
| 01/02/2024 | Jane S. | content | content | content | content | content | content | content | content | content |


### Parameters

- `--csv` | Path to .csv file to process
- `--output` | Output path for resulting .xlsx file
- `--sheet` | Worksheet name where filtered data will land
- `--filter-cols` | Comma-separated list of columns to INCLUDE on new worksheet, other columns are left behind on 'raw' worksheet

## Contributing: Getting Started

Create a virtual environment (in current directory)

`python -m venv .`

Install requirements

`pip install requirements.txt`

## References

openpyxl
- Ref: https://openpyxl.readthedocs.io/en/stable/index.html

string
- Ref: https://docs.python.org/3/library/string.html#module-string
- Useful for `ascii_uppercase`