import pandas as pd

# Prepare data
data = pd.DataFrame({
    "id": [14, 82, 5],
    "name": ["Cookie Corp.", "Chocolate Inc.", "Banana AG"],
    "country": ["FR", "DE", "FR"],
    "stock_price": [152.501, 99.00, 45.12],
})
summary = data.groupby("country", as_index=False)["stock_price"].agg("mean")  # Create a second data frame
summary = summary.rename(columns={"stock_price": "mean_stock_price"})

# Company details to include at the top of the "Company_Report" sheet
company_details = [
    ["Company Name", "ACME Corp."],
    ["Report Date", "2024-05-08"],
    ["Author", "Jane Doe"],
]

output_path = "report.xlsx"

# Set a default format for the whole document
writer = pd.ExcelWriter(output_path, engine='xlsxwriter', datetime_format="%Y%m%d")
workbook = writer.book
default_format = workbook.add_format({
    'font_name': 'Open Sans',
    'font_size': 10,
    'font_color': '#3b3b3b'
})

# Create Front Page with list of contents
cover_link_dict = {"A2": "Company_Report", "A3": "Summary"}
sheet_front_page = workbook.add_worksheet("Front_Page")
sheet_front_page.write("A1", "List of contents:")
sheet_front_page.set_column(0, 0, 20)
for column, sheet_name in cover_link_dict.items():
    sheet_front_page.write_url(column, "internal:{}!A1:A1".format(sheet_name), string=sheet_name)

# Write the company details with a header
sheet_report = workbook.add_worksheet("Company_Report")
header_format = workbook.add_format({
    'font_name': 'Open Sans',
    'bold': True,
    'bg_color': '#d3d3d3',  # Light gray background
    'valign': 'vcenter'
})

# Merging cells for the header and writing the title
sheet_report.merge_range("A1:B1", "Company Details", header_format)

# Writing the individual company details below the header
for row_num, (label, value) in enumerate(company_details, start=1):
    sheet_report.write(row_num, 0, label, default_format)
    sheet_report.write(row_num, 1, value, default_format)

# Start writing the data table after the company details
data_start_row = len(company_details) + 2  # +2 to account for the merged header row
data.to_excel(writer, index=False, sheet_name="Company_Report", startrow=data_start_row)
summary.to_excel(writer, index=False, sheet_name="Summary", startrow=1)
sheet_summary = writer.sheets["Summary"]

# Add links to all sheets to return to the Front Page
def link_to_front_page(sheet):
    """
    Put a clickable link to the Front Page to the first cell of a sheet.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    """
    sheet.write_url("A1", "internal:{}!A1:A1".format("Front_Page"), string="Back_to_Front_Page")


link_to_front_page(sheet_report)
link_to_front_page(sheet_summary)

# Change the format of the numeric columns
custom_format_dict = {
    "font": "Open Sans",
    "italic": True,
    "font_size": 10,
    "font_color": "#3b3b3b",
    "num_format": 0x02,  # always display two decimal places
}
custom_format = workbook.add_format(custom_format_dict)
sheet_report.set_column(3, 3, cell_format=custom_format)
sheet_summary.set_column(1, 1, cell_format=custom_format)

# Format sheets as tables
def format_as_table(sheet, df, start_row=0):
    """
    Format the data in a worksheet as a table, starting at the specified row.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    :param df: data frame with the original data that is in the sheet. Required to write the column names into the
      table and to determine the required table size.
    :param start_row: The starting row of the table
    """
    sheet.add_table(
        first_row=start_row,
        first_col=0,
        last_row=start_row + df.shape[0] - 1,
        last_col=df.shape[1] - 1,
        options={"columns": [{"header": col} for col in df.columns]},
    )


format_as_table(sheet_report, data, start_row=data_start_row)
format_as_table(sheet_summary, summary, start_row=1)

# Change column widths
def set_column_widths(sheet, col_width_dict):
    """
    Change column widths for selected columns of a given sheet. All other columns remain unchanged.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    :param col_width_dict: dictionary with the columns widths where the keys are the column indices and the values are
      the column widths, e.g. {0: 10, 5: 25}
    """
    for col_index, width in col_width_dict.items():
        sheet.set_column(col_index, col_index, width)


set_column_widths(sheet_report, {0: 15, 1: 20, 2: 15, 3: 15})
set_column_widths(sheet_summary, {0: 15, 1: 20})

# Freeze the first few rows in the report sheet (including company details)
def freeze_first_rows(sheet, num_rows):
    """
    Freeze the top rows of a sheet so they remain visible while scrolling down.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    :param num_rows: Number of rows to freeze from the top
    """
    sheet.freeze_panes(row=num_rows, col=0)


freeze_first_rows(sheet_report, data_start_row)
freeze_first_rows(sheet_summary, 2)

# Close the writer and save the Excel file
writer.close()