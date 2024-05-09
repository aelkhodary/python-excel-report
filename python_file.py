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

output_path = "report.xlsx"

# Set a default format for the whole document
writer = pd.ExcelWriter(output_path, engine='xlsxwriter',datetime_format="%Y%m%d")
workbook = writer.book
workbook.add_format({
    'font_name': 'Open Sans',
    'font_size': 10,
    'font_color': '#3b3b3b'
})

# Create Front Page with list of contents
cover_link_dict = {"A2": "Company_Report", "A3": "Summary"}
sheet_front_page = workbook.add_worksheet("Front Page")
sheet_front_page.write("A1", "List of contents:")
sheet_front_page.set_column(0, 0, 15)
for column, sheet_name in cover_link_dict.items():
    sheet_front_page.write_url(column, "internal:{}!A1:A1".format(sheet_name), string=sheet_name)

# Write data to multiple sheets
# Start in the second row (index 1) to leave some space for multiheaders and links
data.to_excel(writer, index=False, sheet_name="Company Report", startrow=1)
sheet_report = writer.sheets["Company Report"]
summary.to_excel(writer, index=False, sheet_name="Summary", startrow=1)
sheet_summary = writer.sheets["Summary"]


# Add links to all sheets to get back to front page
def link_to_front_page(sheet):
    """
    Put a clickable link to the Front Page to the first cell of a sheet.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    """
    sheet.write_url("A1", "internal:{}!A1:A1".format("Front Page"), string="Back to Front Page")


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


# Format sheets as table
def format_as_table(sheet, df):
    """
    Format the data in a worksheet as table, starting in the second row
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    :param df: data frame with the original data that is in the sheet. Required to write the column names into the
      table and to determine the required table size.
    """
    sheet.add_table(
        first_row=1,
        first_col=0,
        last_row=df.shape[0],
        last_col=df.shape[1] - 1,
        options={"columns": [{"header": col} for col in df.columns]},
    )


format_as_table(sheet_report, data)
format_as_table(sheet_summary, summary)


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


# Freeze header in report sheet
def freeze_first_2_rows(sheet):
    """
    Freeze header (first to rows) of a sheet, such that they remain visible when scrolling down in the table.
    :param sheet: object of class xlsxwriter.worksheet.Worksheet
    """
    sheet.freeze_panes(row=2, col=0)


freeze_first_2_rows(sheet_report)
freeze_first_2_rows(sheet_summary)

# Add multiheader to report sheet
multiheader_format_dict = {
    "font": "Open Sans",
    "bg_color": "#c4c4c4",
    "valign": "vcenter",
}
multiheader_format = workbook.add_format(multiheader_format_dict)
sheet_report.merge_range(
    first_row=0,
    last_row=0,
    first_col=1,
    last_col=2,
    data="Company information",
    cell_format=multiheader_format,
)

writer.close()