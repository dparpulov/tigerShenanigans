from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart, Reference

def is_valid_date(date_string):
    try:
        datetime.strptime(date_string, '%d.%m.%Y')
        return True
    except ValueError:
        return False


def read_excel_data(excel_file):
    # Load the workbook
    wb = openpyxl.load_workbook(excel_file)

    # Get the first sheet
    sheet = wb.active

    # Initialize a list to store the data
    data = []

    # Iterate over the rows in the sheet, starting from row 11
    for row in sheet.iter_rows(min_row=11):
        # Get the value of the first cell in the row (column 2)
        date_string = row[1].value

        # Check if the value is a valid date
        if not is_valid_date(date_string):
            # If the value is not a valid date, stop reading the data
            break

        # If the value is a valid date, append the row data to the data list
        data.append([cell.value for cell in row[1:9]])

    return data


def create_excel_file(data, date):
    # Create a new workbook
    wb = openpyxl.Workbook()

    # Get the first sheet and rename it to "Data"
    sheet1 = wb.active
    sheet1.title = "Data"

    # Add the data to the first sheet
    for i, row in enumerate(data, start=1):
        for j, cell in enumerate(row, start=1):
            sheet1.cell(row=i, column=j, value=cell)

    # Create the second sheet and add a bar chart
    sheet2 = wb.create_sheet("Chart")

    # Initialize lists to store the data for the X and Y axis
    x_data = []
    y_data = []

    # Iterate over the rows in the data
    for row in data:
        # Get the value in column 2
        date_string = row[1]
        # Check if the value is a valid date and matches the specified date
        if is_valid_date(date_string) and date_string == date:
            # If the value is a valid date and matches the specified date, append the data from column 3 and 4 to the x_data and y_data lists, respectively
            x_data.append(row[2])
            y_data.append(row[3])

    # Add the data to the second sheet
    for i, (x, y) in enumerate(zip(x_data, y_data), start=1):
        sheet2.cell(row=i, column=1, value=x)
        sheet2.cell(row=i, column=2, value=y)

    # Create a bar chart
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = f"Bar chart for {date}"
    chart.y_axis.title = "Y axis"
    chart.x_axis.title = "X axis"

    # Set the data for the chart
    data = Reference(sheet2, min_col=2, min_row=1, max_col=2, max_row=len(y_data))
    cats = Reference(sheet2, min_col=1, min_row=1, max_row=len(x_data))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Add the chart to the sheet
    sheet2.add_chart(chart, "D1")

    # Save the workbook
    wb.save("output.xlsx")




