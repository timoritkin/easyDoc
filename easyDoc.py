import datetime
import os
import sys
import time

import xlwings as xw
from docxtpl import DocxTemplate


def calc_next_visit(date):
    # Convert the string date to a datetime object
    date_obj = datetime.datetime.strptime(date, "%d-%m-%Y")

    # Add 183 days (for next visit)
    next_date = date_obj + datetime.timedelta(days=183)

    # Convert back to a string in the desired format
    return next_date.strftime("%d-%m-%Y")


def main():
    xw.Book("easyDoc.xlsm").set_mock_caller()  # Adjust to your Excel file
    wb = xw.Book.caller()
    sht_panel = wb.sheets['מילוי טופס']  # Make sure the sheet name is correct
    sht_log = wb.sheets['היסטוריה של מטופלים']  # Sheet where inputs will be logged (create this in your Excel file)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, 'template', 'Clalit mushlam template.docx')
    doc = DocxTemplate(template_path)  # Use the dynamically determined path

    # Ensure the output folder exists
    output_folder = os.path.join(script_dir, 'generated_docs')
    os.makedirs(output_folder, exist_ok=True)  # Creates the folder if it doesn't exist

    # Read values from C6 to C9
    values = sht_panel.range('C6:C9').value

    # Prepare the context dictionary with the correct key-value pairs
    context = {
        'f_name': str(values[0]),  # First Name from C6
        'l_name': str(values[1]),  # Last Name from C7
        'id': str(int(values[2])) if values[2] is not None else '',  # ID from C8 (converted to integer string)
        'age': str(int(values[3])) if values[3] is not None else '',  # ID from C8 (converted to integer string)
    }
    # Generate a unique filename with timestamp
    timestamp = time.strftime("%d-%m-%Y")  # Format: DD/MM/YYYY

    # Save the inputs in the 'Log' sheet
    last_row = sht_log.range('A' + str(sht_log.cells.last_cell.row)).end('up').row + 1
    sht_log.range(f"A{last_row}:D{last_row}").value = [context['f_name'], context['l_name'], context['id'],
                                                       context['age'], timestamp, calc_next_visit(timestamp)]

    # Generate a unique filename
    file_name = f"{context['f_name']}_{context['l_name']}_{context['id']}_{timestamp}.docx"
    output_path = os.path.join(output_folder, file_name)

    # Render and save the document
    doc.render(context)
    doc.save(output_path)  # Save the result in the 'output' folder

    # Ensure the file is saved before attempting to open
    if os.path.exists(output_path):
        os.startfile(output_path)  # Open the document with the default associated application
    else:
        print("Failed to save the document.")


if __name__ == "__main__":
    main()
