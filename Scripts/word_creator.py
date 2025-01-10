from docx import Document
from datetime import datetime


def format_date(date_str):
    """Format the date to the required format."""
    try:
        # Convert the date string from the format "%d/%m/%Y" to "%d %B %Y"
        date_obj = date_str.strftime("%d %B %Y")
        return date_obj
    except ValueError:
        # If the date format is incorrect, return the original string
        return date_str


def convert_to_subscript(num):
    """Convert a number to its subscript equivalent."""
    subscript_map = {
        '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆'
    }
    return subscript_map.get(str(num), str(num))


def create_word_document(data_list, output_path):
    # Create a Word document
    doc = Document()

    # Add a table for the data
    if data_list:
        # Define the table headers
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'

        # Add headers to the table
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Date acquired'
        hdr_cells[1].text = 'Mobile phone'
        hdr_cells[2].text = 'Landline'
        hdr_cells[3].text = 'Name'

        sub = 0
        # Add rows with data
        for idx, output_data in enumerate(data_list):
            if 5 <= idx <= 97:  # Skipping rows 6 to 98, replace with ellipsis
                if idx == 5:
                    row_cells = table.add_row().cells
                    row_cells[0].text = '…'
                    row_cells[1].text = '…'
                    row_cells[2].text = '…'
                    row_cells[3].text = '…'
                continue

            subscript_index = convert_to_subscript(sub + 1)
            sub = sub +1
            # Add the actual data row
            row_cells = table.add_row().cells
            row_cells[0].text = format_date(output_data.date_acquired)  # Date formatted
            row_cells[1].text = str(output_data.mobile_phone)
            row_cells[2].text = str(output_data.landline)
            row_cells[3].text = f"{output_data.name} {subscript_index}"


    # Save the document
    doc.save(output_path)
