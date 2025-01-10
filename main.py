# main.py
from tkinter import filedialog, Tk
from Scripts.excel_loader import load_excel
from Scripts.word_creator import create_word_document


def main():
    # Set up Tkinter file dialog to select the Excel file
    root = Tk()
    root.withdraw()  # Hide the main window
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])

    if not excel_file:
        print("No file selected.")
        return

    # Load the Excel data as InputData objects
    input_data_list = load_excel(excel_file)

    # Ask for the output Word file path
    output_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if not output_file:
        print("No output file selected.")
        return

    # Create the Word document
    create_word_document(input_data_list, output_file)
    print(f"Word document created successfully: {output_file}")


if __name__ == "__main__":
    main()
