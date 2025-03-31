# Required Libraries:
# - openpyxl: To create and manipulate Excel files.
#
# To install the necessary library, you can run the following commands:
# For Linux/macOS:
# pip install openpyxl
#
# For Windows:
# python -m pip install openpyxl
#
# Note: Make sure Python is installed on your system before running the script.

import openpyxl

# Function to create an Excel file with sample data
def create_excel_file(file_name):
    # Create a new workbook and select the active worksheet
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Sample Data"

    # Add column headers
    sheet["A1"] = "Name"
    sheet["B1"] = "Age"
    sheet["C1"] = "Occupation"
    sheet["D1"] = "Location"

    # Sample data to add to the sheet
    data = [
        ("John Doe", 28, "Software Developer", "USA"),
        ("Jane Smith", 34, "Data Scientist", "UK"),
        ("Carlos Diaz", 25, "Graphic Designer", "Mexico"),
        ("Anna Lee", 30, "Product Manager", "Canada")
    ]

    # Add sample data to the Excel file
    for row in data:
        sheet.append(row)

    # Save the Excel file
    wb.save(file_name)
    print(f"Excel file has been created and saved as: {file_name}")

# Main function
def main():
    # Specify the name of the Excel file
    file_name = "Sample_Excel_File.xlsx"

    # Call the function to create the file
    create_excel_file(file_name)

    # End message
    print("\nThank you for using this script!\n\nBest regards,\nLewenn\n\nFeel free to follow me on GitHub: https://github.com/lewennpy")

if __name__ == "__main__":
    main()
