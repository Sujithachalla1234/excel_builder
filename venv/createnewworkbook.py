import aspose.cells as ac

def create_and_save_workbook(file_name, file_format, cell_value, sheet_index=0, row=0, column=0):
    """
    Create a new Excel workbook, set a value in a specified cell, and save the workbook with the specified format.

    :param file_name: Name of the file to save the workbook (without extension)
    :param file_format: File format (xlsx, xls, xlsm)
    :param cell_value: Value to set in the specified cell
    :param sheet_index: Index of the worksheet to modify (default is 0 for the first sheet)
    :param row: Row index of the cell to modify (default is 0)
    :param column: Column index of the cell to modify (default is 0)
    """
    # Create a new workbook
    workbook = ac.Workbook()
    
    # Access the specified worksheet
    worksheet = workbook.worksheets[sheet_index]
    
    # Set the value in the specified cell
    worksheet.cells.get(row, column).put_value(cell_value)
    
    # Determine the save format based on the file format
    if file_format == 'xlsx':
        save_format = ac.SaveFormat.XLSX
    elif file_format == 'xls':
        save_format = ac.SaveFormat.EXCEL_97_TO_2003
    elif file_format == 'xlsm':
        save_format = ac.SaveFormat.XLSM
    else:
        raise ValueError(f"Unsupported file format: {file_format}")
    
    # Construct the full file name with extension
    full_file_name = f"{file_name}.{file_format}"
    
    # Save the workbook
    workbook.save(full_file_name, save_format)

def get_file_name():
    return input("Enter the name of the Excel file to create (without extension): ").strip()

def get_file_format():
    formats = ['xlsx', 'xls', 'xlsm']
    format_choice = input(f"Enter the file format ({formats}): ").strip().lower()
    while format_choice not in formats:
        print("Invalid format. Please choose from xlsx, xls, xlsm.")
        format_choice = input(f"Enter the file format ({formats}): ").strip().lower()
    return format_choice

def main():
    file_name = get_file_name()
    file_format = get_file_format()
    
    cell_value = "Hello, Aspose!"  # You can customize or prompt for this value if needed
    
    create_and_save_workbook(file_name, file_format, cell_value)
    
    print(f'Workbook {file_name}.{file_format} created and saved successfully.')

if __name__ == '__main__':
    main()
