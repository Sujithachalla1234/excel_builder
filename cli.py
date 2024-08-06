import argparse
import aspose.cells as ac
import os
import win32com.client as win32

def create_or_update_workbook(file_name, file_format, sheet_index=0):
    """
    Create a new Excel workbook or add worksheets to an existing workbook, and save the workbook with the specified format.

    :param file_name: Name of the file to save or open the workbook (without extension)
    :param file_format: File format (xlsx, xls, xlsm)
    :param sheet_index: Index of the worksheet to add (default is 0 for the first sheet)
    """
    # Construct the full file name with extension
    full_file_name = f"{file_name}.{file_format}"
    
    # Check if the file already exists
    if os.path.exists(full_file_name):
        # Load the existing workbook
        workbook = ac.Workbook(full_file_name)
    else:
        # Create a new workbook
        workbook = ac.Workbook()

    # Ensure the worksheet index is valid
    if sheet_index >= len(workbook.worksheets):
        # Add additional worksheets if necessary
        for _ in range(sheet_index - len(workbook.worksheets) + 1):
            workbook.worksheets.add()

    # Determine the save format based on the file format
    if file_format == 'xlsx':
        save_format = ac.SaveFormat.XLSX
    elif file_format == 'xls':
        save_format = ac.SaveFormat.EXCEL_97_TO_2003
    elif file_format == 'xlsm':
        save_format = ac.SaveFormat.XLSM
    else:
        raise ValueError(f"Unsupported file format: {file_format}")
    
    # Save the workbook
    workbook.save(full_file_name, save_format)

def add_userform_to_workbook(file_name, file_format, userform_name="MyForm"):
    """
    Add a VBA UserForm to an existing Excel workbook.

    :param file_name: Name of the file to open (without extension)
    :param file_format: File format (xlsx, xls, xlsm)
    :param userform_name: Name of the UserForm to add
    """
    # Construct the full file name with extension
    full_file_name = f"{file_name}.{file_format}"
    
    # Check if the file exists
    if not os.path.exists(full_file_name):
        raise FileNotFoundError(f"The file {full_file_name} does not exist.")
    
    # Open the workbook using Excel COM
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(os.path.abspath(full_file_name))
    vba_project = workbook.VBProject
    
    # Add a new UserForm
    vb_component = vba_project.VBComponents.Add(1)  # 1 indicates a UserForm
    vb_component.Name = userform_name
    
    # Save and close the workbook
    workbook.Save()
    workbook.Close(SaveChanges=True)
    excel.Application.Quit()
    
    print(f'UserForm {userform_name} added to {full_file_name} successfully.')

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Create or update an Excel workbook with specified name and format.')
    parser.add_argument('file_name', type=str, help='Name of the Excel file to create or modify (without extension)')
    parser.add_argument('file_format', type=str, choices=['xlsx', 'xls', 'xlsm'], help='File format (xlsx, xls, xlsm)')
    parser.add_argument('--sheet_index', type=int, default=0, help='Index of the worksheet to add (default is 0 for the first sheet)')
    parser.add_argument('--add_userform', action='store_true', help='Add a UserForm to the workbook')
    parser.add_argument('--userform_name', type=str, default="MyForm", help='Name of the UserForm to add')
    
    args = parser.parse_args()
    
    # Create or update the workbook
    create_or_update_workbook(
        file_name=args.file_name,
        file_format=args.file_format,
        sheet_index=args.sheet_index
    )
    
    if args.add_userform:
        add_userform_to_workbook(
            file_name=args.file_name,
            file_format=args.file_format,
            userform_name=args.userform_name
        )
    
    print(f'Workbook {args.file_name}.{args.file_format} created or updated and saved successfully.')

if __name__ == '__main__':
    main()
