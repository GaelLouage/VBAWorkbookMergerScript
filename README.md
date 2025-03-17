VBA Workbook Data Copier

This VBA script helps automate the process of copying data between multiple workbooks in Excel. The script allows you to select two source files and a target workbook. It then copies specific ranges from the source files and pastes them into the target workbook, preserving only the values. A new worksheet is created in the target workbook for each data set.
Features

    Allows the user to select two source workbooks via file picker dialog.
    Creates a new worksheet in the target workbook.
    Copies data from predefined ranges in the source workbooks to the target workbook.
    Ensures that only values (not formulas or formats) are pasted into the target.
    Saves the target workbook after data is copied.

Prerequisites

    Microsoft Excel with Macro (VBA) support enabled.
    The source workbooks should contain the data you want to copy.
    The target workbook should already exist and be specified in the code.

Code Overview
Functions

    GetFilePath()
        This function opens a file dialog to let the user select a file.
        It returns the selected file's path if a file is chosen, otherwise, it cancels the operation.

    AddSheet()
        This function adds a new worksheet to the active workbook and names it "Sheet2."
        Returns the created worksheet object.

    CopyDataBetweenWorkbooks()
        This is the main subroutine that orchestrates the entire process.
        It:
            Opens two source workbooks and their respective sheets.
            Opens a target workbook and its sheets.
            Copies data from the source sheets (from ranges A1:D100).
            Pastes values into the target workbook, preserving only the data (not formulas).
            Saves the target workbook and closes the source workbooks.
            Cleans up the objects and displays a success message.

File Paths

    Source Workbook 1: You will be prompted to select the first source workbook (contains data for Sheet1).
    Source Workbook 2: You will be prompted to select the second source workbook (contains data for Sheet2).
    Target Workbook: The target workbook is hardcoded to C:\Users\louag\OneDrive\Bureau\Book1.xlsm but can be modified in the code to any other path.

How to Use

    Open Excel and press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
    Insert a new module by selecting Insert > Module.
    Paste the code into the new module.
    Close the VBA editor.
    Run the CopyDataBetweenWorkbooks macro (press Alt + F8, select CopyDataBetweenWorkbooks, and click Run).
    The file picker will prompt you to select two source workbooks.
    The script will copy data from the selected source workbooks into the specified target workbook, creating a new worksheet for each data set.

Important Notes

    The code assumes that the source workbooks contain a sheet named "Sheet1." If your sheets have different names, you will need to modify the script.
    The target workbook path is hardcoded. You can modify it as needed to match your file's location.
