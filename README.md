# Excel LAMBDA Function Editor VBA Package

A slightly more user-friendly Excel LAMBDA function editor compared to the "Name Manager".

For users who can’t install the [Advanced formula environment](https://www.microsoft.com/en-us/garage/profiles/excel-labs) (part of Excel Labs addin) because their Office environment is restricted by administrators.
See also https://github.com/alekrutkowski/ExcelLambdaTools for easy importing and exporting of multiple LAMBDA definitions from/to a text file like e.g. [this one](https://github.com/alekrutkowski/ExcelLambdaTools/blob/main/my_excel_lambda_functions.txt).

# Installation

1. Delete the previous `frmLambdaEditor`, `modLambdaEditor` and `modLambdaStore` VBA modules if they exist.
2. Import `modLambdaEditorInstaller.bas`.
3. Run `InstallLambdaEditor` macro.
4. Run `ShowLambdaEditor` macro.

You should be able to see and use something like:
<img width="2170" height="1440" alt="image" src="https://github.com/user-attachments/assets/41a326db-1483-4896-a136-a7d591d8ca63" />


# Notes

- Existing LAMBDA names in the workbook are preserved.
- See also potentially useful functions at https://gist.github.com/alekrutkowski/7847543aae6676269b300b8d40847fbe.
