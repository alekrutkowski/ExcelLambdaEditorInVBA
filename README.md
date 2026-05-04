# Excel LAMBDA Function Editor VBA Package

A more user-friendly and feature-rich Excel LAMBDA function editor compared to the "Name Manager".

For users who can’t install the [Advanced formula environment](https://www.microsoft.com/en-us/garage/profiles/excel-labs) (part of Excel Labs addin) because their Office environment is restricted by administrators.

# Installation

1. Delete the previous `frmLambdaEditor`, `modLambdaEditor` and `modLambdaStore` VBA modules and `UserForm1` if they exist.
2. Import `modLambdaEditorInstaller.bas`.
3. Run `InstallLambdaEditor` macro.
4. Run `ShowLambdaEditor` macro.

You should be able to see and use something like:
<img width="2530" height="1640" alt="image" src="https://github.com/user-attachments/assets/fd122db8-887e-459a-a33c-f813c3a997cf" />

# Notes

- Existing LAMBDA names in the workbook are preserved.
- See also potentially useful functions at https://github.com/alekrutkowski/ExcelLambdaTools/blob/main/my_excel_lambda_functions.txt and https://gist.github.com/alekrutkowski/7847543aae6676269b300b8d40847fbe.
