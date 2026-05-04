# Excel LAMBDA Function Editor VBA Package

**A more user-friendly and feature-rich Excel [LAMBDA](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) function editor compared to the "Name Manager".**

For users who can’t install the [Advanced formula environment](https://www.microsoft.com/en-us/garage/profiles/excel-labs) (part of Excel Labs addin) because their Office environment is restricted by administrators.

See https://www.xlwings.org/blog/excel-lambda-functions for the discussion of the benefits of LAMBDAs.

# Installation

0. *Delete the previous `frmLambdaEditor`, `modLambdaEditor` and `modLambdaStore` VBA modules and `UserForm1` if they exist.*
1. Download [`modLambdaEditorInstaller.bas`](https://raw.githubusercontent.com/alekrutkowski/ExcelLambdaEditorInVBA/refs/heads/main/modLambdaEditorInstaller.bas) and import it in Excel (see steps 1-4 from [here](https://support.tetcos.com/support/solutions/articles/14000143233-how-to-import-vba-script-bas-file-in-ms-excel-)).
2. Run **once** `InstallLambdaEditor` macro (see [this](https://www.geeksforgeeks.org/excel/how-to-run-a-macro-in-excel/) if you don't know how to run Excel macros).
3. Run `ShowLambdaEditor` macro any time you want to manage or edit your lambdas. I suggest you add this macro to your "Quick Access Toolbar" (see [this](https://support.microsoft.com/en-gb/office/assign-a-macro-to-a-button-728c83ec-61d0-40bd-b6ba-927f84eb5d2c) for instructions).

You should be able to see and use something like:
<img width="2530" height="1640" alt="image" src="https://github.com/user-attachments/assets/fd122db8-887e-459a-a33c-f813c3a997cf" />

# Notes

- Existing LAMBDA names in the workbook are preserved.
- The <kbd>Visualize</kbd> button opens a formulaboost.com's web page with the definition of the current lambda function properly syntax-highlighted and indented for easier analysis and editing. E.g. for the `setdiff` function from the screenshot above it opens [this](https://www.formulaboost.com/parse?f==LAMBDA(Range1,Range2,%20UNIQUE(FILTER(Range1,%20ISNA(MATCH(Range1,%20Range2,%200)))))) page.
- See also potentially useful functions at https://github.com/alekrutkowski/ExcelLambdaTools/blob/main/my_excel_lambda_functions.txt and https://gist.github.com/alekrutkowski/7847543aae6676269b300b8d40847fbe.
