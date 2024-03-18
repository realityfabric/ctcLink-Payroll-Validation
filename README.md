# ctcLink-Payroll-Validation
Payroll Validation macros for CBC HR.

This Macro is intended to be included in a Macro Enabled Excel Workbook with macros assigned to buttons for ease of use.

The workbook must be in the same directory as any Payroll workbooks which are to be analyzed. The ctcLink query `QHC_PY_PAY_CHECK_OTH_EARNS` should be run for the desired Pay Period and saved to Excel as `QHC_PY_PAY_CHECK_OTH_EARNS.xlsx`.

## Workbook Instructions
1. Save this workbook to a folder containing the Payroll workbooks to be processed.
     - Only those workbooks and a single workbook containing the output of QHC_PY_PAY_CHECK_OTH_EARNS should be in this folder.
2. Close all other Excel files.
3. Click "Refresh Data". Other workbooks will open and close during this process, and it may take a while to complete.
   - If a workbook does not contain an Appointed sheet or an Hourly sheet then a prompt will appear asking if you would like to continue.
   - If the workbook does not contain that sheet, click OK.
   - If the workbook does contain that sheet, click Cancel, verify that the sheet is named correctly, and then return to Step 2.
5. Click "Remove Canceled Classes". This should run quickly.
6. Click "Generate Employee/Job Code List". This should run quickly.
7. Click "Generate Payroll Summary". This should run quickly.

![Payroll Validation Macro Enabled Workbook, Version 3, Instructions](https://github.com/realityfabric/ctcLink-Payroll-Validation/assets/8910652/9bb8884e-e295-4ed7-97e0-105a16ab3ddc)
