# Automation-EDI-Clean-Up-Macro
Automated clean up of EDI reports, and export of a new workbook.
Built with VBA, in Excel.

A VBS script named "invoke.vbs" is executed by task scheduler, this script then opens "EdiAuto.xlam" and runs the macro "myMacro".
This macro first checks the inbox folder for a report then retrieves neccesary data from it. The script then moves to the invoice folder and retrieves data from the workbook located here. Finally, the macro opens a new workbook and places the retrieved data in appropiate cells. The new workbook is saved as [current-date]-report.xlsm in the outbox folder. 

Sample: 9-18-2018-report.xlsm. 

