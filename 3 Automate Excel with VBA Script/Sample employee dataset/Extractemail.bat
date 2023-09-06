@echo off
setlocal enabledelayedexpansion

:: Specify the name of the Excel file
set "excelFile=Employeedata.xlsx"

:: Specify the name of the output text file
set "outputFile=SalesTeamEmails.txt"

:: Specify the worksheet name in the Excel file
set "worksheet=Sheet1"

:: Start PowerShell to process the Excel file
powershell -command "Import-Module -Name 'ImportExcel'; $data = Import-Excel -Path '%~dp0%excelFile%' -WorksheetName '%worksheet%'; foreach ($row in $data) { if ($row.Department -eq 'Sales') { Add-Content -Path '%~dp0%outputFile%' -Value ('$($row.FirstName) $($row.LastName): $($row.Email)') } }"