<#
.SYNOPSIS
    This script reads Employee IDs from an Excel file, queries Active Directory for user details, and outputs the results in a left-aligned table.

.DESCRIPTION
    The script performs the following steps:
    1. Opens the Windows Explorer, to select an Excel file.
    2. Reads Employee IDs from the worksheet.
    3. Queries Active Directory for each Employee ID to retrieve user details.
    4. Outputs the results in a table with left-aligned columns.

.NOTES
    Author: Nativ Weiss
    Date: 20-August-2024
    Version: 1.0

.EXAMPLE 
    Run the script and use the dialog box to select the Excel file.
#>

Add-Type -AssemblyName System.Windows.Forms

# Create and configure the OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$openFileDialog.FilterIndex = 1
$openFileDialog.Multiselect = $false

# Show the dialog and get the selected file
$dialogResult = $openFileDialog.ShowDialog()
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $ExcelFile = $openFileDialog.FileName
}
else {
    Write-Host "No file selected. Exiting script."
    exit
}

# Open Excel and load the workbook
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($ExcelFile)

# Select the worksheet
$sheet = $workbook.Sheets.Item("Sheet1")

# Get the number of rows
$rowCount = $sheet.UsedRange.Rows.Count

# Initialize an array to hold the results
$results = @()

# Loop through rows to read EmployeeID and query AD
for ($row = 2; $row -le $rowCount; $row++) {
    # Assuming the first row is headers
    $employeeID = $sheet.Cells.Item($row, 1).Value()
    
    if ($employeeID) {
        # Query Active Directory
        $user = Get-ADUser -Filter "EmployeeID -eq '$employeeID'" -Properties mailNickName, DisplayName, City, co
        
        if ($user) {
            # Add the result to the array
            $results += [PSCustomObject]@{
                EmployeeID = $employeeID
                UserName   = $user.mailNickName
                FullName   = $user.DisplayName
                City       = $user.City
                Country    = $user.co
            }
        }
        else {
            # Add a result indicating the user was not found
            $results += [PSCustomObject]@{
                EmployeeID = $employeeID
                UserName   = "Not found"
                FullName   = "Not found"
                City       = "Not Found"
                Country    = "Not Found"
            }
        }
    }
}

# Close the workbook and quit Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Output the results as a table with left-aligned columns
$results | Format-Table -Property @{Expression = "EmployeeID"; Alignment = "Left" },
@{Expression = "UserNAme"; Alignment = "Left" },
@{Expression = "FullName"; Alignment = "Left" },
@{Expression = "City"; Alignment = "Left" },
@{Expression = "Country"; Alignment = "Left" } -AutoSize
