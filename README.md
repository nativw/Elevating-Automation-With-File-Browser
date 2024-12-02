# Get Terminated Users Details
## Overview
Assuming we have an Excel file with details of employees who left the company and their accounts got terminated.
The structure of the Excel file is similar to this:

|EmployeeID|User Name  |Termination Date|
|----------|-----------|----------------|
| 837461   |john-t     |23-Dec-2024     |
| t72849 | paul-r | 11-Nov-2024 |
| b30293 | george-m | 10-Dec-2024 |
| 837493 | ringo-a | 23-Dec-2024 |

If I'm too curious to know the real display name, I can take the data in the `EmployeeID` column, and query Active Directory to return info such as Display Name, City, Country, etc.

This is where the script file **`Get-TerminatedUsersDetails.ps1`** comes in hand.
The script does the following steps:

 1. Opens the Windows Explorer, allowing you to navigate and select the Excel file in a graphical way.
 2.  Reads Employee IDs from the worksheet.
 3. Queries Active Directory for each Employee ID to retrieve user details.
 4. Outputs the results in a table with left-aligned columns.
