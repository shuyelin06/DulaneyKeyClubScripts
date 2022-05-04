# New Point System
This is the new point / event system as of the end of the '22 - '23 Service Year. The scripts and file structure will be saved here in case they need to be restored in the future.

## Set-Up
To restore this system, do the following:
1) Download all files
2) Import all structure files and folders (keeping the same names). Reformat each spreadsheet as necessary to match the excel format (delete blank columns as necessary). Remove the ".txt" files, as they're just present to tell you the purpose of each folder.
> Forms (represented as PDFs) will manually need to be created and linked with their respective spreadsheets. Link the following and reformat the response sheet to the "Form Responses" sheet in the imported spreadsheets:

| Form Name  | Spreadsheet Name |
| ------------- | ------------- |
| [Service Year] Membership Registration | [Service Year] Membership Information |
| Join the [Service Year] Board! | Board Members |

3) Open the Officer Center spreadsheet. Then, go to Extensions --> Apps Script. Copy and paste every ".gs" file in the Scripts folder, to the opened Scripts Project.
4) For each button in the Officer Center spreadsheet, click --> triple dots --> "Assign Script". Link each button to the following function names (copy and paste the names):

| Button Title  | Script Name |
| ------------- | ------------- |
| New Service Year | newServiceYear |
| Log Write-Up | logWriteup |
| Sync Member Information | syncInfo |
| Sort Data (Manual) | sortData |
| Create New Event | newEvent |
| New Member (Manual) | newMember |

5) Repeat steps 3-4 with the "Add New Volunteer" button in the "[Sample Event] Write-Up" spreadsheet. Copy the script "macro.gs" into the spreadsheet's Apps Script project, and assign the button to function name "NewVolunteer".
6) Go to the "Connections" sheet in the Officer Center spreadsheet. For every filled cell in column B, fill the adjacent column C cell with the corresponding ID.
> To find the ID of files and folders, open them, and check their URL. The ID is the long string of text (between two "/") in the URL.

The system should now be appropriately set up.

## How to Use
To use the system, simply press the corresponding buttons in the officer center to meet your needs.

The new service year and new event folders are based off templates - all files in the  templates will be copied, and any "[Sample Event]" or "[Service Year]" strings in the names will be replaced with the corresponding user input. You may modify the templates to meet your need.
> Note: The sample event script will not copy linked forms, whereas the service year script will.

#### Wishing you the best in service!
