# Excel Custom Functions for querying Companies House

This project adds custom functions to excel that allows for querying company data from the United Kingdom's Company House API.

![excel-custom-functions-video.gif](https://s1.gifyu.com/images/excel-custom-functions-video.gif)

## Current query functions

- `=COMPANY.ACCOUNTINGDATE` - Company's accounting date
- `=COMPANY.NAME` - Company's name
- `=COMPANY.REGADDRESS` - Company's registered address
- `=COMPANY.LIQUIDATED` - Company's liquidation status
- `=COMPANY.SICCODES` - Company's registered [SIC Codes](https://www.gov.uk/government/publications/standard-industrial-classification-of-economic-activities-sic)
- `=COMPANY.OVERDUESTATUS` - Status of company's accounts
- `=COMPANY.INCORPORATIONDATE` - Date of company's incorporation
- `=COMPANY.COMPANYSTATUS` - Company's current status
- `=COMPANY.DIRECTORS` - Company's directors' names
- `=COMPANY.SIGCONTROL` - Company's significant controlling persons' names
- `=COMPANY.LASTMEMBERSLIST` - Date of company's last members list submission

## To use the project

Follow these instructions to use this custom function sample add-in:

1. Publish the code files (HTML, JS, and JSON) to localhost.
2. Replace `http://127.0.0.1:8080` in the manifest file (there are 4 occurrences) with the URL you used, if needed (you might be using a different port number). 
3. Sideload the manifest (Companies-House.yml) using the instructions found at <https://aka.ms/sideload-addins>.
4. Test a custom function by entering `=COMPANY.NAME` and a company number (e.g. NI615002) in a cell.

