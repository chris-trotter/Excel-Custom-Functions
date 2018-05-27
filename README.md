# Excel Custom Functions for querying Companies House

This addin represents Excel Custom functions that are used to query Companies House for a company. 

![excel-custom-functions-video.gif](https://s1.gifyu.com/images/excel-custom-functions-video.gif)

## To use the project

Follow these instructions to use this custom function sample add-in:

1. Publish the code files (HTML, JS, and JSON) to localhost.
2. Replace `http://127.0.0.1:8080` in the manifest file (there are 4 occurrences) with the URL you used, if needed (you might be using a different port number). 
3. Sideload the manifest using the instructions found at <https://aka.ms/sideload-addins>.
4. Test a custom function by entering `=COMPANY.NAME` in a cell.
