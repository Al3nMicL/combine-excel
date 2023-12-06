# Combine Excel

#### Readme by Bing GPT
This PowerShell script appears to be merging multiple Excel files into one. Here's a breakdown of what it does:
- It sets the name of the output file as **merged.xlsx** and the export folder as **export**.
- It checks if the export folder exists in the script's root directory. If it does, it proceeds with the process. If it doesn't, it creates the folder.
- It then sets the full path of the export file.
- It gets all the Excel files (.xlsx) in the script's root directory.
- For each Excel file, it does the following:
    - It gets the filename without the extension and splits it by space to get the worksheet name.
    - It imports the Excel file and exports it to the output file with the worksheet name set to the first part of the filename.
    - 
*Note*: This script uses the `Import-Excel` and `Export-Excel` cmdlets, which are not built-in PowerShell cmdlets. They are part of the `ImportExcel` PowerShell module, which needs to be installed separately. If you haven’t installed this module yet, you can do so by running `Install-Module -Name ImportExcel` in your PowerShell console. Please note that you need to have administrator rights to install modules.


> Written with [StackEdit](https://stackedit.io/).
