# Export-EMLToWord
PowerShell Script to Export EML file content to Word Documents

## Description
This script processes EML files from a specified source directory and converts them into Word documents. It extracts the HTML body and embedded images from the EML files and inserts them into the Word documents.

## Requirements
- Microsoft Word installed
- CDO (Collaboration Data Objects) installed
- `Convert-EmlFile` function from [PowerShell-Functions](https://github.com/PsCustomObject/PowerShell-Functions/blob/master/Convert-EmlFile.ps1)

## Parameters
- `SourceDirectory`: The directory containing the EML files to be processed.
- `OutputDirectory`: The directory where the Word documents will be saved (default is `c:\temp\`).

## Example
```powershell
# Load the Convert-EmlFile function
. .\Convert-EmlFile.ps1

# Run the Export-EmlToWord function
Export-EmlToWord -SourceDirectory "C:\path\to\eml\files" -OutputDirectory "C:\path\to\output\directory"