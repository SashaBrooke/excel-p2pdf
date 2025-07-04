<#
.SYNOPSIS
    Converts Excel sheets to PDF format.

.DESCRIPTION
    This script uses the Excel COM object to open each .xlsx file in the 'sheets' folder
    and export the entire workbook to a PDF. Output files are saved in the 'pdf' folder.

.NOTES
    Author  : Sasha Brooke
    Created : 2025-07-04
    Version : 1.0

.EXAMPLE
    Run from PowerShell:
        PS> .\excel-p2pdf.ps1
#>

# Setup inputs and outputs
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$excelFolder = Join-Path $scriptDir "sheets"
$outputFolder = Join-Path $scriptDir "pdf"

# Create output folder if it doesn't exist
if (!(Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Start Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Loop through all Excel files
Get-ChildItem -Path $excelFolder -Filter *.xlsx | ForEach-Object {
    $workbookPath = $_.FullName
    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    $outputPath = Join-Path $outputFolder "$fileNameWithoutExt.pdf"

    try {
        # Open workbook
        $workbook = $excel.Workbooks.Open($workbookPath)

        # Export as PDF (entire workbook)
        $workbook.ExportAsFixedFormat(
            [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF,
            $outputPath
        )

        # Close workbook
        $workbook.Close($false)
        Write-Host "Exported $($_.Name) to PDF"
    } catch {
        Write-Warning "Failed to process $($_.Name): $_"
    }
}

# Quit Excel
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
