# Excel P2PDF (print to PDF)

## Overview

Excel P2PDF is a simple PowerShell tool that converts Excel files into PDFs. The converted PDFs will share the same name as the Excel files that they were converted from.

## Usage
- Place all Excel files in the `sheets` folder
- Open a PowerShell terminal as administrator
- Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` to allow the script to execute in PowerShell
- Run `.\excel-p2pdf.ps1` to execute the script
- Converted PDFs will be output into the `pdf` folder
