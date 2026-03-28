# Folder containing your XLSX files
$sourceFolder = "C:\Temp"

# Folder to save templates
$templateFolder = "C:\Temp\excel_templates"

if (-not (Test-Path $templateFolder)) {
    New-Item -ItemType Directory -Path $templateFolder | Out-Null
}

# Start Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Run in background
$excel.DisplayAlerts = $false  # Suppress prompts

# Get all XLSX files
$files = Get-ChildItem $sourceFolder -Filter *.xlsx

foreach ($file in $files) {

    Write-Host "Processing $($file.Name)..."

    # Open workbook WITHOUT updating links
    $workbook = $excel.Workbooks.Open($file.FullName, 0) # 0 = Don't update links

    # Build the template path
    $templatePath = Join-Path $templateFolder ($file.BaseName + ".xltx")

    # Save as Excel template
    $workbook.SaveAs([string]$templatePath, [int]54)  # 54 = xlTemplate

    # Close workbook
    $workbook.Close($false)
}

# Quit Excel
$excel.Quit()