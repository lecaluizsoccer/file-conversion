# Folder containing your DOCX files
$sourceFolder = "C:\Temp"

# Folder to save templates
$templateFolder = "C:\Temp\word_templates"

# Create the templates folder if it doesn't exist
if (-not (Test-Path $templateFolder)) {
    New-Item -ItemType Directory -Path $templateFolder | Out-Null
}

# Start Word COM object
$word = New-Object -ComObject Word.Application
$word.Visible = $false  # Run in background

# Get all DOCX files
$files = Get-ChildItem $sourceFolder -Filter *.docx

foreach ($file in $files) {

    Write-Host "Processing $($file.Name)..."

    # Open the DOCX file
    $doc = $word.Documents.Open($file.FullName)

    # Build the template path in the word_templates folder
    $templatePath = Join-Path $templateFolder ($file.BaseName + ".dotx")

    # Save as Word template (.dotx)
    $doc.SaveAs([string]$templatePath, [int]16)  # 16 = wdFormatDocumentDefaultTemplate

    # Close the document
    $doc.Close()
}

# Quit Word
$word.Quit()