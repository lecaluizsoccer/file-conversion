# === CONFIG ===
$sourceFolder = "C:\Temp"
$templateFolder = "C:\Temp\powerpoint_templates"

if (-not (Test-Path $templateFolder)) {
    New-Item -ItemType Directory -Path $templateFolder | Out-Null
}

$powerpoint = New-Object -ComObject PowerPoint.Application
# $powerpoint.Visible = $true      

$ppSaveAsOpenXMLTemplate = 35

$files = Get-ChildItem $sourceFolder -Filter *.pptx

foreach ($file in $files) {
    Write-Host "Processing $($file.Name)..."

    try {
        $presentation = $powerpoint.Presentations.Open(
            $file.FullName,
            $false,
            $false,
            $true
        )

        # ⭐ CRITICAL FIX: Use SaveCopyAs instead of SaveAs
        $potxPath = Join-Path $templateFolder ($file.BaseName + ".potx")
        $presentation.SaveCopyAs($potxPath, $ppSaveAsOpenXMLTemplate)

        Write-Host "Saved POTX (slides preserved): $potxPath"
        $presentation.Close()
    }
    catch {
        Write-Warning "Error processing ${file.FullName}: $_"
    }
}

$powerpoint.Quit()
Write-Host "All PPTX files converted with slides intact."