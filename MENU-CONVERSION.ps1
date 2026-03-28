Add-Type -AssemblyName PresentationFramework

# Path to your script folder
$scriptFolder = "\\vcn.ds.volvo.net\it-got\proj03\015165\Upload\!Leandro\FileConvertionScripts"

$xlsxScript = Join-Path $scriptFolder "xlsxToXltx.ps1"
$docScript  = Join-Path $scriptFolder "docsToDotx.ps1"
$pptxScript = Join-Path $scriptFolder "pptxToPotx.ps1"

# Window
$window = New-Object System.Windows.Window
$window.Title = "File Conversion Tool"
$window.Width = 420
$window.Height = 300
$window.WindowStartupLocation = "CenterScreen"
$window.Background = "#1e1e1e"

# Disable minimize and maximize (keep only X)
$window.ResizeMode = "NoResize"

# Layout container
$stack = New-Object System.Windows.Controls.StackPanel
$stack.HorizontalAlignment = "Center"
$stack.VerticalAlignment = "Center"
$stack.Margin = 20

# Title
$title = New-Object System.Windows.Controls.TextBlock
$title.Text = "File Conversion Tool"
$title.FontSize = 22
$title.Foreground = "White"
$title.Margin = "0,0,0,20"
$title.HorizontalAlignment = "Center"

$stack.Children.Add($title)

function Add-ConversionButton {

    param(
        [string]$label,
        [string]$scriptPath,
        [string]$color
    )

    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $label
    $btn.Width = 240
    $btn.Height = 45
    $btn.Margin = "0,8"
    $btn.FontSize = 14
    $btn.Foreground = "White"
    $btn.Background = $color
    $btn.BorderThickness = 0
    $btn.Cursor = "Hand"

    $btn.Tag = $scriptPath

    # Hover effect
    $btn.Add_MouseEnter({ $this.Opacity = 0.85 })
    $btn.Add_MouseLeave({ $this.Opacity = 1 })

    $btn.Add_Click({

        $sp = $this.Tag

        if (-not (Test-Path $sp)) {
            [System.Windows.MessageBox]::Show("Script not found:`n$sp","Error")
            return
        }

        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -Command `"& '$sp'`""
    })

    return $btn
}

# Buttons
$stack.Children.Add((Add-ConversionButton "Excel Template (xlsx to xltx)" $xlsxScript "#0078D7"))
$stack.Children.Add((Add-ConversionButton "Word Template (doc to dotx)" $docScript "#2B579A"))
$stack.Children.Add((Add-ConversionButton "PowerPoint Template (pptx to potx)" $pptxScript "#D24726"))

$window.Content = $stack
$window.ShowDialog()