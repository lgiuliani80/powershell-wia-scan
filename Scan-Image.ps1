param(
    [Parameter(Mandatory = $true)][ValidateSet("Color", "Grayscale", "Text")][string]$TypeOfScan,
    [Parameter(Mandatory = $true)][ValidateSet("Bmp", "Png", "Tiff", "Jpeg")] [string]$FileFormat,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [Parameter(Mandatory = $false)][int]$ResolutionDpi = 200
)

$ErrorActionPreference = "Stop"

$intentMap = @{
    "Color" = 1
    "Grayscale" = 2
    "Text" = 4
}

Add-Type -AssemblyName System.Drawing
$formatGUID = ([System.Drawing.Imaging.ImageFormat]($FileFormat)).Guid.ToString("b").ToUpper()

$wiaDevMgr = New-Object -ComObject WIA.DeviceManager
$wiaDialogs = New-Object -ComObject WIA.CommonDialog

# Assuming to use the first scanner found
$firstScanner = $wiaDevMgr.DeviceInfos | Where-Object { $_.Type -eq 1 } | Select-Object -First 1
$device = $firstScanner.Connect()

# Assuming to scan from the first item in the scanner
$item = $device.Items.Item(1)

# $item.Properties | ft

# Set the scan intent
$item.Properties.Item("6146").Value = $intentMap[$TypeOfScan]

# Set Horizontal and Vertical resolution
$item.Properties.Item("6147").Value = $ResolutionDpi
$item.Properties.Item("6148").Value = $ResolutionDpi

# Set Horizontal and Vertical extent (A4 sheet size)
$item.Properties.Item("6151").Value = [int] (21 / 2.54 * $ResolutionDpi)   ## Horizontal extent
$item.Properties.Item("6152").Value = [int] (29.5 / 2.54 * $ResolutionDpi) ## Vertical extent

# Set multipage file option
#TODO: $item.Properties.Item("4108").Value = 256 # TYMED_MULTIPAGE_FILE

$scannedImage = $wiaDialogs.ShowTransfer($item, $formatGUID, $true)
$scannedImage

if ($scannedImage.FormatID -ne $formatGUID) {
    # Convert to the expected format
    Write-Output "Converting the scanned image to $FileFormat format ..."
    $imageProcess = New-Object -ComObject WIA.ImageProcess
    $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
    # Only working with string constants! Ugly but working.
    switch ($FileFormat) {
        "Bmp"  { $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}" }
        "Jpeg" { $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" }
        "Png"  { $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}" }
        "Tiff" { $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}" }
    }
    $scannedImage = $imageProcess.Apply($scannedImage)
}

$imageFileName = Join-Path -Path $OutputDirectory -ChildPath ("ScannedImage.{0:yyMMdd\THHmmss}.{1}.{2}" -f [DateTime]::Now, [Guid]::NewGuid().ToString().Substring(0,5), $FileFormat.ToLower())

Write-Output "Saving the scanned image to $imageFileName ..."
$scannedImage
$scannedImage.SaveFile($imageFileName)