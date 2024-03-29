param(
    [Parameter(Mandatory = $true)][ValidateSet("Color", "Grayscale", "Text", "Select")][string]$TypeOfScan,
    [Parameter(Mandatory = $true)][ValidateSet("Bmp", "Png", "Tiff", "Jpeg")] [string]$FileFormat,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [Parameter(Mandatory = $false)][int]$ResolutionDpi = 200,
    [switch]$UseFeeder,
    [switch]$MinimizeConsole,
    [switch]$EnableMultipageTiff
)

Function MinimizeConsoleWindow {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public static class Win32Apis
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
    }
"@ -Language CSharp -PassThru | Out-Null
    $consoleWindow = [Win32Apis]::GetConsoleWindow()
    $SW_MINIMIZE = 6
    [Win32Apis]::ShowWindow($consoleWindow, $SW_MINIMIZE)
}

Function MergeImagesToMultipageTiff {
    param(
        [string[]] $Files,
        [string] $OutputFile
    )
    
    if ($Files.Count -eq 0) {
        return
    }
    
    $ice = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ?{ $_.MimeType -eq "image/tiff" }
    $im0 = [Drawing.Image]::FromFile($Files[0])
    $enc = [Drawing.Imaging.Encoder]::SaveFlag
    $eps = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $eps.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::MultiFrame)
    $im0.Save($OutputFile, $ice, $eps)
    
    for ($i = 1; $i -lt $Files.Count; $i++) {
        $eps.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::FrameDimensionPage)
        $bmp = [Drawing.Bitmap]::FromFile($Files[$i])
        $im0.SaveAdd($bmp, $eps)
    }
    
    $eps.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::Flush)
    $im0.SaveAdd($eps)
}

$FEEDER = 1
$FEED_READY = 1

$intentMap = @{
    "Color" = 1
    "Grayscale" = 2
    "Text" = 4
}

if ($MinimizeConsole) {
    MinimizeConsoleWindow
}

$OutDir = $OutputDirectory

if ($EnableMultipageTiff) {
    $OutDir = [IO.Path]::GetTempPath()
}

Add-Type -AssemblyName System.Drawing
$formatGUID = ([System.Drawing.Imaging.ImageFormat]($FileFormat)).Guid.ToString("b").ToUpper()

$wiaDevMgr = New-Object -ComObject WIA.DeviceManager
$wiaDialogs = New-Object -ComObject WIA.CommonDialog

# Assuming to use the first scanner found
$firstScanner = $wiaDevMgr.DeviceInfos | Where-Object { $_.Type -eq 1 } | Select-Object -First 1
$device = $firstScanner.Connect()

if ($UseFeeder) {
    $device.Properties.Item("3088").Value = 1
    $device.Properties.Item("3096").Value = 1   # Number of pages to scan
}

if ($TypeOfScan -eq "Select") {
    $items = $wiaDialogs.ShowSelectItems($device)
    $item = $items[1]
    $UseFeeder = ($device.Properties.Item("3088").Value -eq 1)
} else {
    # Assuming to scan from the first item in the scanner
    $item = $device.Items.Item(1)

    # Set the scan intent
    $item.Properties.Item("6146").Value = $intentMap[$TypeOfScan]

    # Set Horizontal and Vertical resolution
    $item.Properties.Item("6147").Value = $ResolutionDpi
    $item.Properties.Item("6148").Value = $ResolutionDpi

    # Set Horizontal and Vertical extent (A4 sheet size)
    $item.Properties.Item("6151").Value = [int] (21 / 2.54 * $ResolutionDpi)   ## Horizontal extent
    $item.Properties.Item("6152").Value = [int] (29.5 / 2.54 * $ResolutionDpi) ## Vertical extent
}

#$device.Properties | ft
#$item.Properties | ft

$morePages = $true
$pageNo = 1
$createdFiles = @()

while ($morePages) {

    try {
        Write-Output "[PAGE #$pageNo] Acquiring..."
        $scannedImage = $wiaDialogs.ShowTransfer($item, $formatGUID, $true)
        #$scannedImage
        Write-Output "[PAGE #$pageNo] Scanned image: $($scannedImage.FileExtension) $($scannedImage.Width)x$($scannedImage.Height) $($scannedImage.PixelDepth) bpp [$($scannedImage.HorizontalResolution.ToString("N0"))x$($scannedImage.VerticalResolution.ToString("N0")) dpi]"

        if ($scannedImage.FormatID -ne $formatGUID) {
            # Convert to the expected format
            Write-Output "[PAGE #$pageNo] Converting the scanned image to $FileFormat format ..."
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

        $imageFileName = Join-Path -Path $OutDir -ChildPath ("ScannedImage.{0:yyMMdd\THHmmss}.{1}.{2}.{3}" -f [DateTime]::Now, $pageNo, [Guid]::NewGuid().ToString().Substring(0,5), $FileFormat.ToLower())

        Write-Output "[PAGE #$pageNo] Saving the scanned image to $imageFileName ..."
        $scannedImage.SaveFile($imageFileName)
        $scannedImage = $null
        
        $createdFiles += $imageFileName
        $pageNo++

    } catch {
        if ($_.Exception.HResult -eq 0x80210003) {
            Write-Output "No more pages to scan."
            break
        } else {
            Write-Error $_
            break
        }
    } finally {
        $morePages = $UseFeeder -and ($null -ne $device.Properties.Item("3088")) -and `
                     (($device.Properties.Item("3088").Value -band $FEEDER) -ne 0) -and `
                     ($null -ne $device.Properties.Item("3087")) -and `
                     (($device.Properties.Item("3087").Value -band $FEED_READY) -ne 0)
    }
}

if ($EnableMultipageTiff -and $createdFiles.Count -gt 0) {
    $multiPageTiffFileName = Join-Path -Path $OutputDirectory -ChildPath ("ScannedPages.{0:yyMMdd\THHmmss}.tiff" -f [DateTime]::Now)
    Write-Output "Generating merged multipage tiff to $multiPageTiffFileName ..."
    MergeImagesToMultipageTiff -Files $createdFiles -OutputFile $multiPageTiffFileName
    #$createdFiles | ForEach-Object { Remove-Item -Path $_ -Force }
}
