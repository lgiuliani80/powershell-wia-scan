param(
    [string[]] $Files,
    [string] $OutputFile
)

if ($Files.Count -eq 0) {
    Write-Host "No files specified"
    exit
}

$ice = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ?{ $_.MimeType -eq "image/tiff" }
$im0 = [Drawing.Image]::FromFile($Files[0])
$enc = [Drawing.Imaging.Encoder]::SaveFlag
$eps = New-Object System.Drawing.Imaging.EncoderParameters(1)
$ep = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::MultiFrame)
$eps.Param[0] = $ep
$im0.Save($OutputFile, $ice, $eps)

for ($i = 1; $i -lt $Files.Count; $i++) {
    $ep = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::FrameDimensionPage)
    $eps.Param[0] = $ep
    $bmp = [Drawing.Bitmap]::FromFile($Files[$i])
    $im0.SaveAdd($bmp, $eps)
}

$ep = New-Object System.Drawing.Imaging.EncoderParameter($enc, [long][Drawing.Imaging.EncoderValue]::Flush)
$eps.Param[0] = $ep
$im0.SaveAdd($eps)
