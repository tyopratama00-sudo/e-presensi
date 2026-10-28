Add-Type -AssemblyName System.Drawing
$filePath = "C:\Users\Tio\.gemini\antigravity-ide\brain\abfd8bed-9e05-4c34-8ced-f0a7b0b1713e\.user_uploaded\media_1793161913397.jpg"
$bmp = New-Object System.Drawing.Bitmap($filePath)

$minX = $bmp.Width
$maxX = 0
$minY = $bmp.Height
$maxY = 0

for ($y = [int]($bmp.Height * 0.15); $y -lt $bmp.Height; $y++) {
    for ($x = 0; $x -lt $bmp.Width; $x++) {
        $pixel = $bmp.GetPixel($x, $y)
        if ($pixel.R -lt 200 -and $pixel.G -lt 200 -and $pixel.B -lt 200) {
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
        }
    }
}

$pad = 16
$cropX = [Math]::Max(0, $minX - $pad)
$cropY = [Math]::Max(0, $minY - $pad)
$cropW = [Math]::Min($bmp.Width - $cropX, ($maxX - $minX) + ($pad * 2))
$cropH = [Math]::Min($bmp.Height - $cropY, ($maxY - $minY) + ($pad * 2))

$rect = New-Object System.Drawing.Rectangle($cropX, $cropY, $cropW, $cropH)
$cropped = $bmp.Clone($rect, $bmp.PixelFormat)

# Resize to crisp web display width (e.g. width 480, height proportional)
$targetW = 480
$targetH = [int]($cropH * ($targetW / $cropW))
$resized = New-Object System.Drawing.Bitmap($targetW, $targetH)
$g = [System.Drawing.Graphics]::FromImage($resized)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.DrawImage($cropped, 0, 0, $targetW, $targetH)
$g.Dispose()

$outputPath = "c:\Users\Tio\Project\Absensi\apps-script\logo.png"
$resized.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)

$bmp.Dispose()
$cropped.Dispose()
$resized.Dispose()
Write-Host "Done: $targetW x $targetH"
