# Builds a multi-size .ico from a .png. PNG-compressed entries (Vista+ ICO), so no BMP/AND-mask
# fiddling and the 256px entry stays small.
param([string]$In, [string]$Out)

Add-Type -AssemblyName System.Drawing

$sizes = 16, 24, 32, 48, 64, 128, 256
$src = [System.Drawing.Image]::FromFile((Resolve-Path $In))

$blobs = foreach ($s in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.PixelOffsetMode = 'HighQuality'
    $g.DrawImage($src, 0, 0, $s, $s)
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    , $ms.ToArray()
}
$src.Dispose()

$fs = [System.IO.File]::Create((Join-Path (Get-Location) $Out))
$w = New-Object System.IO.BinaryWriter $fs
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)

$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]
    $w.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))   # 0 means 256
    $w.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))
    $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([uint16]1); $w.Write([uint16]32)
    $w.Write([uint32]$blobs[$i].Length)
    $w.Write([uint32]$offset)
    $offset += $blobs[$i].Length
}
foreach ($b in $blobs) { $w.Write($b) }
$w.Close(); $fs.Close()

"{0}  {1:N0} bytes  ({2} sizes)" -f $Out, (Get-Item $Out).Length, $sizes.Count
