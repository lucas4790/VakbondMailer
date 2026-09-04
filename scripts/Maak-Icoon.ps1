<#
.SYNOPSIS
    Genereert een eenvoudig app-icoon (FNV-blauw met "VM"-monogram) voor VakbondMailer.
#>
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dest = Join-Path $root "src\VakbondMailer\app.ico"

Add-Type -AssemblyName System.Drawing

$size = 256
$bmp = New-Object System.Drawing.Bitmap $size, $size
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

$blue = [System.Drawing.ColorTranslator]::FromHtml("#009CDE")
$g.Clear($blue)

$rect = New-Object System.Drawing.Rectangle 0, 0, $size, $size
$path = New-Object System.Drawing.Drawing2D.GraphicsPath
$radius = 48
$path.AddArc($rect.X, $rect.Y, $radius, $radius, 180, 90)
$path.AddArc($rect.Right - $radius, $rect.Y, $radius, $radius, 270, 90)
$path.AddArc($rect.Right - $radius, $rect.Bottom - $radius, $radius, $radius, 0, 90)
$path.AddArc($rect.X, $rect.Bottom - $radius, $radius, $radius, 90, 90)
$path.CloseFigure()
$g.SetClip($path)
$g.Clear($blue)
$g.ResetClip()

$font = New-Object System.Drawing.Font("Segoe UI", 96, [System.Drawing.FontStyle]::Bold)
$brush = [System.Drawing.Brushes]::White
$text = "VM"
$textSize = $g.MeasureString($text, $font)
$x = ($size - $textSize.Width) / 2
$y = ($size - $textSize.Height) / 2 - 6
$g.DrawString($text, $font, $brush, $x, $y)

$hIcon = $bmp.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($hIcon)
$fs = New-Object System.IO.FileStream $dest, 'Create'
$icon.Save($fs)
$fs.Close()
$icon.Dispose()
$g.Dispose()
$bmp.Dispose()

Write-Host "Icoon opgeslagen: $dest"
