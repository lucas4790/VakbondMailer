<#
.SYNOPSIS
    Bouwt VakbondMailer als zelfstandige .exe en zet die samen met de voorbeelddocumenten
    (ledenlijst, sjablonen, handleiding) in één map die je zo aan de gebruiker kunt geven.
#>
param(
    [string]$OutputDir = "dist\VakbondMailer"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dest = Join-Path $root $OutputDir

Write-Host "1/3 Publiceren als zelfstandige .exe..."
dotnet publish (Join-Path $root "src\VakbondMailer") `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:DebugType=none `
    -o $dest
if ($LASTEXITCODE -ne 0) { throw "dotnet publish is mislukt." }

Write-Host "2/3 Voorbeelddocumenten en handleiding kopiëren..."
Copy-Item (Join-Path $root "Ledenlijst-voorbeeld.csv") $dest -Force
Copy-Item (Join-Path $root "LEESMIJ.txt") $dest -Force

$sjablonenDest = Join-Path $dest "Sjablonen"
New-Item -ItemType Directory -Force -Path $sjablonenDest | Out-Null
Copy-Item (Join-Path $root "templates-voorbeeld\*.json") $sjablonenDest -Force

Write-Host "3/3 Klaar."
Write-Host ""
Write-Host "Klaar. Uitgave staat in:"
Write-Host "  $dest"
Write-Host ""
Write-Host "Zip deze map (rechtermuisknop > Verzenden naar > Gecomprimeerde map) om hem te delen."
