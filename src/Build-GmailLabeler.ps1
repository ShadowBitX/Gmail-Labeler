
$ErrorActionPreference = "Stop"

# Directory principale del progetto
$ProjectDir = Split-Path $PSScriptRoot -Parent

Set-Location $ProjectDir

Write-Host "Controllo dei file..." -ForegroundColor Cyan

$Installer = Join-Path $PSScriptRoot "Setup-GmailLabeler.ps1"
$Icon = Join-Path $ProjectDir "assets\icon.ico"
$Output = Join-Path $ProjectDir "Setup-GmailLabeler.exe"

if (!(Test-Path $Installer)) {
    throw "Setup-GmailLabeler.ps1 non trovato!"
}

if (!(Test-Path $Icon)) {
    throw "Icona non trovata!"
}

# Controllo PS2EXE
if (!(Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "PS2EXE non installato." -ForegroundColor Yellow

    $Answer = Read-Host "Vuoi installarlo? (S/N)"

    if ($Answer -notmatch "^[sS]") {
        throw "Installazione annullata."
    }

    Install-Module ps2exe -Scope CurrentUser -Force -ErrorAction Stop
}

Import-Module ps2exe -ErrorAction Stop

Write-Host ""
Write-Host "Avvio compilazione..." -ForegroundColor Cyan

Invoke-PS2EXE `
    -InputFile $Installer `
    -OutputFile $Output `
    -NoConsole `
    -STA `
    -RequireAdmin `
    -x64 `
    -IconFile $Icon `
    -Title "Gmail Labeler Setup" `
    -Description "Installazione guidata di Gmail Labeler" `
    -Company "Gmail Labeler" `
    -Product "Gmail Labeler Setup" `
    -Version "1.0.0.0"

if (!(Test-Path $Output)) {
    throw "Compilazione fallita: file EXE non generato."
}

Write-Host ""
Write-Host "COMPILAZIONE COMPLETATA!" -ForegroundColor Green
Write-Host "File creato: $Output" -ForegroundColor Green
