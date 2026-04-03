# Build-VBEAddIn.ps1
# Snel build script voor VBE Add-in en Installer

param(
    [switch]$Release = $true,
    [switch]$Debug = $false
)

$ErrorActionPreference = "Stop"

# Bepaal configuratie
$config = if ($Debug) { "Debug" } else { "Release" }

# Kleuren
$green = "Green"
$yellow = "Yellow"
$red = "Red"

Write-Host "`n=== VBE Add-in Builder ===" -ForegroundColor Cyan
Write-Host "Configuratie: $config`n" -ForegroundColor Gray

# Zoek MSBuild
$msbuild = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"

if (-not (Test-Path $msbuild)) {
    Write-Host "ERROR: MSBuild niet gevonden op $msbuild" -ForegroundColor $red
    exit 1
}

# Ga naar project directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptPath

Write-Host "[1/2] Building VBEAddIn.csproj..." -ForegroundColor $yellow

try {
    & $msbuild VBEAddIn.csproj /p:Configuration=$config /t:Rebuild /v:minimal /nologo
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "`n✗ VBEAddIn build FAILED" -ForegroundColor $red
        exit 1
    }
    
    Write-Host "✓ VBEAddIn.dll gebouwd`n" -ForegroundColor $green
}
catch {
    Write-Host "`n✗ Error: $($_.Exception.Message)" -ForegroundColor $red
    exit 1
}

Write-Host "[2/2] Building Installer..." -ForegroundColor $yellow

try {
    & $msbuild Installer\Installer.csproj /p:Configuration=$config /t:Rebuild /v:minimal /nologo
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "`n✗ Installer build FAILED" -ForegroundColor $red
        exit 1
    }
    
    Write-Host "✓ Installer gebouwd`n" -ForegroundColor $green
}
catch {
    Write-Host "`n✗ Error: $($_.Exception.Message)" -ForegroundColor $red
    exit 1
}

# Toon resultaat
$installerPath = ".\Installer\bin\$config\VBEAddIn-Installer.exe"
$dllPath = ".\bin\$config\VBEAddIn.dll"

if (Test-Path $installerPath) {
    $installerFile = Get-Item $installerPath
    $dllFile = Get-Item $dllPath
    
    Write-Host "=== BUILD SUCCESS ===" -ForegroundColor $green
    Write-Host "`nOutput:" -ForegroundColor Cyan
    Write-Host "  Installer: $installerPath" -ForegroundColor Gray
    Write-Host "             $([math]::Round($installerFile.Length/1KB, 2)) KB - $($installerFile.LastWriteTime)" -ForegroundColor Gray
    Write-Host "  DLL:       $dllPath" -ForegroundColor Gray
    Write-Host "             $([math]::Round($dllFile.Length/1KB, 2)) KB - $($dllFile.LastWriteTime)" -ForegroundColor Gray
    Write-Host "`n  Klaar om te installeren!`n" -ForegroundColor $green
}
else {
    Write-Host "WARNING: Installer niet gevonden" -ForegroundColor $yellow
}
