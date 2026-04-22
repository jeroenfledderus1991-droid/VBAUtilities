# Build-VBEAddIn.ps1
# Bouwt de VBE Add-in + installer (met framework check)

param(
    [switch]$Release = $true,
    [switch]$Debug = $false,
    [int]$KeepInstallerVersions = 9,
    [switch]$AllowRetargetTo48 = $false
)

$ErrorActionPreference = "Stop"

$config = if ($Debug) { "Debug" } else { "Release" }

function Get-CurrentVersion {
    param([string]$ProjectRoot)

    $changelogDataPath = Join-Path $ProjectRoot "ChangelogData.cs"
    if (-not (Test-Path $changelogDataPath)) {
        throw "ChangelogData.cs niet gevonden op $changelogDataPath"
    }

    $match = Select-String -Path $changelogDataPath -Pattern 'CurrentVersion\s*=\s*"([0-9]+\.[0-9]+\.[0-9]+)"' | Select-Object -First 1
    if ($null -eq $match) {
        throw "CurrentVersion niet gevonden in ChangelogData.cs"
    }

    return $match.Matches[0].Groups[1].Value
}

function New-VersionedInstallerCopy {
    param([string]$InstallerPath, [string]$Version)

    $installerDirectory = Split-Path -Parent $InstallerPath
    $versionedInstallerPath = Join-Path $installerDirectory ("VBEAddIn-Installer-{0}.exe" -f $Version)
    Copy-Item -Path $InstallerPath -Destination $versionedInstallerPath -Force
    return $versionedInstallerPath
}

function Remove-OldVersionedInstallers {
    param([string]$InstallerDirectory, [int]$KeepCount)

    if ($KeepCount -lt 1) { return @() }

    $versionedInstallers = Get-ChildItem -Path $InstallerDirectory -Filter 'VBEAddIn-Installer-*.exe' -File |
        Sort-Object LastWriteTime -Descending

    $installersToRemove = $versionedInstallers | Select-Object -Skip $KeepCount
    foreach ($installer in $installersToRemove) {
        Remove-Item -Path $installer.FullName -Force
    }

    return $installersToRemove
}

function Test-TargetingPack {
    param([string]$Version) # "4.8.1" or "4.8"

    $refPath = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v$Version"
    return (Test-Path $refPath)
}

function Ensure-FrameworkTarget {
    param(
        [string]$ProjectPath,
        [string]$DesiredVersion = "4.8.1",
        [switch]$AllowRetargetTo48
    )

    if (Test-TargetingPack $DesiredVersion) { return $DesiredVersion }

    if ($AllowRetargetTo48 -and (Test-TargetingPack "4.8")) {
        # retarget csproj to v4.8
        (Get-Content $ProjectPath) -replace '<TargetFrameworkVersion>v4\.8\.1</TargetFrameworkVersion>', '<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>' |
            Set-Content $ProjectPath -Encoding UTF8
        return "4.8"
    }

    Write-Host "ERROR: .NET Framework $DesiredVersion Targeting Pack ontbreekt." -ForegroundColor Red
    Write-Host "Installeer de Developer Pack (SDK/Targeting Pack): https://aka.ms/msbuild/developerpacks" -ForegroundColor Yellow
    if ($AllowRetargetTo48) {
        Write-Host "Tip: installeer 4.8.1 of zorg dat 4.8 aanwezig is." -ForegroundColor Yellow
    }
    exit 1
}

Write-Host "`n=== VBE Add-in Builder ===" -ForegroundColor Cyan
Write-Host "Configuratie: $config`n" -ForegroundColor Gray

$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    Write-Host "ERROR: MSBuild niet gevonden op $msbuild" -ForegroundColor Red
    exit 1
}
Write-Host "MSBuild gevonden: $msbuild`n" -ForegroundColor Gray

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptPath
Write-Host "Werkmap: $(Get-Location)`n" -ForegroundColor Gray

# Zorg dat target framework beschikbaar is (VBEAddIn + Installer)
$tf1 = Ensure-FrameworkTarget -ProjectPath ".\VBEAddIn.csproj" -AllowRetargetTo48:$AllowRetargetTo48
$tf2 = Ensure-FrameworkTarget -ProjectPath ".\Installer\Installer.csproj" -AllowRetargetTo48:$AllowRetargetTo48
Write-Host "TargetFramework: VBEAddIn=$tf1, Installer=$tf2`n" -ForegroundColor Gray

$currentVersion = Get-CurrentVersion -ProjectRoot $scriptPath
Write-Host "Versie: $currentVersion`n" -ForegroundColor Gray

Write-Host "[1/2] Building VBEAddIn project..." -ForegroundColor Yellow
& $msbuild ".\VBEAddIn.csproj" /p:Configuration=$config /t:Clean,Build /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { Write-Host "`n✗ VBEAddIn project build FAILED" -ForegroundColor Red; exit 1 }
Write-Host "✓ VBEAddIn.dll gebouwd`n" -ForegroundColor Green

Write-Host "[2/2] Building Installer..." -ForegroundColor Yellow
& $msbuild ".\Installer\Installer.csproj" /p:Configuration=$config /t:Rebuild /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { Write-Host "`n✗ Installer build FAILED" -ForegroundColor Red; exit 1 }
Write-Host "✓ Installer gebouwd`n" -ForegroundColor Green

$installerPath = ".\Installer\bin\$config\VBEAddIn-Installer.exe"
$dllPath = ".\bin\$config\VBEAddIn.dll"

Write-Host "=== BUILD SUCCESS ===" -ForegroundColor Green
Write-Host "`nOutput:" -ForegroundColor Cyan

if (Test-Path $installerPath) {
    $installerFile = Get-Item $installerPath
    $versionedInstallerPath = New-VersionedInstallerCopy -InstallerPath $installerFile.FullName -Version $currentVersion
    $removedInstallers = Remove-OldVersionedInstallers -InstallerDirectory $installerFile.DirectoryName -KeepCount $KeepInstallerVersions

    Write-Host "  Installer: $installerPath" -ForegroundColor Gray
    Write-Host "             $([math]::Round($installerFile.Length / 1KB, 2)) KB - $($installerFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    Write-Host "  Versie:    $versionedInstallerPath" -ForegroundColor Gray

    if ($removedInstallers.Count -gt 0) {
        Write-Host "  Opgeschoond: $($removedInstallers.Count) oudere installer-versie(s) verwijderd" -ForegroundColor Gray
    }
} else {
    Write-Host "  WARNING: Installer niet gevonden op $installerPath" -ForegroundColor Yellow
}

if (Test-Path $dllPath) {
    $dllFile = Get-Item $dllPath
    Write-Host "  DLL:       $dllPath" -ForegroundColor Gray
    Write-Host "             $([math]::Round($dllFile.Length / 1KB, 2)) KB - $($dllFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
} else {
    Write-Host "  WARNING: DLL niet gevonden op $dllPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Project gereed!" -ForegroundColor Green
Write-Host ""