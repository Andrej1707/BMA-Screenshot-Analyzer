param(
    [string]$PythonPath = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if (-not $PythonPath) {
    $localBuildPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
    if (Test-Path $localBuildPython) {
        $PythonPath = (Resolve-Path $localBuildPython).Path
    } else {
        $PythonPath = "python"
    }
}

& (Join-Path $projectRoot "build-portable.ps1") -PythonPath $PythonPath -SkipReleasePackage

$isccCandidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "Inno Setup 6 was not found."
}

& $iscc installer.iss
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed."
}

Write-Host "Installer created:"
Write-Host (Join-Path $projectRoot "dist-installer\BMA Screenshot Analyzer Setup.exe")
