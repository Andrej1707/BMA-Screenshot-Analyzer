param(
    [string]$PythonPath = "",
    [switch]$SkipReleasePackage
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

$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = "True"
& $PythonPath -m PyInstaller BmaScreenshotAnalyzer.spec --clean --noconfirm
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

$distDir = Join-Path $projectRoot "dist\BMA Screenshot Analyzer"
Copy-Item -LiteralPath (Join-Path $projectRoot "MelderService Grundexcel.xlsx") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $projectRoot "Anleitung.txt") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $projectRoot "Beispiel - korrektes Screenshot-Format.png") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $projectRoot "LICENSE") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $projectRoot "THIRD_PARTY_NOTICES.md") -Destination $distDir
Copy-Item -LiteralPath (Join-Path $projectRoot "assets") -Destination (Join-Path $distDir "Dependencies") -Recurse

$modelsTarget = Join-Path $distDir "OCR-Modelle"
New-Item -ItemType Directory -Path $modelsTarget -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot "models\PP-OCRv5_server_det") -Destination $modelsTarget -Recurse
Copy-Item -LiteralPath (Join-Path $projectRoot "models\PP-OCRv5_server_rec") -Destination $modelsTarget -Recurse

if (-not $SkipReleasePackage) {
    $releaseDir = Join-Path $projectRoot "release"
    New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
    $zipPath = Join-Path $releaseDir "BMA-Screenshot-Analyzer-Windows-x64-v1.0.1.zip"
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Host "Release package created:"
    Write-Host $zipPath
}
