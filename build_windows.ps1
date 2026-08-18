$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

Write-Host "Installing the exact project and build requirements..."
python -m pip install -r requirements-dev.txt

Write-Host "Running code checks and automated tests..."
python -m ruff check .
python -m unittest discover -s tests -v

Write-Host "Building the Windows application..."
python -m PyInstaller --noconfirm --clean book-reader-gui.spec

$builtApp = Join-Path $projectRoot "dist\book-reader.exe"
$releaseApp = Join-Path $projectRoot "book-reader.exe"
$legacyApp = Join-Path $projectRoot "dist\book-reader-gui.exe"
if (-not (Test-Path -LiteralPath $builtApp -PathType Leaf)) {
    throw "Build finished without creating $builtApp"
}
Copy-Item -LiteralPath $builtApp -Destination $releaseApp -Force
# Keep the former filename working for anyone who has an old shortcut.
Copy-Item -LiteralPath $builtApp -Destination $legacyApp -Force

$innoCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)
$innoCompiler = $innoCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ($innoCompiler) {
    Write-Host "Building the Windows installer..."
    & $innoCompiler (Join-Path $projectRoot "packaging\windows\book-to-latex.iss")
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup could not build the installer." }
}

Write-Host "Finished: $releaseApp"
