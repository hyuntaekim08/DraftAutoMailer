$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectDir

$python = Join-Path $projectDir ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the build virtual environment."
    }
}

& $python -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) {
    throw "Could not install the application and packaging dependencies."
}

& $python -m PyInstaller --noconfirm --clean outlook_draft_mailer.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed to package the application."
}

$compilerPath = (Get-Command ISCC.exe -ErrorAction SilentlyContinue).Source
if (-not $compilerPath) {
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $compilerPath = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $compilerPath) {
    throw "The app was packaged, but Inno Setup 6 is required to create the installer. Install Inno Setup 6, then rerun .\build.ps1."
}

& $compilerPath setup.iss
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup failed to create the installer."
}

Write-Output "Installer created: dist\installer\OutlookDraftMailerSetup.exe"
