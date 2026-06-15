# Submittal Tracker — watches the "Submittals Inbox" folder and logs to the Excel log.
# Usage:  .\run.ps1               (uses existing $env:ANTHROPIC_API_KEY)
#         .\run.ps1 -Key sk-ant-... (sets the key for this run)
param([string]$Key)

if ($Key) { $env:ANTHROPIC_API_KEY = $Key }

if (-not $env:ANTHROPIC_API_KEY) {
    Write-Host "ANTHROPIC_API_KEY is not set." -ForegroundColor Yellow
    Write-Host "Set it once permanently:  setx ANTHROPIC_API_KEY `"sk-ant-...`"" -ForegroundColor Cyan
    Write-Host "Or pass it this run:      .\run.ps1 -Key sk-ant-..." -ForegroundColor Cyan
    exit 1
}

$projectDir = $PSScriptRoot
$venvPython = Join-Path $projectDir "venv\Scripts\python.exe"

if (-not (Test-Path $venvPython)) {
    Write-Host "Setting up virtual environment..."
    python -m venv "$projectDir\venv"
    & "$projectDir\venv\Scripts\pip.exe" install -r "$projectDir\requirements.txt" --quiet
    Write-Host "Done."
}

& $venvPython "$projectDir\app\watcher.py"
