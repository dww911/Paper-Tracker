# Research Radar Web UI — run from project root
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$venvActivate = Join-Path $Root ".venv\Scripts\Activate.ps1"
if (Test-Path $venvActivate) {
    & $venvActivate
}

Write-Host "Starting Research Radar at http://127.0.0.1:8000" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray
python -m uvicorn web.web_app:app --reload --host 127.0.0.1 --port 8000
