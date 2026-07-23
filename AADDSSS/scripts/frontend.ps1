$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$Root = Split-Path -Parent $PSScriptRoot
Set-Location -Path $Root
Write-Host "Starting ADS frontend..." -ForegroundColor Cyan
Write-Host "Open: http://127.0.0.1:8765" -ForegroundColor Green
$Python = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $Python) {
    $Python = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $Python) {
    Write-Host "Python is not installed. Please install Python first." -ForegroundColor Red
} else {
    & $Python app.py
}
Read-Host "Press Enter to close"
