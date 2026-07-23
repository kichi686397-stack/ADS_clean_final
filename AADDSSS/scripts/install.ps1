$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$Root = Split-Path -Parent $PSScriptRoot
Set-Location -Path $Root
Write-Host "Installing Python packages..." -ForegroundColor Cyan
$Python = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $Python) {
    $Python = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $Python) {
    Write-Host "Python is not installed. Please install Python first." -ForegroundColor Red
} else {
    & $Python -m pip install --upgrade requests openpyxl
}
Read-Host "Press Enter to close"
