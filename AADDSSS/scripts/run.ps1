$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$Root = Split-Path -Parent $PSScriptRoot
Set-Location -Path $Root
Write-Host "Running ADS task..." -ForegroundColor Cyan
py -u run.py
Read-Host "Press Enter to close"
