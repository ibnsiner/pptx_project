# Free TCP 8000 (LISTEN only). Run: powershell -ExecutionPolicy Bypass -File .\stop-parser-8000.ps1
$ErrorActionPreference = "SilentlyContinue"
$conns = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue
if (-not $conns) {
    Write-Host "No LISTEN on port 8000."
    exit 0
}
$conns | ForEach-Object {
    $p = $_.OwningProcess
    Write-Host "Stopping PID $p"
    Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
}
Write-Host "Done."
