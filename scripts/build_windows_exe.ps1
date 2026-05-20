param(
    [string]$Name = "INANDBenchmark"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "== Build Windows GUI executable =="
uv run pyinstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $Name `
    --add-data "scripts;scripts" `
    main.py

$exe = Join-Path "dist" ($Name + ".exe")
if (-not (Test-Path $exe)) {
    throw "Build failed: $exe not found"
}

Write-Host "Built: $((Resolve-Path $exe).Path)"
