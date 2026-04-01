# ============================================================
# Quick rebuild THBIM.Logic only (for hot-reload workflow)
# ============================================================
# Usage: powershell -ExecutionPolicy Bypass -File rebuild-logic.ps1
#
# Workflow:
# 1. Edit code in THBIM.Logic
# 2. Run this script (rebuilds + copies to deploy folder)
# 3. Click "Reload Plugin" button in Revit ribbon
# ============================================================

param(
    [string]$Config = "Release",
    [string]$DeployRoot = "$env:ProgramData\Autodesk\Revit\Addins\2025"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[1/2] Building THBIM.Logic..." -ForegroundColor Yellow
dotnet build "$ScriptDir\THBIM.Logic\THBIM.Logic.csproj" -c $Config --no-restore
if ($LASTEXITCODE -ne 0) {
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    exit 1
}

$LogicDir = Join-Path $DeployRoot "TH Tools\Logic"
$LogicOut = Join-Path $ScriptDir "THBIM.Logic\bin\$Config\net8.0-windows"

Write-Host "[2/2] Copying Logic DLLs..." -ForegroundColor Yellow
$files = @("THBIM.Logic.dll", "THBIM.Logic.pdb", "THBIM.Logic.deps.json",
           "THBIM.Licensing.dll", "THBIM.Licensing.pdb")
foreach ($f in $files) {
    $src = Join-Path $LogicOut $f
    if (Test-Path $src) { Copy-Item $src -Destination $LogicDir -Force }
}

Write-Host "Done! Click 'Reload Plugin' in Revit." -ForegroundColor Green
