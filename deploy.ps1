# ============================================================
# THBIMv2 Deploy Script for Revit 2025
# ============================================================
# Usage: powershell -ExecutionPolicy Bypass -File deploy.ps1
#
# This script:
# 1. Builds the solution in Release mode
# 2. Copies files to Revit's Addins folder with correct structure
# 3. Creates the .addin manifest
# ============================================================

param(
    [string]$Config = "Release",
    [string]$DeployRoot = "$env:ProgramData\Autodesk\Revit\Addins\2025"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  THBIMv2 Deploy to Revit 2025" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# ---- Step 1: Build ----
Write-Host "`n[1/4] Building solution..." -ForegroundColor Yellow
dotnet build "$ScriptDir\THBIMv2.sln" -c $Config --no-restore
if ($LASTEXITCODE -ne 0) {
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    exit 1
}
Write-Host "Build succeeded." -ForegroundColor Green

# ---- Step 2: Prepare deploy folder ----
$PluginDir = Join-Path $DeployRoot "TH Tools"
$LogicDir  = Join-Path $PluginDir "Logic"
$ResDir    = Join-Path $PluginDir "Resources"

Write-Host "`n[2/4] Preparing deploy folder: $PluginDir" -ForegroundColor Yellow

# Create directories
New-Item -ItemType Directory -Force -Path $PluginDir | Out-Null
New-Item -ItemType Directory -Force -Path $LogicDir  | Out-Null
New-Item -ItemType Directory -Force -Path $ResDir     | Out-Null

# ---- Step 3: Copy files ----
Write-Host "`n[3/4] Copying files..." -ForegroundColor Yellow

$LoaderOut = Join-Path $ScriptDir "THBIM.Loader\bin\$Config\net8.0-windows"
$LogicOut  = Join-Path $ScriptDir "THBIM.Logic\bin\$Config\net8.0-windows"

# 3a: Loader files → TH Tools/
$loaderFiles = @(
    "THBIM.Loader.dll",
    "THBIM.Loader.pdb",
    "THBIM.Loader.deps.json",
    "THBIM.Contracts.dll",
    "THBIM.Contracts.pdb"
)
foreach ($f in $loaderFiles) {
    $src = Join-Path $LoaderOut $f
    if (Test-Path $src) {
        Copy-Item $src -Destination $PluginDir -Force
        Write-Host "  [Loader] $f" -ForegroundColor Gray
    }
}

# 3b: Logic files → TH Tools/Logic/
# Only copy OUR DLLs and necessary NuGet deps, NOT Revit's DLLs
$logicInclude = @(
    "THBIM.Logic.dll",
    "THBIM.Logic.pdb",
    "THBIM.Logic.deps.json",
    "THBIM.Contracts.dll",
    "THBIM.Contracts.pdb",
    "THBIM.Licensing.dll",
    "THBIM.Licensing.pdb",
    # NuGet dependencies
    "EPPlus.dll",
    "EPPlus.Interfaces.dll",
    "EPPlus.System.Drawing.dll",
    "ClosedXML.dll",
    "DocumentFormat.OpenXml.dll",
    "DocumentFormat.OpenXml.Framework.dll",
    "ExcelNumberFormat.dll",
    "Irony.dll",
    "MiniExcel.dll",
    "RBush.dll",
    "SixLabors.Fonts.dll",
    "SixLabors.ImageSharp.dll",
    "XLParser.dll",
    "System.IO.Packaging.dll",
    "System.Management.dll"
)
foreach ($f in $logicInclude) {
    $src = Join-Path $LogicOut $f
    if (Test-Path $src) {
        Copy-Item $src -Destination $LogicDir -Force
        Write-Host "  [Logic]  $f" -ForegroundColor Gray
    }
}

# 3c: Resources → TH Tools/Resources/
$ResSrc = Join-Path $ScriptDir "THBIM.Logic\Resources"
if (Test-Path $ResSrc) {
    Copy-Item "$ResSrc\*" -Destination $ResDir -Recurse -Force
    Write-Host "  [Res]    Resources/ copied" -ForegroundColor Gray
}

# 3d: RFA families → TH Tools/rfa/
$RfaSrc = Join-Path $ScriptDir "THBIM.Logic\rfa"
if (Test-Path $RfaSrc) {
    $RfaDst = Join-Path $PluginDir "rfa"
    New-Item -ItemType Directory -Force -Path $RfaDst | Out-Null
    Copy-Item "$RfaSrc\*" -Destination $RfaDst -Recurse -Force
    Write-Host "  [RFA]    rfa/ copied" -ForegroundColor Gray
}

# ---- Step 4: Create .addin file ----
Write-Host "`n[4/4] Creating .addin manifest..." -ForegroundColor Yellow

$addinContent = @"
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
    <AddIn Type="Application">
        <Name>THBIM Tools v2</Name>
        <Assembly>TH Tools\THBIM.Loader.dll</Assembly>
        <AddInId>B28C6442-9985-484F-9314-E8504B768239</AddInId>
        <FullClassName>THBIM.Loader.LoaderApp</FullClassName>
        <VendorId>THBIM</VendorId>
        <VendorDescription>THBIM Tools for Structural Engineering</VendorDescription>
    </AddIn>
</RevitAddIns>
"@

$addinPath = Join-Path $DeployRoot "THBIMv2.addin"
Set-Content -Path $addinPath -Value $addinContent -Encoding UTF8
Write-Host "  Created: $addinPath" -ForegroundColor Gray

# ---- Done ----
Write-Host "`n============================================" -ForegroundColor Green
Write-Host "  Deploy completed successfully!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Deployed to:" -ForegroundColor White
Write-Host "  .addin:   $addinPath" -ForegroundColor Gray
Write-Host "  Plugin:   $PluginDir" -ForegroundColor Gray
Write-Host "  Loader:   $PluginDir\THBIM.Loader.dll" -ForegroundColor Gray
Write-Host "  Logic:    $LogicDir\THBIM.Logic.dll" -ForegroundColor Gray
Write-Host ""
Write-Host "Structure:" -ForegroundColor White
Write-Host "  $DeployRoot\" -ForegroundColor Gray
Write-Host "    THBIMv2.addin" -ForegroundColor Gray
Write-Host "    TH Tools\" -ForegroundColor Gray
Write-Host "      THBIM.Loader.dll" -ForegroundColor Gray
Write-Host "      THBIM.Contracts.dll" -ForegroundColor Gray
Write-Host "      Resources\" -ForegroundColor Gray
Write-Host "      rfa\" -ForegroundColor Gray
Write-Host "      Logic\" -ForegroundColor Gray
Write-Host "        THBIM.Logic.dll" -ForegroundColor Gray
Write-Host "        THBIM.Licensing.dll" -ForegroundColor Gray
Write-Host "        EPPlus.dll, ClosedXML.dll, ..." -ForegroundColor Gray
Write-Host "      LogicShadow\  (created at runtime)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Next: Start Revit 2025. The plugin will appear in 'TH Tools' tab." -ForegroundColor Cyan
Write-Host "Use the 'Reload Plugin' button (Dev panel) for hot-reload." -ForegroundColor Cyan
