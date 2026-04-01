---
name: THBIMv2 Release Workflow
description: Step-by-step guide to check changes, build zip, update manifest, and release THBIMv2 addin + app to GitHub
type: reference
---

## Repos

- **Addin**: `lethienhieu/revit-tool` — THBIM Revit Tools (DLL + addin)
- **App**: `lethienhieu/thbim-autoupdate` — THBIM AutoUpdate app (exe)
- **Source code**: `E:\THBIM-CODE\2025\CODE\THBIMv2`
- **AutoUpdate app**: `E:\THBIM-CODE\2025\THBIM-AutoUpdate`

## gh CLI

```bash
export PATH="/c/Users/Le_ThienHieu/tools/gh/bin:$PATH"
```

## THBIMv2 Project Structure

```
THBIMv2/
├── THBIM.Loader/          → Loader DLL (Revit entry point, FullClassName: THBIM.Loader.LoaderApp)
│   └── bin/Release/net8.0-windows/
│       ├── THBIM.Loader.dll
│       ├── THBIM.Contracts.dll
│       └── Resources/
├── THBIM.Logic/           → Main logic DLL (all tools)
│   └── bin/Release/net8.0-windows/
│       ├── THBIM.Logic.dll
│       ├── THBIM.Contracts.dll
│       ├── THBIM.Licensing.dll
│       └── ... (dependencies)
│   ├── Resources/         → Icons + MEP/
│   ├── rfa/               → Revit families
│   └── SheetLink/Resources/
├── THBIM.Contracts/       → Shared interfaces
├── THBIM.addin            → Addin manifest (Assembly: TH Tools\THBIM.dll)
└── THBIMv2.sln
```

**net48 + net8** (Revit 2023, 2024, 2025, 2026).

## Release Addin (revit-tool)

### Step 1: Update manifest.json (version + changelog)

```bash
node -e "
const d=JSON.parse(require('fs').readFileSync('E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json','utf8'));
d.version='X.Y.Z';
d.changelog=['Change 1','Change 2'];
d.tools.forEach(t=>t.version='X.Y.Z');
require('fs').writeFileSync('E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json',JSON.stringify(d,null,2));
"
```

To keep existing changelog (minor fix), only update version:
```bash
node -e "
const d=JSON.parse(require('fs').readFileSync('E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json','utf8'));
d.version='X.Y.Z';
d.tools.forEach(t=>t.version='X.Y.Z');
require('fs').writeFileSync('E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json',JSON.stringify(d,null,2));
"
```

### Step 2: Rebuild ZIP

```bash
TMPDIR=$(mktemp -d)
LOADER="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.Loader/bin/x64/Release/net8.0-windows"
N48="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.Logic/bin/x64/Release/net48"
N8="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.Logic/bin/x64/Release/net8.0-windows"
RES="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.Logic/Resources"
RFA="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.Logic/rfa"
ADDIN="E:/THBIM-CODE/2025/CODE/THBIMv2/THBIM.addin"
MF="E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json"

mkdir -p "$TMPDIR/net48/Resources" "$TMPDIR/net48/rfa"
mkdir -p "$TMPDIR/net8/Resources" "$TMPDIR/net8/rfa"

# --- net48 (Revit 2023-2024): THBIM.dll = Logic assembly (AssemblyName=THBIM) ---
cp "$N48"/THBIM.dll "$TMPDIR/net48/"; cp "$N48"/THBIM.pdb "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/THBIM.Contracts.dll "$TMPDIR/net48/"; cp "$N48"/THBIM.Contracts.pdb "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/THBIM.Licensing.dll "$TMPDIR/net48/"; cp "$N48"/THBIM.Licensing.pdb "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/ClosedXML.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/DocumentFormat.OpenXml.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/EPPlus.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/EPPlus.Interfaces.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/EPPlus.System.Drawing.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/ExcelNumberFormat.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/IndexRange.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/Irony.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/Microsoft.Bcl.AsyncInterfaces.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/Microsoft.IO.RecyclableMemoryStream.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/SixLabors.Fonts.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.Buffers.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/System.CodeDom.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.ComponentModel.Annotations.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.IO.Packaging.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/System.Memory.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.Numerics.Vectors.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.Runtime.CompilerServices.Unsafe.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.Text.Encodings.Web.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/System.Text.Json.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.Threading.Tasks.Extensions.dll "$TMPDIR/net48/" 2>/dev/null
cp "$N48"/System.ValueTuple.dll "$TMPDIR/net48/" 2>/dev/null; cp "$N48"/XLParser.dll "$TMPDIR/net48/" 2>/dev/null
cp "$RES"/* "$TMPDIR/net48/Resources/" 2>/dev/null; cp -r "$RES/MEP" "$TMPDIR/net48/Resources/" 2>/dev/null
cp "$RFA"/* "$TMPDIR/net48/rfa/" 2>/dev/null

# --- net8 (Revit 2025-2026): THBIM.dll = Loader entry point + THBIM.Logic.dll ---
cp "$LOADER"/THBIM.dll "$TMPDIR/net8/"; cp "$LOADER"/THBIM.pdb "$TMPDIR/net8/" 2>/dev/null
cp "$LOADER"/THBIM.deps.json "$TMPDIR/net8/" 2>/dev/null
cp "$LOADER"/THBIM.Contracts.dll "$TMPDIR/net8/"; cp "$LOADER"/THBIM.Contracts.pdb "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/THBIM.Logic.dll "$TMPDIR/net8/"; cp "$N8"/THBIM.Logic.pdb "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/THBIM.Logic.deps.json "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/THBIM.Licensing.dll "$TMPDIR/net8/"; cp "$N8"/THBIM.Licensing.pdb "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/ClosedXML.dll "$TMPDIR/net8/" 2>/dev/null; cp "$N8"/DocumentFormat.OpenXml.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/EPPlus.dll "$TMPDIR/net8/" 2>/dev/null; cp "$N8"/EPPlus.Interfaces.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/EPPlus.System.Drawing.dll "$TMPDIR/net8/" 2>/dev/null; cp "$N8"/Irony.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/Microsoft.IO.RecyclableMemoryStream.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/SixLabors.Fonts.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/System.CodeDom.dll "$TMPDIR/net8/" 2>/dev/null; cp "$N8"/System.Management.dll "$TMPDIR/net8/" 2>/dev/null
cp "$N8"/XLParser.dll "$TMPDIR/net8/" 2>/dev/null
cp -r "$N8/runtimes" "$TMPDIR/net8/" 2>/dev/null
cp "$RES"/* "$TMPDIR/net8/Resources/" 2>/dev/null; cp -r "$RES/MEP" "$TMPDIR/net8/Resources/" 2>/dev/null
cp "$RFA"/* "$TMPDIR/net8/rfa/" 2>/dev/null

# Addin + manifest
cp "$ADDIN" "$TMPDIR/"; cp "$MF" "$TMPDIR/"

rm -f "E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/THBIM-RevitTools.zip"
cd "$TMPDIR" && powershell.exe -Command "Compress-Archive -Path '.\*' -DestinationPath 'E:\THBIM-CODE\2025\THBIM-AutoUpdate\releases\THBIM-RevitTools.zip' -Force"
```

### Step 3: Create GitHub release

```bash
export PATH="/c/Users/Le_ThienHieu/tools/gh/bin:$PATH"

gh release create vX.Y.Z \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json" \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/THBIM-RevitTools.zip" \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/publish/THBIM.AutoUpdate.exe" \
  --repo lethienhieu/revit-tool \
  --title "THBIM Revit Tools vX.Y.Z" \
  --notes "- Change 1
- Change 2"
```

To update existing release assets (same version):
```bash
gh release upload vX.Y.Z \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/manifest.json" \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/releases/THBIM-RevitTools.zip" \
  --repo lethienhieu/revit-tool --clobber
```

## Release App (thbim-autoupdate)

### Step 1: Bump version in 2 files

- `E:\THBIM-CODE\2025\THBIM-AutoUpdate\src\THBIM.AutoUpdate\THBIM.AutoUpdate.csproj` → Version, FileVersion, AssemblyVersion
- `E:\THBIM-CODE\2025\THBIM-AutoUpdate\src\THBIM.AutoUpdate\ViewModels\MainViewModel.cs` → `public const string AppVersion = "X.Y.Z";`

### Step 2: Build

```bash
dotnet publish "E:/THBIM-CODE/2025/THBIM-AutoUpdate/src/THBIM.AutoUpdate/THBIM.AutoUpdate.csproj" \
  -c Release -r win-x64 --self-contained true \
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true \
  -o "E:/THBIM-CODE/2025/THBIM-AutoUpdate/publish"
```

### Step 3: Release to thbim-autoupdate repo

```bash
gh release create vX.Y.Z \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/publish/THBIM.AutoUpdate.exe" \
  --repo lethienhieu/thbim-autoupdate \
  --title "THBIM AutoUpdate vX.Y.Z" \
  --notes "- Change 1"
```

### Step 4: Also upload exe to revit-tool latest release (for backward compat)

```bash
gh release upload vX.Y.Z \
  "E:/THBIM-CODE/2025/THBIM-AutoUpdate/publish/THBIM.AutoUpdate.exe" \
  --repo lethienhieu/revit-tool --clobber
```

## Versioning Rules

- Addin version: `revit-tool` tag (e.g. v1.1.0, v1.1.1...)
- App version: `thbim-autoupdate` tag (e.g. v1.0.7, v1.0.8...)
- manifest.json `version` = addin version
- manifest.json `appVersion` = app version
- 10 minor versions (x.x.0 → x.x.9) → bump to next minor (x.y+1.0)

## ZIP Structure (THBIMv2 — net48 + net8)

```
THBIM-RevitTools.zip
├── THBIM.addin              (Assembly: TH Tools\THBIM.dll)
├── manifest.json
├── net48\                   (Revit 2023-2024)
│   ├── THBIM.dll            (THBIM.Logic net48 build, AssemblyName=THBIM)
│   ├── THBIM.Contracts.dll
│   ├── THBIM.Licensing.dll
│   ├── Resources/           (icons + MEP/)
│   ├── rfa/                 (Revit families)
│   └── ... (ClosedXML, EPPlus, compat shims)
└── net8\                    (Revit 2025-2026)
    ├── THBIM.dll            (THBIM.Loader entry point, AssemblyName=THBIM)
    ├── THBIM.Logic.dll      (main logic)
    ├── THBIM.Contracts.dll
    ├── THBIM.Licensing.dll
    ├── Resources/           (icons + MEP/)
    ├── rfa/                 (Revit families)
    ├── runtimes/
    └── ... (ClosedXML, EPPlus, etc.)
```

## Current Versions (as of 2026-03-31)

- Addin: v1.1.3
- App: v1.0.8
