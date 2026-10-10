param(
    [string]$Python = 'C:\Users\Admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe',
    [string]$ReleasePath = (Join-Path $PSScriptRoot '..\Release\Packing List Manager'),
    [string]$CatalogPath = (Join-Path $PSScriptRoot '..\Packing List Manager')
)
$ErrorActionPreference = 'Stop'
$release = [IO.Path]::GetFullPath($ReleasePath)
if (Test-Path -LiteralPath $release) { throw 'Build into a fresh release folder. Existing releases and product data are never overwritten.' }
foreach ($name in @('ProductNames.json', 'SkuCatalog.json')) {
    if (-not (Test-Path -LiteralPath (Join-Path $CatalogPath $name))) { throw "Missing approved catalog: $name" }
}
New-Item -ItemType Directory -Path $release | Out-Null
& $Python (Join-Path $PSScriptRoot 'package_runtime.py') (Join-Path $release 'App Files')
if ($LASTEXITCODE -ne 0) { throw 'Python/Tk runtime packaging failed.' }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ ('/out:' + (Join-Path $release 'Packing List Manager.exe')) ('/win32manifest:' + (Join-Path $PSScriptRoot 'app.manifest')) ('/win32icon:' + (Join-Path $PSScriptRoot 'src\assets\packing-list-icon.ico')) /reference:System.dll /reference:System.Windows.Forms.dll (Join-Path $PSScriptRoot 'PackingListLauncher.cs')
if ($LASTEXITCODE -ne 0) { throw 'Windows launcher compilation failed.' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'src\assets\packing-list-icon.ico') -Destination (Join-Path $release 'Packing List Manager.ico')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Employee Guide.txt') -Destination $release
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Release Notes.txt') -Destination $release
foreach ($name in @('ProductNames.json', 'SkuCatalog.json')) {
    Copy-Item -LiteralPath (Join-Path $CatalogPath $name) -Destination $release
}
$defaults = Join-Path $release 'App Files\Defaults'
New-Item -ItemType Directory -Path $defaults | Out-Null
foreach ($name in @('ProductNames.json', 'SkuCatalog.json')) {
    Copy-Item -LiteralPath (Join-Path $CatalogPath $name) -Destination $defaults
}
Write-Output "Packaged application: $release"
