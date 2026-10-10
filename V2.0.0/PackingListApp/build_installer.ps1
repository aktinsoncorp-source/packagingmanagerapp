param(
    [string]$Python = 'C:\Users\Admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe',
    [string]$MakeNSIS = (Join-Path $PSScriptRoot '.builddeps\nsis-3.13\makensis.exe'),
    [string]$ReleasePath = (Join-Path $PSScriptRoot '..\Release\Packing List Manager'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Release')
)
$ErrorActionPreference = 'Stop'
$release = [IO.Path]::GetFullPath($ReleasePath)
$output = [IO.Path]::GetFullPath($OutputFolder)
$setup = Join-Path $output 'Packing List Manager Setup 2.0.0.exe'
$removal = Join-Path $output 'uninstall-files.nsh'
if (Test-Path -LiteralPath $setup) { throw 'The installer already exists. Choose a fresh output folder.' }
New-Item -ItemType Directory -Path $output -Force | Out-Null
& $Python (Join-Path $PSScriptRoot 'package_manifest.py') $release $removal
if ($LASTEXITCODE -ne 0) { throw 'Package inventory failed.' }
& $MakeNSIS /V2 ("/DAPPDIR=$release") ("/DOUTFILE=$setup") ("/DUNINSTALL_FILES=$removal") (Join-Path $PSScriptRoot 'installer.nsi')
if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed.' }
Write-Output $setup
