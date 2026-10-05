param([string]$Python = 'C:\Users\Admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe', [switch]$UiOnly, [string]$ReleasePath)
$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
$release = if ($ReleasePath) { [System.IO.Path]::GetFullPath($ReleasePath) } else { Join-Path (Split-Path $project -Parent) 'Packing List Manager' }
New-Item -ItemType Directory -Path $release -Force | Out-Null
if (-not $UiOnly) {
$engineFolder = Join-Path $release 'App Files'
if (Test-Path -LiteralPath $engineFolder) { throw 'Use a fresh release folder for a full build, then install the verified App Files folder.' }
& $Python (Join-Path $project 'package_runtime.py') $engineFolder
if ($LASTEXITCODE -ne 0) { throw 'Python runtime packaging failed.' }
$engineCompiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $engineCompiler /nologo /target:exe /platform:x64 /optimize+ ('/out:' + (Join-Path $engineFolder 'PackingEngine.exe')) (Join-Path $project 'src\PackingEngineLauncher.cs')
if ($LASTEXITCODE -ne 0) { throw 'Engine launcher compilation failed.' }
}
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ ('/out:' + (Join-Path $release 'Packing List Manager.exe')) ('/win32manifest:' + (Join-Path $project 'src\app.manifest')) ('/win32icon:' + (Join-Path $project 'src\assets\packing-list-icon.ico')) /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll /reference:Microsoft.CSharp.dll (Join-Path $project 'src\PackingListManager.cs') (Join-Path $project 'src\PackingDialogs.cs') (Join-Path $project 'src\NameDefinitionsForm.cs') (Join-Path $project 'src\NameExcelSession.cs')
if ($LASTEXITCODE -ne 0) { throw 'Windows application compilation failed.' }
Copy-Item -LiteralPath (Join-Path $project 'Employee Guide.txt') -Destination $release -Force
Copy-Item -LiteralPath (Join-Path $project 'src\assets\packing-list-icon.ico') -Destination (Join-Path $release 'Packing List Manager.ico') -Force
Copy-Item -LiteralPath (Join-Path $project 'Release Notes.txt') -Destination $release -Force
Write-Output $release
