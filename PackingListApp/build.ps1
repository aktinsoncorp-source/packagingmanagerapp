param([string]$Python = 'C:\Users\Admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe', [switch]$UiOnly)
$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
$release = Join-Path (Split-Path $project -Parent) 'Packing List Manager'
New-Item -ItemType Directory -Path $release -Force | Out-Null
if (-not $UiOnly) {
$env:PYTHONPATH = Join-Path $project '.builddeps'
& $Python -m PyInstaller --noconfirm --clean --onedir --name PackingEngine --distpath (Join-Path $project 'dist') --workpath (Join-Path $project 'build') --specpath $project --add-data ((Join-Path $project 'src\assets') + ';assets') --exclude-module tkinter --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module notebook (Join-Path $project 'src\engine.py')
if ($LASTEXITCODE -ne 0) { throw 'Python packaging failed.' }
$engineFolder = Join-Path $release 'App Files'
New-Item -ItemType Directory -Path $engineFolder -Force | Out-Null
Copy-Item -Path (Join-Path $project 'dist\PackingEngine\*') -Destination $engineFolder -Recurse -Force
}
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ ('/out:' + (Join-Path $release 'Packing List Manager.exe')) ('/win32manifest:' + (Join-Path $project 'src\app.manifest')) ('/win32icon:' + (Join-Path $project 'src\assets\packing-list-icon.ico')) /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll /reference:Microsoft.CSharp.dll (Join-Path $project 'src\PackingListManager.cs') (Join-Path $project 'src\NameDefinitionsForm.cs') (Join-Path $project 'src\NameExcelSession.cs')
if ($LASTEXITCODE -ne 0) { throw 'Windows application compilation failed.' }
Copy-Item -LiteralPath (Join-Path $project 'Employee Guide.txt') -Destination $release -Force
Copy-Item -LiteralPath (Join-Path $project 'src\assets\packing-list-icon.ico') -Destination (Join-Path $release 'Packing List Manager.ico') -Force
Write-Output $release
