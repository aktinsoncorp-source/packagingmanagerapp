param([Parameter(Mandatory=$true)][string]$WorkbookPath, [switch]$HiddenExcel)
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$excelApp = $null
$ownedWorkbook = $null

function Send-Reply($reply) {
    [Console]::WriteLine(($reply | ConvertTo-Json -Compress -Depth 6))
    [Console]::Out.Flush()
}

function Test-OwnedWorkbook {
    try {
        if ($null -eq $excelApp -or $null -eq $ownedWorkbook -or $excelApp.Workbooks.Count -eq 0) { return $false }
        return [String]::Equals([string]$ownedWorkbook.FullName, $WorkbookPath, [StringComparison]::OrdinalIgnoreCase)
    } catch {
        $code = $_.Exception.HResult
        if ($code -eq -2147418111 -or $code -eq -2147417846) {
            throw 'Excel is busy. Finish editing the cell or close its dialog, then click Apply Excel edits again.'
        }
        return $false
    }
}

function Close-OwnedWorkbook {
    if ($null -ne $script:ownedWorkbook) {
        try { $script:ownedWorkbook.Close($false) } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:ownedWorkbook) } catch { }
        $script:ownedWorkbook = $null
    }
    if ($null -ne $script:excelApp) {
        # A user may have opened another workbook in this instance. Leave it open.
        try { if ($script:excelApp.Workbooks.Count -eq 0) { $script:excelApp.Quit() } } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excelApp) } catch { }
        $script:excelApp = $null
    }
}

try {
    try {
        $excelType = [Type]::GetTypeFromProgID('Excel.Application')
        if ($null -eq $excelType) { throw 'Microsoft Excel is required for Edit in Excel. Install desktop Excel and try again.' }
        $excelApp = [Activator]::CreateInstance($excelType)
        $ownedWorkbook = $excelApp.Workbooks.Open($WorkbookPath)
        $sheet = $ownedWorkbook.Worksheets.Item('Names')
        # Excel blocks AutoFit on a protected sheet. Never remove row protection.
        if (-not $sheet.ProtectContents) { [void]$sheet.Rows.AutoFit() }
        [void]$sheet.Range('C2').Select()
        $excelApp.Visible = -not $HiddenExcel
        Send-Reply @{ok=$true;protected=[bool]$sheet.ProtectContents;rows=[int]$sheet.UsedRange.Rows.Count}
    } catch {
        Send-Reply @{ok=$false;error=$_.Exception.Message}
        exit 1
    }
    while ($null -ne ($line = [Console]::ReadLine())) {
        try {
            $request = $line | ConvertFrom-Json
            if ($request.action -eq 'close') {
                Close-OwnedWorkbook
                Send-Reply @{ok=$true}
                break
            }
            if (-not (Test-OwnedWorkbook)) {
                Send-Reply @{ok=$false;missing=$true}
                continue
            }
            if ($request.action -eq 'save') {
                $ownedWorkbook.Save()
                Send-Reply @{ok=$true}
            } else {
                Send-Reply @{ok=$false;error='Unknown Excel action.'}
            }
        } catch {
            Send-Reply @{ok=$false;error=$_.Exception.Message}
        }
    }
} finally {
    Close-OwnedWorkbook
}
