param([Parameter(Mandatory=$true)][string]$RequestFile)
$ErrorActionPreference = 'Stop'
$request = Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$documents = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($path in $request.paths) {
        $doc = $null
        try {
            $doc = $word.Documents.Open([string]$path, $false, $true, $false)
            $doc.Repaginate()
            $doc.Fields.Update() | Out-Null
            $setup = $doc.Sections.Item(1).PageSetup
            if ([Math]::Abs($setup.PageWidth - 595.3) -gt 1 -or [Math]::Abs($setup.PageHeight - 841.9) -gt 1) {
                throw 'Document paper size is not A4.'
            }
            $orderCount = 0
            $summaryCount = 0
            $lastOrderPage = 0
            foreach ($table in $doc.Tables) {
                $start = $doc.Range($table.Range.Start, $table.Range.Start)
                $end = $doc.Range($table.Range.End - 1, $table.Range.End - 1)
                $heading = ($table.Cell(1,1).Range.Text -replace '[\r\a]', ' ').Trim()
                if ($heading -eq 'Done') {
                    $summaryCount++
                    if ($start.Information(3) -le $lastOrderPage) { throw 'The product picking summary must start on a new page.' }
                    foreach ($row in $table.Rows) {
                        $rowStart = $doc.Range($row.Range.Start, $row.Range.Start)
                        $rowEnd = $doc.Range($row.Range.End - 1, $row.Range.End - 1)
                        if ($rowStart.Information(3) -ne $rowEnd.Information(3)) { throw 'A product summary row is too large to fit on one A4 page.' }
                    }
                    continue
                }
                if ($heading -notmatch '^Order (ID|Number):' -or $summaryCount -ne 0) { throw 'Unexpected table in packing document.' }
                $orderCount++
                $lastOrderPage = $end.Information(3)
                if ($start.Information(3) -ne $end.Information(3)) {
                    throw "An order is too large to fit together on one A4 page: $heading. Review this order before printing."
                }
            }
            if ($summaryCount -ne 1) { throw 'Expected one product picking summary per document.' }
            $pages = $doc.ComputeStatistics(2)
            if ($request.export_pdf) {
                $doc.ExportAsFixedFormat([IO.Path]::ChangeExtension([string]$path, '.pdf'), 17)
            }
            $documents += @{path=[string]$path;pages=$pages;orders=$orderCount;summaries=$summaryCount}
        } finally {
            if ($null -ne $doc) { $doc.Close($false) }
        }
    }
    $response = @{ok=$true;documents=@($documents)}
} catch {
    $response = @{ok=$false;error=$_.Exception.Message}
} finally {
    if ($null -ne $word) { $word.Quit() }
}
$response | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $request.result -Encoding UTF8
