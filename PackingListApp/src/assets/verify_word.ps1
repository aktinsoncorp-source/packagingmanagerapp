param([Parameter(Mandatory=$true)][string]$RequestFile)
$ErrorActionPreference = 'Stop'
$request = Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$documents = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.ScreenUpdating = $false
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
                # Cache COM values instead of repeatedly crossing into Word.
                $tableRange = $table.Range
                $tableStart = $tableRange.Start
                $tableEnd = $tableRange.End - 1
                $start = $doc.Range($tableStart, $tableStart)
                $end = $doc.Range($tableEnd, $tableEnd)
                $startPage = $start.Information(3)
                $endPage = $end.Information(3)
                $heading = ($table.Cell(1,1).Range.Text -replace '[\r\a]', ' ').Trim()
                if ($heading -eq 'Done') {
                    $summaryCount++
                    if ($startPage -le $lastOrderPage) { throw 'The product picking summary must start on a new page.' }
                    # A summary wholly on one page also proves all its rows stay together.
                    # Multi-page summaries still receive the original per-row checks.
                    if ($startPage -ne $endPage) {
                        foreach ($row in $table.Rows) {
                            $rowRange = $row.Range
                            $rowStartPosition = $rowRange.Start
                            $rowEndPosition = $rowRange.End - 1
                            $rowStart = $doc.Range($rowStartPosition, $rowStartPosition)
                            $rowEnd = $doc.Range($rowEndPosition, $rowEndPosition)
                            if ($rowStart.Information(3) -ne $rowEnd.Information(3)) { throw 'A product summary row is too large to fit on one A4 page.' }
                        }
                    }
                    continue
                }
                if ($heading -notmatch '^Order (ID|Number):' -or $summaryCount -ne 0) { throw 'Unexpected table in packing document.' }
                $orderCount++
                $lastOrderPage = $endPage
                if ($startPage -ne $endPage) {
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
