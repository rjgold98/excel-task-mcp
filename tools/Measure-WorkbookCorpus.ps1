<#
.SYNOPSIS
Audits a folder of real workbooks and writes ONE anonymized shape report.

.DESCRIPTION
Answers "what do my actual exhibits look like" with evidence instead of recollection, so the
roadmap's open questions - which of range_format is used, whether tables and Power Query matter,
whether multi-workbook link mapping would pay - get decided by a count rather than an argument.
That is the same standard the 46-session demand study set.

It calls the shipped AuditWorkbookFlows operation, which is read-only and proves the workbook's
size and timestamp were identical before and after. Nothing is written to your workbooks.

WHAT THE REPORT CONTAINS, which is what makes it safe to share off the machine:
  - per workbook: a stable pseudonym (workbook-01), size band, sheet count, and counts by item kind
  - per worksheet: a pseudonym (sheet-03), used-range dimensions, and whether it is hidden
  - totals across the corpus
WHAT IT NEVER CONTAINS:
  - cell values or formulas
  - real file names, sheet names, table names, or defined names
  - connection strings, server names, or query text
  - full paths, user names, or machine names
The name map that would let you re-identify anything stays on this machine, in a separate file you
choose whether to keep.

.EXAMPLE
.\tools\Measure-WorkbookCorpus.ps1 -Folder "C:\Work\Exhibits" -OutputFolder "$HOME\Documents\corpus"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Folder,
    [Parameter(Mandatory = $true)][string]$OutputFolder,
    [string]$Server,
    [switch]$Recurse,
    [switch]$KeepNameMap
)

$ErrorActionPreference = "Stop"

if (-not $Server) {
    $Server = Join-Path (Split-Path -Parent $PSScriptRoot) "src\ExcelTask.McpServer\bin\Release\net10.0-windows10.0.17763.0\excel-task-mcp.exe"
}
if (-not (Test-Path $Server)) { throw "Server not found at $Server. Publish or build it first, or pass -Server." }

New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$workbooks = Get-ChildItem -Path $Folder -Include *.xlsx, *.xlsm -File -Recurse:$Recurse
if ($workbooks.Count -eq 0) { throw "No .xlsx or .xlsm files found under $Folder." }
Write-Host "Auditing $($workbooks.Count) workbook(s). Each takes a few seconds - it drives real Excel."

function Invoke-Audit([string]$Path) {
    $argsFile = Join-Path $env:TEMP "corpus-args-$([guid]::NewGuid().ToString('N')).json"
    @{ request = @{ targetWorkbookPath = $Path; operation = @{ kind = "AuditWorkbookFlows"; auditWorkbookFlows = @{} }; mode = "Plan"; workbookBinding = "Isolated" } } |
        ConvertTo-Json -Depth 10 | Set-Content $argsFile -Encoding utf8

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Server
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $proc.StandardInput.WriteLine('{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"corpus","version":"1"}}}')
        $proc.StandardInput.Flush(); Start-Sleep -Milliseconds 500
        $proc.StandardInput.WriteLine('{"jsonrpc":"2.0","method":"notifications/initialized"}')
        $proc.StandardInput.Flush()
        $call = @{ jsonrpc = "2.0"; id = 2; method = "tools/call"; params = @{ name = "excel_task"; arguments = (Get-Content $argsFile -Raw | ConvertFrom-Json) } } | ConvertTo-Json -Depth 40 -Compress
        $proc.StandardInput.WriteLine($call); $proc.StandardInput.Flush()

        $deadline = (Get-Date).AddSeconds(180)
        while ((Get-Date) -lt $deadline) {
            $line = $proc.StandardOutput.ReadLine()
            if ($null -eq $line) { break }
            if ($line -match '"id":2') { return ($line | ConvertFrom-Json).result.structuredContent }
        }
        return $null
    }
    finally {
        try { $proc.StandardInput.Close() } catch {}
        try { if (-not $proc.WaitForExit(8000)) { $proc.Kill($true) } } catch {}
        Remove-Item $argsFile -Force -ErrorAction SilentlyContinue
    }
}

$nameMap = [ordered]@{}
$rows = @()
$index = 0

foreach ($workbook in $workbooks) {
    $index++
    $pseudonym = "workbook-{0:D2}" -f $index
    $nameMap[$pseudonym] = $workbook.FullName
    Write-Host "  [$index/$($workbooks.Count)] $pseudonym"

    $receipt = Invoke-Audit $workbook.FullName
    if ($null -eq $receipt -or $receipt.status -notin @("Planned", "Completed")) {
        $rows += [pscustomobject]@{
            rowType = "workbook"; workbook = $pseudonym; sheetId = ""; sizeBand = ""
            status = $(if ($null -eq $receipt) { "no response" } else { $receipt.status })
            sheets = ""; itemsFound = ""; truncated = ""; kinds = ""; usedRange = ""; unchanged = ""
            note = $(if ($null -eq $receipt) { "the audit did not return" } else { $receipt.summary })
        }
        continue
    }

    $items = @($receipt.audit.items)
    $sizeMb = [math]::Round($workbook.Length / 1MB, 1)
    $band = if ($sizeMb -lt 1) { "<1 MB" } elseif ($sizeMb -lt 10) { "1-10 MB" } elseif ($sizeMb -lt 50) { "10-50 MB" } elseif ($sizeMb -lt 200) { "50-200 MB" } else { ">200 MB" }
    $byKind = $items | Group-Object kind | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }

    $rows += [pscustomobject]@{
        rowType = "workbook"; workbook = $pseudonym; sheetId = ""; sizeBand = $band
        status = $receipt.status
        sheets = @($items | Where-Object { $_.kind -eq "worksheet" }).Count
        itemsFound = $receipt.audit.totalFound
        truncated = $receipt.audit.truncated
        kinds = ($byKind -join " ")
        usedRange = ""
        unchanged = ([bool](@($receipt.checks) | Where-Object { $_.name -eq "workbook-unchanged" -and $_.passed }))
        note = ""
    }

    # Worksheet rows keep the SHAPE and drop the name: used-range dimensions are what decide whether
    # a 400-cell read bound and a 10,000-cell mutation cap are the right size for this kind of work.
    $sheetIndex = 0
    foreach ($item in ($items | Where-Object { $_.kind -eq "worksheet" })) {
        $sheetIndex++
        $sheetId = "sheet-{0:D2}" -f $sheetIndex
        $nameMap["$pseudonym/$sheetId"] = $item.name
        $rows += [pscustomobject]@{
            rowType = "worksheet"; workbook = $pseudonym; sheetId = $sheetId; sizeBand = $band
            status = ""; sheets = ""; itemsFound = ""; truncated = ""; kinds = ""
            usedRange = $(if ($item.detail -match '\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+') { $Matches[0] } else { "(none reported)" })
            unchanged = ""
            note = $(if ($item.detail -match 'hidden') { "hidden" } else { "" })
        }
    }
}

$reportPath = Join-Path $OutputFolder "corpus-shape.csv"
$rows | Export-Csv $reportPath -NoTypeInformation
Write-Host ""
Write-Host "Shape report: $reportPath"

if ($KeepNameMap) {
    $mapPath = Join-Path $OutputFolder "name-map.local.json"
    $nameMap | ConvertTo-Json -Depth 5 | Set-Content $mapPath -Encoding utf8
    Write-Host "Name map (KEEP LOCAL, do not share): $mapPath"
} else {
    Write-Host "Name map discarded. Pass -KeepNameMap to write it locally."
}

Write-Host ""
Write-Host "Totals across the corpus:"
$summaryRows = $rows | Where-Object { $_.itemsFound -ne "" }
$summaryRows | Group-Object status | ForEach-Object { "  status $($_.Name): $($_.Count) workbook(s)" }
$allKinds = @{}
foreach ($row in $summaryRows) {
    foreach ($pair in ($row.kinds -split " " | Where-Object { $_ })) {
        $parts = $pair -split "="
        $allKinds[$parts[0]] = ($allKinds[$parts[0]] ?? 0) + [int]$parts[1]
    }
}
$allKinds.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "  {0,-24} {1}" -f $_.Key, $_.Value }
