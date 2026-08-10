<#
.SYNOPSIS
    Validates ExcelTask on a managed work computer and records comparable measurements.

.DESCRIPTION
    Runs the published excel-task-mcp.exe over real MCP stdio against disposable
    workbooks, then writes one JSON and one Markdown report. Nothing outside the
    output directory and the system temp folder is touched, no workbook of yours is
    opened, and no Excel or security setting is changed.

    Needs only the release ZIP and desktop Excel. No repository, SDK, or Copilot.

.EXAMPLE
    .\Invoke-FieldCheck.ps1 -ServerPath C:\Tools\ExcelTask\excel-task-mcp.exe
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ServerPath,
    [string]$CompareServerPath,
    [string[]]$CompareServerArguments = @(),
    [string]$OutputDirectory = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'ExcelTask-FieldCheck')
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $ServerPath)) { throw "Server executable not found: $ServerPath" }
$ServerPath = (Resolve-Path $ServerPath).Path
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("ExcelTaskField-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null

$report = [ordered]@{
    startedUtc  = (Get-Date).ToUniversalTime().ToString('o')
    serverPath  = $ServerPath
    environment = [ordered]@{}
    toolSurface = [ordered]@{}
    operations  = @()
    notes       = @()
}

function Get-ExcelProcessIds { @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }

# ---------------------------------------------------------------- environment
Write-Host "[1/4] Reading environment (read-only)..." -ForegroundColor Cyan
$env0 = $report.environment
$env0.computerName   = $env:COMPUTERNAME
$env0.osVersion      = [Environment]::OSVersion.VersionString
$env0.dotnetRuntimes = @(& dotnet --list-runtimes 2>$null | Select-Object -First 12) -join '; '
if (-not $env0.dotnetRuntimes) { $env0.dotnetRuntimes = 'dotnet CLI not on PATH (fine: the release is self-contained)' }

foreach ($ver in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $key) {
        $sec = Get-ItemProperty $key
        $env0.officeVersion = $ver
        # 1 = trusted. Reading or editing VBA is impossible without it.
        $env0.accessVBOM    = $sec.AccessVBOM
        # 1 enable all, 2 disable w/ notification, 3 signed only, 4 disable all.
        $env0.vbaWarnings   = $sec.VBAWarnings
    }
    $policy = "HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $policy) {
        $p = Get-ItemProperty $policy
        if ($null -ne $p.AccessVBOM)  { $env0.policyAccessVBOM  = $p.AccessVBOM }
        if ($null -ne $p.VBAWarnings) { $env0.policyVBAWarnings = $p.VBAWarnings }
        $report.notes += 'Group Policy sets Excel macro security here; the organization controls it, not this product.'
    }
}
if ($null -eq $env0.accessVBOM) { $env0.accessVBOM = 'not set' }

$preExisting = Get-ExcelProcessIds
$env0.excelRunningBefore = $preExisting.Count
try {
    $probe = New-Object -ComObject Excel.Application
    $env0.excelVersion  = $probe.Version
    $env0.excelBuild    = $probe.Build
    $env0.excelBitness  = if ([Environment]::Is64BitProcess) { $probe.OperatingSystem } else { 'n/a' }
    $addins = @()
    foreach ($a in $probe.COMAddIns) { if ($a.Connect) { $addins += $a.ProgId } }
    $env0.connectedComAddins = $addins -join '; '
    $probe.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($probe)
} catch {
    $env0.excelVersion = "COM probe failed: $($_.Exception.Message)"
}

# ------------------------------------------------------------------ MCP layer
$script:proc = $null
$script:nextId = 0

function Start-Server([string]$path = $ServerPath, [string[]]$arguments = @()) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $path
    foreach ($a in $arguments) { [void]$psi.ArgumentList.Add($a) }
    $psi.WorkingDirectory = $work
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardInputEncoding = [System.Text.Encoding]::UTF8
    $script:proc = [System.Diagnostics.Process]::Start($psi)
}

function Send-Rpc([string]$method, $params, [switch]$Notification) {
    $msg = [ordered]@{ jsonrpc = '2.0'; method = $method }
    if (-not $Notification) { $script:nextId++; $msg.id = $script:nextId }
    if ($null -ne $params) { $msg.params = $params }
    $script:proc.StandardInput.WriteLine(($msg | ConvertTo-Json -Depth 40 -Compress))
    $script:proc.StandardInput.Flush()
    if ($Notification) { return $null }

    # MCP stdio is newline-delimited JSON; skip anything that is not our reply.
    $deadline = (Get-Date).AddSeconds(180)
    while ((Get-Date) -lt $deadline) {
        $line = $script:proc.StandardOutput.ReadLine()
        if ($null -eq $line) { throw "Server closed stdout while awaiting '$method'." }
        if ($line.Trim() -eq '') { continue }
        try { $parsed = $line | ConvertFrom-Json } catch { continue }
        if ($parsed.PSObject.Properties.Name -contains 'id' -and $parsed.id -eq $script:nextId) { return $parsed }
    }
    throw "Timed out awaiting '$method'."
}

# Handshake with any MCP stdio server and measure the tool surface it advertises.
# The tools/list payload is what a client must carry every session, so its size is the
# fairest like-for-like context cost between servers.
function Measure-ToolSurface([string]$path, [string[]]$arguments, [string]$label) {
    $saved = $script:proc; $savedId = $script:nextId
    $script:proc = $null; $script:nextId = 0
    $surface = [ordered]@{ label = $label; path = $path }
    try {
        Start-Server $path $arguments
        $init = Send-Rpc 'initialize' @{
            protocolVersion = '2025-06-18'; capabilities = @{}
            clientInfo = @{ name = 'ExcelTask-FieldCheck'; version = '1.0.0' }
        }
        Send-Rpc 'notifications/initialized' $null -Notification | Out-Null
        $tools = Send-Rpc 'tools/list' @{}
        $json = ($tools.result | ConvertTo-Json -Depth 60 -Compress)
        $surface.serverName    = $init.result.serverInfo.name
        $surface.serverVersion = $init.result.serverInfo.version
        $surface.toolCount     = @($tools.result.tools).Count
        $surface.toolListBytes = [System.Text.Encoding]::UTF8.GetByteCount($json)
    } catch {
        $surface.error = $_.Exception.Message
    } finally {
        if ($script:proc -and -not $script:proc.HasExited) {
            try { $script:proc.StandardInput.Close() } catch { }
            if (-not $script:proc.WaitForExit(5000)) { try { $script:proc.Kill() } catch { } }
        }
        $script:proc = $saved; $script:nextId = $savedId
    }
    return $surface
}

function Invoke-ExcelTask([hashtable]$request, [string]$label) {
    $before = Get-ExcelProcessIds
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $status = 'Error'; $summary = ''; $checks = @(); $errorText = $null
    try {
        $response = Send-Rpc 'tools/call' @{ name = 'excel_task'; arguments = @{ request = $request } }
        if ($response.PSObject.Properties.Name -contains 'error') {
            $errorText = ($response.error | ConvertTo-Json -Depth 10 -Compress)
        } else {
            $sc = $response.result.structuredContent
            $status  = $sc.status
            $summary = $sc.summary
            $checks  = @($sc.checks | ForEach-Object { "$($_.name)=$($_.passed)" })
        }
    } catch {
        $errorText = $_.Exception.Message
    }
    $timer.Stop()
    Start-Sleep -Milliseconds 700
    $leaked = @(Get-ExcelProcessIds | Where-Object { $_ -notin $before })

    $script:report.operations += [ordered]@{
        label          = $label
        status         = $status
        elapsedSeconds = [math]::Round($timer.Elapsed.TotalSeconds, 2)
        leakedExcel    = $leaked.Count
        summary        = $summary
        checks         = ($checks -join '; ')
        error          = $errorText
    }
    $colour = if ($status -in @('Completed', 'Planned')) { 'Green' } else { 'Yellow' }
    Write-Host ("      {0,-34} {1,-16} {2,6}s  leaked={3}" -f $label, $status, [math]::Round($timer.Elapsed.TotalSeconds,1), $leaked.Count) -ForegroundColor $colour
    return $response
}

# ------------------------------------------------------------------- fixtures
# PowerShell releases COM references lazily, so a fixture Excel can outlive Quit().
# These are this script's own processes: they are tracked and force-closed, and they
# are excluded from the leak figures so the report never blames the product for them.
$script:fixturePids = @()

function New-FixtureExcel {
    $before = Get-ExcelProcessIds
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false; $app.DisplayAlerts = $false
    foreach ($id in (Get-ExcelProcessIds | Where-Object { $_ -notin $before })) {
        if ($id -notin $script:fixturePids) { $script:fixturePids += $id }
    }
    return $app
}

function Close-FixtureExcel($app) {
    try { $app.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 400
    foreach ($id in $script:fixturePids) {
        Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function New-Fixtures {
    $app = New-FixtureExcel
    try {
        $t = $app.Workbooks.Add()
        $t.Worksheets.Item(1).Name = 'Model'
        $sheet = $t.Worksheets.Item(1)
        $sheet.Range('A1:D1').Formula = @(@('=1','=2','=3','=4'))
        $sheet.Range('A2').Formula = '=A1*2'
        $sheet.Range('B2').Formula = '=B1*2'
        $t.SaveAs((Join-Path $work 'target.xlsx'))
        $t.Close($false)

        $r = $app.Workbooks.Add()
        $r.Worksheets.Item(1).Name = 'Reference'
        $r.Worksheets.Item(1).Range('A1:A3').Formula = @(@('=ROW()'), @(''), @('=ROW()'))
        $r.SaveAs((Join-Path $work 'reference.xlsx'))
        $r.Close($false)
    } finally {
        Close-FixtureExcel $app
    }
}

function New-MacroFixture {
    $app = New-FixtureExcel
    try {
        $wb = $app.Workbooks.Add()
        $wb.SaveAs((Join-Path $work 'macro-target.xlsm'), 52)
        $mod = $wb.VBProject.VBComponents.Add(1)
        $mod.Name = 'FieldModule'
        $mod.CodeModule.AddFromString("Public Sub FieldMarker()`r`n    ThisWorkbook.Worksheets(1).Range(""A1"").Value2 = ""before""`r`nEnd Sub")
        $wb.Save(); $wb.Close($false)
        return $true
    } catch {
        $script:report.notes += "Macro fixture could not be created, so macro editing was not measured: $($_.Exception.Message)"
        return $false
    } finally {
        Close-FixtureExcel $app
    }
}

# ----------------------------------------------------------------------- run
try {
    Write-Host "[2/4] Building disposable workbooks..." -ForegroundColor Cyan
    New-Fixtures
    $macroReady = New-MacroFixture

    Write-Host "[3/4] Starting the server and measuring the tool surface..." -ForegroundColor Cyan
    Start-Server
    $initTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $init = Send-Rpc 'initialize' @{
        protocolVersion = '2025-06-18'
        capabilities    = @{}
        clientInfo      = @{ name = 'ExcelTask-FieldCheck'; version = '1.0.0' }
    }
    Send-Rpc 'notifications/initialized' $null -Notification | Out-Null
    $initTimer.Stop()

    $tools = Send-Rpc 'tools/list' @{}
    $toolsJson = ($tools.result | ConvertTo-Json -Depth 60 -Compress)
    $report.toolSurface.serverName        = $init.result.serverInfo.name
    $report.toolSurface.serverVersion     = $init.result.serverInfo.version
    $report.toolSurface.handshakeSeconds  = [math]::Round($initTimer.Elapsed.TotalSeconds, 2)
    $report.toolSurface.toolCount         = @($tools.result.tools).Count
    $report.toolSurface.toolNames         = (@($tools.result.tools | ForEach-Object { $_.name }) -join ', ')
    # The whole tools/list payload is what a client must carry in context every session,
    # so its size is the fairest like-for-like cost comparison between MCP servers.
    $report.toolSurface.toolListBytes     = [System.Text.Encoding]::UTF8.GetByteCount($toolsJson)
    Write-Host ("      {0} tool(s), tools/list payload = {1:N0} bytes" -f $report.toolSurface.toolCount, $report.toolSurface.toolListBytes) -ForegroundColor Green

    if ($CompareServerPath) {
        Write-Host "      Measuring the comparison server's tool surface..." -ForegroundColor Cyan
        $other = Measure-ToolSurface $CompareServerPath $CompareServerArguments 'comparison server'
        $report.comparison = $other
        if ($other.error) {
            Write-Host "      Comparison server failed to report: $($other.error)" -ForegroundColor Yellow
        } else {
            $ratio = if ($report.toolSurface.toolListBytes -gt 0) { [math]::Round($other.toolListBytes / $report.toolSurface.toolListBytes, 1) } else { 0 }
            Write-Host ("      Comparison: {0} tool(s), {1:N0} bytes  ({2}x ExcelTask)" -f $other.toolCount, $other.toolListBytes, $ratio) -ForegroundColor Green
        }
    }

    Write-Host "[4/4] Exercising each operation on disposable workbooks..." -ForegroundColor Cyan
    $target = Join-Path $work 'target.xlsx'
    $reference = Join-Path $work 'reference.xlsx'

    Invoke-ExcelTask @{
        targetWorkbookPath = $target
        operation = @{ kind = 'CopyExhibit'; copyExhibit = @{
            referenceWorkbookPath = $reference; referenceWorksheet = 'Reference'
            newWorksheetName = 'Exhibit A'; repairRanges = @('A1:A3') } }
        mode = 'Plan'; workbookBinding = 'Isolated'; save = 'Copy'
        outputWorkbookPath = (Join-Path $work 'out-copy.xlsx'); overwriteConfirmed = $false
    } 'CopyExhibit (Plan)' | Out-Null

    Invoke-ExcelTask @{
        targetWorkbookPath = $target
        operation = @{ kind = 'CopyExhibit'; copyExhibit = @{
            referenceWorkbookPath = $reference; referenceWorksheet = 'Reference'
            newWorksheetName = 'Exhibit A'; repairRanges = @('A1:A3') } }
        mode = 'Apply'; workbookBinding = 'Isolated'; save = 'Copy'
        outputWorkbookPath = (Join-Path $work 'out-copy.xlsx'); overwriteConfirmed = $true
    } 'CopyExhibit (Apply)' | Out-Null

    Invoke-ExcelTask @{
        targetWorkbookPath = $target
        operation = @{ kind = 'ExtendFormulaSeries'; extendFormulaSeries = @{
            worksheetName = 'Model'; direction = 'Right'
            evidenceRange = 'A2:B2'; destinationRange = 'C2:D2' } }
        mode = 'Apply'; workbookBinding = 'Isolated'; save = 'Copy'
        outputWorkbookPath = (Join-Path $work 'out-extend.xlsx'); overwriteConfirmed = $true
    } 'ExtendFormulaSeries (Apply)' | Out-Null

    if ($macroReady) {
        $macroTarget = Join-Path $work 'macro-target.xlsm'
        $plan = Invoke-ExcelTask @{
            targetWorkbookPath = $macroTarget
            operation = @{ kind = 'EditMacroProcedure'; editMacroProcedure = @{
                componentName = 'FieldModule'; procedureName = 'FieldMarker' } }
            mode = 'Plan'; workbookBinding = 'Isolated'; save = 'Copy'
            outputWorkbookPath = (Join-Path $work 'out-macro.xlsm'); overwriteConfirmed = $false
        } 'EditMacroProcedure (Plan)'

        $hash = $null
        try { $hash = $plan.result.structuredContent.macroProcedure.sha256 } catch { }
        if ($hash) {
            Invoke-ExcelTask @{
                targetWorkbookPath = $macroTarget
                operation = @{ kind = 'EditMacroProcedure'; editMacroProcedure = @{
                    componentName = 'FieldModule'; procedureName = 'FieldMarker'
                    expectedProcedureSha256 = $hash
                    replacementSource = "Public Sub FieldMarker()`n    ThisWorkbook.Worksheets(1).Range(""A1"").Value2 = ""after""`nEnd Sub"
                    runAfterEdit = $true } }
                mode = 'Apply'; workbookBinding = 'Isolated'; save = 'Copy'
                outputWorkbookPath = (Join-Path $work 'out-macro.xlsm'); overwriteConfirmed = $true
            } 'EditMacroProcedure (Apply+Run)' | Out-Null
        } else {
            $report.notes += 'Macro Plan returned no procedure hash, so the Apply step was skipped. The check detail in the Plan row explains why; a blocked Trust Center setting is the usual cause.'
        }
    }
}
finally {
    if ($script:proc -and -not $script:proc.HasExited) {
        try { $script:proc.StandardInput.Close() } catch { }
        if (-not $script:proc.WaitForExit(5000)) { $script:proc.Kill() }
    }
    foreach ($id in $script:fixturePids) {
        Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 800
    $remaining = Get-ExcelProcessIds
    $report.environment.excelRunningAfter = $remaining.Count
    # Excludes this script's own fixture processes; what is left is attributable to the product.
    $report.environment.excelLeakedByProduct = @($remaining | Where-Object { $_ -notin $preExisting -and $_ -notin $script:fixturePids }).Count
    if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}

# -------------------------------------------------------------------- report
$report.completedUtc = (Get-Date).ToUniversalTime().ToString('o')
$stamp = Get-Date -Format 'yyyy-MM-dd-HHmm'
$jsonPath = Join-Path $OutputDirectory "exceltask-fieldcheck-$stamp.json"
$mdPath   = Join-Path $OutputDirectory "exceltask-fieldcheck-$stamp.md"
$report | ConvertTo-Json -Depth 12 | Set-Content -Path $jsonPath -Encoding UTF8

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# ExcelTask field check")
[void]$md.AppendLine()
[void]$md.AppendLine("Run $($report.startedUtc) on $($report.environment.computerName).")
[void]$md.AppendLine()
[void]$md.AppendLine("## Environment")
[void]$md.AppendLine()
[void]$md.AppendLine('| Item | Value |')
[void]$md.AppendLine('|---|---|')
foreach ($k in $report.environment.Keys) { [void]$md.AppendLine("| $k | $($report.environment[$k]) |") }
[void]$md.AppendLine()
[void]$md.AppendLine("## Tool surface")
[void]$md.AppendLine()
[void]$md.AppendLine('| Item | Value |')
[void]$md.AppendLine('|---|---|')
foreach ($k in $report.toolSurface.Keys) { [void]$md.AppendLine("| $k | $($report.toolSurface[$k]) |") }
if ($report.Contains('comparison')) {
    [void]$md.AppendLine()
    [void]$md.AppendLine("## Comparison server")
    [void]$md.AppendLine()
    [void]$md.AppendLine('| Item | Value |')
    [void]$md.AppendLine('|---|---|')
    foreach ($k in $report.comparison.Keys) { [void]$md.AppendLine("| $k | $($report.comparison[$k]) |") }
    if (-not $report.comparison.error -and $report.toolSurface.toolListBytes -gt 0) {
        $r = [math]::Round($report.comparison.toolListBytes / $report.toolSurface.toolListBytes, 1)
        [void]$md.AppendLine()
        [void]$md.AppendLine("Its tools/list payload is **${r}x** the size of ExcelTask's. That payload is carried in context every session, before any work is requested.")
    }
}
[void]$md.AppendLine()
[void]$md.AppendLine("## Operations")
[void]$md.AppendLine()
[void]$md.AppendLine('| Operation | Status | Seconds | Leaked Excel | Summary |')
[void]$md.AppendLine('|---|---|---|---|---|')
foreach ($op in $report.operations) {
    $s = "$($op.summary)"; if ($op.error) { $s = "ERROR: $($op.error)" }
    [void]$md.AppendLine("| $($op.label) | $($op.status) | $($op.elapsedSeconds) | $($op.leakedExcel) | $s |")
}
if ($report.notes.Count) {
    [void]$md.AppendLine()
    [void]$md.AppendLine("## Notes")
    [void]$md.AppendLine()
    foreach ($n in $report.notes) { [void]$md.AppendLine("- $n") }
}
[void]$md.AppendLine()
[void]$md.AppendLine("## Checks in full")
[void]$md.AppendLine()
foreach ($op in $report.operations) { [void]$md.AppendLine("- **$($op.label)**: $($op.checks)") }
$md.ToString() | Set-Content -Path $mdPath -Encoding UTF8

Write-Host ""
Write-Host "Report written:" -ForegroundColor Cyan
Write-Host "  $mdPath"
Write-Host "  $jsonPath"
$failed = @($report.operations | Where-Object { $_.status -notin @('Completed', 'Planned') })
if ($failed.Count -or $report.environment.excelLeakedByProduct -gt 0) {
    Write-Host "Some operations did not complete, or an Excel process was left behind. Send both files back." -ForegroundColor Yellow
} else {
    Write-Host "All operations completed and no Excel process was left behind." -ForegroundColor Green
}
