# Two questions about Application.Calculation, one of them a safety question.
#
# 1. Under the write strategy ExcelTask actually ships - repairs grouped by identical R1C1 formula
#    and written as multi-area ranges - is manual calculation worth anything? The per-cell loop in
#    excel-config-probe.ps1 is not what this code does, so its 17% is not a number we can bank.
#
# 2. Does the setting leak into the saved file? Excel records a calculation mode in the workbook.
#    If setting manual for speed means handing the user back a workbook that no longer recalculates,
#    the setting is not a trade worth making at any speed - that is precisely the quiet damage this
#    project exists not to do.
#
# Trials are interleaved rather than grouped, because the first probe showed launch time climbing
# monotonically with run order: measuring all of one variant before the next attributes that drift
# to the variant.

$ErrorActionPreference = 'Stop'

$Rows = 2000
$Trials = 4
$Automatic = -4105
$Manual = -4135

$root = Join-Path $env:TEMP ("exceltask-calc-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null

function New-Fixture {
    param([string]$Path)

    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    try {
        $book = $app.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        $sheet.Name = 'Data'
        $values = New-Object 'object[,]' $Rows, 1
        for ($i = 0; $i -lt $Rows; $i++) { $values[$i, 0] = $i + 1 }
        $sheet.Range("A1:A$Rows").Value2 = $values
        # A dependency chain, so recalculation has real work to do rather than one isolated cell.
        $sheet.Range("C1").Formula = '=SUM(A1:A' + $Rows + ')'
        $sheet.Range("D1").Formula = '=C1*2'
        $book.SaveAs($Path, 51)
        $book.Close($false)
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

# The shipping strategy: one identical R1C1 formula across many cells, written as multi-area ranges
# whose joined address stays under Excel's 255-character argument limit.
function Write-BatchedRepairs {
    param($Sheet, [int]$Rows)

    $builder = New-Object Text.StringBuilder
    $written = 0
    for ($i = 1; $i -le $Rows; $i++) {
        $address = "B$i"
        if ($builder.Length -gt 0 -and $builder.Length + 1 + $address.Length -gt 255) {
            $Sheet.Range($builder.ToString()).FormulaR1C1 = '=RC[-1]*2+1'
            $written++
            [void]$builder.Clear()
        }
        if ($builder.Length -gt 0) { [void]$builder.Append(',') }
        [void]$builder.Append($address)
    }
    if ($builder.Length -gt 0) {
        $Sheet.Range($builder.ToString()).FormulaR1C1 = '=RC[-1]*2+1'
        $written++
    }
    $written
}

function Measure-Run {
    param([bool]$UseManual, [string]$FixturePath)

    $app = New-Object -ComObject Excel.Application
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.EnableEvents = $false
        $app.AutomationSecurity = 3
        $book = $app.Workbooks.Open($FixturePath, 0, $false)
        $sheet = $book.Worksheets.Item('Data')

        $watch = [Diagnostics.Stopwatch]::StartNew()
        if ($UseManual) { $app.Calculation = -4135 }
        $batches = Write-BatchedRepairs -Sheet $sheet -Rows $Rows
        $writeMs = $watch.ElapsedMilliseconds

        $watch.Restart()
        $app.CalculateFull()
        $calcMs = $watch.ElapsedMilliseconds

        # Restored before the save, so the mode the file records is the one it arrived with.
        # Read-back of the prior value is skipped here only because PowerShell's COM adapter
        # mis-binds the setter when handed it; the shipping code captures and restores it.
        $watch.Restart()
        if ($UseManual) { $app.Calculation = -4105 }
        $restoreMs = $watch.ElapsedMilliseconds

        $watch.Restart()
        $book.Save()
        $saveMs = $watch.ElapsedMilliseconds
        $book.Close($false)

        [PSCustomObject]@{
            Manual = $UseManual
            Batches = $batches
            WriteMs = $writeMs
            CalcMs = $calcMs
            RestoreMs = $restoreMs
            SaveMs = $saveMs
            TotalMs = $writeMs + $calcMs + $restoreMs + $saveMs
        }
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

# Question 2, asked directly: save a workbook while the application is in manual mode without
# restoring, then reopen it in a fresh instance and ask what mode that instance came up in.
function Test-CalcModeLeak {
    param([string]$FixturePath)

    $app = New-Object -ComObject Excel.Application
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Open($FixturePath, 0, $false)
        $app.Calculation = $Manual
        $book.Worksheets.Item('Data').Range('B1').FormulaR1C1 = '=RC[-1]*2+1'
        $book.Save()
        $book.Close($false)
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }

    Start-Sleep -Milliseconds 500
    $verify = New-Object -ComObject Excel.Application
    try {
        $verify.Visible = $false
        $verify.DisplayAlerts = $false
        $book = $verify.Workbooks.Open($FixturePath, 0, $true)
        $mode = $verify.Calculation
        $book.Close($false)
        if ($mode -eq $Manual) { 'LEAKS - the saved file reopens in manual calculation' }
        elseif ($mode -eq $Automatic) { 'safe - the saved file reopens in automatic calculation' }
        else { "unexpected mode: $mode" }
    }
    finally {
        $verify.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($verify)
    }
}

try {
    $results = @()
    for ($trial = 1; $trial -le $Trials; $trial++) {
        foreach ($manual in @($false, $true)) {
            $fixture = Join-Path $root ("f-" + [Guid]::NewGuid().ToString('N').Substring(0, 8) + ".xlsx")
            New-Fixture -Path $fixture
            $results += Measure-Run -UseManual $manual -FixturePath $fixture
            Remove-Item $fixture -Force
        }
    }

    Write-Host ("workload: {0} batched R1C1 repairs over {1} rows, {2} interleaved trials" -f `
        ($results[0].Batches), $Rows, $Trials)
    Write-Host ""
    $results | Group-Object Manual | ForEach-Object {
        $runs = $_.Group
        [PSCustomObject]@{
            Calculation = if ($_.Name -eq 'True') { 'manual during writes' } else { 'automatic (today)' }
            WriteMs   = [int](($runs.WriteMs   | Sort-Object)[[int]($Trials / 2)])
            CalcMs    = [int](($runs.CalcMs    | Sort-Object)[[int]($Trials / 2)])
            RestoreMs = [int](($runs.RestoreMs | Sort-Object)[[int]($Trials / 2)])
            SaveMs    = [int](($runs.SaveMs    | Sort-Object)[[int]($Trials / 2)])
            TotalMs   = [int](($runs.TotalMs   | Sort-Object)[[int]($Trials / 2)])
        }
    } | Format-Table -AutoSize | Out-String -Width 200 | Write-Host

    $leakFixture = Join-Path $root "leak.xlsx"
    New-Fixture -Path $leakFixture
    Write-Host ("calculation mode leak into saved file: " + (Test-CalcModeLeak -FixturePath $leakFixture))
}
finally {
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}
