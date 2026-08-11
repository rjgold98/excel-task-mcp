# Measures what each Excel Application setting is actually worth on this machine.
#
# ExcelTask configures a private Excel instance before it does any work. The settings were chosen
# from reputation rather than measurement, and reputation is a poor guide here: some of them cost
# nothing, some save seconds, and at least one is widely recommended and does nothing at all when
# the instance is invisible. This probe runs the same workload under each variant so the choice can
# be made from numbers.
#
# Runs against real desktop Excel. No workbook contents leave this machine: the fixture is generated
# here and deleted at the end.

$ErrorActionPreference = 'Stop'

$FormulaCount = 600
$Trials = 3

$root = Join-Path $env:TEMP ("exceltask-probe-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
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
        # Seed values the formulas will depend on, written as one array so the fixture itself is
        # not what we end up measuring.
        $values = New-Object 'object[,]' $FormulaCount, 1
        for ($i = 0; $i -lt $FormulaCount; $i++) { $values[$i, 0] = $i + 1 }
        $sheet.Range("A1:A$FormulaCount").Value2 = $values
        $book.SaveAs($Path, 51)
        $book.Close($false)
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

# Each variant is the baseline plus one change, so a difference is attributable to that change.
$variants = [ordered]@{
    'baseline (shipping today)' = { param($app) }
    '+ Calculation = manual'    = { param($app) $app.Calculation = -4135 }
    '+ ScreenUpdating = false'  = { param($app) $app.ScreenUpdating = $false }
    '+ PrintCommunication=false'= { param($app) $app.PrintCommunication = $false }
    '+ DisplayStatusBar = false'= { param($app) $app.DisplayStatusBar = $false }
    'all of the above'          = { param($app)
        $app.Calculation = -4135
        $app.ScreenUpdating = $false
        $app.PrintCommunication = $false
        $app.DisplayStatusBar = $false }
}

function Measure-Variant {
    param([string]$Name, [scriptblock]$Configure, [string]$FixturePath)

    $watch = [Diagnostics.Stopwatch]::StartNew()
    $app = New-Object -ComObject Excel.Application
    $launchMs = $watch.ElapsedMilliseconds

    try {
        # The settings ExcelTask always applies, before any workbook is open.
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.EnableEvents = $false
        $app.AutomationSecurity = 3

        $watch.Restart()
        # UpdateLinks 0 and ReadOnly false, matching the shipping open. The remaining arguments are
        # left off rather than passed as Missing, which the COM binder here refuses.
        $book = $app.Workbooks.Open($FixturePath, 0, $false)
        $openMs = $watch.ElapsedMilliseconds

        # Calculation cannot be assigned with no workbook open, so every variant is applied here
        # rather than beside the settings above.
        & $Configure $app

        $sheet = $book.Worksheets.Item('Data')
        $watch.Restart()
        for ($i = 1; $i -le $FormulaCount; $i++) {
            $sheet.Range("B$i").Formula = "=A$i*2+1"
        }
        $writeMs = $watch.ElapsedMilliseconds

        $watch.Restart()
        $app.CalculateFull()
        $calcMs = $watch.ElapsedMilliseconds

        $watch.Restart()
        $book.Save()
        $saveMs = $watch.ElapsedMilliseconds

        $book.Close($false)
        [PSCustomObject]@{
            Variant = $Name
            LaunchMs = $launchMs
            OpenMs = $openMs
            WriteMs = $writeMs
            CalcMs = $calcMs
            SaveMs = $saveMs
            TotalMs = $launchMs + $openMs + $writeMs + $calcMs + $saveMs
        }
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

try {
    $results = @()
    foreach ($name in $variants.Keys) {
        $runs = @()
        for ($trial = 1; $trial -le $Trials; $trial++) {
            $fixture = Join-Path $root ("fixture-" + [Guid]::NewGuid().ToString('N').Substring(0, 8) + ".xlsx")
            New-Fixture -Path $fixture
            $runs += Measure-Variant -Name $name -Configure $variants[$name] -FixturePath $fixture
            Remove-Item $fixture -Force
        }
        # Median, not mean: one antivirus scan or one background save should not decide this.
        $results += [PSCustomObject]@{
            Variant = $name
            LaunchMs = [int](($runs.LaunchMs | Sort-Object)[[int]($Trials / 2)])
            OpenMs   = [int](($runs.OpenMs   | Sort-Object)[[int]($Trials / 2)])
            WriteMs  = [int](($runs.WriteMs  | Sort-Object)[[int]($Trials / 2)])
            CalcMs   = [int](($runs.CalcMs   | Sort-Object)[[int]($Trials / 2)])
            SaveMs   = [int](($runs.SaveMs   | Sort-Object)[[int]($Trials / 2)])
            TotalMs  = [int](($runs.TotalMs  | Sort-Object)[[int]($Trials / 2)])
        }
        Write-Host ("done: {0}" -f $name)
    }

    Write-Host ""
    Write-Host ("workload: {0} formulas written one cell at a time, median of {1} trials" -f $FormulaCount, $Trials)
    $results | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
}
finally {
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}
