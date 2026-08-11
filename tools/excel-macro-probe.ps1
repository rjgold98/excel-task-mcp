# Where the time goes in the macro path.
#
# A formula Apply costs about a second end to end. The macro path was measured at 28.1s against the
# original server's 26.4s and recorded as "the one measured regression", attributed to Plan and
# Apply each opening their own Excel. Two launches is under a second, so that attribution cannot be
# the whole story and the rest has never been measured.
#
# This times each step the macro path actually performs, so the next optimization is aimed at
# whatever is actually expensive rather than at the thing that was easiest to name.
#
# Requires "Trust access to the VBA project object model". Real desktop Excel; the fixture is
# generated here and deleted at the end.

$ErrorActionPreference = 'Stop'

$Trials = 3
$root = Join-Path $env:TEMP ("exceltask-macro-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null

$source = @'
Public Sub WriteMarker()
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "marker"
End Sub
'@

function New-MacroFixture {
    param([string]$Path)

    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    try {
        $book = $app.Workbooks.Add()
        $component = $book.VBProject.VBComponents.Add(1)   # vbext_ct_StdModule
        $component.Name = 'SafeModule'
        $component.CodeModule.AddFromString($source)
        $book.SaveAs($Path, 52)                             # xlOpenXMLWorkbookMacroEnabled
        $book.Close($false)
    }
    finally {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

function Measure-MacroRun {
    param([string]$FixturePath, [string]$OutputPath)

    $steps = [ordered]@{}
    $watch = [Diagnostics.Stopwatch]::StartNew()

    $app = New-Object -ComObject Excel.Application
    $steps['launch'] = $watch.ElapsedMilliseconds

    try {
        $watch.Restart()
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.EnableEvents = $false
        $app.AutomationSecurity = 1        # Low, required before open for Application.Run to work
        $steps['configure'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $book = $app.Workbooks.Open($FixturePath, 0, $false)
        $steps['open-xlsm'] = $watch.ElapsedMilliseconds

        # The suspect: the first touch of VBProject materializes the VBA editor.
        $watch.Restart()
        $project = $book.VBProject
        $steps['first-VBProject'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $component = $project.VBComponents.Item('SafeModule')
        $module = $component.CodeModule
        $steps['find-component'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $start = $module.ProcStartLine('WriteMarker', 0)
        $count = $module.ProcCountLines('WriteMarker', 0)
        $text = $module.Lines($start, $count)
        $steps['read-procedure'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $module.DeleteLines($start, $count)
        $module.InsertLines($start, $text)
        $steps['replace-procedure'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        [void]$app.Run('WriteMarker')
        $steps['Application.Run'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $book.SaveAs($OutputPath, 52)
        $steps['save-as-xlsm'] = $watch.ElapsedMilliseconds

        $watch.Restart()
        $book.Close($false)
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        $app = $null
        $steps['close-and-quit'] = $watch.ElapsedMilliseconds

        # The verification reopen: a second Excel, and a second first-touch of VBProject.
        $watch.Restart()
        $verify = New-Object -ComObject Excel.Application
        $verify.Visible = $false
        $verify.DisplayAlerts = $false
        $verify.AutomationSecurity = 3
        $steps['verify-launch'] = $watch.ElapsedMilliseconds

        try {
            $watch.Restart()
            $verifyBook = $verify.Workbooks.Open($OutputPath, 0, $true)
            $steps['verify-open'] = $watch.ElapsedMilliseconds

            $watch.Restart()
            $verifyModule = $verifyBook.VBProject.VBComponents.Item('SafeModule').CodeModule
            $verifyStart = $verifyModule.ProcStartLine('WriteMarker', 0)
            [void]$verifyModule.Lines($verifyStart, $verifyModule.ProcCountLines('WriteMarker', 0))
            $steps['verify-read-procedure'] = $watch.ElapsedMilliseconds

            $verifyBook.Close($false)
        }
        finally {
            $verify.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($verify)
        }
    }
    finally {
        if ($app) {
            try { $app.Quit() } catch { }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        }
    }

    [PSCustomObject]$steps
}

try {
    $runs = @()
    for ($trial = 1; $trial -le $Trials; $trial++) {
        $fixture = Join-Path $root ("m-" + [Guid]::NewGuid().ToString('N').Substring(0, 8) + ".xlsm")
        $output = Join-Path $root ("o-" + [Guid]::NewGuid().ToString('N').Substring(0, 8) + ".xlsm")
        New-MacroFixture -Path $fixture
        # Filter: several COM calls in the function return values that would otherwise land in
        # the pipeline alongside the result object.
        $runs += Measure-MacroRun -FixturePath $fixture -OutputPath $output |
            Where-Object { $_ -is [PSCustomObject] } | Select-Object -Last 1
        Remove-Item $fixture, $output -Force -ErrorAction SilentlyContinue
        Write-Host "trial $trial done"
    }

    Write-Host ""
    Write-Host ("median of {0} trials, milliseconds" -f $Trials)
    $names = $runs[0].PSObject.Properties.Name
    $table = foreach ($name in $names) {
        $values = $runs | ForEach-Object { $_.$name } | Sort-Object
        [PSCustomObject]@{ Step = $name; Ms = [int]$values[[int]($Trials / 2)] }
    }
    $total = ($table | Measure-Object Ms -Sum).Sum
    $table | ForEach-Object {
        [PSCustomObject]@{ Step = $_.Step; Ms = $_.Ms; Share = ('{0,5:P0}' -f ($_.Ms / $total)) }
    } | Format-Table -AutoSize | Out-String -Width 120 | Write-Host
    Write-Host ("total {0} ms" -f $total)
}
finally {
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}
