# Runs the full local gate, including the real-Excel tests.
#
# One project at a time, deliberately. `dotnet test` on the solution runs test assemblies in
# parallel, and two of these drive real desktop Excel and then assert that no Excel process was
# left behind. Run concurrently, each one sees the other's Excel and reports a leak that is not
# there - and worse, a run that happened not to overlap would have reported a pass it had not
# earned. The assertion is only meaningful when nothing else is starting Excel.
#
# CI cannot run these at all: hosted runners have no Office. This script is the gate.

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$projects = @(
    'tests/ExcelTask.Core.Tests/ExcelTask.Core.Tests.csproj'
    'tests/ExcelTask.Excel.Tests/ExcelTask.Excel.Tests.csproj'
    'tests/ExcelTask.McpServer.Tests/ExcelTask.McpServer.Tests.csproj'
)

$failed = @()
foreach ($project in $projects) {
    Write-Host ""
    Write-Host "=== $project ===" -ForegroundColor Cyan
    dotnet test (Join-Path $root $project) -c Release -p:NuGetAudit=false --nologo
    if ($LASTEXITCODE -ne 0) { $failed += $project }
}

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "gate passed" -ForegroundColor Green
    exit 0
}

Write-Host ("gate failed: " + ($failed -join ', ')) -ForegroundColor Red
exit 1
