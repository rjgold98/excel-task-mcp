[CmdletBinding()]
param(
    [switch]$IncludeExcel
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

dotnet test (Join-Path $repo 'ExcelTask.slnx') `
    --filter 'RunType!=OnDemand' `
    -p:NuGetAudit=false
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($IncludeExcel) {
    dotnet test (Join-Path $repo 'tests\ExcelTask.Excel.Tests\ExcelTask.Excel.Tests.csproj') `
        --filter 'RunType=OnDemand' `
        -p:NuGetAudit=false
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    # Keep desktop Excel acceptance serial. This second project proves the
    # same workflow through the actual one-tool MCP protocol boundary.
    dotnet test (Join-Path $repo 'tests\ExcelTask.McpServer.Tests\ExcelTask.McpServer.Tests.csproj') `
        --filter 'RunType=OnDemand' `
        -p:NuGetAudit=false
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
