[CmdletBinding()]
param(
    [ValidateSet("test", "run")]
    [string]$Task = "test",
    [string]$Python = "python",
    [string]$Config = "config/config.yaml",
    [string]$Output = "workbooks/current_roster_workbook.xlsx"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Command
    )

    Write-Host ">> $($Command -join ' ')"
    & $Command[0] $Command[1..($Command.Length - 1)]
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE"
    }
}

switch ($Task) {
    "test" {
        Invoke-Step @($Python, "-m", "unittest", "discover", "-s", "tests", "-v")
    }
    "run" {
        Invoke-Step @($Python, "src/excel-template.py", "--config", $Config, "--output", $Output)
    }
}
