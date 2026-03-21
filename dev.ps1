[CmdletBinding()]
param(
    [ValidateSet(
        "help",
        "phase1-test",
        "phase1-run",
        "phase1-build",
        "phase2-test",
        "phase2-run",
        "phase2-build",
        "test",
        "run",
        "build",
        "clean"
    )]
    [string]$Task = "help",
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

$AppName = "excel-template-generator"
$GuiEntry = "src/gui_app.py"
$CliEntry = "src/excel-template.py"
$ConfigPath = "config/config.yaml"
$Phase1Smoke = "workbooks/phase1_make_smoke.xlsx"
$Phase2Smoke = "workbooks/phase2_make_smoke.xlsx"
$GuiSpec = "$AppName.spec"
$PyInstallerConfigDir = "build/pyinstaller-config"
$PyInstallerWorkPath = "build/pyinstaller-work"
$PyInstallerDistPath = "dist"

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

function Invoke-Phase2Test {
    $previous = $env:QT_QPA_PLATFORM
    $env:QT_QPA_PLATFORM = "offscreen"
    try {
        Invoke-Step @($Python, "-m", "unittest", "discover", "-s", "tests", "-v")
    }
    finally {
        if ($null -eq $previous) {
            Remove-Item Env:QT_QPA_PLATFORM -ErrorAction SilentlyContinue
        }
        else {
            $env:QT_QPA_PLATFORM = $previous
        }
    }
}

function Invoke-Phase2Build {
    $previous = $env:PYINSTALLER_CONFIG_DIR
    $env:PYINSTALLER_CONFIG_DIR = (Join-Path $RepoRoot $PyInstallerConfigDir)
    try {
        Invoke-Step @(
            $Python,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--windowed",
            "--distpath",
            $PyInstallerDistPath,
            "--workpath",
            $PyInstallerWorkPath,
            "--specpath",
            ".",
            "--name",
            $AppName,
            $GuiEntry
        )
    }
    finally {
        if ($null -eq $previous) {
            Remove-Item Env:PYINSTALLER_CONFIG_DIR -ErrorAction SilentlyContinue
        }
        else {
            $env:PYINSTALLER_CONFIG_DIR = $previous
        }
    }
}

switch ($Task) {
    "help" {
        @(
            "Available tasks:",
            "  phase1-test   Run the Phase 1 backend smoke tests.",
            "  phase1-run    Run the CLI workbook generator and write a smoke workbook.",
            "  phase1-build  Byte-compile project sources for an early build checkpoint.",
            "  phase2-test   Run the full Phase 1 + Phase 2 checkpoint suite.",
            "  phase2-run    Run the prototype PyQt GUI on the local machine.",
            "  phase2-build  Package the prototype GUI with PyInstaller.",
            "  test          Alias for phase2-test.",
            "  run           Alias for phase2-run.",
            "  build         Alias for phase2-build.",
            "  clean         Remove generated smoke workbooks, caches, and build artifacts."
        ) | ForEach-Object { Write-Host $_ }
    }
    "phase1-test" {
        Invoke-Step @($Python, "-m", "unittest", "discover", "-s", "tests", "-p", "test_template_generator.py", "-v")
    }
    "phase1-run" {
        Invoke-Step @($Python, $CliEntry, "--config", $ConfigPath, "--output", $Phase1Smoke)
    }
    "phase1-build" {
        Invoke-Step @($Python, "-m", "compileall", "src", "tests")
    }
    "phase2-test" {
        Invoke-Phase2Test
    }
    "phase2-run" {
        Invoke-Step @($Python, $GuiEntry, "--config", $ConfigPath)
    }
    "phase2-build" {
        Invoke-Phase2Build
    }
    "test" {
        Invoke-Phase2Test
    }
    "run" {
        Invoke-Step @($Python, $GuiEntry, "--config", $ConfigPath)
    }
    "build" {
        Invoke-Phase2Build
    }
    "clean" {
        @(
            "build",
            "dist",
            $GuiSpec,
            "__pycache__",
            "src/__pycache__",
            "tests/__pycache__",
            ".pytest_cache",
            ".coverage",
            $Phase1Smoke,
            $Phase2Smoke
        ) | ForEach-Object {
            if (Test-Path $_) {
                Remove-Item $_ -Recurse -Force
            }
        }
    }
}
