<#
.SYNOPSIS
    CI entry point (Jenkins / Azure DevOps / GitLab runner on a Windows agent
    with UFT installed). Run it from the checked-out repository:

        powershell -ExecutionPolicy Bypass -File Scripts\run_uft.ps1

.DESCRIPTION
    1. If .env does not exist, builds it from .env.example, taking each value
       from a CI environment variable of the same name when one is set
       (inject secrets such as APP_PASSWORD from the CI credential store).
    2. Runs the Driver test through Scripts\RunTest.vbs.
    3. Returns RunTest.vbs' exit code (0 = passed, non-zero = failed/error).
#>
param(
    [string]$ResultsDir = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
Write-Host "==================== UFT Runner ===================="
Write-Host "Framework root: $root"

# ---- 1. .env -----------------------------------------------------------------
$envFile = Join-Path $root ".env"
$template = Join-Path $root ".env.example"
if (-not (Test-Path $envFile)) {
    if (-not (Test-Path $template)) { throw ".env and .env.example are both missing" }
    Write-Host "Creating .env from .env.example + CI environment variables"
    $lines = foreach ($line in Get-Content $template) {
        if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=') {
            $key = $Matches[1]
            $val = [Environment]::GetEnvironmentVariable($key)
            if ($val) { "$key=$val" } else { $line }
        } else { $line }
    }
    $lines | Set-Content -Path $envFile -Encoding ASCII
}

# ---- 2. run ------------------------------------------------------------------
$runner = Join-Path $root "Scripts\RunTest.vbs"
$testPath = Join-Path $root "Driver"
$cliArgs = @("//nologo", $runner, $testPath)
if ($ResultsDir) { $cliArgs += $ResultsDir }

& cscript.exe @cliArgs
$code = $LASTEXITCODE

Write-Host "RunTest.vbs exit code: $code"
Write-Host "Reports: $(Join-Path $root 'Reports')"
Write-Host "==================== UFT Runner finished ===================="
exit $code
