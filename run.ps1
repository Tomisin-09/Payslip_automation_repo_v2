$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

# ==============================
# 1) Find Python (prefer py, fallback to python)
# ==============================
function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return "py" }
    if (Get-Command python -ErrorAction SilentlyContinue) { return "python" }
    return $null
}

$PyCmd = Get-PythonCommand
if (-not $PyCmd) {
    Write-Error "Python 3.10+ is required but was not found. Install from https://www.python.org (tick: Add to PATH)."
    exit 1
}

# ==============================
# 2) Enforce minimum Python version
# ==============================
$minVersion = [Version]"3.10"

try {
    $versionString = & $PyCmd -c "import sys; print(sys.version.split()[0])"
    $currentVersion = [Version]$versionString
} catch {
    Write-Error "Python was found ('$PyCmd') but version could not be detected. Try running: $PyCmd --version"
    exit 1
}

if ($currentVersion -lt $minVersion) {
    Write-Error "Python $($minVersion)+ is required. Found $currentVersion"
    exit 1
}

Write-Host "Python version OK: $currentVersion (via $PyCmd)"

# ==============================
# 3) venv + deps + run
# ==============================
if (!(Test-Path ".venv")) {
    & $PyCmd -m venv .venv
}

& .\.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m src.main