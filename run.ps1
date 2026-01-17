$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== Payslip Automation Runner ==="

# ----------------------------------------
# 1. Detect Python (python or py)
# ----------------------------------------
$PY = $null

if (Get-Command python -ErrorAction SilentlyContinue) {
    $PY = "python"
}
elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $PY = "py"
}

if (-not $PY) {
    Write-Error @"
Python is not installed or not on PATH.

Please install Python 3.10 or newer from:
https://www.python.org/downloads/windows/

IMPORTANT:
- Tick 'Add Python to PATH'
- Restart PowerShell after installation
"@
    exit 1
}

# ----------------------------------------
# 2. Enforce minimum Python version
# ----------------------------------------
$version = & $PY - << 'EOF'
import sys
print(f"{sys.version_info.major}.{sys.version_info.minor}")
EOF

if ([version]$version -lt [version]"3.10") {
    Write-Error "Python 3.10+ required. Found $version."
    exit 1
}

& $PY --version
Write-Host "Using Python via '$PY'"

# ----------------------------------------
# 3. Create virtual environment
# ----------------------------------------
if (!(Test-Path ".venv")) {
    Write-Host "Creating virtual environment..."
    & $PY -m venv .venv
}

# ----------------------------------------
# 4. Activate virtual environment
# ----------------------------------------
Write-Host "Activating virtual environment..."
& .\.venv\Scripts\Activate.ps1

# ----------------------------------------
# 5. Install dependencies
# ----------------------------------------
Write-Host "Upgrading pip..."
python -m pip install --upgrade pip

Write-Host "Installing requirements..."
python -m pip install -r requirements.txt

# ----------------------------------------
# 6. Run application
# ----------------------------------------
Write-Host "Running application..."
python -m src.main

Write-Host "=== Run complete ==="
