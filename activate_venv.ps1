$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPath = Join-Path $ProjectRoot ".venv"
$ActivateScript = Join-Path $ProjectRoot ".venv\Scripts\Activate.ps1"
$TempPath = Join-Path $ProjectRoot ".tmp"

if (-not (Test-Path $TempPath)) {
    New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
}

function Get-BasePythonPath {
    $Resolved = ""
    try {
        $Resolved = (& python -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1).Trim()
    }
    catch {
        $Resolved = ""
    }

    if ($Resolved -and (Test-Path $Resolved)) {
        return $Resolved
    }

    $PythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($PythonCmd -and $PythonCmd.Source -and (Test-Path $PythonCmd.Source)) {
        return $PythonCmd.Source
    }

    throw "A working Python interpreter was not found. Install Python or add it to PATH."
}

if (-not (Test-Path $ActivateScript)) {
    Write-Host "Creating virtual environment..."
    $env:TEMP = $TempPath
    $env:TMP = $TempPath
    $BasePython = Get-BasePythonPath
    if (Test-Path $VenvPath) {
        $ResolvedTarget = [System.IO.Path]::GetFullPath($VenvPath)
        Remove-Item -LiteralPath $ResolvedTarget -Recurse -Force
    }
    & $BasePython -m venv --system-site-packages --without-pip $VenvPath
}

if (-not (Test-Path $ActivateScript)) {
    throw "Virtual environment activation script not found: $ActivateScript"
}

. $ActivateScript
Write-Host "Virtual environment activated."
