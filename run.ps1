param(
    [switch]$AnalyzeOnly,
    [switch]$Mock,
    [switch]$GoogleSafe,
    [string]$PromptFile,
    [string]$CurriculumFile,
    [string]$TemplateFile,
    [string]$Model,
    [string]$lecture,
    [string]$page
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
trap [System.Management.Automation.PipelineStoppedException] {
    Write-Host "[취소] Ctrl+C로 실행이 중단되었습니다."
    exit 130
}

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPath = Join-Path $ProjectRoot ".venv"
$PythonExe = Join-Path $VenvPath "Scripts\\python.exe"
$RequirementsFile = Join-Path $ProjectRoot "requirements.txt"
$AppFile = Join-Path $ProjectRoot "app.py"
$TempPath = Join-Path $ProjectRoot ".tmp"

if (-not (Test-Path $TempPath)) {
    New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
}

$env:TEMP = $TempPath
$env:TMP = $TempPath
$env:PYTHONIOENCODING = "utf-8"

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

function New-ProjectVenv {
    param([string]$TargetPath)

    $ResolvedTarget = [System.IO.Path]::GetFullPath($TargetPath)
    if (Test-Path $ResolvedTarget) {
        Write-Host "[1/4] Recreating broken virtual environment..."
        Remove-Item -LiteralPath $ResolvedTarget -Recurse -Force
    }
    else {
        Write-Host "[1/4] Creating virtual environment..."
    }

    $BasePython = Get-BasePythonPath
    & $BasePython -m venv --system-site-packages --without-pip $ResolvedTarget
}

function Test-VenvDependencies {
    param([string]$InterpreterPath)

    try {
        $null = & $InterpreterPath -c "import openai, pptx" 2>$null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

if (-not (Test-Path $PythonExe)) {
    New-ProjectVenv -TargetPath $VenvPath
}

if (-not (Test-Path $PythonExe)) {
    throw "Virtual environment Python not found after creation: $PythonExe"
}

Write-Host "[2/4] Installing or updating dependencies..."
$BasePython = Get-BasePythonPath
$HasDependencies = Test-VenvDependencies -InterpreterPath $PythonExe

if (-not $HasDependencies) {
    Write-Host "Dependencies not visible in the virtual environment. Recreating the environment once..."
    New-ProjectVenv -TargetPath $VenvPath
    $HasDependencies = Test-VenvDependencies -InterpreterPath $PythonExe
}

if (-not $HasDependencies) {
    Write-Host "Dependencies still missing. Trying installation..."
    try {
        & $BasePython -m pip --python $PythonExe install -r $RequirementsFile
        if ($LASTEXITCODE -ne 0) {
            throw "Dependency installation failed."
        }
    }
    catch {
        throw "Could not prepare the virtual environment dependencies. Check Python/network settings or use a machine with package access."
    }
}

Write-Host "[3/4] Running PPT generator..."
$ArgsList = @($AppFile)

if ($AnalyzeOnly) {
    $ArgsList += "--analyze-only"
}

if ($Mock) {
    $ArgsList += "--mock"
}

if ($GoogleSafe) {
    $ArgsList += "--google-safe"
}

if ($PromptFile) {
    $ArgsList += "--prompt-file"
    $ArgsList += $PromptFile
}

if ($CurriculumFile) {
    $ArgsList += "--curriculum-file"
    $ArgsList += $CurriculumFile
}

if ($TemplateFile) {
    $ArgsList += "--template"
    $ArgsList += $TemplateFile
}

if ($Model) {
    $ArgsList += "--model"
    $ArgsList += $Model
}

if ($lecture) {
    $ArgsList += "--lecture"
    $ArgsList += $lecture
}

if ($page) {
    $ArgsList += "--page"
    $ArgsList += $page
}

Push-Location $ProjectRoot
try {
    & $PythonExe @ArgsList
    if ($LASTEXITCODE -eq 130) {
        Write-Host "[취소] Python 작업이 중단되었습니다."
        exit 130
    }
}
finally {
    Pop-Location
}

Write-Host "[4/4] Done."
