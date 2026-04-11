param(
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [Parameter(Mandatory = $true)]
    [string]$InputListFile
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $InputListFile)) {
    throw "Input list file not found: $InputListFile"
}

$InputFiles = Get-Content -LiteralPath $InputListFile -Encoding UTF8 | Where-Object { $_ -and $_.Trim() }

if (-not $InputFiles -or $InputFiles.Count -eq 0) {
    throw "No input PPT files were provided in the input list file."
}

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Add()

    foreach ($inputFile in $InputFiles) {
        if (-not (Test-Path $inputFile)) {
            throw "Input PPT not found: $inputFile"
        }

        $insertIndex = $presentation.Slides.Count
        $presentation.Slides.InsertFromFile($inputFile, $insertIndex)
    }

    $presentation.SaveAs($OutputFile)
}
finally {
    if ($presentation) {
        try {
            $presentation.Saved = $true
            $presentation.Close()
        }
        catch {
            Write-Warning ("Presentation close warning: " + $_.Exception.Message)
        }
    }
    if ($powerPoint) {
        try {
            $powerPoint.Quit()
        }
        catch {
            Write-Warning ("PowerPoint quit warning: " + $_.Exception.Message)
        }
    }
    if ($presentation) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
    }
    if ($powerPoint) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
