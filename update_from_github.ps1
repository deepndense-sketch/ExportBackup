param(
    [string]$RepoZipUrl = "https://github.com/deepndense-sketch/ExportBackup/archive/refs/heads/main.zip",
    [string]$Destination = "C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\ExportBackup"
)

$ErrorActionPreference = "Stop"

function Write-Step($message) {
    Write-Host "[ExportBackup Updater] $message"
}

$tempRoot = Join-Path $env:TEMP ("ExportBackupUpdate_" + [guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempRoot "ExportBackup-main.zip"
$extractPath = Join-Path $tempRoot "extract"

try {
    Write-Step "Preparing temporary workspace."
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
    New-Item -ItemType Directory -Force -Path $extractPath | Out-Null

    Write-Step "Downloading latest package from GitHub."
    Invoke-WebRequest -Uri $RepoZipUrl -OutFile $zipPath

    Write-Step "Extracting package."
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    $sourceRoot = Join-Path $extractPath "ExportBackup-main"
    if (-not (Test-Path $sourceRoot)) {
        throw "Could not find extracted ExportBackup-main folder."
    }

    Write-Step "Ensuring CEP extensions destination exists."
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null

    Write-Step "Copying updated extension files."
    $robocopyOutput = robocopy $sourceRoot $Destination /MIR /XD .git /XF deploy_extension.bat update_from_github.ps1
    $robocopyExit = $LASTEXITCODE
    if ($robocopyExit -ge 8) {
        throw "Robocopy failed with exit code $robocopyExit."
    }

    Write-Step "Update completed successfully."
    Write-Host "Restart Premiere Pro if the panel is already open."
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Recurse -Force $tempRoot -ErrorAction SilentlyContinue
    }
}
