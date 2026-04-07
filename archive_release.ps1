param(
    [string]$Version,
    [string]$DestinationRoot
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$projectName = Split-Path $repoRoot -Leaf
$versionFile = Join-Path $repoRoot "version.json"

if (-not (Test-Path $versionFile)) {
    throw "Could not find version.json at $versionFile"
}

$versionInfo = Get-Content $versionFile -Raw | ConvertFrom-Json

if (-not $Version -or -not $Version.Trim()) {
    $Version = [string]$versionInfo.version
}

if (-not $DestinationRoot -or -not $DestinationRoot.Trim()) {
    $DestinationRoot = Join-Path (Split-Path $repoRoot -Parent) ($projectName + "_Versions")
}

$destination = Join-Path $DestinationRoot $Version

New-Item -ItemType Directory -Force $destination | Out-Null

$robocopyArgs = @(
    $repoRoot,
    $destination,
    "/MIR",
    "/XD", ".git",
    "/R:2",
    "/W:1"
)

& robocopy @robocopyArgs | Out-Null
$robocopyExitCode = $LASTEXITCODE
if ($robocopyExitCode -ge 8) {
    throw "robocopy failed with exit code $robocopyExitCode"
}

$commit = ""
try {
    $commit = (& 'C:\Program Files\Git\cmd\git.exe' -C $repoRoot rev-parse HEAD).Trim()
} catch {}

$releaseInfo = [ordered]@{
    version = $Version
    archivedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    sourceRoot = $repoRoot
    archiveRoot = $destination
    commit = $commit
}

$releaseInfo | ConvertTo-Json | Set-Content (Join-Path $destination "release-info.json") -Encoding UTF8

Write-Host "Archived $projectName version $Version to $destination"
