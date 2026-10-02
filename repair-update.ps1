param([string]$PackagePath = "")
$ErrorActionPreference = 'Stop'
try {
    if (Get-Process -Name 'Adobe Premiere Pro' -ErrorAction SilentlyContinue) {
        throw 'Close Premiere Pro before running this repair, then try again.'
    }
    $work = Join-Path $env:TEMP ('BackupProjectRepair-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $work | Out-Null
    if (-not $PackagePath) {
        Write-Host 'Downloading the latest Backup Project...'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $PackagePath = Join-Path $work 'latest.zip'
        Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/deepndense-sketch/ExportBackup/archive/refs/heads/main.zip' -OutFile $PackagePath
    }
    Expand-Archive -LiteralPath $PackagePath -DestinationPath (Join-Path $work 'extracted')
    $source = Join-Path $work 'extracted\ExportBackup-main'
    $bundleId = 'com.deepndense.exportbackup'
    [xml]$manifest = Get-Content -LiteralPath (Join-Path $source 'CSXS\manifest.xml') -Raw
    $version = (Get-Content -LiteralPath (Join-Path $source 'version.json') -Raw | ConvertFrom-Json).version
    if ($manifest.ExtensionManifest.ExtensionBundleId -ne $bundleId -or
        $manifest.ExtensionManifest.ExtensionBundleVersion -ne $version -or
        $manifest.ExtensionManifest.ExtensionList.Extension.Version -ne $version) {
        throw 'Downloaded plugin identity or version is invalid.'
    }
    $cep = [IO.Path]::GetFullPath((Join-Path $env:APPDATA 'Adobe\CEP\extensions'))
    $destination = Join-Path $cep 'Backup Project'
    if (Test-Path -LiteralPath $destination) {
        [xml]$existing = Get-Content -LiteralPath (Join-Path $destination 'CSXS\manifest.xml') -Raw
        if ($existing.ExtensionManifest.ExtensionBundleId -ne $bundleId) { throw 'Destination belongs to another plugin.' }
    }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $source -Recurse -File) {
        $relative = $file.FullName.Substring($source.Length).TrimStart('\')
        $target = Join-Path $destination $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
        Copy-Item -LiteralPath $file.FullName -Destination $target -Force
        if ((Get-FileHash -LiteralPath $file.FullName).Hash -ne (Get-FileHash -LiteralPath $target).Hash) {
            throw "Copy verification failed: $relative"
        }
    }
    foreach ($oldName in @('Backup','ExportBackup')) {
        $old = Join-Path $cep $oldName
        if (-not (Test-Path -LiteralPath (Join-Path $old 'CSXS\manifest.xml'))) { continue }
        [xml]$oldManifest = Get-Content -LiteralPath (Join-Path $old 'CSXS\manifest.xml') -Raw
        if ($oldManifest.ExtensionManifest.ExtensionBundleId -ne $bundleId) { continue }
        $resolvedOld = (Resolve-Path -LiteralPath $old).Path
        if ((Split-Path -Parent $resolvedOld) -ine $cep) { throw 'Unexpected old installation path.' }
        $archiveRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'BackupProject\DisabledExtensions'))
        if ($archiveRoot.StartsWith($cep + '\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Archive must be outside CEP.' }
        New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null
        $archive = Join-Path $archiveRoot ($oldName + '-' + [guid]::NewGuid().ToString('N'))
        Move-Item -LiteralPath $resolvedOld -Destination $archive
        Write-Host "Duplicate preserved outside CEP: $archive"
    }
    Write-Host "Installed and verified Backup Project $version. Open Premiere Pro. Future updates use the plugin's Update button."
} catch {
    Write-Host "Repair failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
