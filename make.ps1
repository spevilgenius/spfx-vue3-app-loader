function log {
  param (
    $Message
  )
  Write-Host $Message -ForegroundColor Yellow -BackgroundColor Black
  Write-Host " "
}

$versiontype = 'Draft'

function Show-VersioningMenu {
  Clear-Host
  Write-Host "=== Version Selection ==="
  Write-Host  "1. Build Draft Version"
  Write-Host  "1. Build Minor Version"
  Write-Host  "1. Build Major Version"
  Write-Host  "1. Cancel"
}

Show-VersioningMenu

$choice = Read-Host "Select a versioning option (1-4)"
if ($choice -match '^[1-4]$') {
  switch ($choice) {
    1 {
      $versiontype = 'Draft'
    }
    2 {
      $versiontype = 'Minor'
    }
    3 {
      $versiontype = 'Major'
    }
    4 {
      $versiontype = 'NONE'
    }
  }
}
else {
  Write-Host "Invalid Choice"
  Start-Sleep -Seconds 5
  Show-VersioningMenu
}

if ($versiontype -ne 'NONE') {
  log "Building Web Part"
  $script = ".\create-package-version.js"
  $cmd = 'node'
  $path = @($script)
  $version = @($versiontype)
  & $cmd  $path  $version

  log "Packaging Web Part"

  gulp clean
  gulp bundle --ship
  gulp package-solution --ship

  log "Web Part Built. Upload to SharePoint or run install.ps1"
}