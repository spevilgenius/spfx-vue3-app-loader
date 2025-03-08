function log {
  param (
    $Message
  )
  Write-Host $Message -ForegroundColor Yellow -BackgroundColor Black
  Write-Host " "
}

log "Uploading to SharePoint"

$SiteURL = "https://legodan.sharepoint.com/sites/LegoDanDev"
$FilesPath = "D:\Repos\spfx-vue3-app-loader\sharepoint\solution"
$serverRelativePath = "/sites/LegoDanDev/AppCatalog"
$AppName = "SPFX Vue App Loader"

Connect-PnPOnline -Url $SiteURL -Interactive -Verbose

$PackageFiles = Get-ChildItem -Path $FilesPath -Force -File "spfx-vue-3-app-loader.sppkg"

ForEach ($File in $PackageFiles) {
  Write-Host "Uploading $($File.Directory)\$($File.Name)"
  Add-PnPFile -Path "$($File.Directory)\$($File.Name)" -Folder $serverRelativePath -Values @{ "Title" = $($File.Name) }
  Start-Sleep -Seconds 5
  $App = Get-PnPApp -Scope Site | Where-Object { $_.Title -eq $AppName }
  Publish-PnPApp -Identity $App.Id -Scope Site
}

Start-Sleep -Seconds 5

Function UpdateAppFromCatalog {
  [cmdletbinding()]
  Param(
    [parameter(Mandatory = $true, ValueFromPipeline = $true)] $web
  )
  Try {
    $web.Title
    Connect-PnPOnline -Url $web.Url -Interactive
    $App = Get-PnPApp -Scope Site | Where-Object { $_.Title -eq $AppName }
    Update-PnPApp -Identity $App.Id -Scope Site
  }
  Catch {
    Write-Host -f Red "ERROR"
  }
}

Get-PnPSubWeb -IncludeRootWeb -Recurse | ForEach-Object { UpdateAppFromCatalog $_ }