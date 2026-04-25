$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$targets = @(
  @{ Role = "cliente"; Dist = "dist-cliente"; Preset = "role-config.cliente.js" },
  @{ Role = "staff"; Dist = "dist-staff"; Preset = "role-config.staff.js" },
  @{ Role = "struttura"; Dist = "dist-struttura"; Preset = "role-config.struttura.js" }
)

function Copy-ProjectFiles {
  param(
    [string]$Destination
  )

  if (Test-Path $Destination) {
    Remove-Item -Recurse -Force $Destination
  }
  New-Item -ItemType Directory -Path $Destination | Out-Null

  $items = Get-ChildItem -Force $projectRoot | Where-Object {
    $name = $_.Name
    if ($name -eq ".git") { return $false }
    if ($name -like "dist-*") { return $false }
    if ($name -eq "build-distribuzioni.ps1") { return $false }
    if ($name -eq "build-distribuzioni.bat") { return $false }
    return $true
  }

  foreach ($item in $items) {
    $destPath = Join-Path $Destination $item.Name
    Copy-Item -Path $item.FullName -Destination $destPath -Recurse -Force
  }
}

Write-Host "Generazione pacchetti distribuzione..." -ForegroundColor Cyan

foreach ($target in $targets) {
  $distPath = Join-Path $projectRoot $target.Dist
  $presetPath = Join-Path $projectRoot $target.Preset
  $roleConfigPath = Join-Path $distPath "role-config.js"

  if (!(Test-Path $presetPath)) {
    throw "Preset non trovato: $($target.Preset)"
  }

  Copy-ProjectFiles -Destination $distPath
  Copy-Item -Path $presetPath -Destination $roleConfigPath -Force

  Write-Host ("Creato: {0} (ruolo: {1})" -f $target.Dist, $target.Role) -ForegroundColor Green
}

Write-Host "Completato." -ForegroundColor Cyan
