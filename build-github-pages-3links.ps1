$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$webRoot = Join-Path $projectRoot "web-links"
if (Test-Path $webRoot) {
  Remove-Item -Recurse -Force $webRoot
}
New-Item -ItemType Directory -Path $webRoot | Out-Null

$targets = @(
  @{ Role = "cliente"; Folder = "cliente"; Preset = "role-config.cliente.js" },
  @{ Role = "staff"; Folder = "staff"; Preset = "role-config.staff.js" },
  @{ Role = "struttura"; Folder = "struttura"; Preset = "role-config.struttura.js" }
)

function Copy-WebFiles {
  param(
    [string]$Destination
  )

  New-Item -ItemType Directory -Path $Destination -Force | Out-Null

  $files = @(
    "index.html",
    "styles.css",
    "app.js",
    "role-config.js"
  )

  foreach ($file in $files) {
    $src = Join-Path $projectRoot $file
    if (!(Test-Path $src)) {
      throw "File mancante: $file"
    }
    Copy-Item -Path $src -Destination (Join-Path $Destination $file) -Force
  }
}

Write-Host "Generazione struttura web-links..." -ForegroundColor Cyan

foreach ($t in $targets) {
  $dest = Join-Path $webRoot $t.Folder
  Copy-WebFiles -Destination $dest
  Copy-Item -Path (Join-Path $projectRoot $t.Preset) -Destination (Join-Path $dest "role-config.js") -Force
  Write-Host ("Creato percorso /{0} (ruolo: {1})" -f $t.Folder, $t.Role) -ForegroundColor Green
}

$indexRoot = @"
<!doctype html>
<html lang="it">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Distribuzioni Evento Tavoli</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 2rem; line-height: 1.45; }
      a { display: inline-block; margin: 0.35rem 0; font-size: 1.05rem; }
    </style>
  </head>
  <body>
    <h1>Distribuzioni Evento Tavoli</h1>
    <p>Scegli la versione da aprire:</p>
    <a href="./cliente/">Versione Cliente</a><br />
    <a href="./staff/">Versione Staff</a><br />
    <a href="./struttura/">Versione Struttura</a>
  </body>
</html>
"@

Set-Content -Path (Join-Path $webRoot "index.html") -Value $indexRoot -Encoding UTF8

Write-Host "Completato: cartella web-links pronta per GitHub Pages." -ForegroundColor Cyan
