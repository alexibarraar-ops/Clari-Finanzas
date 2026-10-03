param(
  [string]$RepoPath = "C:\Users\deel\Clari-Finanzas"
)

$ErrorActionPreference = "Stop"

Set-Location $RepoPath

Write-Host "== Clari Finanzas: actualizando GitHub =="
git fetch origin main
git switch main
git pull --ff-only origin main

$claspCmd = Get-Command clasp.cmd -ErrorAction SilentlyContinue
if (-not $claspCmd) {
  throw "No se encontro 'clasp.cmd'. Instalar con: npm install -g @google/clasp"
}

Write-Host "== Enviando archivos a Apps Script =="
& $claspCmd.Source push -f

Write-Host ""
Write-Host "Listo. Codigo enviado a Apps Script."
Write-Host "Ahora abrir: https://script.google.com/d/1QUyGNawFp0wEP88wTlxX5zE2gykb2klwor1idqoHbHw4aqm7vMMm5mNZ/edit"
