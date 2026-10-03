param(
  [string]$RepoPath = "C:\Users\deel\Clari-Finanzas"
)

$ErrorActionPreference = "Stop"

Set-Location $RepoPath

Write-Host "== Clari Finanzas: actualizando GitHub =="
git fetch origin main
git switch main
git pull --ff-only origin main

if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
  throw "No se encontro 'clasp'. Instalar con: npm install -g @google/clasp"
}

Write-Host "== Enviando archivos a Apps Script =="
clasp push -f

Write-Host ""
Write-Host "Listo. Codigo enviado a Apps Script."
Write-Host "Ahora abrir: https://script.google.com/d/1QUyGNawFp0wEP88wTlxX5zE2gykb2klwor1idqoHbHw4aqm7vMMm5mNZ/edit"
