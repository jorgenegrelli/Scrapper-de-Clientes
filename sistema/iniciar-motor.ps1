# Liga o motor do Google Maps (gosom/google-maps-scraper) em segundo plano, ouvindo so neste computador.
$raiz = Split-Path -Parent $PSScriptRoot
$exe  = Join-Path $raiz 'motor\google-maps-scraper.exe'
$data = Join-Path $raiz 'motor\dados'
$api  = 'http://127.0.0.1:8080/api/v1/jobs'

function Motor-Ligado {
    try { Invoke-RestMethod -Uri $api -TimeoutSec 3 | Out-Null; return $true } catch { return $false }
}

if (Motor-Ligado) { Write-Host 'Motor ja esta ligado: http://localhost:8080'; exit 0 }
if (-not (Test-Path $exe)) { Write-Host 'Motor nao instalado. Rode o INSTALAR.bat primeiro.'; exit 1 }

New-Item -ItemType Directory -Force -Path $data | Out-Null
# -addr 127.0.0.1: o motor nao tem senha, entao fica acessivel so neste computador
Start-Process -FilePath $exe -ArgumentList '-web', '-data-folder', "`"$data`"", '-addr', '127.0.0.1:8080' `
    -WorkingDirectory (Split-Path $exe) -WindowStyle Minimized

Write-Host -NoNewline 'Ligando o motor'
for ($i = 0; $i -lt 90; $i++) {
    if (Motor-Ligado) { Write-Host ' pronto. (Para desligar, feche a janela minimizada do google-maps-scraper.)'; exit 0 }
    Write-Host -NoNewline '.'
    Start-Sleep -Seconds 2
}
Write-Host ''
Write-Host 'O motor nao respondeu em 3 minutos. Veja a janela minimizada do google-maps-scraper para o erro.'
exit 1
