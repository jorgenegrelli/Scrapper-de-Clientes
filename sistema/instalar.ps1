# Instala tudo que o sistema precisa (Windows 10/11, 64 bits). Pode rodar de novo quantas vezes quiser:
# o que ja estiver instalado e pulado.
#   1. Python 3 (via winget, se faltar) + openpyxl
#   2. Motor gosom/google-maps-scraper v1.18.1 (exe oficial do GitHub)
#   3. Driver do Playwright 1.61.1 (playwright-core + node), montado a mao porque o download
#      automatico do playwright-go aponta para um servidor que a Microsoft desativou
#   4. Navegador Chromium usado pelo motor
# Todo download e conferido pelo SHA-256 antes de ser usado.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$raiz = Split-Path -Parent $PSScriptRoot
$sistema = $PSScriptRoot

$MOTOR_URL = 'https://github.com/gosom/google-maps-scraper/releases/download/v1.18.1/google_maps_scraper-1.18.1-windows-amd64.exe'
$MOTOR_SHA = 'C124FAB30F12E4AAE25EF52F38EB2BEA5422ECE708B3171CA0F34BE56CF7848F'
$PW_VERSAO = '1.61.1'
$PW_URL    = "https://registry.npmjs.org/playwright-core/-/playwright-core-$PW_VERSAO.tgz"
$PW_SHA    = '3740BBA3A3E93B4C7BD7C0F6E3C0E5F8CDCE498BEF13511E4D534CC99A8CE198'
$NODE_URL  = 'https://nodejs.org/dist/v24.18.0/win-x64/node.exe'
$NODE_SHA  = '9A4EB5F1C29C6A2E93852EAD46B999E284A6A5CA8BAB4D4E241D587D025A52DE'

function Passo($texto) { Write-Host ''; Write-Host "== $texto" -ForegroundColor Cyan }
function Ok($texto)    { Write-Host "   OK  $texto" -ForegroundColor Green }
function Falha($texto) { Write-Host ''; Write-Host "   ERRO: $texto" -ForegroundColor Red; exit 1 }

function Baixar($url, $destino, $sha) {
    $tmp = "$destino.baixando"
    Write-Host "   baixando $url"
    try { Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing }
    catch { Falha "nao consegui baixar $url ($($_.Exception.Message)). Confira a internet e rode de novo." }
    $h = (Get-FileHash $tmp -Algorithm SHA256).Hash
    if ($h -ne $sha) { Remove-Item $tmp -Force; Falha "o arquivo baixado de $url nao confere (SHA-256 $h). Nada foi instalado." }
    Move-Item -Force $tmp $destino
}

function Achar-Python {
    foreach ($tentativa in @(@('py', '-3'), @('python'), @('python3'))) {
        $cmd = $tentativa[0]; $args_ = @($tentativa | Select-Object -Skip 1)
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) { continue }
        try {
            $exe = & $cmd @args_ -c "import sys; print(sys.executable if sys.version_info >= (3, 9) else '')" 2>$null
            if ($LASTEXITCODE -eq 0 -and $exe -and (Test-Path $exe)) { return "$exe".Trim() }
        } catch {}
    }
    return $null
}

Write-Host 'Instalando o sistema Clientes para Sites...' -ForegroundColor White

# 1. Python ------------------------------------------------------------------------------------------
Passo 'Python'
$py = Achar-Python
if (-not $py) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host '   Python nao encontrado. Instalando o Python 3.12 pelo winget (pode levar alguns minutos)...'
        winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements | Out-Host
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
        $py = Achar-Python
        if (-not $py) {
            $padrao = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
            if (Test-Path $padrao) { $py = $padrao }
        }
    }
    if (-not $py) { Falha 'instale o Python em https://www.python.org/downloads/ (marque "Add python.exe to PATH") e rode o INSTALAR.bat de novo.' }
}
Set-Content -Path (Join-Path $sistema 'python.txt') -Value $py -Encoding ASCII
Ok $py

Write-Host '   instalando openpyxl (para gerar o Excel)...'
& $py -m pip install --user --quiet --disable-pip-version-check openpyxl | Out-Host
$temOpenpyxl = $false
try { & $py -c "import openpyxl" 2>$null; $temOpenpyxl = ($LASTEXITCODE -eq 0) } catch {}
if (-not $temOpenpyxl) { Falha 'nao consegui instalar o openpyxl. Rode: python -m pip install openpyxl' }
Ok 'openpyxl'

# 2. Motor --------------------------------------------------------------------------------------------
Passo 'Motor do Google Maps (gosom/google-maps-scraper v1.18.1)'
$motorDir = Join-Path $raiz 'motor'
$motorExe = Join-Path $motorDir 'google-maps-scraper.exe'
New-Item -ItemType Directory -Force -Path $motorDir | Out-Null
if ((Test-Path $motorExe) -and ((Get-FileHash $motorExe -Algorithm SHA256).Hash -eq $MOTOR_SHA)) {
    Ok 'ja instalado'
} else {
    Baixar $MOTOR_URL $motorExe $MOTOR_SHA
    Ok 'instalado'
}

# 3. Driver do Playwright -----------------------------------------------------------------------------
Passo "Driver do navegador (Playwright $PW_VERSAO)"
$pwDir = Join-Path $env:LOCALAPPDATA "ms-playwright-go\$PW_VERSAO"
$node = Join-Path $pwDir 'node.exe'
$cli = Join-Path $pwDir 'package\cli.js'
New-Item -ItemType Directory -Force -Path $pwDir | Out-Null
if (-not (Test-Path $cli)) {
    $tgz = Join-Path $env:TEMP "playwright-core-$PW_VERSAO.tgz"
    $tmpDir = Join-Path $env:TEMP "playwright-core-$PW_VERSAO"
    Baixar $PW_URL $tgz $PW_SHA
    if (Test-Path $tmpDir) { Remove-Item -Recurse -Force $tmpDir }
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
    tar.exe -xzf $tgz -C $tmpDir
    if ($LASTEXITCODE -ne 0) { Falha 'nao consegui extrair o playwright-core.' }
    $pkg = Join-Path $pwDir 'package'
    if (Test-Path $pkg) { Remove-Item -Recurse -Force $pkg }
    Move-Item (Join-Path $tmpDir 'package') $pkg
    Remove-Item -Recurse -Force $tmpDir, $tgz
}
if (-not ((Test-Path $node) -and ((Get-FileHash $node -Algorithm SHA256).Hash -eq $NODE_SHA))) {
    Baixar $NODE_URL $node $NODE_SHA
}
$versao = & $node $cli --version
if ($versao -notmatch [regex]::Escape($PW_VERSAO)) { Falha "driver do Playwright nao respondeu como esperado ($versao)." }
Ok $versao

# 4. Chromium -----------------------------------------------------------------------------------------
Passo 'Navegador Chromium (so na primeira vez, cerca de 250 MB)'
& $node $cli install chromium | Out-Host
if ($LASTEXITCODE -ne 0) { Falha 'nao consegui baixar o Chromium. Confira a internet e rode de novo.' }
Ok 'Chromium'

# 5. Configuracao pessoal -----------------------------------------------------------------------------
Passo 'Configuracao'
$config = Join-Path $raiz 'configuracao.txt'
if (Test-Path $config) {
    Ok 'configuracao.txt ja existe (edite no Bloco de Notas se quiser mudar)'
} else {
    $nome = Read-Host '   Seu nome, como vai aparecer nas mensagens ("Aqui e o ...")'
    $modelo = Get-Content (Join-Path $raiz 'configuracao.exemplo.txt') -Encoding UTF8
    $modelo = $modelo -replace '^seu_nome = .*$', "seu_nome = $nome"
    [IO.File]::WriteAllLines($config, $modelo, (New-Object Text.UTF8Encoding $false))
    Ok "configuracao.txt criado (nome: $nome)"
}
& $py (Join-Path $sistema 'buscar_clientes.py') --modelo | Out-Host

# 6. Teste --------------------------------------------------------------------------------------------
Passo 'Ligando o motor para testar'
& (Join-Path $sistema 'iniciar-motor.ps1')
if ($LASTEXITCODE -ne 0) { Falha 'o motor nao ligou.' }

Write-Host ''
Write-Host 'Tudo pronto! Agora e so dar duplo clique em "BUSCAR CLIENTES.bat".' -ForegroundColor Green
