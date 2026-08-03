$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURAÇÃO
# ============================================================

# Pasta final onde será colocado o executável
$DESTINO = "C:\Users\EDEC\Elaboração de Programas\envio-facturas-email\Programa_Terminado\EnvioEmailPython\Envio_Facturas_Email"

# Pasta atual do projeto
$PROJETO = $PSScriptRoot

# ============================================================
# CRIAR PASTA DE DESTINO
# ============================================================

if (!(Test-Path $DESTINO)) {
    New-Item -ItemType Directory -Path $DESTINO -Force | Out-Null
}

Write-Host ""
Write-Host "============================================="
Write-Host "     EnvioFacturasEDEC - BUILD"
Write-Host "============================================="
Write-Host ""

Write-Host "Projeto:"
Write-Host $PROJETO

Write-Host ""
Write-Host "Destino:"
Write-Host $DESTINO
Write-Host ""

# ============================================================
# INSTALAR DEPENDÊNCIAS
# ============================================================

Write-Host "Instalando/verificando dependencias..."
Write-Host ""

python -m pip install -r "$PROJETO\requirements.txt"

# Garantir que PyInstaller está instalado
python -m pip install pyinstaller

# ============================================================
# LIMPAR BUILD ANTERIOR
# ============================================================

Write-Host ""
Write-Host "Limpando build anterior..."

if (Test-Path "$PROJETO\build") {
    Remove-Item "$PROJETO\build" -Recurse -Force
}

if (Test-Path "$PROJETO\EnvioFacturasEDEC.spec") {
    Remove-Item "$PROJETO\EnvioFacturasEDEC.spec" -Force
}

# Remover executável anterior
if (Test-Path "$DESTINO\EnvioFacturasEDEC.exe") {
    Remove-Item "$DESTINO\EnvioFacturasEDEC.exe" -Force
}

# ============================================================
# GERAR EXECUTÁVEL
# ============================================================

Write-Host ""
Write-Host "Gerando executavel..."
Write-Host ""

Set-Location $PROJETO

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "EnvioFacturasEDEC" `
    --distpath "$DESTINO" `
    --workpath "$PROJETO\build" `
    "$PROJETO\main.py"

# ============================================================
# COPIAR CONFIG.INI
# ============================================================

Write-Host ""
Write-Host "Copiando config.ini..."

if (Test-Path "$PROJETO\config.ini") {

    Copy-Item `
        "$PROJETO\config.ini" `
        "$DESTINO\config.ini" `
        -Force

    Write-Host "config.ini copiado."
}
else {
    Write-Host "AVISO: config.ini nao encontrado."
}

# ============================================================
# COPIAR ASSETS
# ============================================================

if (Test-Path "$PROJETO\assets") {

    Write-Host ""
    Write-Host "Copiando pasta assets..."

    if (Test-Path "$DESTINO\assets") {
        Remove-Item "$DESTINO\assets" -Recurse -Force
    }

    Copy-Item `
        "$PROJETO\assets" `
        "$DESTINO\assets" `
        -Recurse `
        -Force

    Write-Host "Assets copiados."
}

# ============================================================
# FINALIZAÇÃO
# ============================================================

Write-Host ""
Write-Host "============================================="
Write-Host " EXECUTAVEL CRIADO COM SUCESSO!"
Write-Host "============================================="
Write-Host ""

Write-Host "Executavel:"
Write-Host "$DESTINO\EnvioFacturasEDEC.exe"

Write-Host ""
Write-Host "Configuracao:"
Write-Host "$DESTINO\config.ini"

Write-Host ""
Write-Host "============================================="