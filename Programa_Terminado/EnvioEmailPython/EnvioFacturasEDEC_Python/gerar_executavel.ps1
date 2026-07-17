$ErrorActionPreference = "Stop"
py -m pip install -r requirements.txt
pyinstaller --noconfirm --clean --windowed --onedir `
  --name "EnvioFacturasEDEC" `
  --add-data "config.ini;." `
  main.py
Write-Host "Executável criado em dist\EnvioFacturasEDEC\EnvioFacturasEDEC.exe"
