# Inventário ASI — execução via GitHub RAW

Abaixo estão as formas corretas de chamada para cada script.

## Inventory.ps1 (CSV + ZIP + FTP) — v2.0.3

Comando recomendado (download para `%TEMP%` e execução):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $u='https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/Inventory.ps1'; $p=Join-Path $env:TEMP 'Inventory.ps1'; iwr -UseBasicParsing $u -OutFile $p; Unblock-File $p -ErrorAction SilentlyContinue; & $p"
```

Comando alternativo via `iex`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iex ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/Inventory.ps1'))"
```

## invent_asi.ps1 (CSV legado + ZIP + FTP)

Comando recomendado (download para `%TEMP%` e execução):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $u='https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/invent_asi.ps1'; $p=Join-Path $env:TEMP 'invent_asi.ps1'; iwr -UseBasicParsing $u -OutFile $p; Unblock-File $p -ErrorAction SilentlyContinue; & $p"
```

Comando alternativo via `iex`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iex ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/invent_asi.ps1'))"
```

## invent_asi_m2.ps1 (JSON + ZIP + FTP) — v2.0.2

Comando recomendado (download para `%TEMP%` e execução):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $u='https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/invent_asi_m2.ps1'; $p=Join-Path $env:TEMP 'invent_asi_m2.ps1'; iwr -UseBasicParsing $u -OutFile $p; Unblock-File $p -ErrorAction SilentlyContinue; & $p"
```

Comando alternativo via `iex`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iex ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/invent_asi_m2.ps1'))"
```

## Saídas

`Inventory.ps1` e `invent_asi.ps1`:
- `C:\TI\<NOME_DA_MAQUINA>\*.csv`
- `C:\TI\<NOME_DA_MAQUINA>.zip`
- Log: `C:\TI\<NOME_DA_MAQUINA>\logs\inventario.log`

`invent_asi_m2.ps1`:
- `C:\TI\<NOME_DA_MAQUINA>\inventory.json`
- `C:\TI\<NOME_DA_MAQUINA>.zip`
- Log: `C:\TI\<NOME_DA_MAQUINA>\logs\inventario.log`

