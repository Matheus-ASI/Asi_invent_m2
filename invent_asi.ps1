# ==================================================================================================
# SCRIPT: Inventario_Completo_FTP.ps1
# Autor: Marcelo Gamito / Kawan Randoli Monegatto / Matheus Lucizano

# DESCRICAO:
#   Script responsavel por coletar informacoes da maquina local, gerar relatorios em CSV,
#   compactar os arquivos em uma pasta nomeada com o nome da maquina e realizar upload via FTP.
#
# OBS:
#   - Utiliza o MESMO metodo de autenticacao, servidor e protocolo do script_asi
#   - Gera uma pasta C:\TI\<NOME_DA_MAQUINA>
#   - Compacta a pasta em C:\TI\<NOME_DA_MAQUINA>.zip
#   - Realiza upload do ZIP via FTP
#
# ==================================================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        "INFO"  { $Color = "Cyan" }
        "OK"    { $Color = "Green" }
        "WARN"  { $Color = "Yellow" }
        "ERROR" { $Color = "Red" }
    }

    Write-Host "[$Timestamp] [$Level] $Message" -ForegroundColor $Color
}

# ==================================================================================================
# FUNCAO: FtpConnection
# DESCRICAO:
#   Realiza upload de um arquivo local para o servidor FTP utilizando FtpWebRequest.
# ==================================================================================================
function FtpConnection {
    param (
        $localFile,
        $fileName
    )

    # Credenciais FTP (padrao script_asi)
    $Username = "suporte"
    $Password = "arrayservic3"

    # Endereco FTP
    $RemoteFile = "ftp://app01.arrayservice.com.br/$fileName"

    $FTPRequest = [System.Net.FtpWebRequest]::Create("$RemoteFile")
    $FTPRequest = [System.Net.FtpWebRequest]$FTPRequest
    $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $FTPRequest.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
    $FTPRequest.UseBinary = $true
    $FTPRequest.UsePassive = $true
    $FTPRequest.EnableSsl = $false

    $FileContent = [System.IO.File]::ReadAllBytes($localFile)
    $FTPRequest.ContentLength = $FileContent.Length

    $Run = $FTPRequest.GetRequestStream()
    $Run.Write($FileContent, 0, $FileContent.Length)
    $Run.Close()
    $Run.Dispose()
}


# ==================================================================================================
# VARIAVEIS GERAIS
# ==================================================================================================
Write-Log "Iniciando coleta de inventario da maquina local." "INFO"

$BasePath     = "C:\TI"
$NomeMaquina  = $env:COMPUTERNAME
$MachinePath = Join-Path $BasePath $NomeMaquina

New-Item -ItemType Directory -Path $MachinePath -Force | Out-Null
Write-Log "Pasta de trabalho criada: $MachinePath" "OK"


# ==================================================================================================
# ARQUIVOS DE SAIDA
# ==================================================================================================
$ArquivoConsolidado = Join-Path $MachinePath "$NomeMaquina.csv"
$ArquivoSoftwares   = Join-Path $MachinePath "${NomeMaquina}_SW.csv"
$ArquivoSO          = Join-Path $MachinePath "${NomeMaquina}_SO.csv"
$ArquivoUsuarios    = Join-Path $MachinePath "${NomeMaquina}_USERS.csv"
$ArquivoHardware    = Join-Path $MachinePath "${NomeMaquina}_HW.csv"
$ArquivoDispositivos= Join-Path $MachinePath "${NomeMaquina}_DEVICES.csv"



# ==================================================================================================
# INVENTARIO DE SOFTWARES
# ==================================================================================================
Write-Log "Coletando inventario de softwares..." "INFO"

"NOME_DA_MAQUINA;NOME_DO_SOFTWARE;VERSAO;FABRICANTE;DATA_INSTALACAO;CHAVE" |
    Out-File -Encoding UTF8 $ArquivoSoftwares

$Softwares = Get-ItemProperty `
    HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
    HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName } |
    Sort-Object DisplayName

foreach ($App in $Softwares) {

    $DataInstalacao = ""
    if ($App.InstallDate -match '^\d{8}$') {
        try {
            $DataInstalacao = ([datetime]::ParseExact(
                $App.InstallDate,'yyyyMMdd',$null)).ToString('dd/MM/yyyy')
        } catch {}
    }

    "$NomeMaquina;$($App.DisplayName);$($App.DisplayVersion);$($App.Publisher);$DataInstalacao;" |
        Out-File -Append -Encoding UTF8 $ArquivoSoftwares
}
Write-Log "Arquivo de softwares gerado: $ArquivoSoftwares" "OK"


# ==================================================================================================
# INFORMACOES DO SISTEMA OPERACIONAL
# ==================================================================================================
Write-Log "Coletando informacoes do sistema operacional..." "INFO"

"NOME_DA_MAQUINA;NOME_SO;VERSAO;FABRICANTE;DATA_INSTALACAO;PRODUCT_ID" |
    Out-File -Encoding UTF8 $ArquivoSO

try {
    $OS = Get-CimInstance Win32_OperatingSystem
    $Reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

    $DataSO = ([System.Management.ManagementDateTimeConverter]::ToDateTime(
        $OS.InstallDate)).ToString('dd/MM/yyyy')

    "$NomeMaquina;$($OS.Caption);$($Reg.DisplayVersion);$($OS.Manufacturer);$DataSO;$($Reg.ProductId)" |
        Out-File -Append -Encoding UTF8 $ArquivoSO
} catch {
    "$NomeMaquina;ERRO;;;;" | Out-File -Append -Encoding UTF8 $ArquivoSO
    Write-Log "Falha ao coletar informacoes do sistema operacional." "WARN"
}
Write-Log "Arquivo de sistema operacional gerado: $ArquivoSO" "OK"


# ==================================================================================================
# USUARIOS LOCAIS
# ==================================================================================================
Write-Log "Coletando usuarios locais..." "INFO"

"NOME_DA_MAQUINA;USUARIO;STATUS;ULTIMO_LOGIN" |
    Out-File -Encoding UTF8 $ArquivoUsuarios

try {
    Get-LocalUser | ForEach-Object {
        $Status = if ($_.Enabled) { "Ativo" } else { "Desativado" }
        $Login  = if ($_.LastLogon) { $_.LastLogon } else { "Nunca" }
        "$NomeMaquina;$($_.Name);$Status;$Login" |
            Out-File -Append -Encoding UTF8 $ArquivoUsuarios
    }
} catch {
    "$NomeMaquina;ERRO;;" | Out-File -Append -Encoding UTF8 $ArquivoUsuarios
    Write-Log "Falha ao coletar usuarios locais." "WARN"
}
Write-Log "Arquivo de usuarios gerado: $ArquivoUsuarios" "OK"


# ==================================================================================================
# HARDWARE
# ==================================================================================================
Write-Log "Coletando dados de hardware..." "INFO"

"NOME_DA_MAQUINA;COMPONENTE;VALOR" |
    Out-File -Encoding UTF8 $ArquivoHardware

try {
    $Sys = Get-CimInstance Win32_ComputerSystem
    "$NomeMaquina;Modelo;$($Sys.Model)" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Memoria_GB;$([math]::Round($Sys.TotalPhysicalMemory / 1GB,2))" |
        Out-File -Append -Encoding UTF8 $ArquivoHardware
} catch {
    Write-Log "Falha ao coletar dados de hardware." "WARN"
}
Write-Log "Arquivo de hardware gerado: $ArquivoHardware" "OK"


# ==================================================================================================
# DISPOSITIVOS
# ==================================================================================================
Write-Log "Coletando lista de dispositivos..." "INFO"

"NOME_DA_MAQUINA;CLASSE;DISPOSITIVO;STATUS" |
    Out-File -Encoding UTF8 $ArquivoDispositivos

try {
    Get-CimInstance Win32_PnPEntity | Where-Object { $_.Name } | ForEach-Object {
        "$NomeMaquina;$($_.PNPClass);$($_.Name);$($_.Status)" |
            Out-File -Append -Encoding UTF8 $ArquivoDispositivos
    }
} catch {
    "$NomeMaquina;ERRO;;" | Out-File -Append -Encoding UTF8 $ArquivoDispositivos
    Write-Log "Falha ao coletar dispositivos." "WARN"
}
Write-Log "Arquivo de dispositivos gerado: $ArquivoDispositivos" "OK"


# ==================================================================================================
# CSV CONSOLIDADO (NOME EXATO DA MAQUINA)
# DESCRICAO:
#   Concatena todos os CSVs gerados no arquivo <NOME_DA_MAQUINA>.csv
# ==================================================================================================
Write-Log "Montando CSV consolidado da maquina: $ArquivoConsolidado" "INFO"

if (Test-Path $ArquivoConsolidado) {
    Remove-Item $ArquivoConsolidado -Force -ErrorAction SilentlyContinue
}

$ArquivosOrigem = @(
    $ArquivoSoftwares,
    $ArquivoSO,
    $ArquivoUsuarios,
    $ArquivoHardware,
    $ArquivoDispositivos
)

foreach ($Arquivo in $ArquivosOrigem) {
    if (Test-Path $Arquivo) {
        Get-Content $Arquivo | Add-Content -Encoding UTF8 $ArquivoConsolidado
    } else {
        Write-Log "Arquivo nao encontrado para concatenacao: $Arquivo" "WARN"
    }
}

Write-Log "Arquivo consolidado gerado com sucesso: $ArquivoConsolidado" "OK"


# ==================================================================================================
# COMPACTACAO DA PASTA DA MAQUINA
# ==================================================================================================
Write-Log "Compactando pasta de inventario..." "INFO"

$ZipPath  = Join-Path $BasePath "$NomeMaquina.zip"
$ZipName  = "$NomeMaquina.zip"

if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

Compress-Archive -Path $MachinePath -DestinationPath $ZipPath -Force
Write-Log "Arquivo ZIP criado: $ZipPath" "OK"


# ==================================================================================================
# UPLOAD FTP
# ==================================================================================================
Write-Log "Enviando inventario via FTP para o servidor..." "INFO"
$UploadOK = $false

try {
    FtpConnection $ZipPath $ZipName
    $UploadOK = $true
    Write-Log "Upload FTP concluido: $ZipName" "OK"
} catch {
    Write-Log "Falha no upload FTP: $($_.Exception.Message)" "ERROR"
}


# ==================================================================================================
# LIMPEZA LOCAL
# ==================================================================================================
if ($UploadOK) {
    Write-Log "Removendo arquivos locais apos upload bem-sucedido..." "INFO"
    Remove-Item -Path $MachinePath -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue
    Write-Log "Limpeza local finalizada." "OK"
    Write-Log "Inventario enviado com sucesso: $ZipName" "OK"
} else {
    Write-Log "Arquivos locais mantidos para analise devido falha no upload." "WARN"
    Write-Log "Execucao finalizada com erro." "ERROR"
}
