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
    $Run.Dispose
}


# ==================================================================================================
# VARIAVEIS GERAIS
# ==================================================================================================
$BasePath     = "C:\TI"
$NomeMaquina  = $env:COMPUTERNAME
$MachinePath = Join-Path $BasePath $NomeMaquina

New-Item -ItemType Directory -Path $MachinePath -Force | Out-Null


# ==================================================================================================
# ARQUIVOS DE SAIDA
# ==================================================================================================
$ArquivoSoftwares   = Join-Path $MachinePath "${NomeMaquina}_SW.csv"
$ArquivoSO          = Join-Path $MachinePath "${NomeMaquina}_SO.csv"
$ArquivoUsuarios    = Join-Path $MachinePath "${NomeMaquina}_USERS.csv"
$ArquivoHardware    = Join-Path $MachinePath "${NomeMaquina}_HW.csv"
$ArquivoDispositivos= Join-Path $MachinePath "${NomeMaquina}_DEVICES.csv"


# ==================================================================================================
# INVENTARIO DE SOFTWARES
# ==================================================================================================
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


# ==================================================================================================
# INFORMACOES DO SISTEMA OPERACIONAL
# ==================================================================================================
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
}


# ==================================================================================================
# USUARIOS LOCAIS
# ==================================================================================================
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
}


# ==================================================================================================
# HARDWARE
# ==================================================================================================
"NOME_DA_MAQUINA;COMPONENTE;VALOR" |
    Out-File -Encoding UTF8 $ArquivoHardware

try {
    $Sys = Get-CimInstance Win32_ComputerSystem
    "$NomeMaquina;Modelo;$($Sys.Model)" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Memoria_GB;$([math]::Round($Sys.TotalPhysicalMemory / 1GB,2))" |
        Out-File -Append -Encoding UTF8 $ArquivoHardware
} catch {}


# ==================================================================================================
# DISPOSITIVOS
# ==================================================================================================
"NOME_DA_MAQUINA;CLASSE;DISPOSITIVO;STATUS" |
    Out-File -Encoding UTF8 $ArquivoDispositivos

try {
    Get-CimInstance Win32_PnPEntity | Where-Object { $_.Name } | ForEach-Object {
        "$NomeMaquina;$($_.PNPClass);$($_.Name);$($_.Status)" |
            Out-File -Append -Encoding UTF8 $ArquivoDispositivos
    }
} catch {
    "$NomeMaquina;ERRO;;" | Out-File -Append -Encoding UTF8 $ArquivoDispositivos
}


# ==================================================================================================
# COMPACTACAO DA PASTA DA MAQUINA
# ==================================================================================================
$ZipPath  = Join-Path $BasePath "$NomeMaquina.zip"
$ZipName  = "$NomeMaquina.zip"

if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

Compress-Archive -Path $MachinePath -DestinationPath $ZipPath -Force


# ==================================================================================================
# UPLOAD FTP
# ==================================================================================================
FtpConnection $ZipPath $ZipName


# ==================================================================================================
# LIMPEZA LOCAL
# ==================================================================================================
Remove-Item -Path $MachinePath -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue

Write-Host "Inventario enviado com sucesso: $ZipName"
