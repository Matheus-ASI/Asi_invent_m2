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

"NOME_DA_MAQUINA;WINDOWS_FAMILY;NOME_SO;DISPLAY_VERSION;VERSION;BUILD_COMPLETO;ARQUITETURA;EDITION_ID;INSTALLATION_TYPE;DATA_INSTALACAO;PRODUCT_ID" |
    Out-File -Encoding UTF8 $ArquivoSO

try {
    $OS = Get-CimInstance Win32_OperatingSystem
    $Reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

    $NomeSO = $OS.Caption
    $DisplayVersion = if ($Reg.PSObject.Properties.Name -contains "DisplayVersion") { $Reg.DisplayVersion } else { "" }
    $Version = $OS.Version
    $Arquitetura = $OS.OSArchitecture
    $EditionId = if ($Reg.PSObject.Properties.Name -contains "EditionID") { $Reg.EditionID } else { "" }
    $InstallationType = if ($Reg.PSObject.Properties.Name -contains "InstallationType") { $Reg.InstallationType } else { "" }
    $ProductId = if ($Reg.PSObject.Properties.Name -contains "ProductId") { $Reg.ProductId } else { "" }

    $DataSO = ""
    try {
        $DataSO = ([System.Management.ManagementDateTimeConverter]::ToDateTime(
            $OS.InstallDate)).ToString('dd/MM/yyyy')
    } catch {}

    $BuildBase = ""
    if ($Reg.PSObject.Properties.Name -contains "CurrentBuild") {
        $BuildBase = "$($Reg.CurrentBuild)"
    } elseif ($Reg.PSObject.Properties.Name -contains "CurrentBuildNumber") {
        $BuildBase = "$($Reg.CurrentBuildNumber)"
    } elseif ($OS.BuildNumber) {
        $BuildBase = "$($OS.BuildNumber)"
    }

    $BuildCompleto = $BuildBase
    if ($Reg.PSObject.Properties.Name -contains "UBR" -and $BuildBase) {
        $BuildCompleto = "$BuildBase.$($Reg.UBR)"
    }

    # Windows Family (Win10 vs Win11) por build
    $BuildNum = $null

    try {
        if ($OS.BuildNumber) { $BuildNum = [int]$OS.BuildNumber }
    } catch {}

    if (-not $BuildNum) {
        try {
            if ($Reg.PSObject.Properties.Name -contains "CurrentBuild") {
                $BuildNum = [int]$Reg.CurrentBuild
            } elseif ($Reg.PSObject.Properties.Name -contains "CurrentBuildNumber") {
                $BuildNum = [int]$Reg.CurrentBuildNumber
            }
        } catch {}
    }

    $WindowsFamily = "UNKNOWN"
    if ($BuildNum) {
        if ($BuildNum -ge 22000) { $WindowsFamily = "WINDOWS_11" }
        else { $WindowsFamily = "WINDOWS_10" }
    }

    "$NomeMaquina;$WindowsFamily;$NomeSO;$DisplayVersion;$Version;$BuildCompleto;$Arquitetura;$EditionId;$InstallationType;$DataSO;$ProductId" |
        Out-File -Append -Encoding UTF8 $ArquivoSO
} catch {
    (@($NomeMaquina, "UNKNOWN", "", "", "", "", "", "", "", "", "") -join ";") |
        Out-File -Append -Encoding UTF8 $ArquivoSO
    Write-Log "Falha ao coletar informacoes do sistema operacional." "WARN"
}
Write-Log "Arquivo de sistema operacional gerado: $ArquivoSO" "OK"


# ==================================================================================================
# USUARIOS (USUARIO ATIVO + PERFIS EM C:\Users)
# DESCRICAO:
#   - Usuario ativo: Win32_ComputerSystem.UserName (fallback: quser)
#   - Perfis: pastas em C:\Users
#   - Datas: LastWriteTime da pasta do perfil e do arquivo NTUSER.DAT
# ==================================================================================================
Write-Log "Coletando usuario ativo e perfis em C:\Users..." "INFO"

"NOME_DA_MAQUINA;USUARIO;ATIVO;SESSAO;PERFIL;DATA_ALTERACAO_PERFIL;DATA_ALTERACAO_NTUSER" |
    Out-File -Encoding UTF8 $ArquivoUsuarios

# Usuario ativo (principal)
$UsuarioAtivo = $null
try {
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $UsuarioAtivo = $cs.UserName
} catch {}

# Fallback: tentar quser (sessao interativa)
$Sessoes = @()
try {
    $raw = (quser 2>$null) | Select-Object -Skip 1
    foreach ($line in $raw) {
        $t = ($line -replace '\s{2,}', '|').Trim()
        if ($t) {
            $parts = $t.Split('|')
            if ($parts.Count -ge 2) {
                $u = ($parts[0]).Trim()
                if ($u -and $u -ne '>') {
                    $Sessoes += $u
                }
            }
        }
    }
} catch {}

function Get-ShortUser {
    param([string]$UserName)
    if (-not $UserName) { return $null }
    if ($UserName -match '\\') { return ($UserName.Split('\')[-1]).Trim() }
    return $UserName.Trim()
}

$UsuarioAtivoShort = Get-ShortUser $UsuarioAtivo

# Pastas de perfil em C:\Users
$Perfis = @()
try {
    $Perfis = Get-ChildItem "C:\Users" -Directory -ErrorAction Stop |
        Where-Object {
            $_.Name -notin @("Public","Default","Default User","All Users","Administrador","Administrator")
        }
} catch {
    Write-Log "Falha ao enumerar C:\Users." "WARN"
}

foreach ($p in $Perfis) {

    $UserFolder = $p.Name
    $PerfilPath = $p.FullName

    $DataPerfil = ""
    try { $DataPerfil = (Get-Item $PerfilPath -ErrorAction Stop).LastWriteTime.ToString("dd/MM/yyyy HH:mm:ss") } catch {}

    $NtUserPath = Join-Path $PerfilPath "NTUSER.DAT"
    $DataNtUser = ""
    if (Test-Path $NtUserPath) {
        try { $DataNtUser = (Get-Item $NtUserPath -ErrorAction Stop).LastWriteTime.ToString("dd/MM/yyyy HH:mm:ss") } catch {}
    }

    $Ativo = "Nao"
    if ($UsuarioAtivoShort -and ($UserFolder -ieq $UsuarioAtivoShort)) { $Ativo = "Sim" }

    $Sessao = "Desconhecida"
    if ($Sessoes.Count -gt 0) {
        $Sessao = if ($Sessoes -contains $UserFolder) { "Interativa" } else { "Nao listada" }
    }

    "$NomeMaquina;$UserFolder;$Ativo;$Sessao;$PerfilPath;$DataPerfil;$DataNtUser" |
        Out-File -Append -Encoding UTF8 $ArquivoUsuarios
}

# Se nao achou perfis, ainda registra o usuario ativo (se existir)
if (($Perfis.Count -eq 0) -and $UsuarioAtivoShort) {
    "$NomeMaquina;$UsuarioAtivoShort;Sim;Interativa;;;" |
        Out-File -Append -Encoding UTF8 $ArquivoUsuarios
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
    $Bios = Get-CimInstance Win32_BIOS
    $ChassiTypes = Get-WmiObject -Class Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes
    $cpuInfo = (Get-CimInstance Win32_Processor).Name -join " | "
    $netInfo = Get-NetAdapter -Physical
    $videoInfo = Get-CimInstance Win32_VideoController
    $memoryInfo = Get-CimInstance Win32_PhysicalMemory
    $diskInfo = Get-Disk
    $monitorInfo = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID |
        Select-Object @{n="Model";e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)}}

    $chassiType = if ($ChassiTypes -is [array]) { $ChassiTypes[0] } else { $ChassiTypes }
    switch ($chassiType) {
        "3"  { $computerType = "Desktop" }
        "6"  { $computerType = "Mini Tower" }
        "7"  { $computerType = "Tower" }
        "9"  { $computerType = "Laptop" }
        "10" { $computerType = "Notebook" }
        "13" { $computerType = "All in One" }
        Default { $computerType = "Outros" }
    }

    $memoryTypeCode = ($memoryInfo | Select-Object -ExpandProperty SMBIOSMemoryType -First 1)
    switch ($memoryTypeCode) {
        "24" { $memoryType = "DDR3" }
        "26" { $memoryType = "DDR4" }
        "34" { $memoryType = "DDR5" }
        Default { $memoryType = "Desconhecido" }
    }

    $memoryFormFactor = ($memoryInfo | Select-Object -ExpandProperty FormFactor -First 1)
    switch ($memoryFormFactor) {
        "8"  { $memoryFormat = "DIMM" }
        "12" { $memoryFormat = "SODIMM" }
        Default { $memoryFormat = "Desconhecido" }
    }

    $netInterface = $netInfo.InterfaceDescription -join " | "
    $netMac = $netInfo.MacAddress -join " | "
    $videoModel = $videoInfo.Name -join " | "
    $videoProcessor = $videoInfo.VideoProcessor -join " | "
    $videoCapacity = ($videoInfo.AdapterRAM | ForEach-Object { [math]::Round($_ / 1GB) }) -join " | "
    $memoryModule = $memoryInfo.DeviceLocator -join " | "
    $memoryCapacity = ($memoryInfo.Capacity | ForEach-Object { [math]::Round($_ / 1GB) }) -join " | "
    $memoryFrequency = ($memoryInfo.ConfiguredClockSpeed | Where-Object { $_ }) -join " | "
    $memoryTotal = [math]::Round((($memoryInfo | Measure-Object -Property Capacity -Sum).Sum / 1GB), 2)
    $diskType = ($diskInfo.BusType | Where-Object { $_ -in @("SATA","NVMe") }) -join " | "
    $diskModel = $diskInfo.Model -join " | "
    $diskSize = ($diskInfo.Size | ForEach-Object { [math]::Round($_ / 1GB) }) -join " | "
    $monitorModel = ($monitorInfo | ForEach-Object { $_.Model -replace "`0","" }) -join " | "

    "$NomeMaquina;Modelo;$($Sys.Model)" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Memoria_GB;$([math]::Round($Sys.TotalPhysicalMemory / 1GB,2))" |
        Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Tipo;$computerType" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Fabricante;$($Sys.Manufacturer)" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Serial;$($Bios.SerialNumber)" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;CPU;$cpuInfo" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;NetworkAdapter;$netInterface" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MacAddress;$netMac" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;VideoBoard;$videoModel" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;VideoProcessor;$videoProcessor" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;VideoCapacity_GB;$videoCapacity" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryType;$memoryType" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryFormat;$memoryFormat" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryFrequency_MHz;$memoryFrequency" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryModule;$memoryModule" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryCapacity_GB;$memoryCapacity" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;MemoryTotal_GB;$memoryTotal" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;DiskType;$diskType" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;DiskModel;$diskModel" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;DiskSize_GB;$diskSize" | Out-File -Append -Encoding UTF8 $ArquivoHardware
    "$NomeMaquina;Monitor;$monitorModel" | Out-File -Append -Encoding UTF8 $ArquivoHardware
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
    $ArquivoHardware,
    $ArquivoSO,
    $ArquivoUsuarios,
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
