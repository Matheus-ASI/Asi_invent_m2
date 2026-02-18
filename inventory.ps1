#requires -version 5.1
param(
    [switch]$TestSoftware
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


# ==================================================================================================
# SCRIPT: Inventory.ps1
# Autores: Marcelo Gamito / Kawan Randoli Monegatto / Matheus Lucizano
# ======================================================================================
# INVENTORY (single-file) — v2.0.3
# - Execução via RAW do GitHub (DownloadString | iex) conforme fluxo previamente definido.
# - Será alocado junto aos scripts padrão da empresa
# - Gera uma pasta C:\TI\<NOME_DA_MAQUINA>
# - Compacta a pasta em C:\TI\<NOME_DA_MAQUINA>.zip
# - Realiza upload do ZIP via FTP
# - Logs em console + arquivo.
# ======================================================================================

function Write-Log {
    param(

        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")][string]$Level = "INFO",
        [Parameter(Mandatory=$false)]$Context
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "INFO"  { "Cyan" }
        "OK"    { "Green" }
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }

    $line = "[$ts] [$Level] $Message"
    Write-Host $line -ForegroundColor $color

    if ($Context -and $Context.LogFile) {
        try {
            $logDir = Split-Path -Parent $Context.LogFile
            if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
            Add-Content -Path $Context.LogFile -Value $line -Encoding UTF8
        } catch {
            # Não quebrar por falha de log
        }
    }
}

function Write-CsvLine {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Line,
        [switch]$Append
    )

    if ($Append) {
        Out-File -FilePath $Path -Encoding UTF8 -Append -InputObject $Line
    } else {
        Out-File -FilePath $Path -Encoding UTF8 -InputObject $Line
    }
}

function Format-DatePtBr {
    param([string]$IsoOrNull, [string]$Format = "dd/MM/yyyy")
    if (-not $IsoOrNull) { return "" }
    try {
        return ([datetime]$IsoOrNull).ToString($Format)
    } catch {
        return ""
    }
}

function Format-DateTimePtBr {
    param([string]$IsoOrNull)
    return (Format-DatePtBr -IsoOrNull $IsoOrNull -Format "dd/MM/yyyy HH:mm:ss")
}

function Get-ShortUserName {
    param([string]$UserName)
    if (-not $UserName) { return $null }
    if ($UserName -match '\\') { return ($UserName.Split('\')[-1]).Trim() }
    return $UserName.Trim()
}

function Join-Values {
    param([object[]]$Values)
    if (-not $Values) { return "" }
    return (($Values | Where-Object { $_ -ne $null -and $_ -ne "" }) -join " | ")
}

function Initialize-InventoryContext {
    param(
        [Parameter(Mandatory=$true)][string]$BasePath,
        [Parameter(Mandatory=$true)][string]$Version
    )

    $machine = $env:COMPUTERNAME
    if (-not $machine) { throw "COMPUTERNAME não encontrado." }

    if (-not ($BasePath -match '^[A-Za-z]:\\')) {
        throw "BasePath inválido: $BasePath"
    }

    $machinePath = Join-Path $BasePath $machine
    $zipName = "$machine.zip"
    $zipPath = Join-Path $BasePath $zipName
    $logFile = Join-Path $machinePath "logs\inventario.log"

    New-Item -ItemType Directory -Path $machinePath -Force | Out-Null

    return [pscustomobject]@{
        Version     = $Version
        BasePath    = $BasePath
        MachineName = $machine
        MachinePath = $machinePath
        ZipName     = $zipName
        ZipPath     = $zipPath
        LogFile     = $logFile
        StartedAt   = (Get-Date)
    }
}

function Get-MetaInfo {
    param([Parameter(Mandatory=$true)]$Context)

    $now = Get-Date
    $user = $env:USERNAME
    $domain = $env:USERDOMAIN

    return [ordered]@{
        script_version = $Context.Version
        generated_at   = $now.ToString("o")
        started_at     = $Context.StartedAt.ToString("o")
        hostname       = $Context.MachineName
        run_as         = if ($domain) { "$domain\$user" } else { $user }
        base_path      = $Context.BasePath
    }
}

function Get-OSInfo {
    param([Parameter(Mandatory=$true)]$Context)

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

        $buildBase = $null
        if ($reg.PSObject.Properties.Name -contains "CurrentBuild") { $buildBase = [string]$reg.CurrentBuild }
        elseif ($reg.PSObject.Properties.Name -contains "CurrentBuildNumber") { $buildBase = [string]$reg.CurrentBuildNumber }
        elseif ($os.BuildNumber) { $buildBase = [string]$os.BuildNumber }

        $buildFull = $buildBase
        if (($reg.PSObject.Properties.Name -contains "UBR") -and $buildBase) {
            $buildFull = "$buildBase.$($reg.UBR)"
        }

        $buildNum = $null
        try { if ($os.BuildNumber) { $buildNum = [int]$os.BuildNumber } } catch {}
        if (-not $buildNum) { try { if ($buildBase) { $buildNum = [int]$buildBase } } catch {} }

        $family = "UNKNOWN"
        if ($buildNum) { if ($buildNum -ge 22000) { $family = "WINDOWS_11" } else { $family = "WINDOWS_10" } }

        $installDate = $null
        try { $installDate = ([System.Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)).ToString("o") } catch {}

        return [ordered]@{
            windows_family     = $family
            caption            = $os.Caption
            display_version    = if ($reg.PSObject.Properties.Name -contains "DisplayVersion") { $reg.DisplayVersion } else { $null }
            version            = $os.Version
            build              = $buildFull
            architecture       = $os.OSArchitecture
            edition_id         = if ($reg.PSObject.Properties.Name -contains "EditionID") { $reg.EditionID } else { $null }
            installation_type  = if ($reg.PSObject.Properties.Name -contains "InstallationType") { $reg.InstallationType } else { $null }
            product_id         = if ($reg.PSObject.Properties.Name -contains "ProductId") { $reg.ProductId } else { $null }
            install_date       = $installDate
        }
    }
    catch {
        Write-Log ("Falha ao coletar OS: " + $_.Exception.Message) "WARN" $Context
        return [ordered]@{ error = $_.Exception.Message }
    }
}

function Get-SoftwareInventory {
    param([Parameter(Mandatory=$true)]$Context)

    $items = New-Object System.Collections.Generic.List[object]
    $toStr = {
        param($v)
        if ($null -eq $v) { return $null }
        if ($v -is [System.Array]) { return ($v -join "; ") }
        return [string]$v
    }
    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($p in $paths) {
        try {
            $apps = Get-ItemProperty $p -ErrorAction SilentlyContinue |
                Where-Object { ($_.PSObject.Properties.Name -contains "DisplayName") -and $_.DisplayName } |
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, PSChildName, UninstallString

            $sourceHive = if ($p.StartsWith("HKCU")) { "HKCU" } else { "HKLM" }
            foreach ($a in $apps) {
                $dt = $null
                if ($a.InstallDate -and ($a.InstallDate -match '^\d{8}$')) {
                    try { $dt = ([datetime]::ParseExact($a.InstallDate,'yyyyMMdd',$null)).ToString("o") } catch {}
                }

                $items.Add([pscustomobject]@{
                    name         = & $toStr $a.DisplayName
                    version      = & $toStr $a.DisplayVersion
                    publisher    = & $toStr $a.Publisher
                    install_date = $dt
                    registry_key = & $toStr $a.PSChildName
                    uninstall    = & $toStr $a.UninstallString
                    source_hive  = $sourceHive
                }) | Out-Null
            }
        } catch {
            Write-Log ("Falha ao coletar softwares em {0}: {1}" -f $p, $_.Exception.Message) "WARN" $Context
        }
    }

    return ($items |
        Group-Object -Property name, version, publisher |
        ForEach-Object { $_.Group | Select-Object -First 1 } |
        Sort-Object name)
}

function Get-UsersInfo {
    param([Parameter(Mandatory=$true)]$Context)

    $active = New-Object System.Collections.Generic.List[object]
    $sessions = New-Object System.Collections.Generic.List[object]
    $profiles = New-Object System.Collections.Generic.List[object]

    # Sessões via CIM (preferência)
    try {
        $logons = Get-CimInstance Win32_LogonSession
        $links  = Get-CimInstance Win32_LoggedOnUser

        $interesting = $logons | Where-Object { $_.LogonType -in 2,10 }  # 2=Interactive, 10=RemoteInteractive

        foreach ($ls in $interesting) {
            $assoc = $links | Where-Object { $_.Dependent -match "LogonId=`"$($ls.LogonId)`"" }
            foreach ($a in $assoc) {
                $m = [regex]::Match($a.Antecedent, 'Domain=`"(?<d>[^`"]+)`".*Name=`"(?<n>[^`"]+)`"')
                if ($m.Success) {
                    $d = $m.Groups['d'].Value
                    $n = $m.Groups['n'].Value
                    $sessions.Add([pscustomobject]@{
                        user       = "$d\$n"
                        logon_id   = $ls.LogonId
                        logon_type = $ls.LogonType
                        start_time = try { ([System.Management.ManagementDateTimeConverter]::ToDateTime($ls.StartTime)).ToString("o") } catch { $null }
                    }) | Out-Null
                }
            }
        }
    } catch {
        Write-Log ("Falha ao coletar sessões CIM: " + $_.Exception.Message) "WARN" $Context
    }

    # Usuário principal (fallback)
    try {
        $cs = Get-CimInstance Win32_ComputerSystem
        if ($cs.UserName) {
            $active.Add([pscustomobject]@{ user = $cs.UserName; source = "Win32_ComputerSystem" }) | Out-Null
        }
    } catch {}

    # Fallback quser se sessões vazias
    if ($sessions.Count -eq 0) {
        try {
            $raw = (quser 2>$null) | Select-Object -Skip 1
            foreach ($line in $raw) {
                $t = ($line -replace '\s{2,}', '|').Trim()
                if ($t) {
                    $parts = $t.Split('|')
                    if ($parts.Count -ge 1) {
                        $u = ($parts[0]).Trim()
                        if ($u -and $u -ne '>') {
                            $sessions.Add([pscustomobject]@{ user = $u; logon_type = "quser"; start_time=$null; logon_id=$null }) | Out-Null
                        }
                    }
                }
            }
        } catch {
            Write-Log ("Falha ao rodar quser: " + $_.Exception.Message) "WARN" $Context
        }
    }

    # Perfis via Win32_UserProfile
    try {
        $ups = Get-CimInstance Win32_UserProfile -ErrorAction Stop | Where-Object {
            $_.LocalPath -and ($_.LocalPath -like "C:\Users\*") -and (-not $_.Special)
        }

        foreach ($p in $ups) {
            $lastUse = $null
            try {
                if ($p.LastUseTime) {
                    $lastUse = ([System.Management.ManagementDateTimeConverter]::ToDateTime($p.LastUseTime)).ToString("o")
                }
            } catch {}

            $ntUser = Join-Path $p.LocalPath "NTUSER.DAT"
            $ntUserLast = $null
            if (Test-Path -LiteralPath $ntUser) {
                try { $ntUserLast = (Get-Item -LiteralPath $ntUser).LastWriteTime.ToString("o") } catch {}
            }

            $profiles.Add([pscustomobject]@{
                local_path        = $p.LocalPath
                sid               = $p.SID
                last_use_time     = $lastUse
                loaded            = [bool]$p.Loaded
                ntuser_last_write = $ntUserLast
            }) | Out-Null
        }
    } catch {
        Write-Log ("Falha ao coletar perfis via Win32_UserProfile: " + $_.Exception.Message) "WARN" $Context
        # fallback: enumerar C:\Users
        try {
            Get-ChildItem "C:\Users" -Directory -ErrorAction Stop |
                Where-Object { $_.Name -notin @("Public","Default","Default User","All Users","Administrador","Administrator") } |
                ForEach-Object {
                    $ntUser = Join-Path $_.FullName "NTUSER.DAT"
                    $profiles.Add([pscustomobject]@{
                        local_path        = $_.FullName
                        sid               = $null
                        last_use_time     = $null
                        loaded            = $null
                        ntuser_last_write = if (Test-Path -LiteralPath $ntUser) { (Get-Item -LiteralPath $ntUser).LastWriteTime.ToString("o") } else { $null }
                    }) | Out-Null
                }
        } catch {
            Write-Log ("Falha ao enumerar C:\Users: " + $_.Exception.Message) "WARN" $Context
        }
    }

    # Marcar ativos com base em sessões (preferência)
    if ($sessions.Count -gt 0) {
        $uniq = $sessions | Group-Object user | ForEach-Object { $_.Group | Select-Object -First 1 }
        foreach ($s in $uniq) {
            $active.Add([pscustomobject]@{ user = $s.user; source = "sessions" }) | Out-Null
        }
    }

    $activeDedup = $active | Group-Object user | ForEach-Object { $_.Group | Select-Object -First 1 }

    return [ordered]@{
        active_users = $activeDedup
        sessions     = $sessions
        profiles     = $profiles | Sort-Object local_path
    }
}

function Get-HardwareInfo {
    param([Parameter(Mandatory=$true)]$Context)

    try {
        $sys  = Get-CimInstance Win32_ComputerSystem
        $bios = Get-CimInstance Win32_BIOS
        $cpu  = Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors
        $video = Get-CimInstance Win32_VideoController | Select-Object Name, VideoProcessor, AdapterRAM
        $mem  = Get-CimInstance Win32_PhysicalMemory | Select-Object DeviceLocator, Capacity, ConfiguredClockSpeed, SMBIOSMemoryType, FormFactor
        $net  = Get-NetAdapter -ErrorAction SilentlyContinue | Select-Object Name, InterfaceDescription, MacAddress, Status, LinkSpeed
        $disk = Get-CimInstance Win32_DiskDrive | Select-Object Model, InterfaceType, MediaType, Size, SerialNumber
        $enclosure = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue

        $chassiCode = $null
        try { if ($enclosure -and $enclosure.ChassisTypes) { $chassiCode = @($enclosure.ChassisTypes)[0] } } catch {}

        $computerType = switch ($chassiCode) {
            3  { "Desktop" }
            6  { "Mini Tower" }
            7  { "Tower" }
            9  { "Laptop" }
            10 { "Notebook" }
            13 { "All in One" }
            default { "Outros" }
        }

        $memTotalGB = [math]::Round((($mem | Measure-Object -Property Capacity -Sum).Sum / 1GB), 2)

        $memoryTypeCode = ($mem | Select-Object -ExpandProperty SMBIOSMemoryType -First 1)
        $memoryType = switch ($memoryTypeCode) {
            24 { "DDR3" }
            26 { "DDR4" }
            34 { "DDR5" }
            default { "Desconhecido" }
        }

        $memoryFormFactor = ($mem | Select-Object -ExpandProperty FormFactor -First 1)
        $memoryFormat = switch ($memoryFormFactor) {
            8  { "DIMM" }
            12 { "SODIMM" }
            default { "Desconhecido" }
        }

        return [ordered]@{
            manufacturer       = $sys.Manufacturer
            model              = $sys.Model
            system_type        = $computerType
            serial_bios        = $bios.SerialNumber
            total_memory_gb    = [math]::Round($sys.TotalPhysicalMemory / 1GB, 2)
            memory_total_gb    = $memTotalGB
            memory_type        = $memoryType
            memory_form_factor = $memoryFormat
            cpu                = $cpu
            video              = ($video | ForEach-Object {
                                    [pscustomobject]@{
                                        name      = $_.Name
                                        processor = $_.VideoProcessor
                                        vram_gb   = if ($_.AdapterRAM) { [math]::Round($_.AdapterRAM / 1GB) } else { $null }
                                    }
                                  })
            memory_modules     = ($mem | ForEach-Object {
                                    [pscustomobject]@{
                                        locator     = $_.DeviceLocator
                                        capacity_gb = if ($_.Capacity) { [math]::Round($_.Capacity / 1GB) } else { $null }
                                        clock_mhz   = $_.ConfiguredClockSpeed
                                    }
                                  })
            network_adapters   = ($net | ForEach-Object {
                                    [pscustomobject]@{
                                        name        = $_.Name
                                        description = $_.InterfaceDescription
                                        mac         = $_.MacAddress
                                        status      = $_.Status
                                        link_speed  = $_.LinkSpeed
                                    }
                                  })
            disks              = ($disk | ForEach-Object {
                                    [pscustomobject]@{
                                        model     = $_.Model
                                        interface = $_.InterfaceType
                                        media_type= $_.MediaType
                                        size_gb   = if ($_.Size) { [math]::Round($_.Size / 1GB) } else { $null }
                                        serial    = $_.SerialNumber
                                    }
                                  })
        }
    }
    catch {
        Write-Log ("Falha ao coletar hardware: " + $_.Exception.Message) "WARN" $Context
        return [ordered]@{ error = $_.Exception.Message }
    }
}

function Get-DevicesInventory {
    param([Parameter(Mandatory=$true)]$Context)

    $list = New-Object System.Collections.Generic.List[object]
    try {
        Get-CimInstance Win32_PnPEntity | Where-Object { $_.Name } | ForEach-Object {
            $list.Add([pscustomobject]@{
                pnp_class = $_.PNPClass
                name      = $_.Name
                status    = $_.Status
                device_id = $_.DeviceID
            }) | Out-Null
        }
    } catch {
        Write-Log ("Falha ao coletar dispositivos: " + $_.Exception.Message) "WARN" $Context
    }
    return $list
}

function New-InventoryZip {
    param([Parameter(Mandatory=$true)]$Context)

    if (Test-Path -LiteralPath $Context.ZipPath) {
        Remove-Item -LiteralPath $Context.ZipPath -Force -ErrorAction SilentlyContinue
    }
    Compress-Archive -Path $Context.MachinePath -DestinationPath $Context.ZipPath -Force
}

function Cleanup-InventoryLocal {
    param([Parameter(Mandatory=$true)]$Context)

    $base = (Resolve-Path $Context.BasePath).Path.TrimEnd("\")
    $mp   = (Resolve-Path $Context.MachinePath).Path

    if (-not ($mp.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase))) {
        throw "Caminho de limpeza inválido (fora do BasePath): $mp"
    }

    Remove-Item -LiteralPath $Context.MachinePath -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $Context.ZipPath -Force -ErrorAction SilentlyContinue
}

function Send-FtpFile {
    param(
        [Parameter(Mandatory=$true)][string]$LocalPath,
        [Parameter(Mandatory=$true)][string]$RemoteFileName,
        [Parameter(Mandatory=$false)]$Context
    )

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Arquivo local não encontrado: $LocalPath"
    }

    # Credenciais FTP (hardcoded por exigência atual)
    $Username = "suporte"
    $Password = "arrayservic3"

    # Endpoint FTP (hardcoded por exigência atual)
    $RemoteUri = "ftp://app01.arrayservice.com.br/$RemoteFileName"

    $req = [System.Net.FtpWebRequest]::Create($RemoteUri)
    $req = [System.Net.FtpWebRequest]$req
    $req.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $req.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
    $req.UseBinary = $true
    $req.UsePassive = $true
    $req.EnableSsl = $false   # mantido (não recomendado)
    $req.KeepAlive = $false

    # Timeouts básicos
    $req.Timeout = 120000
    $req.ReadWriteTimeout = 120000

    $fileStream = $null
    $reqStream  = $null
    $resp       = $null

    try {
        $fileInfo = Get-Item -LiteralPath $LocalPath
        $req.ContentLength = $fileInfo.Length

        $fileStream = [System.IO.File]::OpenRead($LocalPath)
        $reqStream  = $req.GetRequestStream()

        $buffer = New-Object byte[] (64KB)
        while (($read = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $reqStream.Write($buffer, 0, $read)
        }
        # Importante: fechar o stream antes do GetResponse para finalizar o envio
        $reqStream.Flush()
        $reqStream.Close()
        $reqStream.Dispose()
        $reqStream = $null

        $resp = [System.Net.FtpWebResponse]$req.GetResponse()
        if (-not $resp) { throw "Sem resposta do servidor FTP." }

        $status = $resp.StatusDescription
        if ($Context) { Write-Log ("Resposta FTP: " + $status.Trim()) "INFO" $Context }
        return $true
    }
    finally {
        if ($resp) { $resp.Close() }
        if ($reqStream) { $reqStream.Close(); $reqStream.Dispose() }
        if ($fileStream) { $fileStream.Close(); $fileStream.Dispose() }
    }
}

# ============================ MAIN ============================
$invVersion = "2.0.3"
$ctx = $null

if ($TestSoftware) {
    $ctx = [pscustomobject]@{ LogFile = $null }
    Write-Log "TESTE: coletando softwares..." "INFO" $ctx
    $sw = Get-SoftwareInventory -Context $ctx
    $count = if ($sw) { $sw.Count } else { 0 }
    Write-Host ("TESTE: softwares encontrados: {0}" -f $count) -ForegroundColor Green
    if ($sw) { $sw | Select-Object -First 5 | Format-Table -AutoSize }
    exit 0
}

try {
    $ctx = Initialize-InventoryContext -BasePath "C:\TI" -Version $invVersion
    Write-Log "Iniciando inventário (v$invVersion)..." "INFO" $ctx

    $result = [ordered]@{
        meta     = Get-MetaInfo -Context $ctx
        os       = Get-OSInfo -Context $ctx
        users    = Get-UsersInfo -Context $ctx
        hardware = Get-HardwareInfo -Context $ctx
        software = Get-SoftwareInventory -Context $ctx
        devices  = Get-DevicesInventory -Context $ctx
    }

    # Arquivos de saída (mesmos do invent_asi.ps1)
    $arquivoConsolidado = Join-Path $ctx.MachinePath "$($ctx.MachineName).csv"
    $arquivoSoftwares   = Join-Path $ctx.MachinePath "$($ctx.MachineName)_SW.csv"
    $arquivoSO          = Join-Path $ctx.MachinePath "$($ctx.MachineName)_SO.csv"
    $arquivoUsuarios    = Join-Path $ctx.MachinePath "$($ctx.MachineName)_USERS.csv"
    $arquivoHardware    = Join-Path $ctx.MachinePath "$($ctx.MachineName)_HW.csv"
    $arquivoDispositivos= Join-Path $ctx.MachinePath "$($ctx.MachineName)_DEVICES.csv"

    # SOFTWARES
    Write-CsvLine -Path $arquivoSoftwares -Line "NOME_DA_MAQUINA;NOME_DO_SOFTWARE;VERSAO;FABRICANTE;DATA_INSTALACAO;CHAVE"
    foreach ($app in $result.software) {
        $dataInst = Format-DatePtBr $app.install_date
        $line = "$($ctx.MachineName);$($app.name);$($app.version);$($app.publisher);$dataInst;$($app.registry_key)"
        Write-CsvLine -Path $arquivoSoftwares -Line $line -Append
    }
    Write-Log "CSV de softwares gerado: $arquivoSoftwares" "OK" $ctx

    # SISTEMA OPERACIONAL
    Write-CsvLine -Path $arquivoSO -Line "NOME_DA_MAQUINA;WINDOWS_FAMILY;NOME_SO;DISPLAY_VERSION;VERSION;BUILD_COMPLETO;ARQUITETURA;EDITION_ID;INSTALLATION_TYPE;DATA_INSTALACAO;PRODUCT_ID"
    $os = $result.os
    $dataSO = Format-DatePtBr $os.install_date
    $lineSO = "$($ctx.MachineName);$($os.windows_family);$($os.caption);$($os.display_version);$($os.version);$($os.build);$($os.architecture);$($os.edition_id);$($os.installation_type);$dataSO;$($os.product_id)"
    Write-CsvLine -Path $arquivoSO -Line $lineSO -Append
    Write-Log "CSV de sistema operacional gerado: $arquivoSO" "OK" $ctx

    # USUARIOS
    Write-CsvLine -Path $arquivoUsuarios -Line "NOME_DA_MAQUINA;USUARIO;ATIVO;SESSAO;PERFIL;DATA_ALTERACAO_PERFIL;DATA_ALTERACAO_NTUSER"
    $activeShort = @()
    if ($result.users.active_users) {
        $activeShort = $result.users.active_users | ForEach-Object { Get-ShortUserName $_.user }
    }
    $sessionShort = @()
    if ($result.users.sessions) {
        $sessionShort = $result.users.sessions | ForEach-Object { Get-ShortUserName $_.user }
    }

    if ($result.users.profiles -and $result.users.profiles.Count -gt 0) {
        foreach ($p in $result.users.profiles) {
            $userFolder = if ($p.local_path) { Split-Path -Leaf $p.local_path } else { "" }
            $ativo = if ($activeShort -contains $userFolder) { "Sim" } else { "Nao" }
            $sessao = "Desconhecida"
            if ($sessionShort.Count -gt 0) {
                $sessao = if ($sessionShort -contains $userFolder) { "Interativa" } else { "Nao listada" }
            }
            $dataPerfil = Format-DateTimePtBr $p.last_use_time
            $dataNtUser = Format-DateTimePtBr $p.ntuser_last_write
            $lineU = "$($ctx.MachineName);$userFolder;$ativo;$sessao;$($p.local_path);$dataPerfil;$dataNtUser"
            Write-CsvLine -Path $arquivoUsuarios -Line $lineU -Append
        }
    } else {
        # Se não houver perfis, ainda registra usuários ativos (se existirem)
        foreach ($u in $activeShort) {
            $lineU = "$($ctx.MachineName);$u;Sim;Interativa;;;"
            Write-CsvLine -Path $arquivoUsuarios -Line $lineU -Append
        }
    }
    Write-Log "CSV de usuarios gerado: $arquivoUsuarios" "OK" $ctx

    # HARDWARE
    Write-CsvLine -Path $arquivoHardware -Line "NOME_DA_MAQUINA;COMPONENTE;VALOR"
    $hw = $result.hardware
    $cpuNames = Join-Values ($hw.cpu | ForEach-Object { $_.Name })
    $netDesc = Join-Values ($hw.network_adapters | ForEach-Object { $_.description })
    $netMac  = Join-Values ($hw.network_adapters | ForEach-Object { $_.mac })
    $videoName = Join-Values ($hw.video | ForEach-Object { $_.name })
    $videoProc = Join-Values ($hw.video | ForEach-Object { $_.processor })
    $videoVram = Join-Values ($hw.video | ForEach-Object { $_.vram_gb })
    $memLoc = Join-Values ($hw.memory_modules | ForEach-Object { $_.locator })
    $memCap = Join-Values ($hw.memory_modules | ForEach-Object { $_.capacity_gb })
    $memClk = Join-Values ($hw.memory_modules | ForEach-Object { $_.clock_mhz })
    $diskType = Join-Values ($hw.disks | ForEach-Object { $_.interface })
    $diskModel = Join-Values ($hw.disks | ForEach-Object { $_.model })
    $diskSize = Join-Values ($hw.disks | ForEach-Object { $_.size_gb })

    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Modelo;$($hw.model)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Memoria_GB;$($hw.total_memory_gb)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Tipo;$($hw.system_type)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Fabricante;$($hw.manufacturer)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Serial;$($hw.serial_bios)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);CPU;$cpuNames" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);NetworkAdapter;$netDesc" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MacAddress;$netMac" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);VideoBoard;$videoName" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);VideoProcessor;$videoProc" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);VideoCapacity_GB;$videoVram" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryType;$($hw.memory_type)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryFormat;$($hw.memory_form_factor)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryFrequency_MHz;$memClk" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryModule;$memLoc" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryCapacity_GB;$memCap" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);MemoryTotal_GB;$($hw.memory_total_gb)" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);DiskType;$diskType" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);DiskModel;$diskModel" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);DiskSize_GB;$diskSize" -Append
    Write-CsvLine -Path $arquivoHardware -Line "$($ctx.MachineName);Monitor;" -Append
    Write-Log "CSV de hardware gerado: $arquivoHardware" "OK" $ctx

    # DISPOSITIVOS
    Write-CsvLine -Path $arquivoDispositivos -Line "NOME_DA_MAQUINA;CLASSE;DISPOSITIVO;STATUS"
    foreach ($d in $result.devices) {
        $lineD = "$($ctx.MachineName);$($d.pnp_class);$($d.name);$($d.status)"
        Write-CsvLine -Path $arquivoDispositivos -Line $lineD -Append
    }
    Write-Log "CSV de dispositivos gerado: $arquivoDispositivos" "OK" $ctx

    # CSV consolidado
    if (Test-Path $arquivoConsolidado) {
        Remove-Item $arquivoConsolidado -Force -ErrorAction SilentlyContinue
    }
    $arquivosOrigem = @(
        $arquivoSoftwares,
        $arquivoHardware,
        $arquivoSO,
        $arquivoUsuarios,
        $arquivoDispositivos
    )
    foreach ($arquivo in $arquivosOrigem) {
        if (Test-Path $arquivo) {
            Get-Content $arquivo | Add-Content -Encoding UTF8 $arquivoConsolidado
        } else {
            Write-Log "Arquivo nao encontrado para concatenacao: $arquivo" "WARN" $ctx
        }
    }
    Write-Log "CSV consolidado gerado: $arquivoConsolidado" "OK" $ctx

    New-InventoryZip -Context $ctx
    Write-Log "ZIP criado: $($ctx.ZipPath)" "OK" $ctx

    Write-Log "Iniciando upload FTP..." "INFO" $ctx
    Send-FtpFile -LocalPath $ctx.ZipPath -RemoteFileName $ctx.ZipName -Context $ctx
    Write-Log "Upload FTP concluído: $($ctx.ZipName)" "OK" $ctx

    Write-Log "Removendo arquivos locais após upload bem-sucedido..." "INFO" $ctx
    Cleanup-InventoryLocal -Context $ctx
    Write-Log "Limpeza local finalizada." "OK" $ctx
    Write-Log "Inventário finalizado com sucesso." "OK" $ctx
    exit 0
}
catch {
    if ($ctx) {
        Write-Log ("Falha geral: " + $_.Exception.Message) "ERROR" $ctx
        Write-Log "Arquivos locais mantidos para análise." "WARN" $ctx
    } else {
        Write-Host ("[ERROR] Falha antes de inicializar contexto: " + $_.Exception.Message) -ForegroundColor Red
    }
    exit 1
}
