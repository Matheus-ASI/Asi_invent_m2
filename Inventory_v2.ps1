#requires -version 5.1
param(
    [switch]$TestSoftware
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


# ==================================================================================================
# SCRIPT: Inventory_v2.ps1
# Autores: Marcelo Gamito / Kawan Randoli Monegatto / Matheus Lucizano
# ======================================================================================
# INVENTORY (single-file) — v2.1.0
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
    if ($Format -eq "dd/MM/yyyy") {
        return (Normalize-Date -IsoOrNull $IsoOrNull)
    }
    if ($Format -eq "dd/MM/yyyy HH:mm:ss") {
        return (Normalize-Date -IsoOrNull $IsoOrNull -WithTime)
    }
    if (-not $IsoOrNull) { return "" }
    try { return ([datetime]$IsoOrNull).ToString($Format) } catch { return "" }
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

function Get-Utf8BomEncoding {
    return (New-Object System.Text.UTF8Encoding($true))
}

function Write-AllLinesUtf8 {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Lines
    )

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    [System.IO.File]::WriteAllLines($Path, $Lines, (Get-Utf8BomEncoding))
}

function Normalize-Text {
    param(
        $Value,
        [string]$NullValue = ""
    )

    if ($null -eq $Value) { return $NullValue }

    $text = $null
    if ($Value -is [System.Array]) {
        $text = (($Value | ForEach-Object { [string]$_ }) -join " | ")
    } else {
        $text = [string]$Value
    }

    $text = $text -replace '[\r\n]+', ' '
    $text = $text.Replace(";", ",")
    $text = $text.Replace('"', "'")
    return $text.Trim()
}

function Normalize-Date {
    param(
        [string]$IsoOrNull,
        [switch]$WithTime
    )

    if (-not $IsoOrNull) { return "" }
    try {
        $dt = [datetime]$IsoOrNull
        if ($WithTime) { return $dt.ToString("dd/MM/yyyy HH:mm:ss") }
        return $dt.ToString("dd/MM/yyyy")
    } catch {
        return ""
    }
}

function Normalize-User {
    param(
        [string]$UserName,
        [switch]$Lower
    )

    $short = Get-ShortUserName $UserName
    if (-not $short) { return "" }
    $short = Normalize-Text $short
    if ($Lower) { return $short.ToLowerInvariant() }
    return $short
}

function New-CsvSchemaMap {
    return [ordered]@{
        software = @(
            "NOME_DA_MAQUINA","NOME_DO_SOFTWARE","VERSAO","FABRICANTE","DATA_INSTALACAO","CHAVE"
        )
        os = @(
            "NOME_DA_MAQUINA","WINDOWS_FAMILY","NOME_SO","DISPLAY_VERSION","VERSION","BUILD_COMPLETO","ARQUITETURA","EDITION_ID","INSTALLATION_TYPE","DATA_INSTALACAO","PRODUCT_ID"
        )
        users = @(
            "NOME_DA_MAQUINA","USUARIO","ATIVO","SESSAO","PERFIL","DATA_ALTERACAO_PERFIL","DATA_ALTERACAO_NTUSER"
        )
        hardware = @(
            "NOME_DA_MAQUINA","COMPONENTE","VALOR"
        )
        devices = @(
            "NOME_DA_MAQUINA","CLASSE","DISPOSITIVO","STATUS"
        )
    }
}

function New-CsvBuffer {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Columns
    )

    return [pscustomobject]@{
        Path    = $Path
        Columns = @($Columns)
        Header  = (@($Columns | ForEach-Object { Normalize-Text $_ }) -join ";")
        Rows    = (New-Object System.Collections.Generic.List[string])
    }
}

function Add-CsvRowSafe {
    param(
        [Parameter(Mandatory=$true)]$Buffer,
        [Parameter(Mandatory=$true)][hashtable]$Row
    )

    $fields = New-Object System.Collections.Generic.List[string]
    foreach ($col in $Buffer.Columns) {
        $val = $null
        if ($Row.ContainsKey($col)) { $val = $Row[$col] }
        $fields.Add((Normalize-Text $val)) | Out-Null
    }

    if ($fields.Count -ne $Buffer.Columns.Count) {
        throw ("Quantidade de colunas inválida para {0}: esperado={1}, atual={2}" -f $Buffer.Path, $Buffer.Columns.Count, $fields.Count)
    }

    $Buffer.Rows.Add(($fields -join ";")) | Out-Null
}

function Validate-CsvRowCount {
    param(
        [Parameter(Mandatory=$true)]$Buffer
    )

    $expected = $Buffer.Columns.Count
    $all = New-Object System.Collections.Generic.List[string]
    $all.Add($Buffer.Header) | Out-Null
    foreach ($r in $Buffer.Rows) { $all.Add($r) | Out-Null }

    for ($i = 0; $i -lt $all.Count; $i++) {
        $line = [string]$all[$i]
        $actual = ([regex]::Matches($line, ';')).Count + 1
        if ($actual -ne $expected) {
            throw ("Linha inválida no CSV {0} (linha {1}): esperado={2}, atual={3}" -f $Buffer.Path, ($i + 1), $expected, $actual)
        }
    }
    return $true
}

function Flush-CsvBuffer {
    param(
        [Parameter(Mandatory=$true)]$Buffer,
        [Parameter(Mandatory=$false)]$Context
    )

    [void](Validate-CsvRowCount -Buffer $Buffer)

    $allLines = New-Object System.Collections.Generic.List[string]
    $allLines.Add($Buffer.Header) | Out-Null
    foreach ($line in $Buffer.Rows) { $allLines.Add($line) | Out-Null }

    Write-AllLinesUtf8 -Path $Buffer.Path -Lines $allLines.ToArray()
    if ($Context) {
        Write-Log ("CSV validado e gravado ({0} linhas): {1}" -f $allLines.Count, $Buffer.Path) "OK" $Context
    }
}

function Get-StepItemCount {
    param($Value)

    if ($null -eq $Value) { return 0 }

    if ($Value -is [System.Collections.IDictionary]) { return $Value.Count }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        $count = 0
        foreach ($item in $Value) { $count++ }
        return $count
    }

    if ($Value.PSObject -and ($Value.PSObject.Properties.Name -contains "Count")) {
        try { return [int]$Value.Count } catch {}
    }

    return 1
}

function Validate-InventoryResult {
    param(
        [Parameter(Mandatory=$true)]$Result,
        [Parameter(Mandatory=$false)]$Context
    )

    $safeUsers = if ($Result.users) { $Result.users } else { [ordered]@{} }
    $activeUsers = @()
    $sessions = @()
    $profiles = @()

    if ($safeUsers -is [System.Collections.IDictionary]) {
        if ($safeUsers.Contains("active_users")) { $activeUsers = @($safeUsers["active_users"]) }
        if ($safeUsers.Contains("sessions")) { $sessions = @($safeUsers["sessions"]) }
        if ($safeUsers.Contains("profiles")) { $profiles = @($safeUsers["profiles"]) }
    } else {
        if ($safeUsers.PSObject.Properties["active_users"]) { $activeUsers = @($safeUsers.active_users) }
        if ($safeUsers.PSObject.Properties["sessions"]) { $sessions = @($safeUsers.sessions) }
        if ($safeUsers.PSObject.Properties["profiles"]) { $profiles = @($safeUsers.profiles) }
    }

    $validated = [ordered]@{
        meta     = if ($Result.meta) { $Result.meta } else { [ordered]@{} }
        os       = if ($Result.os) { $Result.os } else { [ordered]@{} }
        users    = [ordered]@{
            active_users = @($activeUsers | Where-Object { $_ -and $_.user } | Sort-Object user)
            sessions     = @($sessions | Where-Object { $_ -and $_.user } | Sort-Object user, logon_id)
            profiles     = @($profiles | Where-Object { $_ } | Sort-Object local_path)
        }
        hardware = if ($Result.hardware) { $Result.hardware } else { [ordered]@{} }
        software = @($Result.software | Where-Object { $_ } | Sort-Object name, version, publisher)
        devices  = @($Result.devices | Where-Object { $_ } | Sort-Object pnp_class, name)
    }

    if ($Context) {
        Write-Log ("Validação de payload concluída: software={0}, sessions={1}, profiles={2}, devices={3}" -f `
            $validated.software.Count, $validated.users.sessions.Count, $validated.users.profiles.Count, $validated.devices.Count) "INFO" $Context
    }

    return $validated
}

function Invoke-InventoryStep {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][scriptblock]$Script,
        [Parameter(Mandatory=$true)]$Context,
        $Fallback
    )

    $started = Get-Date
    $status = "ok"
    $result = $null

    try {
        Write-Log ("Coletando {0}..." -f $Name) "INFO" $Context
        $result = & $Script
    } catch {
        $status = "fallback"
        Write-Log ("Falha ao coletar {0}: {1}" -f $Name, $_.Exception.Message) "WARN" $Context
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
            $pos = ($_.InvocationInfo.PositionMessage -replace '\r?\n', ' ').Trim()
            if ($pos) { Write-Log ("Local: " + $pos) "WARN" $Context }
        }
        $result = $Fallback
    }

    if ($status -eq "ok") {
        if (($result -is [System.Collections.IDictionary]) -and $result.Contains("error")) {
            $status = "warn"
        } elseif ($result -and $result.PSObject -and $result.PSObject.Properties["error"]) {
            $status = "warn"
        }
    }

    $duration = [int]([math]::Round(((Get-Date) - $started).TotalMilliseconds, 0))
    $itemCount = Get-StepItemCount $result

    if ($Context -and $Context.PSObject.Properties["StepMetrics"] -and $null -ne $Context.StepMetrics) {
        $Context.StepMetrics.Add([pscustomobject]@{
            step        = $Name
            status      = $status
            duration_ms = $duration
            item_count  = $itemCount
            captured_at = (Get-Date).ToString("o")
        }) | Out-Null
    }

    Write-Log ("Etapa '{0}' finalizada: status={1}; duracao_ms={2}; itens={3}" -f $Name, $status, $duration, $itemCount) "INFO" $Context
    return $result
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
        StepMetrics = (New-Object System.Collections.Generic.List[object])
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

    $dedupe = @{}
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

                $name = Normalize-Text (& $toStr $a.DisplayName)
                if (-not $name) { continue }

                $version = Normalize-Text (& $toStr $a.DisplayVersion)
                $publisher = Normalize-Text (& $toStr $a.Publisher)
                $registryKey = Normalize-Text (& $toStr $a.PSChildName)
                $uninstall = Normalize-Text (& $toStr $a.UninstallString)

                $key = ("{0}|{1}|{2}" -f $name.ToLowerInvariant(), $version.ToLowerInvariant(), $publisher.ToLowerInvariant())
                if (-not $dedupe.ContainsKey($key)) {
                    $dedupe[$key] = [pscustomobject]@{
                    name         = $name
                    version      = $version
                    publisher    = $publisher
                    install_date = $dt
                    registry_key = $registryKey
                    uninstall    = $uninstall
                    source_hive  = $sourceHive
                    }
                }
            }
        } catch {
            Write-Log ("Falha ao coletar softwares em {0}: {1}" -f $p, $_.Exception.Message) "WARN" $Context
        }
    }

    return @(
        $dedupe.GetEnumerator() |
        ForEach-Object { $_.Value } |
        Sort-Object name, version, publisher
    )
}

function Get-UsersInfo {
    param([Parameter(Mandatory=$true)]$Context)

    $active = New-Object System.Collections.Generic.List[object]
    $sessions = New-Object System.Collections.Generic.List[object]
    $profiles = New-Object System.Collections.Generic.List[object]
    $logonUsersById = @{}

    # Sessões via CIM (preferência)
    try {
        $logons = Get-CimInstance Win32_LogonSession -Property LogonId, LogonType, StartTime
        $links  = Get-CimInstance Win32_LoggedOnUser -Property Antecedent, Dependent

        foreach ($link in $links) {
            $depMatch = [regex]::Match([string]$link.Dependent, 'LogonId=`"(?<id>[^`"]+)`"')
            if (-not $depMatch.Success) { continue }

            $userMatch = [regex]::Match([string]$link.Antecedent, 'Domain=`"(?<d>[^`"]+)`".*Name=`"(?<n>[^`"]+)`"')
            if (-not $userMatch.Success) { continue }

            $id = $depMatch.Groups['id'].Value
            $user = Normalize-Text ("{0}\{1}" -f $userMatch.Groups['d'].Value, $userMatch.Groups['n'].Value)
            if (-not $user) { continue }

            if (-not $logonUsersById.ContainsKey($id)) {
                $logonUsersById[$id] = (New-Object System.Collections.Generic.List[string])
            }
            if (-not $logonUsersById[$id].Contains($user)) {
                $logonUsersById[$id].Add($user) | Out-Null
            }
        }

        $interesting = $logons | Where-Object { $_.LogonType -in 2,10 }  # 2=Interactive, 10=RemoteInteractive

        foreach ($ls in $interesting) {
            $logonId = [string]$ls.LogonId
            if ($logonUsersById.ContainsKey($logonId)) {
                foreach ($user in $logonUsersById[$logonId]) {
                    $sessions.Add([pscustomobject]@{
                        user       = $user
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
        $cs = Get-CimInstance Win32_ComputerSystem -Property UserName
        if ($cs.UserName) {
            $active.Add([pscustomobject]@{ user = (Normalize-Text $cs.UserName); source = "Win32_ComputerSystem" }) | Out-Null
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
                        $u = Normalize-Text (($parts[0]).Trim().TrimStart('>'))
                        if ($u) {
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
        $ups = Get-CimInstance Win32_UserProfile -Property LocalPath, SID, LastUseTime, Loaded, Special -ErrorAction Stop | Where-Object {
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
                local_path        = Normalize-Text $p.LocalPath
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
                        local_path        = Normalize-Text $_.FullName
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
        $uniq = $sessions | Where-Object { $_.user } | Group-Object user | ForEach-Object { $_.Group | Select-Object -First 1 }
        foreach ($s in $uniq) {
            $active.Add([pscustomobject]@{ user = $s.user; source = "sessions" }) | Out-Null
        }
    }

    $activeDedup = @(
        $active |
        Where-Object { $_ -and $_.user } |
        Group-Object user |
        ForEach-Object { $_.Group | Select-Object -First 1 }
    )

    $sessionArray = @($sessions.ToArray() | Where-Object { $_ -and $_.user })
    $profileArray = @($profiles.ToArray() | Sort-Object local_path)

    return [ordered]@{
        active_users = @($activeDedup | Sort-Object user)
        sessions     = @($sessionArray | Sort-Object user, logon_id, start_time)
        profiles     = $profileArray
    }
}

function Get-HardwareInfo {
    param([Parameter(Mandatory=$true)]$Context)

    try {
        $sys  = Get-CimInstance Win32_ComputerSystem -Property Manufacturer, Model, TotalPhysicalMemory
        $bios = Get-CimInstance Win32_BIOS -Property SerialNumber
        $cpu  = Get-CimInstance Win32_Processor -Property Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors | Select-Object Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors
        $video = Get-CimInstance Win32_VideoController -Property Name, VideoProcessor, AdapterRAM | Select-Object Name, VideoProcessor, AdapterRAM
        $mem  = Get-CimInstance Win32_PhysicalMemory -Property DeviceLocator, Capacity, ConfiguredClockSpeed, SMBIOSMemoryType, FormFactor | Select-Object DeviceLocator, Capacity, ConfiguredClockSpeed, SMBIOSMemoryType, FormFactor
        $net  = Get-NetAdapter -ErrorAction SilentlyContinue | Select-Object Name, InterfaceDescription, MacAddress, Status, LinkSpeed
        $disk = Get-CimInstance Win32_DiskDrive -Property Model, InterfaceType, MediaType, Size, SerialNumber | Select-Object Model, InterfaceType, MediaType, Size, SerialNumber
        $enclosure = Get-CimInstance Win32_SystemEnclosure -Property ChassisTypes -ErrorAction SilentlyContinue

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
            manufacturer       = Normalize-Text $sys.Manufacturer
            model              = Normalize-Text $sys.Model
            system_type        = $computerType
            serial_bios        = Normalize-Text $bios.SerialNumber
            total_memory_gb    = [math]::Round($sys.TotalPhysicalMemory / 1GB, 2)
            memory_total_gb    = $memTotalGB
            memory_type        = $memoryType
            memory_form_factor = $memoryFormat
            cpu                = $cpu
            video              = ($video | ForEach-Object {
                                    [pscustomobject]@{
                                        name      = Normalize-Text $_.Name
                                        processor = Normalize-Text $_.VideoProcessor
                                        vram_gb   = if ($_.AdapterRAM) { [math]::Round($_.AdapterRAM / 1GB) } else { $null }
                                    }
                                  })
            memory_modules     = ($mem | ForEach-Object {
                                    [pscustomobject]@{
                                        locator     = Normalize-Text $_.DeviceLocator
                                        capacity_gb = if ($_.Capacity) { [math]::Round($_.Capacity / 1GB) } else { $null }
                                        clock_mhz   = $_.ConfiguredClockSpeed
                                    }
                                  })
            network_adapters   = ($net | ForEach-Object {
                                    [pscustomobject]@{
                                        name        = Normalize-Text $_.Name
                                        description = Normalize-Text $_.InterfaceDescription
                                        mac         = Normalize-Text $_.MacAddress
                                        status      = Normalize-Text $_.Status
                                        link_speed  = Normalize-Text $_.LinkSpeed
                                    }
                                  })
            disks              = ($disk | ForEach-Object {
                                    [pscustomobject]@{
                                        model     = Normalize-Text $_.Model
                                        interface = Normalize-Text $_.InterfaceType
                                        media_type= Normalize-Text $_.MediaType
                                        size_gb   = if ($_.Size) { [math]::Round($_.Size / 1GB) } else { $null }
                                        serial    = Normalize-Text $_.SerialNumber
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
        Get-CimInstance Win32_PnPEntity -Property PNPClass, Name, Status, DeviceID |
            Where-Object { $_.Name } |
            ForEach-Object {
            $list.Add([pscustomobject]@{
                pnp_class = Normalize-Text $_.PNPClass
                name      = Normalize-Text $_.Name
                status    = Normalize-Text $_.Status
                device_id = Normalize-Text $_.DeviceID
            }) | Out-Null
        }
    } catch {
        Write-Log ("Falha ao coletar dispositivos: " + $_.Exception.Message) "WARN" $Context
    }
    return @($list.ToArray() | Sort-Object pnp_class, name)
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
$invVersion = "2.1.0"
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

    $meta = Invoke-InventoryStep -Name "metadados" -Script { Get-MetaInfo -Context $ctx } -Context $ctx `
        -Fallback ([ordered]@{ error = "Falha ao coletar metadados." })
    $os = Invoke-InventoryStep -Name "sistema operacional" -Script { Get-OSInfo -Context $ctx } -Context $ctx `
        -Fallback ([ordered]@{ error = "Falha ao coletar sistema operacional." })
    $users = Invoke-InventoryStep -Name "usuarios" -Script { Get-UsersInfo -Context $ctx } -Context $ctx `
        -Fallback ([ordered]@{ active_users = @(); sessions = @(); profiles = @(); error = "Falha ao coletar usuarios." })
    $hardware = Invoke-InventoryStep -Name "hardware" -Script { Get-HardwareInfo -Context $ctx } -Context $ctx `
        -Fallback ([ordered]@{ cpu=@(); network_adapters=@(); video=@(); memory_modules=@(); disks=@(); error="Falha ao coletar hardware." })
    $software = Invoke-InventoryStep -Name "softwares" -Script { Get-SoftwareInventory -Context $ctx } -Context $ctx `
        -Fallback @()
    $devices = Invoke-InventoryStep -Name "dispositivos" -Script { Get-DevicesInventory -Context $ctx } -Context $ctx `
        -Fallback @()

    $result = [ordered]@{
        meta     = $meta
        os       = $os
        users    = $users
        hardware = $hardware
        software = $software
        devices  = $devices
    }
    $result = Validate-InventoryResult -Result $result -Context $ctx

    # Arquivos de saída (mesmos do invent_asi.ps1)
    $arquivoConsolidado = Join-Path $ctx.MachinePath "$($ctx.MachineName).csv"
    $arquivoSoftwares   = Join-Path $ctx.MachinePath "$($ctx.MachineName)_SW.csv"
    $arquivoSO          = Join-Path $ctx.MachinePath "$($ctx.MachineName)_SO.csv"
    $arquivoUsuarios    = Join-Path $ctx.MachinePath "$($ctx.MachineName)_USERS.csv"
    $arquivoHardware    = Join-Path $ctx.MachinePath "$($ctx.MachineName)_HW.csv"
    $arquivoDispositivos= Join-Path $ctx.MachinePath "$($ctx.MachineName)_DEVICES.csv"

    $schemas = New-CsvSchemaMap
    $buffers = [ordered]@{
        software = New-CsvBuffer -Path $arquivoSoftwares -Columns $schemas.software
        os       = New-CsvBuffer -Path $arquivoSO -Columns $schemas.os
        users    = New-CsvBuffer -Path $arquivoUsuarios -Columns $schemas.users
        hardware = New-CsvBuffer -Path $arquivoHardware -Columns $schemas.hardware
        devices  = New-CsvBuffer -Path $arquivoDispositivos -Columns $schemas.devices
    }

    # SOFTWARES
    foreach ($app in $result.software) {
        $dataInst = Format-DatePtBr $app.install_date
        Add-CsvRowSafe -Buffer $buffers.software -Row @{
            NOME_DA_MAQUINA = $ctx.MachineName
            NOME_DO_SOFTWARE = $app.name
            VERSAO = $app.version
            FABRICANTE = $app.publisher
            DATA_INSTALACAO = $dataInst
            CHAVE = $app.registry_key
        }
    }
    Flush-CsvBuffer -Buffer $buffers.software -Context $ctx
    Write-Log "CSV de softwares gerado: $arquivoSoftwares" "OK" $ctx

    # SISTEMA OPERACIONAL
    $os = $result.os
    $dataSO = Format-DatePtBr $os.install_date
    Add-CsvRowSafe -Buffer $buffers.os -Row @{
        NOME_DA_MAQUINA = $ctx.MachineName
        WINDOWS_FAMILY = $os.windows_family
        NOME_SO = $os.caption
        DISPLAY_VERSION = $os.display_version
        VERSION = $os.version
        BUILD_COMPLETO = $os.build
        ARQUITETURA = $os.architecture
        EDITION_ID = $os.edition_id
        INSTALLATION_TYPE = $os.installation_type
        DATA_INSTALACAO = $dataSO
        PRODUCT_ID = $os.product_id
    }
    Flush-CsvBuffer -Buffer $buffers.os -Context $ctx
    Write-Log "CSV de sistema operacional gerado: $arquivoSO" "OK" $ctx

    # USUARIOS
    $activeShort = @($result.users.active_users | ForEach-Object { Normalize-User $_.user })
    $sessionShort = @($result.users.sessions | ForEach-Object { Normalize-User $_.user })
    $profiles = @($result.users.profiles)

    $activeMap = @{}
    foreach ($u in $activeShort) {
        if ($u) { $activeMap[$u.ToLowerInvariant()] = $true }
    }
    $sessionMap = @{}
    foreach ($u in $sessionShort) {
        if ($u) { $sessionMap[$u.ToLowerInvariant()] = $true }
    }

    if ($profiles.Count -gt 0) {
        foreach ($p in $profiles) {
            $userFolder = if ($p.local_path) { Normalize-Text (Split-Path -Leaf $p.local_path) } else { "" }
            $userKey = if ($userFolder) { $userFolder.ToLowerInvariant() } else { "" }
            $ativo = if ($userKey -and $activeMap.ContainsKey($userKey)) { "Sim" } else { "Nao" }
            $sessao = "Desconhecida"
            if ($sessionMap.Count -gt 0) {
                $sessao = if ($userKey -and $sessionMap.ContainsKey($userKey)) { "Interativa" } else { "Nao listada" }
            }
            $dataPerfil = Format-DateTimePtBr $p.last_use_time
            $dataNtUser = Format-DateTimePtBr $p.ntuser_last_write
            Add-CsvRowSafe -Buffer $buffers.users -Row @{
                NOME_DA_MAQUINA = $ctx.MachineName
                USUARIO = $userFolder
                ATIVO = $ativo
                SESSAO = $sessao
                PERFIL = $p.local_path
                DATA_ALTERACAO_PERFIL = $dataPerfil
                DATA_ALTERACAO_NTUSER = $dataNtUser
            }
        }
    } else {
        # Se não houver perfis, ainda registra usuários ativos (se existirem)
        $activeUnique = New-Object System.Collections.Generic.List[string]
        $activeSeen = @{}
        foreach ($u in $activeShort) {
            if (-not $u) { continue }
            $k = $u.ToLowerInvariant()
            if (-not $activeSeen.ContainsKey($k)) {
                $activeSeen[$k] = $true
                $activeUnique.Add($u) | Out-Null
            }
        }
        foreach ($u in $activeUnique) {
            Add-CsvRowSafe -Buffer $buffers.users -Row @{
                NOME_DA_MAQUINA = $ctx.MachineName
                USUARIO = $u
                ATIVO = "Sim"
                SESSAO = "Interativa"
                PERFIL = ""
                DATA_ALTERACAO_PERFIL = ""
                DATA_ALTERACAO_NTUSER = ""
            }
        }
    }
    Flush-CsvBuffer -Buffer $buffers.users -Context $ctx
    Write-Log "CSV de usuarios gerado: $arquivoUsuarios" "OK" $ctx

    # HARDWARE
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

    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Modelo"; VALOR = $hw.model }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Memoria_GB"; VALOR = $hw.total_memory_gb }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Tipo"; VALOR = $hw.system_type }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Fabricante"; VALOR = $hw.manufacturer }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Serial"; VALOR = $hw.serial_bios }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "CPU"; VALOR = $cpuNames }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "NetworkAdapter"; VALOR = $netDesc }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MacAddress"; VALOR = $netMac }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "VideoBoard"; VALOR = $videoName }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "VideoProcessor"; VALOR = $videoProc }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "VideoCapacity_GB"; VALOR = $videoVram }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryType"; VALOR = $hw.memory_type }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryFormat"; VALOR = $hw.memory_form_factor }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryFrequency_MHz"; VALOR = $memClk }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryModule"; VALOR = $memLoc }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryCapacity_GB"; VALOR = $memCap }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "MemoryTotal_GB"; VALOR = $hw.memory_total_gb }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "DiskType"; VALOR = $diskType }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "DiskModel"; VALOR = $diskModel }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "DiskSize_GB"; VALOR = $diskSize }
    Add-CsvRowSafe -Buffer $buffers.hardware -Row @{ NOME_DA_MAQUINA = $ctx.MachineName; COMPONENTE = "Monitor"; VALOR = "" }
    Flush-CsvBuffer -Buffer $buffers.hardware -Context $ctx
    Write-Log "CSV de hardware gerado: $arquivoHardware" "OK" $ctx

    # DISPOSITIVOS
    foreach ($d in $result.devices) {
        Add-CsvRowSafe -Buffer $buffers.devices -Row @{
            NOME_DA_MAQUINA = $ctx.MachineName
            CLASSE = $d.pnp_class
            DISPOSITIVO = $d.name
            STATUS = $d.status
        }
    }
    Flush-CsvBuffer -Buffer $buffers.devices -Context $ctx
    Write-Log "CSV de dispositivos gerado: $arquivoDispositivos" "OK" $ctx

    # CSV consolidado
    $arquivosOrigem = @(
        $arquivoSoftwares,
        $arquivoHardware,
        $arquivoSO,
        $arquivoUsuarios,
        $arquivoDispositivos
    )
    $linhasConsolidadas = New-Object System.Collections.Generic.List[string]
    foreach ($arquivo in $arquivosOrigem) {
        if (Test-Path -LiteralPath $arquivo) {
            $lines = [System.IO.File]::ReadAllLines($arquivo)
            foreach ($line in $lines) { $linhasConsolidadas.Add($line) | Out-Null }
        } else {
            Write-Log "Arquivo nao encontrado para concatenacao: $arquivo" "WARN" $ctx
        }
    }
    Write-AllLinesUtf8 -Path $arquivoConsolidado -Lines $linhasConsolidadas.ToArray()
    Write-Log "CSV consolidado gerado: $arquivoConsolidado" "OK" $ctx

    if ($ctx.StepMetrics -and $ctx.StepMetrics.Count -gt 0) {
        foreach ($m in $ctx.StepMetrics) {
            Write-Log ("Metrica: etapa={0}; status={1}; duracao_ms={2}; itens={3}" -f $m.step, $m.status, $m.duration_ms, $m.item_count) "INFO" $ctx
        }
    }

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
