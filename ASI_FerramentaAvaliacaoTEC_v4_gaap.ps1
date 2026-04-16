clear
Write-Host "========== MMP - GAAP ==========" -ForegroundColor Cyan
Write-Host "======= DIAGNOSTICO AVANCADO WORKSTATION ======" -ForegroundColor Cyan
Write-Host "======= Aguarde processo, pode demorar ========" -ForegroundColor Yellow

function Step-Log {
    param(
        [string]$Mensagem,
        [string]$Cor = "Green"
    )
    $hora = Get-Date -Format "HH:mm:ss"
    Write-Host "[$hora] $Mensagem" -ForegroundColor $Cor
}

Step-Log "Iniciando processo..."

# =========================
# CONFIGURACOES
# =========================

$reportFolder   = "C:\TI\Relatorios_TI"
Step-Log "Pasta de relatórios: $reportFolder"
$basePath       = "C:\TI\Relatorios_TI"

$redeDestino    = "\\127.0.0.1\c$"   # AJUSTAR PARA SHARE REAL
Step-Log "Pasta de rede: [$redeDestino]" 
Step-Log "Pasta de rede: * Caso uso apenas local altere para \\127.0.0.1\c$"

$smtpServer     = "smtp.office365.com"
$smtpPort       = 587
$emailFrom      = "suporte@gaaplocacao.onmicrosoft.com"
$emailTo        = "suporte@gaaplocacao.onmicrosoft.com"
$emailPassword  = 'Arr@y$ervice'

$systemDrive = $env:SystemDrive;    


# =========================
# CRIAR PASTAS
# =========================

Step-Log "Validando pastas de trabalho..."

if (!(Test-Path $reportFolder)) { New-Item -ItemType Directory $reportFolder -Force | Out-Null }
if (!(Test-Path $basePath))     { New-Item -ItemType Directory $basePath -Force | Out-Null }

# =========================
# IDENTIDADE DO ARQUIVO
# =========================

$hostname = $env:COMPUTERNAME
$usuario  = $env:USERNAME
$dominio  = $env:USERDOMAIN

# tentativa de pegar e-mail (AD / Azure)
try {
    $email = (whoami /upn) 2>$null
    if (-not $email) { $email = $usuario }
} catch {
    $email = $usuario
}

$data = Get-Date -Format "yyyyMMdd_HHmmss"

# PREFIXO PADRAO GLOBAL
$prefixo = "$dominio_$hostname`_$usuario`_$data"

# $report = Join-Path $reportFolder ("Relatorio_PC_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$report = Join-Path $reportFolder ("${prefixo}_relatorio.html")

# =========================
# LISTAS DE APOIO
# =========================

$processosExecutados = New-Object System.Collections.Generic.List[string]
$motivos             = New-Object System.Collections.Generic.List[string]
$recomendacoes       = New-Object System.Collections.Generic.List[string]

# =========================
# INFORMACOES DO SISTEMA
# =========================

Step-Log "Coletando informacoes do sistema..."

$hostname = $env:COMPUTERNAME
$usuario  = $env:USERNAME
$dominio  = $env:USERDOMAIN

$cs   = Get-CimInstance Win32_ComputerSystem
$cpu  = Get-CimInstance Win32_Processor
$ram  = Get-CimInstance Win32_PhysicalMemory
$os   = Get-CimInstance Win32_OperatingSystem
$gpu  = Get-CimInstance Win32_VideoController
$bios = Get-CimInstance Win32_BIOS
$mb   = Get-CimInstance Win32_BaseBoard
$gpuName = ($gpu | Select-Object -First 1).Name
$cpuName = ($cpu | Select-Object -First 1).Name

$fabricante = $cs.Manufacturer
$modeloPC   = $cs.Model
$serial     = $bios.SerialNumber
$biosVer    = $bios.SMBIOSBIOSVersion
$placaMae   = $mb.Product

$ramTotal   = [math]::Round((($ram | Measure-Object Capacity -Sum).Sum / 1GB), 2)
# $ramLivre   = [math]::Round(($os.FreePhysicalMemory / 1MB), 2)
#$ramLivre = [math]::Round(($os.FreePhysicalMemory / 1KB / 1024),2)
$ramLivre = [math]::Round(($os.FreePhysicalMemory / 1KB / 1024),2)
$ramUsada   = [math]::Round(($ramTotal - $ramLivre), 2)

if ($ramTotal -gt 0) {
    $ramPercent = [math]::Round((($ramUsada / $ramTotal) * 100), 1)
} else {
    $ramPercent = 0
}

$cpuLoad = [math]::Round(((Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average), 1)
$uptime  = (Get-Date) - $os.LastBootUpTime

$processosExecutados.Add("Coleta de inventario de sistema e hardware")

Step-Log "Desabilitando desfragmentacao automatica..." "Yellow"

try {
    Disable-ScheduledTask -TaskName "ScheduledDefrag" -TaskPath "\Microsoft\Windows\Defrag\" -ErrorAction Stop
    $processosExecutados.Add("Desativacao da desfragmentacao automatica")
}
catch {
    Step-Log "Falha ao desabilitar defrag: $_" "Red"
    $motivos.Add("Falha ao desativar desfragmentacao automatica")
}

# =========================
# DATA INSTALACAO SO
# =========================
try {
    $installDate = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").InstallDate
    $installDate = (Get-Date "1970-01-01").AddSeconds($installDate)
    $installDateFormat = $installDate.ToString("dd/MM/yyyy HH:mm")
}
catch {
    $installDateFormat = "Nao identificado"
}
# =========================
# DATA FORMATACAO DISCO C
# =========================
try {

    $formatDateFormat = "Nao identificado"

    # 🔹 Método 1 - Win32_Volume
    $volume = Get-CimInstance Win32_Volume -Filter "DriveLetter = 'C:'"

    if ($volume -and $volume.CreationDate) {

        $formatDate = [Management.ManagementDateTimeConverter]::ToDateTime($volume.CreationDate)
        $formatDateFormat = $formatDate.ToString("dd/MM/yyyy HH:mm")

    }

    # 🔹 Método 2 - fallback via instalação (se igual, assume mesma data)
    if ($formatDateFormat -eq "Nao identificado" -and $installDate) {

        $formatDateFormat = $installDate.ToString("dd/MM/yyyy HH:mm") + " (estimado)"

    }

}
catch {

    $formatDateFormat = "Nao identificado"

}
# =========================
# LIMPEZA WINDOWS
# =========================

Step-Log "Limpando temporarios do Windows..."

Get-ChildItem $env:TEMP -Recurse -ErrorAction SilentlyContinue |
    Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

Get-ChildItem "C:\Windows\Temp" -Recurse -ErrorAction SilentlyContinue |
    Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

$processosExecutados.Add("Limpeza de temporarios do Windows")

# =========================
# LIMPEZA NAVEGADORES
# =========================

Step-Log "Fechando e limpando cache de navegadores..."

Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$chromeCache = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
$edgeCache   = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"

if (Test-Path $chromeCache) {
    Remove-Item "$chromeCache\*" -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path $edgeCache) {
    Remove-Item "$edgeCache\*" -Recurse -Force -ErrorAction SilentlyContinue
}

$processosExecutados.Add("Limpeza de cache Chrome/Edge")

# =========================
# LIMPEZA WHATSAPP
# =========================

Step-Log "Validando cache do WhatsApp Desktop..."

$waCache = "$env:APPDATA\WhatsApp"
if (Test-Path $waCache) {
    Remove-Item "$waCache\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$waCache\Code Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
    $processosExecutados.Add("Limpeza de cache WhatsApp Desktop")
}

# =========================
# WINDOWS VERSION
# =========================

Step-Log "Coletando versao do Windows..."

$osInfo = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

$windowsBuild   = [int]$osInfo.CurrentBuild
$windowsRelease = $osInfo.DisplayVersion
$windowsEdition = $osInfo.EditionID

if ($windowsBuild -ge 22000) {
    $windowsVersion = "Windows 11 $windowsEdition"
} else {
    $windowsVersion = "Windows 10 $windowsEdition"
}

$processosExecutados.Add("Coleta da versao do Windows")



# =========================
# WINDOWS UPDATE
# =========================
Step-Log "Consultando atualizacoes pendentes do Windows..."
$updatesHTML  = ""
$updatesCount = 0

try {
    $updateSession  = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $pendingUpdates = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
    $updatesCount   = $pendingUpdates.Updates.Count

    foreach ($update in $pendingUpdates.Updates) {
        $updatesHTML += "<li>$($update.Title)</li>"
    }

    if ($updatesCount -gt 0) {
        $motivos.Add("$updatesCount atualizacoes pendentes do Windows")
        $recomendacoes.Add("Executar Windows Update para aplicar atualizacoes pendentes")
    }

    $processosExecutados.Add("Consulta de Windows Update")
}
catch {
    $updatesHTML = "<li>Falha ao consultar Windows Update</li>"
    $motivos.Add("Nao foi possivel validar Windows Update")
}

# =========================
# LIMPEZA ChatGPT
# =========================
Step-Log "Limpeza de cache ChatGPT..."

Stop-Process -Name ChatGPT -Force -ErrorAction SilentlyContinue

Remove-Item "$env:APPDATA\ChatGPT\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\ChatGPT\Code Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\ChatGPT\GPUCache\*" -Recurse -Force -ErrorAction SilentlyContinue

Remove-Item "$env:LOCALAPPDATA\ChatGPT\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\ChatGPT\GPUCache\*" -Recurse -Force -ErrorAction SilentlyContinue

# =========================
# LIMPEZA WINDOWS UPDATE
# =========================
Step-Log "Limpando cache Windows Update..."

try {

$service = Get-Service wuauserv -ErrorAction Stop

if ($service.Status -eq "Running") {

    Stop-Service wuauserv -Force -ErrorAction Stop
    Start-Sleep -Seconds 2

}

Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue

Start-Service wuauserv -ErrorAction Stop

$processosExecutados.Add("Limpeza de cache Windows Update")

}
catch {

Step-Log "Falha limpeza Windows Update: $_" "Yellow"
$motivos.Add("Falha ao limpar cache Windows Update")

}

# =========================
# WINDOWS DEFENDER
# =========================

Step-Log "Consultando status do Windows Defender..."
$defenderStatus = "Desconhecido"

try {
    $defender = Get-MpComputerStatus
    if ($defender.AntivirusEnabled) {
        $defenderStatus = "Ativo"
    } else {
        $defenderStatus = "Desativado"
        $motivos.Add("Windows Defender desativado")
        $recomendacoes.Add("Ativar o Windows Defender")
    }
    $processosExecutados.Add("Consulta do status do Windows Defender")
}
catch {
    $defenderStatus = "Nao foi possivel consultar"
    $motivos.Add("Falha na consulta do Windows Defender")
}


# =========================
# BITLOCKER STATUS
# =========================
Step-Log "Coletando status do BitLocker..."
$bitlockerFiles = @()

$bitlockerFolder = Join-Path $reportFolder "BitLocker_Keys"

if (!(Test-Path $bitlockerFolder)) {
    New-Item -ItemType Directory $bitlockerFolder -Force | Out-Null
}


$bitlockerHTML = ""

# 🔐 CONTROLE DE SEGURANÇA
# $exportarBitlocker = $false
# $modoSeguro = $true

$exportarBitlocker = $true
$modoSeguro = $false   # cuidado: isso libera export real

try {
    $volumes = Get-BitLockerVolume -ErrorAction Stop
}
catch {
    Step-Log "BitLocker não disponível nesta máquina" "Yellow"
    $volumes = @()
}

foreach ($v in $volumes) {

    if (-not $v) { continue }

    $protection = if ($v.ProtectionStatus -eq 1) {
        "<span style='color:green'>Ativo</span>"
    } else {
        "<span style='color:red'>Desativado</span>"
    }

    $encryption = $v.EncryptionPercentage
    $keyIdShort = "-"
    $hasRecovery = $false

    foreach ($kp in $v.KeyProtector) {

        if ($kp.KeyProtectorType -eq "RecoveryPassword") {

            $hasRecovery = $true

            $keyId = $kp.KeyProtectorId
            $keyIdShort = "****" + $keyId.Substring($keyId.Length - 6)


            # 🔐 EXPORT CONTROLADO
            if ($exportarBitlocker -and -not $modoSeguro) {
				$processosExecutados.Add("Exportacao BitLocker: $($v.MountPoint)")
                $data = Get-Date -Format "yyyyMMdd_HHmmss"
                # $file = "$reportFolder\BitLocker_$($v.MountPoint.Trim(':'))_$data.txt"

				$file = Join-Path $bitlockerFolder "BL_$hostname`_$($v.MountPoint.Trim(':'))_$data.txt"

@"
Chave de recuperação de Criptografia de Unidade de Disco BitLocker

Para verificar se esta é a chave de recuperação correta, compare o início do identificador a seguir com o valor do identificador exibido no computador.

Identificador:

    $keyId

Se o identificador acima corresponder ao que é exibido no computador, use a chave a seguir para desbloquear a unidade.

Chave de Recuperação:

    $($kp.RecoveryPassword)

Se o identificador acima não corresponder ao que é exibido no computador, significa que esta não é a chave correta para desbloquear a unidade.
Tente usar outra chave de recuperação ou consulte https://go.microsoft.com/fwlink/?LinkID=260589 para obter assistência.
"@ | Out-File $file -Encoding utf8

                Step-Log "RecoveryKey exportada: $($v.MountPoint)" "Yellow"
				$bitlockerFiles += $file
            }
        }
    }

if (-not $hasRecovery) {
    $motivos.Add("CRITICO: BitLocker sem chave de recuperação na unidade de sistema")
}

    if (-not $hasRecovery) {
        $keyIdShort = "<span style='color:red'>Sem RecoveryKey</span>"
    }

    if ($v.ProtectionStatus -ne 1) {
        $motivos.Add("BitLocker desativado na unidade $($v.MountPoint)")
        $recomendacoes.Add("Ativar BitLocker na unidade $($v.MountPoint)")
    }

    $bitlockerHTML += @"
<tr>
<td>$($v.MountPoint)</td>
<td>$protection</td>
<td>$encryption %</td>
<td>$($v.VolumeStatus)</td>
<td>$keyIdShort</td>
<td>$(if($hasRecovery){"<span style='color:green'>OK</span>"}else{"<span style='color:red'>Sem chave</span>"})</td>
</tr>
"@
}

$bitlockerExportStatus = if ($exportarBitlocker -and -not $modoSeguro) {
    "Exportacao habilitada"
} else {
    "Modo seguro (sem exportacao)"
}
# =========================
# registro Azure AD
# =========================

Step-Log "Coletando status de registro Azure AD..."

$dsregRaw = dsregcmd /status | Out-String

function Get-DSValue {
    param($text, $key)

    if ($text -match "$key\s*:\s*(.+)") {
        return $matches[1].Trim()
    }
    return "N/A"
}

$azureJoined   = Get-DSValue $dsregRaw "AzureAdJoined"
$domainJoined  = Get-DSValue $dsregRaw "DomainJoined"
$deviceId      = Get-DSValue $dsregRaw "DeviceId"
$tenantId      = Get-DSValue $dsregRaw "TenantId"
$tenantName    = Get-DSValue $dsregRaw "TenantName"
$mdmUrl        = Get-DSValue $dsregRaw "MdmUrl"

$processosExecutados.Add("Coleta de status Azure AD (dsregcmd)")

if ($azureJoined -ne "YES") {
    $motivos.Add("Dispositivo NÃO integrado ao Azure AD")
    $recomendacoes.Add("Ingressar dispositivo no Azure AD / Entra ID")
}

# =========================
# ESPACO EM DISCO
# =========================

Step-Log "Analisando espaco em disco..."

$systemDriveFree = 0
$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
$diskTableHTML = ""

foreach ($d in $disks) {

    $total = [math]::Round(($d.Size / 1GB), 2)
    $free  = [math]::Round(($d.FreeSpace / 1GB), 2)

    if ($d.Size -gt 0) {
        $percentFree = [math]::Round((($d.FreeSpace / $d.Size) * 100), 1)
    } else {
        $percentFree = 0
    }

    if ($d.DeviceID -eq $systemDrive) {
        $systemDriveFree = $percentFree
    }

    $percentUsed = [math]::Round((100 - $percentFree), 1)

    # 🎯 COR POR STATUS
    if ($percentFree -lt 10) {
        $color = "#f44336"   # vermelho crítico
        $motivos.Add("CRITICO: Disco com menos de 10% livre na unidade $($d.DeviceID)")
        $recomendacoes.Add("Limpeza urgente ou upgrade de armazenamento na unidade $($d.DeviceID)")
    }
    elseif ($percentFree -lt 15) {
        $color = "#ff9800"   # laranja alerta
        $motivos.Add("Espaco em disco baixo na unidade $($d.DeviceID)")
        $recomendacoes.Add("Avaliar limpeza de disco na unidade $($d.DeviceID)")
    }
    else {
        $color = "#4CAF50"   # verde OK
    }

    $bar = "<div style='background:#ddd;width:300px'>
<div style='background:$color;width:$percentUsed%;color:white'>$percentUsed %</div>
</div>"

    $diskTableHTML += @"
<tr>
<td>$($d.DeviceID)</td>
<td>$total GB</td>
<td>$free GB</td>
<td>$bar</td>
</tr>
"@
}

# ✅ recomendação global (uma vez só)
if ($disks.Count -gt 0) {
    $recomendacoes.Add("Executar limpeza de disco (cleanmgr) ou analisar uso com WinDirStat")
}

$processosExecutados.Add("Analise de espaco em disco")







# =========================
# BENCHMARK DE DISCO
# =========================

Step-Log "Executando benchmark simples de disco..."

$testFile = "C:\TI\disk_test.tmp"
$diskSpeed = 0

try {
$testFile = "$env:TEMP\disk_test.bin"
$size = 200MB

$buffer = New-Object byte[] $size

$sw = [Diagnostics.Stopwatch]::StartNew()
[System.IO.File]::WriteAllBytes($testFile,$buffer)
$sw.Stop()

$diskSpeed = [math]::Round(($size/1MB)/$sw.Elapsed.TotalSeconds,2)

Remove-Item $testFile -Force

    if ($sw.Elapsed.TotalSeconds -gt 0) {
        $diskSpeed = [math]::Round((200 / $sw.Elapsed.TotalSeconds), 2)
    }

    Remove-Item $testFile -ErrorAction SilentlyContinue

    if ($diskSpeed -lt 80) {
        $motivos.Add("Desempenho de disco abaixo do ideal ($diskSpeed MB/s) na unidade do sistema")
        $recomendacoes.Add("Considerar migracao para SSD na unidade do sistema")
    }

    $processosExecutados.Add("Benchmark de disco")
}
catch {
    $motivos.Add("Falha no benchmark de disco na unidade na unidade do sistema")
}

# =========================
# TESTE REAL DE REDE
# =========================

Step-Log "Executando teste real de rede..."

$netSpeedReal = "Nao testado"

if (Test-Path $redeDestino) {
    try {
        $arquivoTeste = "$env:TEMP\teste_rede.bin"
        fsutil file createnew $arquivoTeste 50000000 | Out-Null

        $sw = [Diagnostics.Stopwatch]::StartNew()
        Copy-Item $arquivoTeste $redeDestino -Force -ErrorAction Stop
        $sw.Stop()

        if ($sw.Elapsed.TotalSeconds -gt 0) {
            $netSpeedReal = [math]::Round((50 / $sw.Elapsed.TotalSeconds), 2)
        }

        Remove-Item $arquivoTeste -ErrorAction SilentlyContinue

        if (($netSpeedReal -is [double]) -and ($netSpeedReal -lt 60)) {
            $motivos.Add("Velocidade de rede abaixo do esperado ($netSpeedReal MB/s)")
            $recomendacoes.Add("Verificar cabeamento, switch, placa de rede ou compartilhamento")
        }

        $processosExecutados.Add("Teste real de rede por copia")
    }
    catch {
        $netSpeedReal = "Falha no teste"
        $motivos.Add("Falha no teste real de rede")
    }
}
else {
    $motivos.Add("Pasta de teste de rede nao encontrada")
}

# =========================
# STARTUP APPS
# =========================

Step-Log "Levantando aplicativos na inicializacao do Windows..."

$startupHTML = ""

try {
    $startupItems = Get-CimInstance Win32_StartupCommand

    foreach ($item in $startupItems) {
        $recomendacaoStartup = "Manter"

        if ($item.Name -match "Teams|OneDrive|Spotify|Adobe|Updater|Update|Discord|Zoom|Skype") {
            $recomendacaoStartup = "Opcional - pode desativar se nao for uso diario"
        }

        if ($item.Name -match "Defender|Security|Antivirus") {
            $recomendacaoStartup = "Manter ativo"
        }

        $startupHTML += @"
<tr>
<td>$($item.Name)</td>
<td>$($item.Location)</td>
<td>$recomendacaoStartup</td>
</tr>
"@
    }

    if ([string]::IsNullOrWhiteSpace($startupHTML)) {
        $startupHTML = "<tr><td colspan='3'>Nenhum item de inicializacao encontrado</td></tr>"
    }

    $processosExecutados.Add("Levantamento de aplicativos na inicializacao")
}
catch {
    $startupHTML = "<tr><td colspan='3'>Falha ao consultar inicializacao</td></tr>"
    $motivos.Add("Falha ao consultar startup apps")
}

# =========================
# SOFTWARE INSTALADO
# =========================

Step-Log "Levantando softwares instalados..."

$softwareHTML = ""

try {
    $software64 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName }

    $software32 = Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName }

    $software = @($software64 + $software32) | Sort-Object DisplayName -Unique

    foreach ($s in $software) {
        $softwareHTML += @"
<tr>
<td>$($s.DisplayName)</td>
<td>$($s.DisplayVersion)</td>
</tr>
"@
    }

    if ([string]::IsNullOrWhiteSpace($softwareHTML)) {
        $softwareHTML = "<tr><td colspan='2'>Nenhum software encontrado</td></tr>"
    }

    $processosExecutados.Add("Inventario de softwares instalados")
}
catch {
    $softwareHTML = "<tr><td colspan='2'>Falha ao consultar softwares instalados</td></tr>"
    $motivos.Add("Falha ao consultar softwares instalados")
}

# =========================
# DIAGNOSTICO INTELIGENTE
# =========================

Step-Log "Gerando diagnostico automatico..."

$diagnostico = "Sistema operando normalmente."

if ($ramPercent -gt 85) {
    $diagnostico = "Uso elevado de memoria. Possivel excesso de aplicacoes abertas ou necessidade de upgrade."
}
elseif ($cpuLoad -gt 85) {
    $diagnostico = "CPU com alta utilizacao. Verificar processos pesados ou travados."
}
elseif ($diskSpeed -lt 80) {
    $diagnostico = "Disco com desempenho abaixo do ideal. Possivel gargalo de armazenamento."
}
elseif (($netSpeedReal -is [double]) -and ($netSpeedReal -lt 60)) {
    $diagnostico = "Possivel lentidao de rede."
}



$processosExecutados.Add("Diagnostico automatico gerado")

# =========================
# RECOMENDACOES TECNICAS
# =========================

Step-Log "Consolidando recomendacoes tecnicas..."

if ($ramPercent -gt 85) { $recomendacoes.Add("Fechar aplicativos em excesso ou considerar upgrade de memoria") }
if ($cpuLoad -gt 85)   { $recomendacoes.Add("Verificar processos com alto consumo de CPU") }

$recomendacaoHTML = ""
foreach ($r in ($recomendacoes | Sort-Object -Unique)) {
    $recomendacaoHTML += "<li>$r</li>"
}

if ([string]::IsNullOrWhiteSpace($recomendacaoHTML)) {
    $recomendacaoHTML = "<li>Nenhuma recomendacao adicional</li>"
}


Step-Log "Executando detecção automática de PC lento..."

$lentidaoScore = 0
$lentidaoMotivos = @()

$acoesTecnicas = New-Object System.Collections.Generic.List[string]


$nivelProblema = "NORMAL"

if ($lentidaoScore -ge 3 -or $systemDriveFree -lt 10)
{
    $nivelProblema = "CRITICO"
}
elseif ($lentidaoScore -ge 1) {
    $nivelProblema = "MODERADO"
}


# =========================
# CPU
# =========================

if ($cpuLoad -gt 85) {
    $lentidaoScore++
    $lentidaoMotivos += "CPU acima de 85%"
}

# =========================
# RAM
# =========================

if ($ramPercent -gt 85) {
    $lentidaoScore++
    $lentidaoMotivos += "Uso de RAM acima de 85%"
}

# =========================
# DISCO - VELOCIDADE
# =========================

if ($diskSpeed -lt 80) {
    $lentidaoScore++
    $lentidaoMotivos += "Disco com velocidade baixa ($diskSpeed MB/s)"
}

# =========================
# DISCO - USO REAL
# =========================

try{

$diskLoad = (Get-Counter '\PhysicalDisk(_Total)\% Disk Time' -ErrorAction Stop).CounterSamples.CookedValue
$diskLoad = [math]::Round($diskLoad,1)

if($diskLoad -gt 90){

$lentidaoScore++
$lentidaoMotivos += "Disco saturado ($diskLoad %)"

}

}
catch{

Step-Log "Contador de disco não disponível neste Windows" Yellow

}

# =========================
# PROCESSO PESADO
# =========================

try{

$topProcess = Get-Process | Sort-Object CPU -Descending | Select-Object -First 1

if($topProcess.CPU -gt 300){

$lentidaoScore++
$lentidaoMotivos += "Processo pesado detectado: $($topProcess.ProcessName)"

}

}catch{}

# =========================
# STARTUP
# =========================

try{

$startupCount = (Get-CimInstance Win32_StartupCommand).Count

if($startupCount -gt 15){

$lentidaoScore++
$lentidaoMotivos += "Muitos programas na inicializacao ($startupCount)"

}

}catch{}

# =========================
# REDE
# =========================

if (($netSpeedReal -is [double]) -and ($netSpeedReal -lt 60)) {

$lentidaoScore++
$lentidaoMotivos += "Rede abaixo de 60 MB/s"

}

# =========================
# CLASSIFICACAO
# =========================

if ($lentidaoScore -eq 0){

$pcLento = $false
$status = "SAUDAVEL"

}

elseif ($lentidaoScore -le 2){

$pcLento = $false
$status = "OBSERVACAO"

}

else{

$pcLento = $true
$status = "CRITICO"

}

$lentidaoMotivosHTML = ""

foreach ($m in $lentidaoMotivos){

$lentidaoMotivosHTML += "<li>$m</li>"

}



# =========================
# STATUS GERAL
# =========================

Step-Log "Calculando status geral da maquina..."

$status = "OK"

if ($ramPercent -gt 85) { $status = "ATENCAO" }
if ($cpuLoad -gt 85)    { $status = "ATENCAO" }
if ($diskSpeed -lt 80)  { $status = "ATENCAO" }

if ($motivos.Count -eq 0) {
    $motivos.Add("Nenhum problema relevante detectado")
}

$motivoHTML = ""
foreach ($m in ($motivos | Sort-Object -Unique)) {
    $motivoHTML += "<li>$m</li>"
}

# =========================
# HISTORICO
# =========================

Step-Log "Gravando historico CSV..."

$historyFile = "$basePath\historico_$hostname.csv"

$historyData = [PSCustomObject]@{
    DATE  = (Get-Date)
    CPU   = $cpuLoad
    RAM   = $ramPercent
    DISK  = $diskSpeed
    NET   = $netSpeedReal
    STATUS= $status
}

$historyData | Export-Csv $historyFile -Append -NoTypeInformation

$processosExecutados.Add("Gravacao de historico CSV")

# =========================
# RANKING
# =========================

Step-Log "Atualizando ranking de PCs..."

$rankingFile = "$basePath\ranking_pc.csv"
$pcScore = [math]::Round(($cpuLoad + $ramPercent + (100 - [math]::Min($diskSpeed,100))) / 3, 2)

$rankingData = [PSCustomObject]@{
    PC    = $hostname
    CPU   = $cpuLoad
    RAM   = $ramPercent
    DISK  = $diskSpeed
    NET   = $netSpeedReal
    SCORE = $pcScore
    DATE  = (Get-Date)
}

$rankingData | Export-Csv $rankingFile -Append -NoTypeInformation

$topRankingHTML = ""
try {
    $rankingLoaded = Import-Csv $rankingFile | Sort-Object {[double]$_.SCORE} -Descending | Select-Object -First 10
    foreach ($r in $rankingLoaded) {
        $topRankingHTML += @"
<tr>
<td>$($r.PC)</td>
<td>$($r.CPU)</td>
<td>$($r.RAM)</td>
<td>$($r.DISK)</td>
<td>$($r.NET)</td>
<td>$($r.SCORE)</td>
</tr>
"@
    }
}
catch {
    $topRankingHTML = "<tr><td colspan='6'>Falha ao carregar ranking</td></tr>"
}

$processosExecutados.Add("Atualizacao de ranking CSV")


# =========================
# GRAFICOS VISUAIS
# =========================

Step-Log "Montando graficos visuais..."

$cpuBar = "<div style='background:#ddd;width:300px'><div style='background:#4CAF50;width:$cpuLoad%;color:white'>$cpuLoad %</div></div>"
$ramBar = "<div style='background:#ddd;width:300px'><div style='background:#2196F3;width:$ramPercent%;color:white'>$ramPercent %</div></div>"



# =========================
# DIAGNOSTICO INTELIGENTE
# =========================

# =========================
# DIAGNOSTICO AVANCADO V3
# =========================

Step-Log "Executando diagnostico avancado V3..."

# -------------------------
# SSD OU HDD
# -------------------------
$diskType = "Desconhecido"

try {
    $physical = Get-PhysicalDisk -ErrorAction Stop

    foreach ($pd in $physical) {
        if ($pd.MediaType -eq "SSD") {
            $diskType = "SSD ($systemDrive)"
        }
        elseif ($pd.MediaType -eq "HDD") {
            $diskType = "HDD ($systemDrive)"
        }
    }
}
catch {
    $diskType = "Nao identificado ($systemDrive)"
}


# =========================
# DIAGNOSTICO AVANCADO V4
# =========================

Step-Log "Executando diagnostico avancado V4..."


if ($systemDriveFree -lt 10) {

    $motivos.Add("CRITICO: Disco com menos de 10% livre na unidade do sistema")

    $recomendacoes.Add("Limpeza urgente de disco ou expansao de armazenamento na unidade do sistema")

    $acoesTecnicas.Add("Executar limpeza avancada: cleanmgr /sageset:1 e cleanmgr /sagerun:1 unidade do sistema")
    $acoesTecnicas.Add("Remover arquivos temporarios e downloads desnecessarios na unidade do sistema")
    $acoesTecnicas.Add("Avaliar upgrade para SSD ou aumento de capacidade na unidade do sistema")
}
if ($diskSpeed -lt 80) {

    $acoesTecnicas.Add("Executar CHKDSK: chkdsk $systemDrive /f /r")
    $acoesTecnicas.Add("Verificar SMART do disco com ferramenta especializada na unidade do sistema")
    $acoesTecnicas.Add("Se HDD, recomendar migracao para SSD imediatamente na unidade do sistema")
}
if ($ramPercent -gt 85) {

    $acoesTecnicas.Add("Identificar processos com alto consumo de memoria (Task Manager / script)")
    $acoesTecnicas.Add("Executar diagnostico de memoria: mdsched.exe")
    $acoesTecnicas.Add("Considerar upgrade de memoria RAM")
}

if ($cpuLoad -gt 85) {

    $acoesTecnicas.Add("Analisar processos com alto uso de CPU")
    $acoesTecnicas.Add("Verificar malware ou processos travados")
}
if ($updatesCount -gt 5 -or $motivos -match "Falha") {

    $acoesTecnicas.Add("Executar DISM: DISM /Online /Cleanup-Image /RestoreHealth")
    $acoesTecnicas.Add("Executar SFC: sfc /scannow")
}

if (($netSpeedReal -is [double]) -and ($netSpeedReal -lt 60)) {

    $acoesTecnicas.Add("Testar cabo de rede e porta do switch")
    $acoesTecnicas.Add("Atualizar driver da placa de rede")
    $acoesTecnicas.Add("Executar reset de rede: netsh int ip reset")
}

if ($defenderStatus -ne "Ativo") {

    $acoesTecnicas.Add("Ativar protecao em tempo real do Windows Defender")
    $acoesTecnicas.Add("Executar verificacao completa de virus")
}

if ($azureJoined -ne "YES") {

    $acoesTecnicas.Add("Ingressar dispositivo no Entra ID (Azure AD)")
    $acoesTecnicas.Add("Validar politicas de seguranca e compliance")
}

$acoesHTML = ""

foreach ($a in ($acoesTecnicas | Sort-Object -Unique)) {
    $acoesHTML += "<li>$a</li>"
}

if ([string]::IsNullOrWhiteSpace($acoesHTML)) {
    $acoesHTML = "<li>Nenhuma acao tecnica necessaria</li>"
}
# -------------------------
# TEMPERATURA CPU
# -------------------------

$cpuTemp="Nao suportado"

try{

$temp = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction Stop

if($temp){

$cpuTemp = ($temp.CurrentTemperature / 10 - 273.15)

$cpuTemp = [math]::Round($cpuTemp,1)

}

}catch{}

# -------------------------
# SMART DISCO
# -------------------------

$smartStatus="Nao suportado"

try{

$smart = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction Stop

if($smart){

if($smart.PredictFailure){
$smartStatus="ALERTA"
$motivos.Add("SMART do disco indica possivel falha")
}
else{
$smartStatus="OK"
}

}

}catch{}

# -------------------------
# LATENCIA DNS
# -------------------------

$dnsLatency="Nao testado"

try{

$sw = [Diagnostics.Stopwatch]::StartNew()

Resolve-DnsName "google.com" -ErrorAction Stop | Out-Null

$sw.Stop()

$dnsLatency=[math]::Round($sw.Elapsed.TotalMilliseconds,0)

}catch{}

# -------------------------
# LATENCIA INTERNET
# -------------------------

$internetLatency="Nao testado"

try{

$ping = Test-Connection "8.8.8.8" -Count 2 -ErrorAction Stop

$internetLatency = ($ping | Measure-Object ResponseTime -Average).Average

$internetLatency = [math]::Round($internetLatency,0)

}catch{}

# -------------------------
# TEMPO DE BOOT
# -------------------------

$bootTime="Desconhecido"

try{

$bootEvent = Get-WinEvent -FilterHashtable @{
LogName='System'
ID=6005
} -MaxEvents 1

$bootTime=$bootEvent.TimeCreated

}catch{}

# -------------------------
# TOP 5 PROCESSOS PESADOS
# -------------------------

$topProcessHTML=""

try {

$top5 = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5

foreach($p in $top5){

$cpu = [math]::Round($p.CPU,1)
$ram = [math]::Round($p.WorkingSet/1MB,0)

$path = "N/A"
try{
    $path = (Get-Process -Id $p.Id -ErrorAction Stop).Path
}catch{}

$topProcessHTML += @"
<tr>
<td>$($p.ProcessName)</td>
<td>$($p.Id)</td>
<td>$cpu s</td>
<td>$ram MB</td>
<td style='font-size:10px'>$path</td>
</tr>
"@

}

}catch{

$topProcessHTML="<tr><td colspan='5'>Nao foi possivel coletar</td></tr>"

}

# -------------------------
# SCORE DE SAUDE DA MAQUINA
# -------------------------

$healthScore=100

if($cpuLoad -gt 85){ $healthScore -= 20 }
if($ramPercent -gt 85){ $healthScore -= 20 }
if($diskSpeed -lt 80){ $healthScore -= 20 }

if(($netSpeedReal -is [double]) -and ($netSpeedReal -lt 60)){
$healthScore -= 15
}

if($smartStatus -eq "ALERTA"){
$healthScore -= 30
}

if($healthScore -lt 0){ $healthScore=0 }



# diskColor = "#ff5722"

$bitlockerAtivo = $bitlockerHTML -match "Ativo"

if ($percentFree -lt 15) {
    $diskColor = "#f44336"
}
elseif ($bitlockerAtivo) {
    $diskColor = "#2196F3"
}

$bar = "<div style='background:#ddd;width:300px'>
<div style='background:$diskColor;width:$percentUsed%;color:white'>$percentUsed %</div>
</div>"











# =========================
# RELATORIO HTML
# =========================

Step-Log "Gerando relatorio HTML..."

$processosHTML = ""
foreach ($p in $processosExecutados) {
    $processosHTML += "<li>$p</li>"
}

$html = @"
<html>
<head>
<meta charset='UTF-8'>
<title>Relatorio Auditoria TI</title>
<style>
body{font-family:Arial;margin:30px;background:#fafafa;color:#222}
h2,h3{margin-top:25px}
table{border-collapse:collapse;width:900px;margin-bottom:20px;background:#fff}
td,th{border:1px solid #ccc;padding:8px;vertical-align:top}
th{background:#f2f2f2;text-align:left}
.card{background:#fff;border:1px solid #ddd;padding:15px;margin-bottom:20px;width:900px}
.status-ok{color:green;font-weight:bold}
.status-atencao{color:orange;font-weight:bold}
</style>
</head>

<body>

<img src="https://os.arrayservice.com.br/mobile/php/logo.png" height="60">

<h2>Painel de Diagnostico Workstation</h2>

<div class="card">
<h3>Resumo Executivo</h3>
<table>
<tr><th>CPU</th><td>$cpuLoad %</td><th>RAM</th><td>$ramPercent %</td></tr>
<tr><th>Disco</th><td>$diskSpeed MB/s</td><th>Rede</th><td>$netSpeedReal MB/s</td></tr>
<tr><th>Status</th><td colspan="3">$(if($status -eq 'OK'){"<span class='status-ok'>$status</span>"}else{"<span class='status-atencao'>$status</span>"})</td></tr>
</table>
</div>

<h3>Informacoes do Sistema</h3>
<table>
<tr><th>Hostname</th><td>$hostname</td></tr>
<tr><th>Usuario</th><td>$usuario</td></tr>
<tr><th>Dominio</th><td>$dominio</td></tr>
<tr><th>Fabricante</th><td>$fabricante</td></tr>
<tr><th>Modelo</th><td>$modeloPC</td></tr>
<tr><th>Serial</th><td>$serial</td></tr>
<tr><th>BIOS</th><td>$biosVer</td></tr>
<tr><th>Placa Mae</th><td>$placaMae</td></tr>
<tr><th>Uptime</th><td>$($uptime.Days) dias</td></tr>
<tr><th>Instalação do Windows</th><td>$installDateFormat</td></tr>
<tr><th>Formatação Disco C</th><td>$formatDateFormat</td></tr>
<tr><th>Idade do Sistema</th><td>$((New-TimeSpan -Start $installDate -End (Get-Date)).Days) dias</td></tr>
</table>

<h3>Sistema Operacional</h3>
<table>
<tr><th>Sistema</th><td>$windowsVersion</td></tr>
<tr><th>Versao</th><td>$windowsRelease</td></tr>
<tr><th>Build</th><td>$windowsBuild</td></tr>
</table>

<h3>Hardware</h3>
<table>
<tr><th>CPU</th><td>$($cpuName)</td></tr>
<tr><th>RAM Total</th><td>$ramTotal GB</td></tr>
<tr><th>RAM Livre</th><td>$ramLivre GB</td></tr>
<tr><th>GPU</th><td>$($gpu.Name)</td></tr>
<tr><th>GPU</th><td>$($gpuName)</td></tr>
</table>


<h3>Uso de Recursos</h3>
<table>
<tr><th>CPU</th><td>$cpuBar</td></tr>
<tr><th>RAM</th><td>$ramBar</td></tr>
</table>

<h3>Performance</h3>
<table>
<tr><th>Benchmark Disco</th><td>$diskSpeed MB/s</td></tr>
<tr><th>Teste Real de Rede</th><td>$netSpeedReal MB/s</td></tr>
</table>

<h3>BitLocker</h3>
<table>
<tr>
<th>Unidade</th>
<th>Status</th>
<th>Criptografia</th>
<th>Volume</th>
<th>KeyProtector ID</th>
<th>Resumo</th>
</tr>
$bitlockerHTML
</table>

<h3>Backup BitLocker</h3>
<table>
<tr><th>Status</th><td>$bitlockerExportStatus</td></tr>
<tr><th>Pasta</th><td>$bitlockerFolder</td></tr>
</table>

<h3>Espaco em Disco</h3>
<table>
<tr>
<th>Unidade</th>
<th>Total</th>
<th>Livre</th>
<th>Uso</th>
</tr>
$diskTableHTML
</table>

<h3>Atualizacoes Pendentes</h3>
<p>Total: $updatesCount</p>
<ul>$updatesHTML</ul>

<h3>Seguranca</h3>
<table>
<tr><th>Windows Defender</th><td>$defenderStatus</td></tr>
</table>


<h3>Diagnostico Automatico</h3>
<div class="card">$diagnostico</div>


<h3>Motivos do Status</h3>
<ul>$motivoHTML</ul>

<h3>Recomendacoes Tecnicas</h3>
<ul>$recomendacaoHTML</ul>

<h3>Acoes Tecnicas Recomendadas</h3>
<ul>$acoesHTML</ul>

<h3>Nivel do Problema</h3>
<tr><th>Situação/Avaliação: </th><td>$nivelProblema</td></tr>

<h3>Detecção de Lentidão</h3>
<table>
<tr><th>PC Lento Detectado</th><td>$pcLento</td></tr>
<tr><th>Score de Problemas</th><td>$lentidaoScore</td></tr>
</table>

<ul>
$lentidaoMotivosHTML
</ul>


<h3>Top Ranking da Rede</h3>
<table>
<th>PC</th>
<th>CPU</th>
<th>RAM</th>
<th>DISK</th>
<th>NET</th>
<th>SCORE</th>
</tr>
$topRankingHTML
</table>

<h3>Processos Executados</h3>
<ul>$processosHTML</ul>

<h3>Status de Envio do Email</h3>
<p id='emailStatusPlaceholder'>Pendente</p>


<h3>Dashboard de Saúde da Máquina</h3>

<table>
<tr>
<th>Score de Saúde</th>
<td><b>$healthScore / 100</b></td>
</tr>

<tr>
<th>Tipo de Disco</th>
<td>$diskType</td>
</tr>

<tr>
<th>Temperatura CPU</th>
<td>$cpuTemp °C</td>
</tr>

<tr>
<th>SMART Disco</th>
<td>$smartStatus</td>
</tr>

<tr>
<th>Latência DNS</th>
<td>$dnsLatency ms</td>
</tr>

<tr>
<th>Latência Internet</th>
<td>$internetLatency ms</td>
</tr>

<tr>
<th>Último Boot</th>
<td>$bootTime</td>
</tr>

</table>


<h3>Top 5 Processos Pesados</h3>

<table>
<tr>
<th>Processo</th>
<th>PID</th>
<th>CPU (segundos acumulados)</th>
<th>RAM</th>
<th>Caminho</th>
</tr>
$topProcessHTML
</table>

<h3>Registro Azure AD / Entra ID</h3>
<table>
<tr><th>Azure AD Joined</th><td>$azureJoined</td></tr>
<tr><th>Domain Joined</th><td>$domainJoined</td></tr>
<tr><th>Device ID</th><td>$deviceId</td></tr>
<tr><th>Tenant ID</th><td>$tenantId</td></tr>
<tr><th>Tenant Name</th><td>$tenantName</td></tr>
<tr><th>MDM URL</th><td>$mdmUrl</td></tr>
</table>

<h3>Aplicativos na Inicializacao</h3>
<table>
<tr>
<th>Aplicativo</th>
<th>Origem</th>
<th>Recomendacao</th>
</tr>
$startupHTML
</table>

<h3>Softwares Instalados</h3>
<table>
<tr>
<th>Aplicativo</th>
<th>Versao</th>
</tr>
$softwareHTML
</table>


</body>
</html>
"@

$html | Out-File $report -Encoding utf8
$processosExecutados.Add("Geracao do relatorio HTML")



# =========================
# ENVIO EMAIL (ROBUSTO 587 -> 465)
# =========================

function Send-ReportMail {
    param(
        [string]$SmtpServer,
        [int[]]$Ports,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        # [string]$AttachmentPath
		[string[]]$AttachmentPath,
        [pscredential]$Credential
    )

  #  Add-Type -AssemblyName System.Net.Mail

    foreach ($p in $Ports) {
        try {
            # mensagem
            $msg = New-Object System.Net.Mail.MailMessage
            $msg.From = $From
            $msg.To.Add($To)
            $msg.Subject = $Subject
            $msg.Body = $Body
            $msg.IsBodyHtml = $false

foreach ($file in $AttachmentPath) {

    if (Test-Path $file) {
        $null = $msg.Attachments.Add($file)
    } else {
        Step-Log "Anexo não encontrado: $file" "Yellow"
    }

}

            # client SMTP
            $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $p)
            $smtp.EnableSsl = $true
            $smtp.Timeout = 30000
            $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = $Credential.GetNetworkCredential()

            $smtp.Send($msg)

            # dispose
            $msg.Dispose()
            $smtp.Dispose()

            if ($p -eq 587) { return "Email enviado via TLS 587" }
            if ($p -eq 465) { return "Email enviado via SSL 465" }
            return "Email enviado (porta $p)"
        }
        catch {
            $err = $_.Exception.Message
            Step-Log "Falha ao enviar (porta $p): $err" "Yellow"
            try { if ($msg)  { $msg.Dispose()  } } catch {}
            try { if ($smtp) { $smtp.Dispose() } } catch {}
        }
    }

    return "Falha SMTP (587 e 465) - ver logs no console"
}

Step-Log "Enviando relatorio por email..." "Yellow"

# TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Credencial (mantive seu padrão por compatibilidade)
$secure = ConvertTo-SecureString $emailPassword -AsPlainText -Force
$cred   = New-Object System.Management.Automation.PSCredential($emailFrom, $secure)



$emailStatus = Send-ReportMail `
    -SmtpServer $smtpServer `
    -Ports @(587,465) `
    -From $emailFrom `
    -To $emailTo `
    -Subject "ASI - Relatorio Diagnostico $hostname" `
    -Body "ASI - Relatorio automatico em anexo." `
    -AttachmentPath ($bitlockerFiles + $report) `
    -Credential $cred


# Atualiza status do email no HTML
Step-Log "Atualizando status de envio no relatorio..."
$htmlFinal = Get-Content $report -Raw
$htmlFinal = $htmlFinal -replace "<p id='emailStatusPlaceholder'>Pendente</p>", "<p>$emailStatus</p>"
$htmlFinal | Out-File $report -Encoding utf8

# =========================
# ABRIR RELATORIO
# =========================



if(Test-Path $report){

Step-Log "Abrindo relatorio final..." "Cyan"

Start-Process $report

}else{

Step-Log "Relatorio nao encontrado: $report" "Red"

}

Step-Log "STATUS EMAIL: $emailStatus" "Yellow"
Step-Log "Processo concluido com sucesso." "Cyan"
Write-Host "===== DIAGNOSTICO FINALIZADO =====" -ForegroundColor Cyan
