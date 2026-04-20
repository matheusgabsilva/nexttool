#Requires -Version 5.1
# ================================================================
# NextTool v3.0 - Ferramenta de TI da Next (GUI)
# github.com/matheusgabsilva/nexttool
#
# Uso via URL:
#   irm "https://raw.githubusercontent.com/matheusgabsilva/nexttool/main/nexttool.ps1" | iex
# ================================================================

Set-StrictMode -Off
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8
try { chcp 65001 | Out-Null } catch {}

# === ELEVACAO ===
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -Command `"irm 'https://raw.githubusercontent.com/matheusgabsilva/nexttool/main/nexttool.ps1' | iex`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ================================================================
# CONFIGURACAO GLOBAL
# ================================================================
$script:VERSION    = "3.0"
$script:REPORT_DIR = "C:\Next-Relatorios"
$script:SESSION_TS = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:LOG_FILE   = Join-Path $script:REPORT_DIR "nexttool_$($env:COMPUTERNAME)_$script:SESSION_TS.log"
$script:LogQueue   = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

if (-not (Test-Path $script:REPORT_DIR)) {
    New-Item -Path $script:REPORT_DIR -ItemType Directory -Force | Out-Null
}

# ================================================================
# WRITE-LOG  (funciona em runspaces — usa $LogQueue/$LOG_FILE sem $script:)
# ================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","ERRO","AVISO","STEP","PLAIN")]
        [string]$Level = "INFO"
    )
    $ts = [datetime]::Now.ToString("HH:mm:ss")
    switch ($Level) {
        "OK"    { $tag = "[OK]   "; $hex = "#98C379" }
        "ERRO"  { $tag = "[ERRO] "; $hex = "#E06C75" }
        "AVISO" { $tag = "[AVISO]"; $hex = "#E5C07B" }
        "STEP"  { $tag = "[>>]   "; $hex = "#61AFEF" }
        "PLAIN" { $tag = "       "; $hex = "#5C6370" }
        default { $tag = "[INFO] "; $hex = "#ABB2BF" }
    }
    $line = "$ts $tag $Message"
    if ($null -ne $LogQueue) {
        $LogQueue.Enqueue([PSCustomObject]@{ Text = $line; Color = $hex })
    }
    try { Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# ================================================================
# WINGET
# ================================================================
function Test-Winget { return ($null -ne (Get-Command winget -ErrorAction SilentlyContinue)) }

function Install-Winget {
    Write-Log "Instalando winget (App Installer)..." "STEP"
    try {
        $tmp = "$env:TEMP\AppInstaller.msixbundle"
        Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $tmp -UseBasicParsing
        Add-AppxPackage -Path $tmp
        Write-Log "winget instalado." "OK"
    } catch {
        Write-Log "Falha ao instalar winget: $_" "ERRO"
    }
}

function Install-WingetApp {
    param([string]$Id, [string]$Name)
    Write-Log "Instalando $Name..." "STEP"
    $result = winget install --id $Id -e --accept-source-agreements --accept-package-agreements --silent 2>&1
    if ($LASTEXITCODE -eq 0 -or $result -match "instalado|already installed|No applicable") {
        Write-Log "$Name instalado com sucesso." "OK"
    } else {
        Write-Log "Falha ao instalar $Name (codigo: $LASTEXITCODE)." "ERRO"
    }
}

# ================================================================
# OFFICE - ODT
# ================================================================
$script:OfficeXML = @{
    "365" = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="pt-br"/>
    </Product>
  </Add>
  <Updates Enabled="TRUE"/>
  <Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@
    "2021" = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="pt-br"/>
      <ExcludeApp ID="Teams"/>
    </Product>
  </Add>
  <Updates Enabled="FALSE"/>
  <Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@
    "2016" = @"
<Configuration>
  <Add OfficeClientEdition="64" SourcePath="" Version="">
    <Product ID="ProPlus2016Volume">
      <Language ID="pt-br"/>
    </Product>
  </Add>
  <Updates Enabled="FALSE"/>
  <Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@
}

function Install-Office {
    param([string]$Version)
    Write-Log "Preparando Office $Version via ODT..." "STEP"
    $odtDir = "$env:TEMP\NextODT"
    New-Item -Path $odtDir -ItemType Directory -Force | Out-Null
    $odtExe = "$odtDir\setup.exe"
    Write-Log "Baixando Office Deployment Tool..." "INFO"
    try {
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17928-20114.exe" `
            -OutFile "$odtDir\odt.exe" -UseBasicParsing
        Start-Process "$odtDir\odt.exe" -ArgumentList "/quiet /extract:$odtDir" -Wait
        Write-Log "ODT extraido." "OK"
    } catch {
        Write-Log "Falha ao baixar/extrair ODT: $_" "ERRO"
        return
    }
    $xmlPath = "$odtDir\config_$Version.xml"
    $OfficeXML[$Version] | Out-File -FilePath $xmlPath -Encoding UTF8
    Write-Log "Iniciando Office $Version (pode demorar varios minutos)..." "STEP"
    Start-Process $odtExe -ArgumentList "/configure `"$xmlPath`"" -Wait
    Write-Log "Office $Version instalado." "OK"
}

function Install-PadraoNext {
    Write-Log "=== INSTALACAO PADRAO NEXT ===" "STEP"
    if (-not (Test-Winget)) { Install-Winget }
    Install-WingetApp "Google.Chrome"               "Google Chrome"
    Install-WingetApp "RARLab.WinRAR"               "WinRAR"
    Install-WingetApp "Adobe.Acrobat.Reader.64-bit" "Adobe Acrobat Reader"
    Install-WingetApp "AnyDesk.AnyDesk"             "AnyDesk"
    Install-WingetApp "TeamViewer.TeamViewer"        "TeamViewer"
    Write-Log "Padrao Next concluido. Instale o Office separadamente." "OK"
}

# ================================================================
# TWEAKS
# ================================================================
function Invoke-TweakHibernacao {
    Write-Log "Desativando hibernacao..." "STEP"
    powercfg -h off 2>&1 | Out-Null
    Write-Log "Hibernacao desativada." "OK"
}

function Invoke-TweakSmartApp {
    Write-Log "Desativando Smart App Control..." "STEP"
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
    try {
        Set-ItemProperty -Path $path -Name "VerifiedAndReputablePolicyState" -Value 0 -Type DWord -Force
        Write-Log "Smart App Control desativado. Reinicie para aplicar." "OK"
    } catch {
        Write-Log "Nao foi possivel desativar: $_" "AVISO"
    }
}

function Invoke-TweakTelemetria {
    Write-Log "Desativando telemetria..." "STEP"
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "AllowTelemetry" -Value 0 -Type DWord -Force
    Stop-Service  "DiagTrack" -ErrorAction SilentlyContinue
    Set-Service   "DiagTrack" -StartupType Disabled -ErrorAction SilentlyContinue
    Write-Log "Telemetria desativada." "OK"
}

function Invoke-TweakActivityHistory {
    Write-Log "Desativando historico de atividades..." "STEP"
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "EnableActivityFeed"    -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $path -Name "PublishUserActivities" -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $path -Name "UploadUserActivities"  -Value 0 -Type DWord -Force
    Write-Log "Historico de atividades desativado." "OK"
}

function Invoke-TweakLocationTracking {
    Write-Log "Desativando rastreamento de localizacao..." "STEP"
    $path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "Value" -Value "Deny" -Force
    Write-Log "Rastreamento de localizacao desativado." "OK"
}

function Invoke-TweakFileExtensions {
    Write-Log "Exibindo extensoes de arquivo..." "STEP"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "HideFileExt" -Value 0 -Type DWord -Force
    Write-Log "Extensoes de arquivo visiveis." "OK"
}

function Invoke-TweakHiddenFiles {
    Write-Log "Exibindo arquivos ocultos..." "STEP"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "Hidden" -Value 1 -Type DWord -Force
    Write-Log "Arquivos ocultos visiveis." "OK"
}

function Invoke-TweakNumLock {
    Write-Log "Ativando Num Lock na inicializacao..." "STEP"
    Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name "InitialKeyboardIndicators" -Value "2" -Force
    Write-Log "Num Lock ativado." "OK"
}

function Invoke-TweakEndTask {
    Write-Log "Habilitando Finalizar Tarefa no botao direito..." "STEP"
    $path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "TaskbarEndTask" -Value 1 -Type DWord -Force
    Write-Log "Finalizar Tarefa habilitado." "OK"
}

function Invoke-TweakServices {
    Write-Log "Configurando servicos desnecessarios para Manual..." "STEP"
    $svcs = @(
        "DiagTrack","dmwappushservice","lfsvc","MapsBroker",
        "RemoteRegistry","TrkWks","WMPNetworkSvc",
        "XblAuthManager","XblGameSave","XboxGipSvc","XboxNetApiSvc"
    )
    foreach ($svc in $svcs) {
        try {
            if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
                Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue
                Write-Log " - $svc -> Manual" "PLAIN"
            }
        } catch {}
    }
    Write-Log "Servicos configurados." "OK"
}

function Invoke-TweakUltimatePerf {
    Write-Log "Ativando plano Ultimate Performance..." "STEP"
    $result = powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>&1 | Out-String
    $guid   = ([regex]::Match($result, "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")).Value
    if ($guid) {
        powercfg -setactive $guid | Out-Null
        Write-Log "Ultimate Performance ativado (GUID: $guid)." "OK"
    } else {
        Write-Log "Nao foi possivel ativar (pode ja estar ativo): $result" "AVISO"
    }
}

function Invoke-TweakDarkTheme {
    Write-Log "Aplicando tema escuro..." "STEP"
    $path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Set-ItemProperty -Path $path -Name "AppsUseLightTheme"    -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $path -Name "SystemUsesLightTheme" -Value 0 -Type DWord -Force
    Write-Log "Tema escuro aplicado." "OK"
}

function Invoke-TweakWidgets {
    Write-Log "Desativando Widgets do Windows 11..." "STEP"
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "AllowNewsAndInterests" -Value 0 -Type DWord -Force
    Write-Log "Widgets desativados." "OK"
}

function Invoke-TweakVerboseLogon {
    Write-Log "Ativando mensagens detalhadas no logon..." "STEP"
    $path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-ItemProperty -Path $path -Name "VerboseStatus" -Value 1 -Type DWord -Force
    Write-Log "Mensagens detalhadas no logon ativadas." "OK"
}

function Invoke-TweakDrivers {
    # Etapa 1: Windows Update
    Write-Log "=== ETAPA 1 — Windows Update (drivers) ===" "STEP"
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log "Instalando PSWindowsUpdate..." "INFO"
        try {
            Install-Module PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
            Write-Log "Modulo instalado." "OK"
        } catch {
            Write-Log "Falha ao instalar PSWindowsUpdate: $_" "ERRO"
        }
    }
    try {
        Import-Module PSWindowsUpdate -Force
        $updates = Get-WindowsUpdate -Category Drivers -ErrorAction Stop
        if ($updates.Count -eq 0) {
            Write-Log "Windows Update: nenhum driver pendente." "OK"
        } else {
            Write-Log "$($updates.Count) driver(s) encontrado(s) no Windows Update. Instalando..." "INFO"
            $updates | ForEach-Object { Write-Log " - $($_.Title)" "PLAIN" }
            Install-WindowsUpdate -Category Drivers -AcceptAll -IgnoreReboot -Verbose 2>&1 |
                ForEach-Object { Write-Log $_ "PLAIN" }
            Write-Log "Windows Update: drivers instalados. Reinicie para aplicar." "OK"
        }
    } catch {
        Write-Log "Erro no Windows Update: $_" "ERRO"
    }

    # Etapa 2: winget upgrade --all
    Write-Log "=== ETAPA 2 — winget upgrade (todos os pacotes) ===" "STEP"
    if (-not (Test-Winget)) {
        Write-Log "winget nao encontrado, pulando etapa 2." "AVISO"
        return
    }
    try {
        Write-Log "Executando winget upgrade --all (aguarde)..." "INFO"
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements 2>&1 |
            ForEach-Object { if ($_.Trim()) { Write-Log $_ "PLAIN" } }
        Write-Log "winget upgrade concluido." "OK"
    } catch {
        Write-Log "Erro no winget upgrade: $_" "ERRO"
    }
}

# ================================================================
# MANUTENCAO
# ================================================================
function Invoke-OtimizarPC {
    Write-Log "=== OTIMIZACAO DO PC ===" "STEP"
    Write-Log "Limpando arquivos temporarios..." "STEP"
    $paths = @($env:TEMP, "C:\Windows\Temp", "C:\Windows\Prefetch")
    $total = 0
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $files = Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue
            $total += $files.Count
            Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Log "$total arquivo(s) temporario(s) removido(s)." "OK"

    Write-Log "Esvaziando lixeira..." "STEP"
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    Write-Log "Lixeira esvaziada." "OK"

    Write-Log "Executando Limpeza de Disco (cleanmgr)..." "STEP"
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    @("Temporary Files","Internet Cache Files","Recycle Bin","Thumbnail Cache",
      "Downloaded Program Files","Memory Dump Files","Old ChkDsk Files",
      "Setup Log Files","System error memory dump files","Update Cleanup") | ForEach-Object {
        $p = "$regPath\$_"
        if (Test-Path $p) { Set-ItemProperty -Path $p -Name StateFlags0064 -Value 2 -ErrorAction SilentlyContinue }
    }
    $cg = Start-Process cleanmgr -ArgumentList "/sagerun:64" -PassThru -WindowStyle Hidden
    $done = $cg.WaitForExit(60000)
    if (-not $done) { try { $cg.Kill() } catch {}; Write-Log "Limpeza de disco encerrada (timeout 60s)." "AVISO" }
    else { Write-Log "Limpeza de disco concluida." "OK" }

    Write-Log "Limpando cache DNS..." "STEP"
    ipconfig /flushdns | Out-Null
    Write-Log "Cache DNS limpo." "OK"
    Write-Log "Otimizacao concluida." "OK"
}

function Invoke-Diagnostico {
    Write-Log "=== DIAGNOSTICO ===" "STEP"
    Write-Log "Top 10 processos por RAM:" "INFO"
    Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 | ForEach-Object {
        $mem = [math]::Round($_.WorkingSet64 / 1MB, 1)
        Write-Log (" {0} {1} MB" -f $_.Name.PadRight(28), $mem.ToString().PadLeft(8)) "PLAIN"
    }
    Write-Log "Programas na inicializacao:" "INFO"
    @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
      "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run") | ForEach-Object {
        $scope = if ($_ -like "*HKLM*") { "Sistema" } else { "Usuario" }
        if (Test-Path $_) {
            Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue |
                Get-Member -MemberType NoteProperty |
                Where-Object { $_.Name -notlike "PS*" } |
                ForEach-Object { Write-Log " [$scope] $($_.Name)" "PLAIN" }
        }
    }
    Write-Log "Seguranca:" "INFO"
    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        if ($def.AntivirusEnabled) {
            Write-Log "Defender ATIVO | Definicoes: $($def.AntivirusSignatureLastUpdated.ToString('dd/MM/yyyy'))" "OK"
        } else { Write-Log "Defender INATIVO" "ERRO" }
    } catch { Write-Log "Nao foi possivel verificar o Defender." "AVISO" }
    try {
        $fwOut   = netsh advfirewall show allprofiles state 2>&1 | Out-String
        $onCount = ([regex]::Matches($fwOut, "(?i)State\s+ON")).Count
        if ($onCount -gt 0) { Write-Log "Firewall: $onCount perfil(is) ativo(s)" "OK" }
        else { Write-Log "Firewall INATIVO" "ERRO" }
    } catch {}
    Write-Log "Erros criticos (ultimas 24h):" "INFO"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=([datetime]::Now.AddHours(-24))} `
            -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($events -and $events.Count -gt 0) {
            Write-Log "$($events.Count) erro(s) critico(s):" "AVISO"
            $events | ForEach-Object {
                $first = $_.Message.Split("`n")[0]
                $msg   = $first.Substring(0, [Math]::Min(90, $first.Length))
                Write-Log " [$($_.TimeCreated.ToString('HH:mm'))] $($_.ProviderName): $msg" "PLAIN"
            }
        } else { Write-Log "Nenhum erro critico nas ultimas 24h." "OK" }
    } catch { Write-Log "Nao foi possivel verificar Event Viewer." "AVISO" }
    Write-Log "Diagnostico concluido." "OK"
}

function Invoke-SFCDISM {
    Write-Log "=== SFC + DISM ===" "STEP"
    Write-Log "Executando SFC /scannow (aguarde)..." "STEP"
    $sfc = sfc /scannow 2>&1 | Out-String
    if ($sfc -match "encontrou|found")             { Write-Log "SFC: problemas encontrados e corrigidos." "OK" }
    elseif ($sfc -match "nao encontrou|did not find") { Write-Log "SFC: nenhuma violacao encontrada." "OK" }
    else { Write-Log "SFC concluido." "INFO" }
    Write-Log "Executando DISM RestoreHealth (aguarde)..." "STEP"
    $dism = DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
    if ($dism -match "concluida|successfully") { Write-Log "DISM concluido com sucesso." "OK" }
    else { Write-Log "DISM: verifique o resultado manualmente." "AVISO" }
}

# ================================================================
# REDE / DOMINIO
# ================================================================
function Get-NicConfig {
    param([string]$Adapter)
    Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = TRUE" | Where-Object {
        $na = Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex = $($_.InterfaceIndex)" -ErrorAction SilentlyContinue
        $na -and $na.NetConnectionID -like "*$Adapter*"
    } | Select-Object -First 1
}

function Show-Adapters {
    Write-Log "Adaptadores de rede ativos:" "INFO"
    Get-CimInstance Win32_NetworkAdapter -Filter "NetEnabled = TRUE" | ForEach-Object {
        Write-Log " - $($_.NetConnectionID)  [$($_.Name)]" "PLAIN"
    }
}

function Invoke-SetDNS {
    param([string]$Adapter, [string]$DNS1, [string]$DNS2)
    Write-Log "Configurando DNS em '$Adapter' -> $DNS1 / $DNS2..." "STEP"
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { Write-Log "Adaptador '$Adapter' nao encontrado." "ERRO"; return }
        $dns = @($DNS1); if ($DNS2) { $dns += $DNS2 }
        $r = Invoke-CimMethod -InputObject $cfg -MethodName SetDNSServerSearchOrder `
             -Arguments @{ DNSServerSearchOrder = [string[]]$dns }
        if ($r.ReturnValue -eq 0) { Write-Log "DNS configurado." "OK" }
        else { Write-Log "Codigo WMI: $($r.ReturnValue)" "AVISO" }
    } catch { Write-Log $_ "ERRO" }
}

function Invoke-ResetDNS {
    param([string]$Adapter)
    Write-Log "Resetando DNS de '$Adapter' para DHCP..." "STEP"
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { Write-Log "Adaptador '$Adapter' nao encontrado." "ERRO"; return }
        $r = Invoke-CimMethod -InputObject $cfg -MethodName SetDNSServerSearchOrder `
             -Arguments @{ DNSServerSearchOrder = [string[]]$null }
        if ($r.ReturnValue -eq 0) { Write-Log "DNS resetado para DHCP." "OK" }
        else { Write-Log "Codigo WMI: $($r.ReturnValue)" "AVISO" }
    } catch { Write-Log $_ "ERRO" }
}

function Invoke-TestarConectividade {
    Write-Log "Testando conectividade..." "STEP"
    foreach ($target in @("8.8.8.8","1.1.1.1","google.com")) {
        $ping = New-Object System.Net.NetworkInformation.Ping
        try {
            $reply = $ping.Send($target, 3000)
            if ($reply.Status -eq "Success") { Write-Log "Ping ${target}: $($reply.RoundtripTime)ms" "OK" }
            else { Write-Log "Ping ${target}: $($reply.Status)" "ERRO" }
        } catch { Write-Log "Ping ${target} falhou: $_" "ERRO" }
    }
}

function Show-IPConfig {
    Write-Log "=== IPCONFIG ===" "STEP"
    ipconfig /all 2>&1 | ForEach-Object { if ($_.Trim()) { Write-Log $_ "PLAIN" } }
}

function Invoke-JoinDomain {
    param([string]$Domain, [string]$User, [string]$Pass, [string]$NewName)
    if (-not $Domain -or -not $User -or -not $Pass) {
        Write-Log "Preencha Dominio, Usuario e Senha." "ERRO"; return
    }
    Write-Log "Ingressando em $Domain como $User..." "STEP"
    try {
        $cred = New-Object PSCredential("$Domain\$User", (ConvertTo-SecureString $Pass -AsPlainText -Force))
        if ($NewName) {
            Add-Computer -DomainName $Domain -Credential $cred -NewName $NewName -Force
            Write-Log "PC renomeado para '$NewName' e ingressado em $Domain." "OK"
        } else {
            Add-Computer -DomainName $Domain -Credential $cred -Force
            Write-Log "Ingressado em $Domain." "OK"
        }
        Write-Log "Reinicie o computador para aplicar." "AVISO"
    } catch { Write-Log "Falha ao ingressar no dominio: $_" "ERRO" }
}

# ================================================================
# IMPORT / EXPORT DE CONFIGURACAO
# ================================================================
function Export-Config {
    param([hashtable]$State)
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title      = "Exportar configuracao NextTool"
    $dlg.Filter     = "JSON (*.json)|*.json"
    $dlg.FileName   = "nexttool_perfil_$($env:COMPUTERNAME).json"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dlg.ShowDialog() -ne $true) { return }

    $obj = [ordered]@{
        version    = $script:VERSION
        exportedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        tweaks     = $State.Tweaks
        apps       = $State.Apps
    }
    $obj | ConvertTo-Json -Depth 5 | Out-File -FilePath $dlg.FileName -Encoding UTF8
    Write-Log "Configuracao exportada: $($dlg.FileName)" "OK"
}

function Import-Config {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title            = "Importar configuracao NextTool"
    $dlg.Filter           = "JSON (*.json)|*.json"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dlg.ShowDialog() -ne $true) { return $null }

    try {
        $json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
        Write-Log "Configuracao importada: $($dlg.FileName)" "OK"
        return $json
    } catch {
        Write-Log "Falha ao importar: $_" "ERRO"
        return $null
    }
}

# ================================================================
# ASYNC HELPER  — copia funcoes para o runspace e executa em background
# ================================================================
$script:FuncNames = @(
    'Write-Log','Test-Winget','Install-Winget','Install-WingetApp',
    'Install-Office','Install-PadraoNext',
    'Invoke-TweakHibernacao','Invoke-TweakSmartApp','Invoke-TweakDrivers',
    'Invoke-TweakTelemetria','Invoke-TweakActivityHistory','Invoke-TweakLocationTracking',
    'Invoke-TweakFileExtensions','Invoke-TweakHiddenFiles','Invoke-TweakNumLock',
    'Invoke-TweakEndTask','Invoke-TweakServices',
    'Invoke-TweakUltimatePerf','Invoke-TweakDarkTheme','Invoke-TweakWidgets','Invoke-TweakVerboseLogon',
    'Invoke-OtimizarPC','Invoke-Diagnostico','Invoke-SFCDISM',
    'Get-NicConfig','Show-Adapters','Invoke-SetDNS','Invoke-ResetDNS',
    'Invoke-TestarConectividade','Show-IPConfig','Invoke-JoinDomain'
)

function Invoke-Async {
    param([ScriptBlock]$Code, [hashtable]$Vars = @{})

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    Get-ChildItem function: | Where-Object { $_.Name -in $script:FuncNames } | ForEach-Object {
        $fd = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($_.Name, $_.ScriptBlock.ToString())
        $iss.Commands.Add($fd)
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = "MTA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()

    $rs.SessionStateProxy.SetVariable("LogQueue",  $script:LogQueue)
    $rs.SessionStateProxy.SetVariable("LOG_FILE",  $script:LOG_FILE)
    $rs.SessionStateProxy.SetVariable("OfficeXML", $script:OfficeXML)
    foreach ($k in $Vars.Keys) { $rs.SessionStateProxy.SetVariable($k, $Vars[$k]) }

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Code)
    [void]$ps.BeginInvoke()
}

# ================================================================
# XAML
# ================================================================
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="NextTool v3.0 - Ferramenta de TI"
    Height="720" Width="980"
    MinHeight="600" MinWidth="820"
    WindowStartupLocation="CenterScreen"
    Background="#282C34"
    FontFamily="Segoe UI"
    FontSize="13">

  <Window.Resources>

    <Style TargetType="Button">
      <Setter Property="Background" Value="#61AFEF"/>
      <Setter Property="Foreground" Value="#1E2128"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    CornerRadius="4"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Opacity" Value="0.82"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Opacity" Value="0.65"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#ABB2BF"/>
      <Setter Property="Margin" Value="0,5"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#1E2128"/>
      <Setter Property="Foreground" Value="#ABB2BF"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="CaretBrush" Value="#ABB2BF"/>
    </Style>

    <Style TargetType="PasswordBox">
      <Setter Property="Background" Value="#1E2128"/>
      <Setter Property="Foreground" Value="#ABB2BF"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,5"/>
    </Style>

    <Style TargetType="GroupBox">
      <Setter Property="Foreground" Value="#5C6370"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="Margin" Value="0,0,0,14"/>
      <Setter Property="Padding" Value="10,8"/>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="#21252B"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="0,1,0,0"/>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#5C6370"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="18,11"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="TabBorder"
                    Background="{TemplateBinding Background}"
                    BorderThickness="0,0,0,3"
                    BorderBrush="Transparent"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter ContentSource="Header"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#61AFEF"/>
                <Setter Property="Foreground" Value="#61AFEF"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#ABB2BF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#5C6370"/>
      <Setter Property="Padding" Value="0,0,0,2"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>

    <Style TargetType="Separator">
      <Setter Property="Background" Value="#3E4451"/>
      <Setter Property="Margin" Value="0,8"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="54"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="180"/>
    </Grid.RowDefinitions>

    <!-- ============================================================ -->
    <!-- HEADER                                                        -->
    <!-- ============================================================ -->
    <Border Grid.Row="0" Background="#21252B" BorderBrush="#3E4451" BorderThickness="0,0,0,1">
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Next" FontSize="22" FontWeight="Bold" Foreground="#61AFEF"/>
          <TextBlock Text="Tool" FontSize="22" FontWeight="Bold" Foreground="#ABB2BF"/>
          <TextBlock Text="  v3.0" FontSize="11" Foreground="#5C6370" VerticalAlignment="Bottom" Margin="2,0,0,4"/>
          <TextBlock Text="  |  Ferramenta de TI" FontSize="11" Foreground="#5C6370" VerticalAlignment="Bottom" Margin="0,0,0,4"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="BtnImport" Content="Importar" Padding="10,4" Margin="0,0,6,0"
                  Background="#4B5263" Foreground="#ABB2BF" FontSize="11" FontWeight="Normal"/>
          <Button x:Name="BtnExport" Content="Exportar" Padding="10,4" Margin="0,0,16,0"
                  Background="#4B5263" Foreground="#ABB2BF" FontSize="11" FontWeight="Normal"/>
          <TextBlock x:Name="TxtSysInfo" Foreground="#5C6370" FontSize="11" VerticalAlignment="Center"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ============================================================ -->
    <!-- TABS                                                          -->
    <!-- ============================================================ -->
    <TabControl Grid.Row="1" x:Name="MainTabs" Margin="0">

      <!-- ========================================================== -->
      <!-- TAB: INSTALACOES                                            -->
      <!-- ========================================================== -->
      <TabItem Header="  Instalações  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <StackPanel Margin="24,18">

            <Button x:Name="BtnPadraoNext"
                    Content="  ⚡  Instalar Padrão Next  —  Chrome · WinRAR · Adobe Reader · AnyDesk · TeamViewer  "
                    HorizontalAlignment="Left"
                    Background="#98C379"
                    Foreground="#1E2128"
                    FontSize="13"
                    Padding="18,12"
                    Margin="0,0,0,20"/>

            <GroupBox Header="Programas individuais">
              <WrapPanel Margin="4,4">
                <CheckBox x:Name="ChkChrome"   Content="Google Chrome"        Margin="0,5,28,5"/>
                <CheckBox x:Name="ChkWinrar"   Content="WinRAR"               Margin="0,5,28,5"/>
                <CheckBox x:Name="ChkAdobe"    Content="Adobe Acrobat Reader" Margin="0,5,28,5"/>
                <CheckBox x:Name="ChkAnydesk"  Content="AnyDesk"              Margin="0,5,28,5"/>
                <CheckBox x:Name="ChkTeamview" Content="TeamViewer"           Margin="0,5,28,5"/>
              </WrapPanel>
            </GroupBox>

            <Button x:Name="BtnInstalarSelecionados"
                    Content="Instalar Selecionados"
                    HorizontalAlignment="Left"
                    Margin="0,0,0,22"/>

            <GroupBox Header="Microsoft Office  (via ODT — download completo, pode demorar)">
              <StackPanel Orientation="Horizontal" Margin="4,4">
                <Button x:Name="BtnO365"  Content="Microsoft 365"   Margin="0,0,10,0"/>
                <Button x:Name="BtnO2021" Content="Office 2021 VL"  Margin="0,0,10,0"/>
                <Button x:Name="BtnO2016" Content="Office 2016 VL"/>
              </StackPanel>
            </GroupBox>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ========================================================== -->
      <!-- TAB: TWEAKS                                                 -->
      <!-- ========================================================== -->
      <TabItem Header="  Tweaks  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Presets -->
          <Border Grid.Row="0" Background="#1E2128" Padding="16,10">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Selecao rapida:" Foreground="#5C6370" VerticalAlignment="Center" Margin="0,0,12,0" FontSize="11"/>
              <Button x:Name="BtnPresetNext"    Content="Padrao Next" Background="#98C379" Foreground="#1E2128" Padding="14,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnPresetLimpar"  Content="Limpar"      Background="#4B5263" Foreground="#ABB2BF" Padding="14,5"/>
            </StackPanel>
          </Border>

          <!-- Checkboxes -->
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0">
            <Grid Margin="20,14">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <!-- Coluna esquerda: Essenciais -->
              <StackPanel Grid.Column="0">
                <GroupBox Header="Essenciais  (Padrao Next)">
                  <StackPanel Margin="4,4">
                    <CheckBox x:Name="ChkTelemetria"       Content="Desativar Telemetria"/>
                    <CheckBox x:Name="ChkActivityHistory"  Content="Desativar Historico de Atividades"/>
                    <CheckBox x:Name="ChkLocationTracking" Content="Desativar Rastreamento de Localizacao"/>
                    <CheckBox x:Name="ChkFileExtensions"   Content="Exibir Extensoes de Arquivo"/>
                    <CheckBox x:Name="ChkHiddenFiles"      Content="Exibir Arquivos Ocultos"/>
                    <CheckBox x:Name="ChkNumLock"          Content="Num Lock ativo na inicializacao"/>
                    <CheckBox x:Name="ChkEndTask"          Content="Finalizar Tarefa no botao direito"/>
                    <CheckBox x:Name="ChkServices"         Content="Servicos desnecessarios para Manual"/>
                    <CheckBox x:Name="ChkHibernacao"       Content="Desativar Hibernacao"/>
                    <CheckBox x:Name="ChkSmartApp"         Content="Desativar Smart App Control  (Win11)"/>
                  </StackPanel>
                </GroupBox>
              </StackPanel>

              <!-- Coluna direita: Preferencias -->
              <StackPanel Grid.Column="2">
                <GroupBox Header="Preferencias  (opcionais)">
                  <StackPanel Margin="4,4">
                    <CheckBox x:Name="ChkUltimatePerf"  Content="Plano Ultimate Performance"/>
                    <CheckBox x:Name="ChkDarkTheme"     Content="Tema Escuro do Windows"/>
                    <CheckBox x:Name="ChkWidgets"       Content="Desativar Widgets  (Win11)"/>
                    <CheckBox x:Name="ChkVerboseLogon"  Content="Mensagens detalhadas no Logon"/>
                  </StackPanel>
                </GroupBox>
              </StackPanel>

            </Grid>
          </ScrollViewer>

          <!-- Botao aplicar -->
          <Border Grid.Row="2" Background="#1E2128" Padding="16,10">
            <Button x:Name="BtnAplicarTweaks"
                    Content="Aplicar Tweaks Selecionados"
                    HorizontalAlignment="Left"
                    Background="#E5C07B"
                    Foreground="#1E2128"/>
          </Border>

        </Grid>
      </TabItem>

      <!-- ========================================================== -->
      <!-- TAB: MANUTENCAO                                             -->
      <!-- ========================================================== -->
      <TabItem Header="  Manutenção  ">
        <StackPanel Margin="24,18" Background="#21252B">
          <GroupBox Header="Ferramentas">
            <WrapPanel Margin="4,4">
              <Button x:Name="BtnOtimizar"    Content="Otimizar PC"      Width="155" Height="62" Margin="0,0,10,10"/>
              <Button x:Name="BtnDiagnostico" Content="Diagnostico"      Width="155" Height="62" Margin="0,0,10,10"/>
              <Button x:Name="BtnSfcDism"     Content="SFC + DISM"       Width="155" Height="62" Margin="0,0,10,10"/>
              <Button x:Name="BtnFlushDns"    Content="Flush DNS"        Width="155" Height="62" Margin="0,0,10,10"/>
              <Button x:Name="BtnAtualizarDrivers" Content="Atualizar Drivers" Width="155" Height="62" Margin="0,0,10,10"/>
              <Button x:Name="BtnRelatorio"   Content="Abrir Relatorios" Width="155" Height="62"
                      Background="#4B5263" Foreground="#ABB2BF"/>
            </WrapPanel>
          </GroupBox>
        </StackPanel>
      </TabItem>

      <!-- ========================================================== -->
      <!-- TAB: REDE / DOMINIO                                         -->
      <!-- ========================================================== -->
      <TabItem Header="  Rede / Domínio  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <Grid Margin="24,18">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="24"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Coluna esquerda -->
            <StackPanel Grid.Column="0">
              <GroupBox Header="Ferramentas de rede">
                <WrapPanel Margin="4,4">
                  <Button x:Name="BtnListarAdapters" Content="Listar Adaptadores"   Margin="0,0,8,8"/>
                  <Button x:Name="BtnTestarConect"   Content="Testar Conectividade" Margin="0,0,8,8"/>
                  <Button x:Name="BtnIPConfig"       Content="IPConfig /all"        Margin="0,0,0,8"/>
                </WrapPanel>
              </GroupBox>

              <GroupBox Header="Configurar DNS">
                <StackPanel Margin="4,4">
                  <Label Content="Adaptador  (ex: Ethernet, Wi-Fi)"/>
                  <TextBox x:Name="TxtDnsAdapter" Margin="0,0,0,8"/>
                  <Label Content="DNS Primario"/>
                  <TextBox x:Name="TxtDns1" Text="8.8.8.8" Margin="0,0,0,8"/>
                  <Label Content="DNS Secundario"/>
                  <TextBox x:Name="TxtDns2" Text="8.8.4.4" Margin="0,0,0,12"/>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnSetDns"   Content="Configurar DNS"    Margin="0,0,8,0"/>
                    <Button x:Name="BtnResetDns" Content="Resetar para DHCP" Background="#4B5263" Foreground="#ABB2BF"/>
                  </StackPanel>
                </StackPanel>
              </GroupBox>
            </StackPanel>

            <!-- Coluna direita -->
            <StackPanel Grid.Column="2">
              <GroupBox Header="Ingressar em Dominio AD">
                <StackPanel Margin="4,4">
                  <Label Content="Dominio  (ex: empresa.local)"/>
                  <TextBox x:Name="TxtDominio" Margin="0,0,0,8"/>
                  <Label Content="Usuario com permissao de join"/>
                  <TextBox x:Name="TxtDomUser" Margin="0,0,0,8"/>
                  <Label Content="Senha"/>
                  <PasswordBox x:Name="TxtDomPass" Margin="0,0,0,8"/>
                  <Label Content="Novo nome do PC  (opcional)"/>
                  <TextBox x:Name="TxtDomName" Margin="0,0,0,14"/>
                  <Button x:Name="BtnJoinDomain"
                          Content="Ingressar no Dominio"
                          HorizontalAlignment="Left"
                          Background="#E06C75"
                          Foreground="#1E2128"/>
                </StackPanel>
              </GroupBox>
            </StackPanel>

          </Grid>
        </ScrollViewer>
      </TabItem>

    </TabControl>

    <!-- ============================================================ -->
    <!-- LOG PANEL                                                     -->
    <!-- ============================================================ -->
    <Grid Grid.Row="2" Background="#1E2128">
      <Grid.RowDefinitions>
        <RowDefinition Height="26"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Grid.Row="0" Background="#21252B" BorderBrush="#3E4451" BorderThickness="0,1,0,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0">
          <TextBlock Text="Log de saida" Foreground="#5C6370" FontSize="11" FontWeight="SemiBold"/>
          <Button x:Name="BtnLimparLog" Content="limpar"
                  Padding="8,1" Margin="14,0,0,0"
                  Background="Transparent" Foreground="#5C6370"
                  FontSize="10" FontWeight="Normal"/>
        </StackPanel>
      </Border>
      <ListBox Grid.Row="1" x:Name="LogBox"
               Background="#1E2128"
               BorderThickness="0"
               Padding="10,4"
               ScrollViewer.HorizontalScrollBarVisibility="Disabled"
               VirtualizingPanel.IsVirtualizing="True"
               VirtualizingPanel.VirtualizationMode="Recycling"
               FontFamily="Consolas"
               FontSize="11"/>
    </Grid>

  </Grid>
</Window>
"@

# ================================================================
# CARREGAR JANELA
# ================================================================
$reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($reader)

$BtnPadraoNext           = $Window.FindName("BtnPadraoNext")
$ChkChrome               = $Window.FindName("ChkChrome")
$ChkWinrar               = $Window.FindName("ChkWinrar")
$ChkAdobe                = $Window.FindName("ChkAdobe")
$ChkAnydesk              = $Window.FindName("ChkAnydesk")
$ChkTeamview             = $Window.FindName("ChkTeamview")
$BtnInstalarSelecionados = $Window.FindName("BtnInstalarSelecionados")
$BtnO365                 = $Window.FindName("BtnO365")
$BtnO2021                = $Window.FindName("BtnO2021")
$BtnO2016                = $Window.FindName("BtnO2016")
$BtnImport               = $Window.FindName("BtnImport")
$BtnExport               = $Window.FindName("BtnExport")
$BtnPresetNext           = $Window.FindName("BtnPresetNext")
$BtnPresetLimpar         = $Window.FindName("BtnPresetLimpar")
$ChkTelemetria           = $Window.FindName("ChkTelemetria")
$ChkActivityHistory      = $Window.FindName("ChkActivityHistory")
$ChkLocationTracking     = $Window.FindName("ChkLocationTracking")
$ChkFileExtensions       = $Window.FindName("ChkFileExtensions")
$ChkHiddenFiles          = $Window.FindName("ChkHiddenFiles")
$ChkNumLock              = $Window.FindName("ChkNumLock")
$ChkEndTask              = $Window.FindName("ChkEndTask")
$ChkServices             = $Window.FindName("ChkServices")
$ChkHibernacao           = $Window.FindName("ChkHibernacao")
$ChkSmartApp             = $Window.FindName("ChkSmartApp")
$ChkUltimatePerf         = $Window.FindName("ChkUltimatePerf")
$ChkDarkTheme            = $Window.FindName("ChkDarkTheme")
$ChkWidgets              = $Window.FindName("ChkWidgets")
$ChkVerboseLogon         = $Window.FindName("ChkVerboseLogon")
$BtnAplicarTweaks        = $Window.FindName("BtnAplicarTweaks")
$BtnAtualizarDrivers     = $Window.FindName("BtnAtualizarDrivers")
$BtnOtimizar             = $Window.FindName("BtnOtimizar")
$BtnDiagnostico          = $Window.FindName("BtnDiagnostico")
$BtnSfcDism              = $Window.FindName("BtnSfcDism")
$BtnFlushDns             = $Window.FindName("BtnFlushDns")
$BtnRelatorio            = $Window.FindName("BtnRelatorio")
$TxtDnsAdapter           = $Window.FindName("TxtDnsAdapter")
$TxtDns1                 = $Window.FindName("TxtDns1")
$TxtDns2                 = $Window.FindName("TxtDns2")
$BtnSetDns               = $Window.FindName("BtnSetDns")
$BtnResetDns             = $Window.FindName("BtnResetDns")
$BtnListarAdapters       = $Window.FindName("BtnListarAdapters")
$BtnTestarConect         = $Window.FindName("BtnTestarConect")
$BtnIPConfig             = $Window.FindName("BtnIPConfig")
$TxtDominio              = $Window.FindName("TxtDominio")
$TxtDomUser              = $Window.FindName("TxtDomUser")
$TxtDomPass              = $Window.FindName("TxtDomPass")
$TxtDomName              = $Window.FindName("TxtDomName")
$BtnJoinDomain           = $Window.FindName("BtnJoinDomain")
$LogBox                  = $Window.FindName("LogBox")
$BtnLimparLog            = $Window.FindName("BtnLimparLog")
$TxtSysInfo              = $Window.FindName("TxtSysInfo")

# ================================================================
# TIMER — drena a fila de log dos runspaces para a UI
# ================================================================
$LogTimer = New-Object System.Windows.Threading.DispatcherTimer
$LogTimer.Interval = [TimeSpan]::FromMilliseconds(120)
$LogTimer.Add_Tick({
    $item  = $null
    $count = 0
    while ($script:LogQueue.TryDequeue([ref]$item) -and $count -lt 40) {
        $tb              = New-Object System.Windows.Controls.TextBlock
        $tb.Text         = $item.Text
        $tb.Foreground   = [System.Windows.Media.BrushConverter]::new().ConvertFrom($item.Color)
        $tb.TextWrapping = "Wrap"
        [void]$LogBox.Items.Add($tb)
        $count++
    }
    if ($count -gt 0) { $LogBox.ScrollIntoView($LogBox.Items[$LogBox.Items.Count - 1]) }
})
$LogTimer.Start()

# ================================================================
# SYSINFO NO HEADER
# ================================================================
try {
    $cs  = Get-CimInstance Win32_ComputerSystem
    $os  = Get-CimInstance Win32_OperatingSystem
    $ram = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    $TxtSysInfo.Text = "$env:COMPUTERNAME  |  $($os.Caption -replace 'Microsoft Windows ','Win ')  |  RAM: ${ram}GB  |  $env:USERNAME @ $($cs.Domain)"
} catch {}

# ================================================================
# EVENT HANDLERS
# ================================================================

# --- Instalacoes ---
$BtnPadraoNext.Add_Click({
    Invoke-Async { Install-PadraoNext }
})

$BtnInstalarSelecionados.Add_Click({
    $map = @{
        "Google.Chrome"               = @{ On = $ChkChrome.IsChecked;   Nome = "Google Chrome" }
        "RARLab.WinRAR"               = @{ On = $ChkWinrar.IsChecked;   Nome = "WinRAR" }
        "Adobe.Acrobat.Reader.64-bit" = @{ On = $ChkAdobe.IsChecked;    Nome = "Adobe Acrobat Reader" }
        "AnyDesk.AnyDesk"             = @{ On = $ChkAnydesk.IsChecked;  Nome = "AnyDesk" }
        "TeamViewer.TeamViewer"       = @{ On = $ChkTeamview.IsChecked; Nome = "TeamViewer" }
    }
    $ids   = @($map.Keys | Where-Object { $map[$_].On })
    $nomes = @{}; foreach ($id in $ids) { $nomes[$id] = $map[$id].Nome }

    if ($ids.Count -eq 0) { Write-Log "Nenhum programa selecionado." "AVISO"; return }

    Invoke-Async {
        if (-not (Test-Winget)) { Install-Winget }
        foreach ($id in $Ids) { Install-WingetApp $id $Nomes[$id] }
    } -Vars @{ Ids = $ids; Nomes = $nomes }
})

$BtnO365.Add_Click({  Invoke-Async { Install-Office "365"  } })
$BtnO2021.Add_Click({ Invoke-Async { Install-Office "2021" } })
$BtnO2016.Add_Click({ Invoke-Async { Install-Office "2016" } })

# --- Tweaks ---
$script:AllTweakChks = @(
    $ChkTelemetria,$ChkActivityHistory,$ChkLocationTracking,
    $ChkFileExtensions,$ChkHiddenFiles,$ChkNumLock,
    $ChkEndTask,$ChkServices,$ChkHibernacao,$ChkSmartApp,
    $ChkUltimatePerf,$ChkDarkTheme,$ChkWidgets,$ChkVerboseLogon
)

$BtnPresetNext.Add_Click({
    $ChkTelemetria.IsChecked      = $true
    $ChkActivityHistory.IsChecked = $true
    $ChkLocationTracking.IsChecked= $true
    $ChkFileExtensions.IsChecked  = $true
    $ChkHiddenFiles.IsChecked     = $true
    $ChkNumLock.IsChecked         = $true
    $ChkEndTask.IsChecked         = $true
    $ChkServices.IsChecked        = $true
    $ChkHibernacao.IsChecked      = $true
    $ChkSmartApp.IsChecked        = $true
    $ChkUltimatePerf.IsChecked    = $false
    $ChkDarkTheme.IsChecked       = $false
    $ChkWidgets.IsChecked         = $false
    $ChkVerboseLogon.IsChecked    = $false
})

$BtnPresetLimpar.Add_Click({
    $script:AllTweakChks | ForEach-Object { $_.IsChecked = $false }
})

$BtnAplicarTweaks.Add_Click({
    $v = @{
        Tel  = $ChkTelemetria.IsChecked;      Act  = $ChkActivityHistory.IsChecked
        Loc  = $ChkLocationTracking.IsChecked; Ext  = $ChkFileExtensions.IsChecked
        Hid  = $ChkHiddenFiles.IsChecked;      Num  = $ChkNumLock.IsChecked
        End  = $ChkEndTask.IsChecked;          Svc  = $ChkServices.IsChecked
        Hib  = $ChkHibernacao.IsChecked;       Sap  = $ChkSmartApp.IsChecked
        Perf = $ChkUltimatePerf.IsChecked;     Dark = $ChkDarkTheme.IsChecked
        Wgt  = $ChkWidgets.IsChecked;          Vrb  = $ChkVerboseLogon.IsChecked
    }
    if (-not ($v.Values | Where-Object { $_ })) { Write-Log "Nenhum tweak selecionado." "AVISO"; return }
    Invoke-Async {
        if ($V.Tel)  { Invoke-TweakTelemetria }
        if ($V.Act)  { Invoke-TweakActivityHistory }
        if ($V.Loc)  { Invoke-TweakLocationTracking }
        if ($V.Ext)  { Invoke-TweakFileExtensions }
        if ($V.Hid)  { Invoke-TweakHiddenFiles }
        if ($V.Num)  { Invoke-TweakNumLock }
        if ($V.End)  { Invoke-TweakEndTask }
        if ($V.Svc)  { Invoke-TweakServices }
        if ($V.Hib)  { Invoke-TweakHibernacao }
        if ($V.Sap)  { Invoke-TweakSmartApp }
        if ($V.Perf) { Invoke-TweakUltimatePerf }
        if ($V.Dark) { Invoke-TweakDarkTheme }
        if ($V.Wgt)  { Invoke-TweakWidgets }
        if ($V.Vrb)  { Invoke-TweakVerboseLogon }
        Write-Log "Todos os tweaks aplicados." "OK"
    } -Vars @{ V = $v }
})

# --- Manutencao ---
$BtnAtualizarDrivers.Add_Click({ Invoke-Async { Invoke-TweakDrivers } })
$BtnOtimizar.Add_Click({    Invoke-Async { Invoke-OtimizarPC } })
$BtnDiagnostico.Add_Click({ Invoke-Async { Invoke-Diagnostico } })
$BtnSfcDism.Add_Click({     Invoke-Async { Invoke-SFCDISM } })
$BtnFlushDns.Add_Click({
    Invoke-Async { ipconfig /flushdns | Out-Null; Write-Log "Cache DNS limpo." "OK" }
})
$BtnRelatorio.Add_Click({ Start-Process explorer.exe $script:REPORT_DIR })

# --- Rede ---
$BtnListarAdapters.Add_Click({ Invoke-Async { Show-Adapters } })
$BtnTestarConect.Add_Click({   Invoke-Async { Invoke-TestarConectividade } })
$BtnIPConfig.Add_Click({       Invoke-Async { Show-IPConfig } })

$BtnSetDns.Add_Click({
    $ad = $TxtDnsAdapter.Text.Trim()
    $d1 = $TxtDns1.Text.Trim()
    $d2 = $TxtDns2.Text.Trim()
    if (-not $ad -or -not $d1) { Write-Log "Preencha Adaptador e DNS Primario." "ERRO"; return }
    Invoke-Async { Invoke-SetDNS -Adapter $A -DNS1 $D1 -DNS2 $D2 } -Vars @{ A = $ad; D1 = $d1; D2 = $d2 }
})

$BtnResetDns.Add_Click({
    $ad = $TxtDnsAdapter.Text.Trim()
    if (-not $ad) { Write-Log "Preencha o nome do Adaptador." "ERRO"; return }
    Invoke-Async { Invoke-ResetDNS -Adapter $A } -Vars @{ A = $ad }
})

$BtnJoinDomain.Add_Click({
    $dom  = $TxtDominio.Text.Trim()
    $usr  = $TxtDomUser.Text.Trim()
    $pass = $TxtDomPass.Password
    $name = $TxtDomName.Text.Trim()
    if (-not $dom -or -not $usr -or -not $pass) {
        Write-Log "Preencha Dominio, Usuario e Senha." "ERRO"; return
    }
    Invoke-Async {
        Invoke-JoinDomain -Domain $Dom -User $Usr -Pass $Pass -NewName $Name
    } -Vars @{ Dom = $dom; Usr = $usr; Pass = $pass; Name = $name }
})

# --- Import / Export ---
$BtnExport.Add_Click({
    $tweaks = [ordered]@{
        Telemetria       = [bool]$ChkTelemetria.IsChecked
        ActivityHistory  = [bool]$ChkActivityHistory.IsChecked
        LocationTracking = [bool]$ChkLocationTracking.IsChecked
        FileExtensions   = [bool]$ChkFileExtensions.IsChecked
        HiddenFiles      = [bool]$ChkHiddenFiles.IsChecked
        NumLock          = [bool]$ChkNumLock.IsChecked
        EndTask          = [bool]$ChkEndTask.IsChecked
        Services         = [bool]$ChkServices.IsChecked
        Hibernacao       = [bool]$ChkHibernacao.IsChecked
        SmartApp         = [bool]$ChkSmartApp.IsChecked
        UltimatePerf     = [bool]$ChkUltimatePerf.IsChecked
        DarkTheme        = [bool]$ChkDarkTheme.IsChecked
        Widgets          = [bool]$ChkWidgets.IsChecked
        VerboseLogon     = [bool]$ChkVerboseLogon.IsChecked
    }
    $apps = [ordered]@{
        Chrome    = [bool]$ChkChrome.IsChecked
        Winrar    = [bool]$ChkWinrar.IsChecked
        Adobe     = [bool]$ChkAdobe.IsChecked
        Anydesk   = [bool]$ChkAnydesk.IsChecked
        Teamview  = [bool]$ChkTeamview.IsChecked
    }
    Export-Config -State @{ Tweaks = $tweaks; Apps = $apps }
})

$BtnImport.Add_Click({
    $cfg = Import-Config
    if (-not $cfg) { return }

    # Aplicar tweaks
    if ($cfg.tweaks) {
        $ChkTelemetria.IsChecked       = [bool]$cfg.tweaks.Telemetria
        $ChkActivityHistory.IsChecked  = [bool]$cfg.tweaks.ActivityHistory
        $ChkLocationTracking.IsChecked = [bool]$cfg.tweaks.LocationTracking
        $ChkFileExtensions.IsChecked   = [bool]$cfg.tweaks.FileExtensions
        $ChkHiddenFiles.IsChecked      = [bool]$cfg.tweaks.HiddenFiles
        $ChkNumLock.IsChecked          = [bool]$cfg.tweaks.NumLock
        $ChkEndTask.IsChecked          = [bool]$cfg.tweaks.EndTask
        $ChkServices.IsChecked         = [bool]$cfg.tweaks.Services
        $ChkHibernacao.IsChecked       = [bool]$cfg.tweaks.Hibernacao
        $ChkSmartApp.IsChecked         = [bool]$cfg.tweaks.SmartApp
        $ChkUltimatePerf.IsChecked     = [bool]$cfg.tweaks.UltimatePerf
        $ChkDarkTheme.IsChecked        = [bool]$cfg.tweaks.DarkTheme
        $ChkWidgets.IsChecked          = [bool]$cfg.tweaks.Widgets
        $ChkVerboseLogon.IsChecked     = [bool]$cfg.tweaks.VerboseLogon
    }

    # Aplicar apps
    if ($cfg.apps) {
        $ChkChrome.IsChecked   = [bool]$cfg.apps.Chrome
        $ChkWinrar.IsChecked   = [bool]$cfg.apps.Winrar
        $ChkAdobe.IsChecked    = [bool]$cfg.apps.Adobe
        $ChkAnydesk.IsChecked  = [bool]$cfg.apps.Anydesk
        $ChkTeamview.IsChecked = [bool]$cfg.apps.Teamview
    }

    Write-Log "Perfil aplicado — revise as selecoes e clique em Aplicar." "AVISO"
})

# --- Log ---
$BtnLimparLog.Add_Click({ $LogBox.Items.Clear() })

# ================================================================
# INICIAR
# ================================================================
Write-Log "NextTool v$script:VERSION iniciado em $env:COMPUTERNAME" "INFO"
[void]$Window.ShowDialog()
$LogTimer.Stop()
