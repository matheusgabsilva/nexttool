#Requires -Version 5.1
# ================================================================
# NextTool v4.0 - Ferramenta de TI da Next (GUI)
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
    $scriptPath = $MyInvocation.MyCommand.Path
    if ($scriptPath) {
        Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    } else {
        Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -Command `"irm 'https://raw.githubusercontent.com/matheusgabsilva/nexttool/main/nexttool.ps1' | iex`"" -Verb RunAs
    }
    exit
}

# === OCULTAR JANELA DO CONSOLE (sem relançar — funciona com arquivo e com irm|iex) ===
try {
    Add-Type -Name NativeWindow -Namespace NextTool -MemberDefinition '
        [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    ' -ErrorAction SilentlyContinue
    [NextTool.NativeWindow]::ShowWindow([NextTool.NativeWindow]::GetConsoleWindow(), 0) | Out-Null
} catch {}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ================================================================
# CONFIGURACAO GLOBAL
# ================================================================
$script:VERSION    = "4.0"
$script:REPORT_DIR = "C:\Next-Relatorios"
$script:SESSION_TS = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:LOG_FILE   = Join-Path $script:REPORT_DIR "nexttool_$($env:COMPUTERNAME)_$script:SESSION_TS.log"
$script:LogQueue   = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$script:FolderQueue= [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$script:FileQueue  = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$script:UserQueue  = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

if (-not (Test-Path $script:REPORT_DIR)) {
    New-Item -Path $script:REPORT_DIR -ItemType Directory -Force | Out-Null
}

# ================================================================
# WRITE-LOG
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
    if ($null -ne $LogQueue) { $LogQueue.Enqueue([PSCustomObject]@{ Text = $line; Color = $hex }) }
    try { Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# ================================================================
# TWEAKS
# ================================================================
function Invoke-TweakTelemetria {
    Write-Log "Desativando telemetria..." "STEP"
    $p = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name AllowTelemetry -Value 0 -Type DWord -Force
    @("DiagTrack","dmwappushservice") | ForEach-Object {
        try { Stop-Service $_ -Force -ErrorAction SilentlyContinue; Set-Service $_ -StartupType Disabled -ErrorAction SilentlyContinue } catch {}
    }
    Write-Log "Telemetria desativada." "OK"
}

function Invoke-TweakActivityHistory {
    Write-Log "Desativando historico de atividades..." "STEP"
    $p = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name EnableActivityFeed      -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $p -Name PublishUserActivities   -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $p -Name UploadUserActivities    -Value 0 -Type DWord -Force
    Write-Log "Historico de atividades desativado." "OK"
}

function Invoke-TweakLocationTracking {
    Write-Log "Desativando rastreamento de localizacao..." "STEP"
    $p = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    if (Test-Path $p) { Set-ItemProperty -Path $p -Name Value -Value "Deny" -Force }
    Write-Log "Rastreamento de localizacao desativado." "OK"
}

function Invoke-TweakFileExtensions {
    Write-Log "Exibindo extensoes de arquivo..." "STEP"
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name HideFileExt -Value 0 -Type DWord -Force
    Write-Log "Extensoes de arquivo visiveis." "OK"
}

function Invoke-TweakHiddenFiles {
    Write-Log "Exibindo arquivos ocultos..." "STEP"
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name Hidden -Value 1 -Type DWord -Force
    Write-Log "Arquivos ocultos visiveis." "OK"
}

function Invoke-TweakNumLock {
    Write-Log "Ativando Num Lock na inicializacao..." "STEP"
    Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name InitialKeyboardIndicators -Value 2 -Force
    Write-Log "Num Lock ativado." "OK"
}

function Invoke-TweakEndTask {
    Write-Log "Habilitando Finalizar Tarefa no botao direito..." "STEP"
    $p = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name TaskbarEndTask -Value 1 -Type DWord -Force
    Write-Log "Finalizar Tarefa habilitado." "OK"
}

function Invoke-TweakServices {
    Write-Log "Configurando servicos desnecessarios para Manual..." "STEP"
    @("DiagTrack","dmwappushservice","lfsvc","MapsBroker","RemoteRegistry",
      "TrkWks","WMPNetworkSvc","XblAuthManager","XblGameSave","XboxGipSvc","XboxNetApiSvc") | ForEach-Object {
        try {
            Set-Service -Name $_ -StartupType Manual -ErrorAction SilentlyContinue
            Write-Log " - $_ -> Manual" "PLAIN"
        } catch {}
    }
    Write-Log "Servicos configurados." "OK"
}

function Invoke-TweakHibernacao {
    Write-Log "Desativando hibernacao..." "STEP"
    powercfg /h off 2>&1 | Out-Null
    Write-Log "Hibernacao desativada." "OK"
}

function Invoke-TweakSmartApp {
    Write-Log "Desativando Smart App Control..." "STEP"
    $p = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name VerifiedAndReputablePolicyState -Value 0 -Type DWord -Force
    Write-Log "Smart App Control desativado. Reinicie para aplicar." "OK"
}

function Invoke-TweakUltimatePerf {
    Write-Log "Ativando plano Ultimate Performance..." "STEP"
    $r = powercfg /duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>&1
    if ($r -match "([0-9a-f\-]{36})") {
        powercfg /setactive $Matches[1] | Out-Null
        Write-Log "Plano Ultimate Performance ativado." "OK"
    } else { Write-Log "Plano ja existente ou SO incompativel." "AVISO" }
}

function Invoke-TweakDarkTheme {
    Write-Log "Ativando tema escuro..." "STEP"
    $p = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Set-ItemProperty -Path $p -Name AppsUseLightTheme    -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $p -Name SystemUsesLightTheme -Value 0 -Type DWord -Force
    Write-Log "Tema escuro ativado." "OK"
}

function Invoke-TweakWidgets {
    Write-Log "Desativando Widgets (Win11)..." "STEP"
    $p = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name AllowNewsAndInterests -Value 0 -Type DWord -Force
    Write-Log "Widgets desativados." "OK"
}

function Invoke-TweakVerboseLogon {
    Write-Log "Ativando mensagens detalhadas no logon..." "STEP"
    $p = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-ItemProperty -Path $p -Name VerboseStatus -Value 1 -Type DWord -Force
    Write-Log "Mensagens detalhadas no logon ativadas." "OK"
}

# ================================================================
# MANUTENCAO
# ================================================================
function Invoke-OtimizarPC {
    Write-Log "=== OTIMIZACAO DO PC ===" "STEP"
    Write-Log "Limpando arquivos temporarios..." "STEP"
    $total = 0
    foreach ($p in @($env:TEMP, "C:\Windows\Temp", "C:\Windows\Prefetch")) {
        if (Test-Path $p) {
            $total += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue).Count
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
        $rp = "$regPath\$_"
        if (Test-Path $rp) { Set-ItemProperty -Path $rp -Name StateFlags0064 -Value 2 -ErrorAction SilentlyContinue }
    }
    $cg = Start-Process cleanmgr -ArgumentList "/sagerun:64" -PassThru -WindowStyle Hidden
    if (-not $cg.WaitForExit(60000)) { try { $cg.Kill() } catch {} ; Write-Log "Limpeza de disco encerrada (timeout)." "AVISO" }
    else { Write-Log "Limpeza de disco concluida." "OK" }

    Write-Log "Limpando cache DNS..." "STEP"
    ipconfig /flushdns | Out-Null
    Write-Log "Cache DNS limpo." "OK"
    Write-Log "Otimizacao concluida." "OK"
}

function Invoke-SFCDISM {
    Write-Log "=== SFC + DISM ===" "STEP"
    Write-Log "Executando SFC /scannow (aguarde)..." "STEP"
    $sfc = sfc /scannow 2>&1 | Out-String
    if     ($sfc -match "encontrou|found")              { Write-Log "SFC: problemas encontrados e corrigidos." "OK" }
    elseif ($sfc -match "nao encontrou|did not find")   { Write-Log "SFC: nenhuma violacao encontrada." "OK" }
    else                                                { Write-Log "SFC concluido." "INFO" }
    Write-Log "Executando DISM RestoreHealth (aguarde)..." "STEP"
    $dism = DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
    if ($dism -match "concluida|successfully") { Write-Log "DISM concluido com sucesso." "OK" }
    else { Write-Log "DISM: verifique o resultado manualmente." "AVISO" }
}

function Invoke-CheckDisk {
    param([string]$Drive = "C:")
    Write-Log "=== VERIFICAR DISCO $Drive ===" "STEP"
    Write-Log "Executando chkdsk /scan (sem reinicializacao)..." "INFO"
    $out = & chkdsk.exe $Drive /scan 2>&1 | Out-String
    $out -split "`n" | ForEach-Object { $l = $_.Trim(); if ($l) { Write-Log $l "PLAIN" } }
    Write-Log "Verificacao de disco concluida." "OK"
}

function Invoke-ResetWinsock {
    Write-Log "=== RESET WINSOCK + IP STACK ===" "STEP"
    netsh winsock reset | Out-Null
    Write-Log "Winsock resetado." "OK"
    netsh int ip reset | Out-Null
    Write-Log "IP Stack resetado." "OK"
    ipconfig /flushdns  | Out-Null
    ipconfig /release   | Out-Null
    ipconfig /renew     | Out-Null
    Write-Log "DNS limpo, DHCP renovado." "OK"
    Write-Log "Reinicie o computador para aplicar completamente." "AVISO"
}

function Invoke-LimparCacheWindowsUpdate {
    Write-Log "=== LIMPAR CACHE WINDOWS UPDATE ===" "STEP"
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Stop-Service bits     -Force -ErrorAction SilentlyContinue
    Write-Log "Servicos parados." "INFO"
    $dir = "C:\Windows\SoftwareDistribution\Download"
    if (Test-Path $dir) {
        Remove-Item "$dir\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Cache limpo." "OK"
    }
    Start-Service wuauserv -ErrorAction SilentlyContinue
    Start-Service bits     -ErrorAction SilentlyContinue
    Write-Log "Servicos reiniciados." "OK"
    Write-Log "Cache do Windows Update limpo." "OK"
}

function Invoke-GpUpdate {
    Write-Log "=== GPUPDATE /FORCE ===" "STEP"
    $out = gpupdate /force 2>&1 | Out-String
    $out -split "`n" | ForEach-Object { $l = $_.Trim(); if ($l) { Write-Log $l "PLAIN" } }
    Write-Log "gpupdate concluido." "OK"
}

function Invoke-RestartExplorer {
    Write-Log "Reiniciando Explorer..." "STEP"
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Start-Process explorer
    Write-Log "Explorer reiniciado." "OK"
}

function Invoke-Diagnostico {
    Write-Log "=== DIAGNOSTICO DO SISTEMA ===" "STEP"

    # Hardware
    Write-Log "-- Hardware --" "INFO"
    try {
        $cs  = Get-CimInstance Win32_ComputerSystem
        $cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1)
        $os  = Get-CimInstance Win32_OperatingSystem
        $gpu = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Caption
        $mb  = Get-CimInstance Win32_BaseBoard
        $bios= Get-CimInstance Win32_BIOS
        $ramGB = [math]::Round($cs.TotalPhysicalMemory/1GB,1)
        $freeGB= [math]::Round($os.FreePhysicalMemory/1MB,1)
        Write-Log " CPU  : $($cpu.Name.Trim())  ($($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads)" "PLAIN"
        Write-Log " RAM  : $ramGB GB total  |  $freeGB GB livre" "PLAIN"
        Write-Log " GPU  : $gpu" "PLAIN"
        Write-Log " Mobo : $($mb.Manufacturer) $($mb.Product)" "PLAIN"
        Write-Log " BIOS : $($bios.Manufacturer) v$($bios.SMBIOSBIOSVersion)  ($($bios.ReleaseDate.ToString('dd/MM/yyyy')))" "PLAIN"
    } catch { Write-Log "Falha ao obter hardware: $_" "AVISO" }

    # Seguranca
    Write-Log "-- Seguranca --" "INFO"
    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        $status = if ($def.AntivirusEnabled) { "ATIVO" } else { "INATIVO" }
        Write-Log " Defender: $status  |  Definicoes: $($def.AntivirusSignatureLastUpdated.ToString('dd/MM/yyyy'))" $(if ($def.AntivirusEnabled) {"OK"} else {"ERRO"})
        Write-Log " RTP: $(if ($def.RealTimeProtectionEnabled){'Ativada'}else{'DESATIVADA'})" $(if ($def.RealTimeProtectionEnabled){"OK"} else {"AVISO"})
    } catch { Write-Log " Nao foi possivel verificar o Defender." "AVISO" }
    try {
        # Usa CIM em vez de netsh para nao depender de encoding do runspace
        $fwProfiles = Get-CimInstance -Namespace "root/StandardCimv2" -ClassName MSFT_NetFirewallProfile -ErrorAction Stop
        $on = @($fwProfiles | Where-Object { $_.Enabled -eq $true }).Count
        $fwNames = ($fwProfiles | Where-Object { $_.Enabled -eq $true } | ForEach-Object {
            switch ($_.Name) { "Domain"{"Dominio"} "Private"{"Privado"} "Public"{"Publico"} default{$_.Name} }
        }) -join ", "
        if ($on -gt 0) { Write-Log " Firewall: $on perfil(is) ativo(s) [$fwNames]" "OK" }
        else { Write-Log " Firewall: TODOS OS PERFIS DESATIVADOS" "ERRO" }
    } catch {
        # Fallback netsh
        try {
            $fw = & netsh.exe advfirewall show allprofiles state 2>&1 | Out-String
            $on = ([regex]::Matches($fw,"(?i)(State|Estado)\s+(ON|Ativado)")).Count
            Write-Log " Firewall: $on perfil(is) ativo(s)" $(if ($on -gt 0){"OK"} else {"ERRO"})
        } catch {}
    }
    try {
        $uac = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System").EnableLUA
        Write-Log " UAC: $(if ($uac -eq 1){'Ativado'}else{'DESATIVADO'})" $(if ($uac -eq 1){"OK"} else {"AVISO"})
    } catch {}
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        Write-Log " Secure Boot: $(if ($sb){'Ativado'}else{'Desativado'})" "PLAIN"
    } catch {}
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        if ($tpm.TpmPresent) {
            # SpecVersion ex: "2.0, 0, 1.59" — pega o primeiro segmento
            $tpmVer = try { ((Get-CimInstance -Namespace "root/cimv2/security/microsofttpm" -ClassName Win32_Tpm -ErrorAction Stop).SpecVersion -split ",")[0].Trim() } catch { "" }
            Write-Log " TPM: Presente$(if ($tpmVer){" (spec $tpmVer)"})" "PLAIN"
        } else { Write-Log " TPM: Nao detectado" "PLAIN" }
    } catch {}

    # Processos
    Write-Log "-- Top 10 processos por RAM --" "INFO"
    Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Log (" {0} {1} MB" -f $_.Name.PadRight(28), ([math]::Round($_.WorkingSet64/1MB,1)).ToString().PadLeft(8)) "PLAIN"
    }

    # Inicializacao
    Write-Log "-- Programas na inicializacao --" "INFO"
    @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
      "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run") | ForEach-Object {
        $scope = if ($_ -like "*HKLM*") {"[Sistema]"} else {"[Usuario]"}
        if (Test-Path $_) {
            Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue |
                Get-Member -MemberType NoteProperty |
                Where-Object { $_.Name -notlike "PS*" } |
                ForEach-Object { Write-Log " $scope $($_.Name)" "PLAIN" }
        }
    }

    # Erros criticos
    Write-Log "-- Erros criticos (ultimas 24h) --" "INFO"
    try {
        $evts = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=([datetime]::Now.AddHours(-24))} `
            -MaxEvents 15 -ErrorAction SilentlyContinue
        if ($evts -and $evts.Count -gt 0) {
            Write-Log "$($evts.Count) erro(s) critico(s):" "AVISO"
            $evts | ForEach-Object {
                $msg = $_.Message.Split("`n")[0]
                $msg = $msg.Substring(0,[Math]::Min(100,$msg.Length))
                Write-Log " [$($_.TimeCreated.ToString('HH:mm'))] $($_.ProviderName): $msg" "PLAIN"
            }
        } else { Write-Log "Nenhum erro critico nas ultimas 24h." "OK" }
    } catch { Write-Log "Nao foi possivel ler o Event Viewer." "AVISO" }

    # Bateria
    try {
        $bat = Get-WmiObject Win32_Battery -ErrorAction Stop
        if ($bat) {
            Write-Log "-- Bateria --" "INFO"
            Write-Log " Carga: $($bat.EstimatedChargeRemaining)%  |  Status: $($bat.Status)" "PLAIN"
        }
    } catch {}

    Write-Log "Diagnostico concluido." "OK"
}

function Invoke-TweakDrivers {
    Write-Log "=== ATUALIZACAO DE DRIVERS ===" "STEP"
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log "Instalando PSWindowsUpdate..." "INFO"
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
            Write-Log "Provedor NuGet instalado." "OK"
        } catch { Write-Log "Aviso NuGet: $_" "AVISO" }
        try {
            Install-Module PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers -ErrorAction Stop
            Write-Log "Modulo PSWindowsUpdate instalado." "OK"
        } catch { Write-Log "Falha ao instalar PSWindowsUpdate: $_" "ERRO" }
    }
    try {
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        $updates = Get-WindowsUpdate -Category Drivers -ErrorAction Stop
        if ($updates.Count -eq 0) {
            Write-Log "Windows Update: nenhum driver pendente." "OK"
        } else {
            Write-Log "$($updates.Count) driver(s) encontrado(s). Instalando..." "INFO"
            $updates | ForEach-Object { Write-Log " - $($_.Title)" "PLAIN" }
            Install-WindowsUpdate -Category Drivers -AcceptAll -IgnoreReboot -Verbose 2>&1 |
                ForEach-Object { Write-Log $_ "PLAIN" }
            Write-Log "Drivers instalados. Reinicie para aplicar." "OK"
        }
    } catch { Write-Log "Erro no Windows Update: $_" "ERRO" }

    Write-Log "=== WINGET UPGRADE (todos os pacotes) ===" "STEP"
    $wingetExe = $null
    foreach ($p in @("$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
                     (Resolve-Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe" -ErrorAction SilentlyContinue | Select-Object -First 1))) {
        if ($p -and (Test-Path "$p")) { $wingetExe = "$p"; break }
    }
    if (-not $wingetExe) { $cmd = Get-Command winget -ErrorAction SilentlyContinue; if ($cmd) { $wingetExe = $cmd.Source } }
    if (-not $wingetExe) { Write-Log "winget nao encontrado, pulando upgrade." "AVISO"; return }
    $tmpOut = [System.IO.Path]::GetTempFileName()
    $proc = Start-Process -FilePath $wingetExe `
        -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" `
        -Wait -PassThru -NoNewWindow -RedirectStandardOutput $tmpOut
    $out = Get-Content $tmpOut -Raw -ErrorAction SilentlyContinue
    Remove-Item $tmpOut -ErrorAction SilentlyContinue
    if ($out) { $out -split "`n" | ForEach-Object { $l=$_.Trim(); if ($l) { Write-Log $l "PLAIN" } } }
    Write-Log "winget upgrade concluido (codigo: $($proc.ExitCode))." "OK"
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
    Write-Log "=== ADAPTADORES DE REDE ===" "STEP"
    Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = TRUE" | ForEach-Object {
        $na = Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex = $($_.InterfaceIndex)" -ErrorAction SilentlyContinue
        $nome = if ($na) { $na.NetConnectionID } else { "Desconhecido" }
        Write-Log " $nome" "INFO"
        Write-Log "   IP   : $($_.IPAddress -join ', ')" "PLAIN"
        Write-Log "   Mask : $($_.IPSubnet -join ', ')" "PLAIN"
        Write-Log "   GW   : $($_.DefaultIPGateway -join ', ')" "PLAIN"
        Write-Log "   DNS  : $($_.DNSServerSearchOrder -join ', ')" "PLAIN"
        Write-Log "   MAC  : $($_.MACAddress)" "PLAIN"
        Write-Log "   DHCP : $(if ($_.DHCPEnabled) {'Sim'} else {'Nao'})" "PLAIN"
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

function Invoke-RenovarDHCP {
    param([string]$Adapter)
    Write-Log "Renovando DHCP em '$Adapter'..." "STEP"
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { Write-Log "Adaptador '$Adapter' nao encontrado." "ERRO"; return }
        Invoke-CimMethod -InputObject $cfg -MethodName ReleaseDHCPLease | Out-Null
        Start-Sleep -Seconds 1
        Invoke-CimMethod -InputObject $cfg -MethodName RenewDHCPLease  | Out-Null
        $cfg2 = Get-NicConfig -Adapter $Adapter
        Write-Log "DHCP renovado. Novo IP: $($cfg2.IPAddress | Select-Object -First 1)" "OK"
    } catch { Write-Log "Falha ao renovar DHCP: $_" "ERRO" }
}

function Invoke-TestarConectividade {
    Write-Log "=== TESTE DE CONECTIVIDADE ===" "STEP"
    Write-Log "-- Ping --" "INFO"
    foreach ($t in @("8.8.8.8","1.1.1.1","google.com","microsoft.com")) {
        try {
            $out = & ping.exe -n 1 -w 2000 $t 2>&1 | Out-String
            if ($out -match "[<=](\d+)\s*ms") { Write-Log "Ping ${t}: $($Matches[1])ms" "OK" }
            else { Write-Log "Ping ${t}: sem resposta" "ERRO" }
        } catch { Write-Log "Ping ${t}: erro" "ERRO" }
    }
    Write-Log "-- DNS --" "INFO"
    foreach ($d in @("google.com","microsoft.com","8.8.8.8")) {
        try {
            $ip = [System.Net.Dns]::GetHostAddresses($d) | Select-Object -First 1
            Write-Log "DNS ${d}: $($ip.IPAddressToString)" "OK"
        } catch { Write-Log "DNS ${d}: falha" "ERRO" }
    }
    Write-Log "-- Tracert (10 saltos -> 8.8.8.8) --" "INFO"
    try {
        & tracert.exe -h 10 -w 1000 8.8.8.8 2>&1 | Where-Object { $_ -match "^\s*\d+" } |
            ForEach-Object { Write-Log $_.Trim() "PLAIN" }
    } catch { Write-Log "Tracert: erro — $_" "ERRO" }
    Write-Log "-- Netstat (conexoes estabelecidas) --" "INFO"
    try {
        & netstat.exe -n 2>&1 | Where-Object { $_ -match "ESTABLISHED" } | Select-Object -First 20 |
            ForEach-Object { Write-Log $_.Trim() "PLAIN" }
    } catch {}
    Write-Log "Teste concluido." "OK"
}

function Show-IPConfig {
    Write-Log "=== IPCONFIG /ALL ===" "STEP"
    ipconfig /all 2>&1 | ForEach-Object { if ($_.Trim()) { Write-Log $_ "PLAIN" } }
}

function Invoke-RenomearPC {
    param([string]$NovoNome)
    if (-not $NovoNome) { Write-Log "Informe o novo nome." "ERRO"; return }
    Write-Log "Renomeando PC para '$NovoNome'..." "STEP"
    try {
        Rename-Computer -NewName $NovoNome -Force -ErrorAction Stop
        Write-Log "PC renomeado para '$NovoNome'. Reinicie para aplicar." "OK"
    } catch { Write-Log "Falha ao renomear: $_" "ERRO" }
}

function Invoke-JoinDomain {
    param([string]$Domain, [string]$User, [string]$Pass, [string]$NewName)
    if (-not $Domain -or -not $User -or -not $Pass) { Write-Log "Preencha Dominio, Usuario e Senha." "ERRO"; return }
    Write-Log "Ingressando em $Domain como $User..." "STEP"
    try {
        $cred = New-Object PSCredential("$Domain\$User", (ConvertTo-SecureString $Pass -AsPlainText -Force))
        if ($NewName) { Add-Computer -DomainName $Domain -Credential $cred -NewName $NewName -Force }
        else          { Add-Computer -DomainName $Domain -Credential $cred -Force }
        Write-Log "Ingressado em $Domain. Reinicie para aplicar." "OK"
    } catch { Write-Log "Falha ao ingressar: $_" "ERRO" }
}

# ================================================================
# USUARIOS LOCAIS
# ================================================================
function Get-LocalUsersInfo {
    Write-Log "=== USUARIOS LOCAIS ===" "STEP"
    $admins = (Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue).Name | ForEach-Object { ($_ -split "\\")[-1] }
    $users = Get-LocalUser -ErrorAction SilentlyContinue | ForEach-Object {
        $isAdmin = $admins -contains $_.Name
        [PSCustomObject]@{
            Nome       = $_.Name
            Descricao  = $_.Description
            Ativo      = if ($_.Enabled) { "Sim" } else { "Nao" }
            Admin      = if ($isAdmin)   { "Sim" } else { "Nao" }
            UltimaSenha= if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("dd/MM/yyyy") } else { "Nunca" }
        }
    }
    $UserQueue.Enqueue($users)
    Write-Log "$(@($users).Count) usuario(s) local(is) listado(s)." "OK"
}

function Invoke-CreateUser {
    param([string]$Nome, [string]$Senha, [bool]$IsAdmin)
    if (-not $Nome -or -not $Senha) { Write-Log "Informe nome e senha." "ERRO"; return }
    Write-Log "Criando usuario '$Nome'..." "STEP"
    try {
        $sec = ConvertTo-SecureString $Senha -AsPlainText -Force
        New-LocalUser -Name $Nome -Password $sec -PasswordNeverExpires:$true -ErrorAction Stop
        Write-Log "Usuario '$Nome' criado." "OK"
        if ($IsAdmin) {
            Add-LocalGroupMember -Group "Administrators" -Member $Nome -ErrorAction Stop
            Write-Log "Usuario adicionado ao grupo Administradores." "OK"
        }
    } catch { Write-Log "Falha ao criar usuario: $_" "ERRO" }
}

function Invoke-SetPassword {
    param([string]$Nome, [string]$Senha)
    if (-not $Nome -or -not $Senha) { Write-Log "Informe usuario e nova senha." "ERRO"; return }
    Write-Log "Alterando senha de '$Nome'..." "STEP"
    try {
        $sec = ConvertTo-SecureString $Senha -AsPlainText -Force
        Set-LocalUser -Name $Nome -Password $sec -ErrorAction Stop
        Write-Log "Senha de '$Nome' alterada." "OK"
    } catch { Write-Log "Falha ao alterar senha: $_" "ERRO" }
}

function Invoke-ToggleUser {
    param([string]$Nome)
    if (-not $Nome) { Write-Log "Selecione um usuario." "ERRO"; return }
    try {
        $u = Get-LocalUser -Name $Nome -ErrorAction Stop
        if ($u.Enabled) { Disable-LocalUser -Name $Nome; Write-Log "Usuario '$Nome' desativado." "OK" }
        else            { Enable-LocalUser  -Name $Nome; Write-Log "Usuario '$Nome' ativado."    "OK" }
    } catch { Write-Log "Falha: $_" "ERRO" }
}

function Invoke-AddToAdmins {
    param([string]$Nome)
    if (-not $Nome) { Write-Log "Selecione um usuario." "ERRO"; return }
    try {
        Add-LocalGroupMember -Group "Administrators" -Member $Nome -ErrorAction Stop
        Write-Log "Usuario '$Nome' adicionado ao grupo Administradores." "OK"
    } catch { Write-Log "Falha: $_" "ERRO" }
}

function Invoke-RemoveUser {
    param([string]$Nome)
    if (-not $Nome) { Write-Log "Selecione um usuario." "ERRO"; return }
    try {
        Remove-LocalUser -Name $Nome -ErrorAction Stop
        Write-Log "Usuario '$Nome' removido." "OK"
    } catch { Write-Log "Falha ao remover usuario: $_" "ERRO" }
}

# ================================================================
# ARMAZENAMENTO
# ================================================================
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes/1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes/1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes/1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes/1KB) }
    return "$Bytes B"
}

function Get-FolderSize {
    param([string]$Path)
    try {
        $size = (Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        return [long]($size)
    } catch { return [long]0 }
}

function Invoke-AnalisarPasta {
    param([string]$Path, [int]$MinMB = 50)
    Write-Log "Analisando pasta: $Path ..." "STEP"
    Write-Log "Calculando tamanho das subpastas..." "INFO"
    $subfolders = Get-ChildItem $Path -Directory -Force -ErrorAction SilentlyContinue
    $folderResults = @($subfolders | ForEach-Object {
        $sz = Get-FolderSize -Path $_.FullName
        [PSCustomObject]@{ Nome=$_.Name; Tamanho=Format-Size $sz; Bytes=$sz; Caminho=$_.FullName }
    } | Sort-Object Bytes -Descending)
    $FolderQueue.Enqueue($folderResults)
    Write-Log "$($folderResults.Count) subpasta(s) encontrada(s)." "OK"

    Write-Log "Buscando arquivos acima de ${MinMB}MB..." "INFO"
    $minBytes = $MinMB * 1MB
    $fileResults = @(Get-ChildItem $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -ge $minBytes } | Sort-Object Length -Descending | Select-Object -First 100 |
        ForEach-Object { [PSCustomObject]@{ Nome=$_.Name; Tamanho=Format-Size $_.Length; Bytes=$_.Length; Caminho=$_.FullName } })
    $FileQueue.Enqueue($fileResults)
    Write-Log "$($fileResults.Count) arquivo(s) acima de ${MinMB}MB encontrado(s)." "OK"
    Write-Log "Analise concluida." "OK"
}

# ================================================================
# ASYNC HELPER
# ================================================================
$script:FuncNames = @(
    'Write-Log',
    'Invoke-TweakTelemetria','Invoke-TweakActivityHistory','Invoke-TweakLocationTracking',
    'Invoke-TweakFileExtensions','Invoke-TweakHiddenFiles','Invoke-TweakNumLock',
    'Invoke-TweakEndTask','Invoke-TweakServices','Invoke-TweakHibernacao','Invoke-TweakSmartApp',
    'Invoke-TweakUltimatePerf','Invoke-TweakDarkTheme','Invoke-TweakWidgets','Invoke-TweakVerboseLogon',
    'Invoke-TweakDrivers','Invoke-OtimizarPC','Invoke-SFCDISM','Invoke-CheckDisk',
    'Invoke-ResetWinsock','Invoke-LimparCacheWindowsUpdate','Invoke-GpUpdate','Invoke-RestartExplorer',
    'Invoke-Diagnostico',
    'Get-NicConfig','Show-Adapters','Invoke-SetDNS','Invoke-ResetDNS','Invoke-RenovarDHCP',
    'Invoke-TestarConectividade','Show-IPConfig','Invoke-RenomearPC','Invoke-JoinDomain',
    'Get-LocalUsersInfo','Invoke-CreateUser','Invoke-SetPassword',
    'Invoke-ToggleUser','Invoke-AddToAdmins','Invoke-RemoveUser',
    'Format-Size','Get-FolderSize','Invoke-AnalisarPasta'
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
    $rs.SessionStateProxy.SetVariable("LogQueue",    $script:LogQueue)
    $rs.SessionStateProxy.SetVariable("LOG_FILE",    $script:LOG_FILE)
    $rs.SessionStateProxy.SetVariable("FolderQueue", $script:FolderQueue)
    $rs.SessionStateProxy.SetVariable("FileQueue",   $script:FileQueue)
    $rs.SessionStateProxy.SetVariable("UserQueue",   $script:UserQueue)
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
    Title="NextTool v4.0 - Ferramenta de TI"
    Height="740" Width="1020"
    MinHeight="620" MinWidth="860"
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
            <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Opacity" Value="0.82"/></Trigger>
              <Trigger Property="IsPressed"   Value="True"><Setter Property="Opacity" Value="0.65"/></Trigger>
              <Trigger Property="IsEnabled"   Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
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
            <Border x:Name="TabBorder" Background="{TemplateBinding Background}"
                    BorderThickness="0,0,0,3" BorderBrush="Transparent"
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
    <Style TargetType="ListView">
      <Setter Property="Background" Value="#1E2128"/>
      <Setter Property="Foreground" Value="#ABB2BF"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="Background" Value="#2D3139"/>
      <Setter Property="Foreground" Value="#61AFEF"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="Padding" Value="8,4"/>
    </Style>
    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="14"/>
      <Setter Property="Background" Value="#2D3139"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="#61AFEF"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="54"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="240"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#21252B" BorderBrush="#3E4451" BorderThickness="0,0,0,1">
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Next" FontSize="22" FontWeight="Bold" Foreground="#61AFEF"/>
          <TextBlock Text="Tool" FontSize="22" FontWeight="Bold" Foreground="#ABB2BF"/>
          <TextBlock Text="  v4.0" FontSize="11" Foreground="#5C6370" VerticalAlignment="Bottom" Margin="2,0,0,4"/>
          <TextBlock Text="  |  Ferramenta de TI" FontSize="11" Foreground="#5C6370" VerticalAlignment="Bottom" Margin="0,0,0,4"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,4,0">
          <TextBlock x:Name="HdrPC"     Foreground="#5C6370" FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
          <TextBlock x:Name="HdrUptime" Foreground="#5C6370" FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
          <Button x:Name="BtnRelatorio" Content="📁 Relatórios"
                  Background="#4B5263" Foreground="#ABB2BF" FontSize="11"
                  Padding="10,5" FontWeight="Normal"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- TABS -->
    <TabControl Grid.Row="1" x:Name="MainTabs" Margin="0">

      <!-- ===== SISTEMA ===== -->
      <TabItem Header="  Sistema  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <StackPanel Margin="20,16">

            <!-- Hardware -->
            <GroupBox Header="Hardware">
              <Grid Margin="4,4">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,12,0">
                  <TextBlock Text="COMPUTADOR"  Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiPC"      Foreground="#ABB2BF" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <TextBlock Text="SISTEMA OPERACIONAL" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiOS"      Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="DOMÍNIO / USUÁRIO" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiUser"    Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                  <TextBlock Text="UPTIME" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiUptime"  Foreground="#ABB2BF" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="0,0,12,0">
                  <TextBlock Text="PROCESSADOR" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiCPU"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="MEMÓRIA RAM"  Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiRAM"     Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                  <TextBlock Text="GPU"          Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiGPU"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="PLACA-MÃE / BIOS" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiMobo"    Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="ARMAZENAMENTO" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiDisk"    Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="SEGURANÇA"    Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiSec"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="TPM / SECURE BOOT" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiTpm"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <!-- Acao -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,14">
              <Button x:Name="BtnDiagnostico" Content="🔍  Diagnóstico Completo"
                      Background="#61AFEF" Foreground="#1E2128" Padding="18,10" Margin="0,0,10,0"/>
              <Button x:Name="BtnAtualizarDrivers" Content="⬆  Atualizar Drivers + winget"
                      Background="#98C379" Foreground="#1E2128" Padding="18,10"/>
            </StackPanel>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ===== MANUTENCAO ===== -->
      <TabItem Header="  Manutenção  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <StackPanel Margin="20,16">

            <GroupBox Header="Limpeza e Reparo">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnOtimizar"    Content="🧹 Otimizar PC"           Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnSfcDism"     Content="🔧 SFC + DISM"            Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnCheckDisk"   Content="💽 Verificar Disco (C:)"  Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnLimparWU"    Content="🗑 Limpar Cache WU"       Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnFlushDns"    Content="🌐 Flush DNS"             Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnResetWinsock" Content="🔄 Reset Winsock/IP"     Width="175" Height="64" Margin="0,0,10,10"/>
              </WrapPanel>
            </GroupBox>

            <GroupBox Header="Sistema">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnGpUpdate"       Content="📋 gpupdate /force"      Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnRestartExplorer" Content="🔃 Reiniciar Explorer"  Width="175" Height="64" Margin="0,0,10,10"
                        Background="#E5C07B" Foreground="#1E2128"/>
              </WrapPanel>
            </GroupBox>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ===== TWEAKS ===== -->
      <TabItem Header="  Tweaks  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Border Grid.Row="0" Background="#1E2128" Padding="16,10">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Seleção rápida:" Foreground="#5C6370" VerticalAlignment="Center" Margin="0,0,12,0" FontSize="11"/>
              <Button x:Name="BtnPresetNext"   Content="Padrão Next" Background="#98C379" Foreground="#1E2128" Padding="14,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnPresetLimpar" Content="Limpar"      Background="#4B5263" Foreground="#ABB2BF" Padding="14,5"/>
            </StackPanel>
          </Border>
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <Grid Margin="20,14">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0">
                <GroupBox Header="Essenciais  (Padrão Next)">
                  <StackPanel Margin="4,4">
                    <CheckBox x:Name="ChkTelemetria"       Content="Desativar Telemetria"/>
                    <CheckBox x:Name="ChkActivityHistory"  Content="Desativar Histórico de Atividades"/>
                    <CheckBox x:Name="ChkLocationTracking" Content="Desativar Rastreamento de Localização"/>
                    <CheckBox x:Name="ChkFileExtensions"   Content="Exibir Extensões de Arquivo"/>
                    <CheckBox x:Name="ChkHiddenFiles"      Content="Exibir Arquivos Ocultos"/>
                    <CheckBox x:Name="ChkNumLock"          Content="Num Lock ativo na inicialização"/>
                    <CheckBox x:Name="ChkEndTask"          Content="Finalizar Tarefa no botão direito"/>
                    <CheckBox x:Name="ChkServices"         Content="Serviços desnecessários para Manual"/>
                    <CheckBox x:Name="ChkHibernacao"       Content="Desativar Hibernação"/>
                    <CheckBox x:Name="ChkSmartApp"         Content="Desativar Smart App Control  (Win11)"/>
                  </StackPanel>
                </GroupBox>
              </StackPanel>
              <StackPanel Grid.Column="2">
                <GroupBox Header="Preferências  (opcionais)">
                  <StackPanel Margin="4,4">
                    <CheckBox x:Name="ChkUltimatePerf" Content="Plano Ultimate Performance"/>
                    <CheckBox x:Name="ChkDarkTheme"    Content="Tema Escuro do Windows"/>
                    <CheckBox x:Name="ChkWidgets"      Content="Desativar Widgets  (Win11)"/>
                    <CheckBox x:Name="ChkVerboseLogon" Content="Mensagens detalhadas no Logon"/>
                  </StackPanel>
                </GroupBox>
              </StackPanel>
            </Grid>
          </ScrollViewer>
          <Border Grid.Row="2" Background="#1E2128" Padding="16,10">
            <Button x:Name="BtnAplicarTweaks" Content="Aplicar Tweaks Selecionados"
                    HorizontalAlignment="Left" Background="#E5C07B" Foreground="#1E2128"/>
          </Border>
        </Grid>
      </TabItem>

      <!-- ===== REDE / DOMINIO ===== -->
      <TabItem Header="  Rede / Domínio  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <Grid Margin="20,16">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="24"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

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
                  <Label Content="DNS Primário"/>
                  <TextBox x:Name="TxtDns1" Text="8.8.8.8" Margin="0,0,0,8"/>
                  <Label Content="DNS Secundário"/>
                  <TextBox x:Name="TxtDns2" Text="8.8.4.4" Margin="0,0,0,12"/>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnSetDns"      Content="Configurar DNS"    Margin="0,0,8,0"/>
                    <Button x:Name="BtnResetDns"    Content="Resetar para DHCP" Background="#4B5263" Foreground="#ABB2BF" Margin="0,0,8,0"/>
                    <Button x:Name="BtnRenovarDHCP" Content="Renovar DHCP"      Background="#4B5263" Foreground="#ABB2BF"/>
                  </StackPanel>
                </StackPanel>
              </GroupBox>
            </StackPanel>

            <StackPanel Grid.Column="2">
              <GroupBox Header="Ingressar em Domínio AD">
                <StackPanel Margin="4,4">
                  <Label Content="Domínio  (ex: empresa.local)"/>
                  <TextBox x:Name="TxtDominio" Margin="0,0,0,8"/>
                  <Label Content="Usuário com permissão de join"/>
                  <TextBox x:Name="TxtDomUser" Margin="0,0,0,8"/>
                  <Label Content="Senha"/>
                  <PasswordBox x:Name="TxtDomPass" Margin="0,0,0,8"/>
                  <Label Content="Novo nome do PC  (opcional)"/>
                  <TextBox x:Name="TxtDomName" Margin="0,0,0,14"/>
                  <Button x:Name="BtnJoinDomain" Content="Ingressar no Domínio"
                          HorizontalAlignment="Left" Background="#E06C75" Foreground="#1E2128"/>
                </StackPanel>
              </GroupBox>
              <GroupBox Header="Renomear PC">
                <StackPanel Margin="4,4" Orientation="Horizontal">
                  <TextBox x:Name="TxtNovoNomePC" Width="180" Margin="0,0,10,0"/>
                  <Button x:Name="BtnRenomearPC" Content="Renomear" Background="#E5C07B" Foreground="#1E2128"/>
                </StackPanel>
              </GroupBox>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ===== USUARIOS ===== -->
      <TabItem Header="  Usuários  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Botoes de acao -->
          <Border Grid.Row="0" Background="#1E2128" Padding="14,10">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="BtnListarUsers" Content="🔄 Atualizar Lista"
                      Background="#4B5263" Foreground="#ABB2BF" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnToggleUser" Content="Ativar/Desativar"
                      Background="#E5C07B" Foreground="#1E2128" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnAddToAdmins" Content="→ Administradores"
                      Background="#61AFEF" Foreground="#1E2128" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnRemoveUser" Content="🗑 Remover Usuário"
                      Background="#E06C75" Foreground="#1E2128" Padding="12,5"/>
            </StackPanel>
          </Border>

          <!-- Lista de usuarios -->
          <ListView x:Name="LvUsers" Grid.Row="1" Margin="16,8,16,8">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="Nome"          DisplayMemberBinding="{Binding Nome}"        Width="160"/>
                <GridViewColumn Header="Ativo"         DisplayMemberBinding="{Binding Ativo}"       Width="60"/>
                <GridViewColumn Header="Admin"         DisplayMemberBinding="{Binding Admin}"       Width="60"/>
                <GridViewColumn Header="Última Senha"  DisplayMemberBinding="{Binding UltimaSenha}" Width="110"/>
                <GridViewColumn Header="Descrição"     DisplayMemberBinding="{Binding Descricao}"   Width="260"/>
              </GridView>
            </ListView.View>
          </ListView>

          <!-- Criar usuario / alterar senha -->
          <Border Grid.Row="2" Background="#1E2128" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="24"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0">
                <TextBlock Text="CRIAR USUÁRIO" Foreground="#5C6370" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal">
                  <TextBox x:Name="TxtNewUserName"  Width="140" Margin="0,0,8,0"/>
                  <PasswordBox x:Name="TxtNewUserPass" Width="130" Margin="0,0,8,0"/>
                  <CheckBox x:Name="ChkNewUserAdmin" Content="Admin" Margin="0,0,8,0" VerticalAlignment="Center"/>
                  <Button x:Name="BtnCreateUser" Content="Criar" Background="#98C379" Foreground="#1E2128" Padding="12,5"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2">
                <TextBlock Text="ALTERAR SENHA" Foreground="#5C6370" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal">
                  <TextBox x:Name="TxtChgUser"  Width="140" Margin="0,0,8,0"/>
                  <PasswordBox x:Name="TxtChgPass" Width="160" Margin="0,0,8,0"/>
                  <Button x:Name="BtnSetPassword" Content="Alterar" Background="#E5C07B" Foreground="#1E2128" Padding="12,5"/>
                </StackPanel>
              </StackPanel>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- ===== ARMAZENAMENTO ===== -->
      <TabItem Header="  Armazenamento  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <GroupBox Grid.Row="0" Header="Discos" Margin="16,12,16,0">
            <StackPanel x:Name="DrivePanel" Margin="4,4"/>
          </GroupBox>
          <Border Grid.Row="1" Background="#1E2128" Margin="16,8,16,0" CornerRadius="4" Padding="12,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="TxtAnalyzePath" Grid.Column="0" Text="C:\" Margin="0,0,8,0"/>
              <Button  x:Name="BtnBrowseFolder" Grid.Column="1" Content="..." Width="32" Margin="0,0,8,0"
                       Background="#4B5263" Foreground="#ABB2BF"/>
              <TextBlock Grid.Column="2" Text="Arquivos acima de" Foreground="#5C6370"
                         VerticalAlignment="Center" Margin="0,0,6,0" FontSize="11"/>
              <TextBox x:Name="TxtMinMB" Grid.Column="3" Text="50" Width="50" Margin="0,0,8,0"/>
              <Button  x:Name="BtnAnalisarPasta" Grid.Column="4" Content="Analisar"
                       Background="#61AFEF" Foreground="#1E2128"/>
            </Grid>
          </Border>
          <Grid Grid.Row="2" Margin="16,8,16,12">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <GroupBox Grid.Column="0" Header="Subpastas  (por tamanho)">
              <ListView x:Name="LvFolders">
                <ListView.View>
                  <GridView>
                    <GridViewColumn Header="Pasta"   DisplayMemberBinding="{Binding Nome}"    Width="180"/>
                    <GridViewColumn Header="Tamanho" DisplayMemberBinding="{Binding Tamanho}" Width="90"/>
                    <GridViewColumn Header="Caminho" DisplayMemberBinding="{Binding Caminho}" Width="260"/>
                  </GridView>
                </ListView.View>
              </ListView>
            </GroupBox>
            <GroupBox Grid.Column="2" Header="Maiores arquivos">
              <ListView x:Name="LvFiles">
                <ListView.View>
                  <GridView>
                    <GridViewColumn Header="Arquivo"  DisplayMemberBinding="{Binding Nome}"    Width="180"/>
                    <GridViewColumn Header="Tamanho"  DisplayMemberBinding="{Binding Tamanho}" Width="90"/>
                    <GridViewColumn Header="Caminho"  DisplayMemberBinding="{Binding Caminho}" Width="260"/>
                  </GridView>
                </ListView.View>
              </ListView>
            </GroupBox>
          </Grid>
        </Grid>
      </TabItem>

    </TabControl>

    <!-- LOG PANEL -->
    <Grid Grid.Row="2" Background="#1E2128" MinHeight="220">
      <Grid.RowDefinitions>
        <RowDefinition Height="26"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Grid.Row="0" Background="#21252B" BorderBrush="#3E4451" BorderThickness="0,1,0,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0">
          <TextBlock Text="Log de saída" Foreground="#5C6370" FontSize="11" FontWeight="SemiBold"/>
          <Button x:Name="BtnLimparLog" Content="limpar"
                  Padding="8,1" Margin="14,0,0,0"
                  Background="Transparent" Foreground="#5C6370"
                  FontSize="10" FontWeight="Normal"/>
          <Button x:Name="BtnExportLog" Content="exportar"
                  Padding="8,1" Margin="6,0,0,0"
                  Background="Transparent" Foreground="#5C6370"
                  FontSize="10" FontWeight="Normal"/>
        </StackPanel>
      </Border>
      <ListBox Grid.Row="1" x:Name="LogBox"
               Background="#1E2128" BorderThickness="0" Padding="10,4"
               ScrollViewer.HorizontalScrollBarVisibility="Disabled"
               VirtualizingPanel.IsVirtualizing="True"
               VirtualizingPanel.VirtualizationMode="Recycling"
               FontFamily="Consolas" FontSize="11"/>
    </Grid>

  </Grid>
</Window>
"@

# ================================================================
# CARREGAR JANELA
# ================================================================
$reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Controles
$HdrPC              = $Window.FindName("HdrPC")
$HdrUptime          = $Window.FindName("HdrUptime")
$BtnRelatorio       = $Window.FindName("BtnRelatorio")
$SiPC               = $Window.FindName("SiPC")
$SiOS               = $Window.FindName("SiOS")
$SiUser             = $Window.FindName("SiUser")
$SiUptime           = $Window.FindName("SiUptime")
$SiCPU              = $Window.FindName("SiCPU")
$SiRAM              = $Window.FindName("SiRAM")
$SiGPU              = $Window.FindName("SiGPU")
$SiMobo             = $Window.FindName("SiMobo")
$SiDisk             = $Window.FindName("SiDisk")
$SiSec              = $Window.FindName("SiSec")
$SiTpm              = $Window.FindName("SiTpm")
$BtnDiagnostico     = $Window.FindName("BtnDiagnostico")
$BtnAtualizarDrivers= $Window.FindName("BtnAtualizarDrivers")
$BtnOtimizar        = $Window.FindName("BtnOtimizar")
$BtnSfcDism         = $Window.FindName("BtnSfcDism")
$BtnCheckDisk       = $Window.FindName("BtnCheckDisk")
$BtnLimparWU        = $Window.FindName("BtnLimparWU")
$BtnFlushDns        = $Window.FindName("BtnFlushDns")
$BtnResetWinsock    = $Window.FindName("BtnResetWinsock")
$BtnGpUpdate        = $Window.FindName("BtnGpUpdate")
$BtnRestartExplorer = $Window.FindName("BtnRestartExplorer")
$ChkTelemetria      = $Window.FindName("ChkTelemetria")
$ChkActivityHistory = $Window.FindName("ChkActivityHistory")
$ChkLocationTracking= $Window.FindName("ChkLocationTracking")
$ChkFileExtensions  = $Window.FindName("ChkFileExtensions")
$ChkHiddenFiles     = $Window.FindName("ChkHiddenFiles")
$ChkNumLock         = $Window.FindName("ChkNumLock")
$ChkEndTask         = $Window.FindName("ChkEndTask")
$ChkServices        = $Window.FindName("ChkServices")
$ChkHibernacao      = $Window.FindName("ChkHibernacao")
$ChkSmartApp        = $Window.FindName("ChkSmartApp")
$ChkUltimatePerf    = $Window.FindName("ChkUltimatePerf")
$ChkDarkTheme       = $Window.FindName("ChkDarkTheme")
$ChkWidgets         = $Window.FindName("ChkWidgets")
$ChkVerboseLogon    = $Window.FindName("ChkVerboseLogon")
$BtnPresetNext      = $Window.FindName("BtnPresetNext")
$BtnPresetLimpar    = $Window.FindName("BtnPresetLimpar")
$BtnAplicarTweaks   = $Window.FindName("BtnAplicarTweaks")
$TxtDnsAdapter      = $Window.FindName("TxtDnsAdapter")
$TxtDns1            = $Window.FindName("TxtDns1")
$TxtDns2            = $Window.FindName("TxtDns2")
$BtnSetDns          = $Window.FindName("BtnSetDns")
$BtnResetDns        = $Window.FindName("BtnResetDns")
$BtnRenovarDHCP     = $Window.FindName("BtnRenovarDHCP")
$BtnListarAdapters  = $Window.FindName("BtnListarAdapters")
$BtnTestarConect    = $Window.FindName("BtnTestarConect")
$BtnIPConfig        = $Window.FindName("BtnIPConfig")
$TxtDominio         = $Window.FindName("TxtDominio")
$TxtDomUser         = $Window.FindName("TxtDomUser")
$TxtDomPass         = $Window.FindName("TxtDomPass")
$TxtDomName         = $Window.FindName("TxtDomName")
$BtnJoinDomain      = $Window.FindName("BtnJoinDomain")
$TxtNovoNomePC      = $Window.FindName("TxtNovoNomePC")
$BtnRenomearPC      = $Window.FindName("BtnRenomearPC")
$LvUsers            = $Window.FindName("LvUsers")
$BtnListarUsers     = $Window.FindName("BtnListarUsers")
$BtnToggleUser      = $Window.FindName("BtnToggleUser")
$BtnAddToAdmins     = $Window.FindName("BtnAddToAdmins")
$BtnRemoveUser      = $Window.FindName("BtnRemoveUser")
$TxtNewUserName     = $Window.FindName("TxtNewUserName")
$TxtNewUserPass     = $Window.FindName("TxtNewUserPass")
$ChkNewUserAdmin    = $Window.FindName("ChkNewUserAdmin")
$BtnCreateUser      = $Window.FindName("BtnCreateUser")
$TxtChgUser         = $Window.FindName("TxtChgUser")
$TxtChgPass         = $Window.FindName("TxtChgPass")
$BtnSetPassword     = $Window.FindName("BtnSetPassword")
$DrivePanel         = $Window.FindName("DrivePanel")
$TxtAnalyzePath     = $Window.FindName("TxtAnalyzePath")
$TxtMinMB           = $Window.FindName("TxtMinMB")
$BtnBrowseFolder    = $Window.FindName("BtnBrowseFolder")
$BtnAnalisarPasta   = $Window.FindName("BtnAnalisarPasta")
$LvFolders          = $Window.FindName("LvFolders")
$LvFiles            = $Window.FindName("LvFiles")
$LogBox             = $Window.FindName("LogBox")
$BtnLimparLog       = $Window.FindName("BtnLimparLog")
$BtnExportLog       = $Window.FindName("BtnExportLog")

# ================================================================
# TIMER — drena filas dos runspaces para a UI
# ================================================================
$LogTimer = New-Object System.Windows.Threading.DispatcherTimer
$LogTimer.Interval = [TimeSpan]::FromMilliseconds(120)
$LogTimer.Add_Tick({
    $item = $null; $count = 0
    while ($script:LogQueue.TryDequeue([ref]$item) -and $count -lt 40) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $item.Text
        $tb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($item.Color)
        $tb.TextWrapping = "Wrap"
        [void]$LogBox.Items.Add($tb)
        $count++
    }
    if ($count -gt 0) { $LogBox.ScrollIntoView($LogBox.Items[$LogBox.Items.Count - 1]) }

    $res = $null
    if ($script:FolderQueue.TryDequeue([ref]$res)) {
        $LvFolders.Items.Clear()
        foreach ($r in $res) { [void]$LvFolders.Items.Add($r) }
    }
    if ($script:FileQueue.TryDequeue([ref]$res)) {
        $LvFiles.Items.Clear()
        foreach ($r in $res) { [void]$LvFiles.Items.Add($r) }
    }
    if ($script:UserQueue.TryDequeue([ref]$res)) {
        $LvUsers.Items.Clear()
        foreach ($r in $res) { [void]$LvUsers.Items.Add($r) }
    }
})
$LogTimer.Start()

# ================================================================
# SYSINFO
# ================================================================
try {
    $cs   = Get-CimInstance Win32_ComputerSystem
    $os   = Get-CimInstance Win32_OperatingSystem
    $cpu  = (Get-CimInstance Win32_Processor | Select-Object -First 1)
    $gpu  = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Caption
    $mb   = Get-CimInstance Win32_BaseBoard
    $bios = Get-CimInstance Win32_BIOS
    $ramT = [math]::Round($cs.TotalPhysicalMemory/1GB,1)
    $ramF = [math]::Round($os.FreePhysicalMemory/1MB,1)
    $ramU = [math]::Round($ramT - $ramF, 1)
    $boot = $os.LastBootUpTime
    $up   = (Get-Date) - $boot

    $SiPC.Text    = $env:COMPUTERNAME
    $SiOS.Text    = "$($os.Caption -replace 'Microsoft ','')  (Build $($os.BuildNumber))"
    $SiUser.Text  = "$env:USERNAME  @  $($cs.Domain)"
    $SiUptime.Text= "$($up.Days)d $($up.Hours)h $($up.Minutes)m  (boot: $($boot.ToString('dd/MM HH:mm')))"
    $SiCPU.Text   = "$($cpu.Name.Trim())  ($($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads)"
    $SiRAM.Text   = "Total: ${ramT} GB   |   Usada: ${ramU} GB   |   Livre: ${ramF} GB"
    $SiGPU.Text   = $gpu
    $SiMobo.Text  = "$($mb.Manufacturer) $($mb.Product)  |  BIOS: $($bios.SMBIOSBIOSVersion)"

    $discos = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
        $t = [math]::Round(($_.Used + $_.Free)/1GB,1)
        $f = [math]::Round($_.Free/1GB,1)
        "$($_.Name):  ${t}GB total  |  ${f}GB livre"
    }
    $SiDisk.Text  = ($discos -join "`n")

    $defSt = "Defender: ?"
    try { $def = Get-MpComputerStatus -ErrorAction Stop; $defSt = if ($def.AntivirusEnabled) {"Defender: ATIVO"} else {"Defender: INATIVO"} } catch {}
    $fwSt = "Firewall: ?"
    try {
        $fwProfiles = Get-CimInstance -Namespace "root/StandardCimv2" -ClassName MSFT_NetFirewallProfile -ErrorAction Stop
        $on = @($fwProfiles | Where-Object { $_.Enabled -eq $true }).Count
        $fwSt = "Firewall: $on perfil(is) ativo(s)"
    } catch {
        try { $fwOut = & netsh.exe advfirewall show allprofiles state 2>&1 | Out-String
              $on = ([regex]::Matches($fwOut,"(?i)(State|Estado)\s+(ON|Ativado)")).Count
              $fwSt = "Firewall: $on perfil(is) ativo(s)" } catch {}
    }
    $SiSec.Text   = "$defSt`n$fwSt"

    $tpmSt = "TPM: ?"
    $sbSt  = "Secure Boot: ?"
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        if ($tpm.TpmPresent) {
            $tpmVer = try { ((Get-CimInstance -Namespace "root/cimv2/security/microsofttpm" -ClassName Win32_Tpm -ErrorAction Stop).SpecVersion -split ",")[0].Trim() } catch { "" }
            $tpmSt = "TPM: Presente$(if ($tpmVer){" (spec $tpmVer)"})"
        } else { $tpmSt = "TPM: Nao detectado" }
    } catch {}
    try { $sb = Confirm-SecureBootUEFI -ErrorAction Stop; $sbSt = if ($sb) {"Secure Boot: Ativo"} else {"Secure Boot: Inativo"} } catch {}
    $SiTpm.Text   = "$tpmSt`n$sbSt"

    $HdrPC.Text     = "🖥  $env:COMPUTERNAME"
    $HdrUptime.Text = "⏱  $($up.Days)d $($up.Hours)h $($up.Minutes)m"
} catch {}

# Barras de disco (aba Armazenamento)
try {
    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
        $total   = $_.Used + $_.Free
        $pct     = [math]::Round(($_.Used / $total) * 100)
        $totalGB = [math]::Round($total/1GB,1)
        $usedGB  = [math]::Round($_.Used/1GB,1)
        $freeGB  = [math]::Round($_.Free/1GB,1)
        $color   = if ($pct -ge 90) {"#E06C75"} elseif ($pct -ge 75) {"#E5C07B"} else {"#61AFEF"}

        $row = New-Object System.Windows.Controls.Grid; $row.Margin = "0,0,0,10"
        $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = "60"
        $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = "*"
        $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = "210"
        $row.ColumnDefinitions.Add($c1); $row.ColumnDefinitions.Add($c2); $row.ColumnDefinitions.Add($c3)

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = "$($_.Name):"; $lbl.FontWeight = "SemiBold"; $lbl.VerticalAlignment = "Center"
        $lbl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ABB2BF")
        [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

        $pb = New-Object System.Windows.Controls.ProgressBar
        $pb.Value = $pct; $pb.Maximum = 100; $pb.VerticalAlignment = "Center"; $pb.Margin = "0,0,12,0"
        $pb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($color)
        [System.Windows.Controls.Grid]::SetColumn($pb, 1)

        $info = New-Object System.Windows.Controls.TextBlock
        $info.Text = "${usedGB}GB / ${totalGB}GB  (${freeGB}GB livre — ${pct}%)"
        $info.FontSize = 11; $info.VerticalAlignment = "Center"
        $info.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#5C6370")
        [System.Windows.Controls.Grid]::SetColumn($info, 2)

        $row.Children.Add($lbl) | Out-Null
        $row.Children.Add($pb)  | Out-Null
        $row.Children.Add($info)| Out-Null
        $DrivePanel.Children.Add($row) | Out-Null
    }
} catch {}

# ================================================================
# EVENT HANDLERS
# ================================================================

# --- Sistema ---
$BtnDiagnostico.Add_Click({      Invoke-Async { Invoke-Diagnostico } })
$BtnAtualizarDrivers.Add_Click({ Invoke-Async { Invoke-TweakDrivers } })
$BtnRelatorio.Add_Click({        Start-Process explorer.exe $script:REPORT_DIR })

# --- Manutencao ---
$BtnOtimizar.Add_Click({         Invoke-Async { Invoke-OtimizarPC } })
$BtnSfcDism.Add_Click({          Invoke-Async { Invoke-SFCDISM } })
$BtnCheckDisk.Add_Click({        Invoke-Async { Invoke-CheckDisk "C:" } })
$BtnLimparWU.Add_Click({         Invoke-Async { Invoke-LimparCacheWindowsUpdate } })
$BtnFlushDns.Add_Click({         Invoke-Async { ipconfig /flushdns | Out-Null; Write-Log "Cache DNS limpo." "OK" } })
$BtnResetWinsock.Add_Click({     Invoke-Async { Invoke-ResetWinsock } })
$BtnGpUpdate.Add_Click({         Invoke-Async { Invoke-GpUpdate } })
$BtnRestartExplorer.Add_Click({  Invoke-Async { Invoke-RestartExplorer } })

# --- Tweaks ---
$script:AllTweakChks = @(
    $ChkTelemetria,$ChkActivityHistory,$ChkLocationTracking,
    $ChkFileExtensions,$ChkHiddenFiles,$ChkNumLock,
    $ChkEndTask,$ChkServices,$ChkHibernacao,$ChkSmartApp,
    $ChkUltimatePerf,$ChkDarkTheme,$ChkWidgets,$ChkVerboseLogon
)

$BtnPresetNext.Add_Click({
    $ChkTelemetria.IsChecked       = $true
    $ChkActivityHistory.IsChecked  = $true
    $ChkLocationTracking.IsChecked = $true
    $ChkFileExtensions.IsChecked   = $true
    $ChkHiddenFiles.IsChecked      = $true
    $ChkNumLock.IsChecked          = $true
    $ChkEndTask.IsChecked          = $true
    $ChkServices.IsChecked         = $true
    $ChkHibernacao.IsChecked       = $true
    $ChkSmartApp.IsChecked         = $true
    $ChkUltimatePerf.IsChecked     = $false
    $ChkDarkTheme.IsChecked        = $false
    $ChkWidgets.IsChecked          = $false
    $ChkVerboseLogon.IsChecked     = $false
})

$BtnPresetLimpar.Add_Click({ $script:AllTweakChks | ForEach-Object { $_.IsChecked = $false } })

$BtnAplicarTweaks.Add_Click({
    $v = @{
        Tel=$ChkTelemetria.IsChecked; Act=$ChkActivityHistory.IsChecked
        Loc=$ChkLocationTracking.IsChecked; Ext=$ChkFileExtensions.IsChecked
        Hid=$ChkHiddenFiles.IsChecked; Num=$ChkNumLock.IsChecked
        End=$ChkEndTask.IsChecked; Svc=$ChkServices.IsChecked
        Hib=$ChkHibernacao.IsChecked; Sap=$ChkSmartApp.IsChecked
        Perf=$ChkUltimatePerf.IsChecked; Dark=$ChkDarkTheme.IsChecked
        Wgt=$ChkWidgets.IsChecked; Vrb=$ChkVerboseLogon.IsChecked
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

# --- Rede ---
$BtnListarAdapters.Add_Click({ Invoke-Async { Show-Adapters } })
$BtnTestarConect.Add_Click({   Invoke-Async { Invoke-TestarConectividade } })
$BtnIPConfig.Add_Click({       Invoke-Async { Show-IPConfig } })

$BtnSetDns.Add_Click({
    $ad=$TxtDnsAdapter.Text.Trim(); $d1=$TxtDns1.Text.Trim(); $d2=$TxtDns2.Text.Trim()
    if (-not $ad -or -not $d1) { Write-Log "Preencha Adaptador e DNS Primario." "ERRO"; return }
    Invoke-Async { Invoke-SetDNS -Adapter $A -DNS1 $D1 -DNS2 $D2 } -Vars @{A=$ad;D1=$d1;D2=$d2}
})
$BtnResetDns.Add_Click({
    $ad=$TxtDnsAdapter.Text.Trim()
    if (-not $ad) { Write-Log "Preencha o nome do Adaptador." "ERRO"; return }
    Invoke-Async { Invoke-ResetDNS -Adapter $A } -Vars @{A=$ad}
})
$BtnRenovarDHCP.Add_Click({
    $ad=$TxtDnsAdapter.Text.Trim()
    if (-not $ad) { Write-Log "Preencha o nome do Adaptador." "ERRO"; return }
    Invoke-Async { Invoke-RenovarDHCP -Adapter $A } -Vars @{A=$ad}
})
$BtnJoinDomain.Add_Click({
    $dom=$TxtDominio.Text.Trim(); $usr=$TxtDomUser.Text.Trim()
    $pass=$TxtDomPass.Password;   $name=$TxtDomName.Text.Trim()
    if (-not $dom -or -not $usr -or -not $pass) { Write-Log "Preencha Dominio, Usuario e Senha." "ERRO"; return }
    Invoke-Async { Invoke-JoinDomain -Domain $Dom -User $Usr -Pass $Pass -NewName $Name } `
        -Vars @{Dom=$dom;Usr=$usr;Pass=$pass;Name=$name}
})
$BtnRenomearPC.Add_Click({
    $nome=$TxtNovoNomePC.Text.Trim()
    if (-not $nome) { Write-Log "Informe o novo nome do PC." "ERRO"; return }
    Invoke-Async { Invoke-RenomearPC -NovoNome $N } -Vars @{N=$nome}
})

# --- Usuarios ---
$BtnListarUsers.Add_Click({ Invoke-Async { Get-LocalUsersInfo } })

$BtnToggleUser.Add_Click({
    $sel = $LvUsers.SelectedItem
    if (-not $sel) { Write-Log "Selecione um usuario na lista." "AVISO"; return }
    Invoke-Async { Invoke-ToggleUser -Nome $N } -Vars @{N=$sel.Nome}
})
$BtnAddToAdmins.Add_Click({
    $sel = $LvUsers.SelectedItem
    if (-not $sel) { Write-Log "Selecione um usuario na lista." "AVISO"; return }
    Invoke-Async { Invoke-AddToAdmins -Nome $N } -Vars @{N=$sel.Nome}
})
$BtnRemoveUser.Add_Click({
    $sel = $LvUsers.SelectedItem
    if (-not $sel) { Write-Log "Selecione um usuario na lista." "AVISO"; return }
    Invoke-Async { Invoke-RemoveUser -Nome $N } -Vars @{N=$sel.Nome}
})
$BtnCreateUser.Add_Click({
    $nome=$TxtNewUserName.Text.Trim(); $senha=$TxtNewUserPass.Password
    $isAdm=[bool]$ChkNewUserAdmin.IsChecked
    if (-not $nome -or -not $senha) { Write-Log "Preencha nome e senha." "ERRO"; return }
    Invoke-Async { Invoke-CreateUser -Nome $N -Senha $S -IsAdmin $A } -Vars @{N=$nome;S=$senha;A=$isAdm}
})
$BtnSetPassword.Add_Click({
    $nome=$TxtChgUser.Text.Trim(); $senha=$TxtChgPass.Password
    if (-not $nome -or -not $senha) { Write-Log "Preencha usuario e nova senha." "ERRO"; return }
    Invoke-Async { Invoke-SetPassword -Nome $N -Senha $S } -Vars @{N=$nome;S=$senha}
})

# --- Armazenamento ---
$BtnBrowseFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecione a pasta para analisar"
    $dlg.SelectedPath = $TxtAnalyzePath.Text
    if ($dlg.ShowDialog() -eq "OK") { $TxtAnalyzePath.Text = $dlg.SelectedPath }
})
$BtnAnalisarPasta.Add_Click({
    $path=$TxtAnalyzePath.Text.Trim()
    $minMB=[int]($TxtMinMB.Text -replace "[^0-9]","")
    if (-not $minMB) { $minMB = 50 }
    if (-not (Test-Path $path)) { Write-Log "Pasta nao encontrada: $path" "ERRO"; return }
    $LvFolders.Items.Clear(); $LvFiles.Items.Clear()
    Invoke-Async { Invoke-AnalisarPasta -Path $P -MinMB $M } -Vars @{P=$path;M=$minMB}
})

# --- Log ---
$BtnLimparLog.Add_Click({ $LogBox.Items.Clear() })
$BtnExportLog.Add_Click({
    $dest = "$script:REPORT_DIR\nexttool_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $LogBox.Items | ForEach-Object { $_.Text } | Out-File -FilePath $dest -Encoding UTF8
    Write-Log "Log exportado: $dest" "OK"
})

# ================================================================
# INICIAR
# ================================================================
Write-Log "NextTool v$script:VERSION iniciado em $env:COMPUTERNAME" "INFO"
[void]$Window.ShowDialog()
$LogTimer.Stop()
