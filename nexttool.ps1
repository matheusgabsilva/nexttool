#Requires -Version 5.1
# ================================================================
# NextTool v2.0 - Ferramenta de TI da Next (CLI)
# github.com/matheusgabsilva/nexttool
#
# Uso via URL:
#   irm "https://raw.githubusercontent.com/matheusgabsilva/nexttool/master/nexttool.ps1" | iex
# ================================================================

Set-StrictMode -Off
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8
try { chcp 65001 | Out-Null } catch {}

# === ELEVACAO ===
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -Command `"irm 'https://raw.githubusercontent.com/matheusgabsilva/nexttool/master/nexttool.ps1' | iex`"" -Verb RunAs
    exit
}

# ================================================================
# CONFIGURACAO GLOBAL
# ================================================================
$script:VERSION     = "2.0"
$script:REPORT_DIR  = "C:\Next-Relatorios"
$script:SESSION_TS  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:LOG_FILE    = Join-Path $script:REPORT_DIR "nexttool_$($env:COMPUTERNAME)_$script:SESSION_TS.log"

if (-not (Test-Path $script:REPORT_DIR)) {
    New-Item -Path $script:REPORT_DIR -ItemType Directory -Force | Out-Null
}

# ================================================================
# LOGGING COLORIDO
# ================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","ERRO","AVISO","STEP","PLAIN")]
        [string]$Level = "INFO"
    )
    $ts = [datetime]::Now.ToString("HH:mm:ss")
    switch ($Level) {
        "OK"    { $tag = "[OK]   "; $color = "Green"       }
        "ERRO"  { $tag = "[ERRO] "; $color = "Red"         }
        "AVISO" { $tag = "[AVISO]"; $color = "Yellow"      }
        "STEP"  { $tag = "[>>]   "; $color = "Cyan"        }
        "PLAIN" { $tag = "       "; $color = "Gray"        }
        default { $tag = "[INFO] "; $color = "White"       }
    }
    $line = "$ts $tag $Message"
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $script:LOG_FILE -Value $line -Encoding UTF8 } catch {}
}

function Write-Header {
    param([string]$Title)
    $w = 66
    $pad = [Math]::Max(0, ($w - $Title.Length - 2) / 2)
    $l = " " * [Math]::Floor($pad)
    $r = " " * [Math]::Ceiling($pad)
    Write-Host ""
    Write-Host ("=" * $w) -ForegroundColor DarkCyan
    Write-Host "$l$Title$r" -ForegroundColor Cyan
    Write-Host ("=" * $w) -ForegroundColor DarkCyan
}

function Write-Sep { Write-Host ("-" * 66) -ForegroundColor DarkGray }

function Pause-Enter {
    Write-Host ""
    Write-Host "Pressione ENTER para continuar..." -ForegroundColor DarkGray -NoNewline
    [void][Console]::ReadLine()
}

function Read-Option {
    param([string]$Prompt = "Opcao")
    Write-Host ""
    Write-Host "$Prompt > " -ForegroundColor Yellow -NoNewline
    return ([Console]::ReadLine()).Trim()
}

function Read-Input {
    param([string]$Prompt, [string]$Default = "")
    $suffix = if ($Default) { " [$Default]" } else { "" }
    Write-Host "$Prompt$suffix" -ForegroundColor White -NoNewline
    Write-Host " > " -ForegroundColor Yellow -NoNewline
    $v = ([Console]::ReadLine()).Trim()
    if (-not $v -and $Default) { return $Default }
    return $v
}

function Confirm-Yes {
    param([string]$Prompt)
    Write-Host "$Prompt (s/N) > " -ForegroundColor Yellow -NoNewline
    $r = ([Console]::ReadLine()).Trim().ToLower()
    return ($r -eq "s" -or $r -eq "sim" -or $r -eq "y" -or $r -eq "yes")
}

# ================================================================
# SYS INFO HEADER
# ================================================================
function Show-SysInfo {
    try {
        $os   = Get-CimInstance Win32_OperatingSystem
        $cs   = Get-CimInstance Win32_ComputerSystem
        $cpu  = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name -replace "\s{2,}"," "
        $ramGB  = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        $freeGB = [math]::Round($os.FreePhysicalMemory / 1MB / 1024, 1)
        $usedGB = [math]::Round($ramGB - $freeGB, 1)

        Write-Host ""
        Write-Host "  PC      : " -NoNewline -ForegroundColor DarkGray; Write-Host $env:COMPUTERNAME -ForegroundColor White
        Write-Host "  Usuario : " -NoNewline -ForegroundColor DarkGray; Write-Host "$env:USERNAME  |  Dominio: $($cs.Domain)" -ForegroundColor White
        Write-Host "  SO      : " -NoNewline -ForegroundColor DarkGray; Write-Host "$($os.Caption) $($os.OSArchitecture)" -ForegroundColor White
        Write-Host "  CPU     : " -NoNewline -ForegroundColor DarkGray; Write-Host $cpu -ForegroundColor White
        Write-Host "  RAM     : " -NoNewline -ForegroundColor DarkGray; Write-Host "${ramGB}GB total | ${usedGB}GB usada | ${freeGB}GB livre" -ForegroundColor White

        Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
            $t = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
            $f = [math]::Round($_.Free / 1GB, 1)
            Write-Host "  Disco $($_.Name)  : " -NoNewline -ForegroundColor DarkGray
            Write-Host "${t}GB total | ${f}GB livre" -ForegroundColor White
        }

        try {
            $def = Get-MpComputerStatus -ErrorAction Stop
            $st  = if ($def.AntivirusEnabled) { "ATIVO" } else { "INATIVO" }
            $cl  = if ($def.AntivirusEnabled) { "Green" } else { "Red" }
            Write-Host "  Defender: " -NoNewline -ForegroundColor DarkGray
            Write-Host $st -ForegroundColor $cl
        } catch {}
        try {
            $fwOut = netsh advfirewall show allprofiles state 2>&1 | Out-String
            $onCount = ([regex]::Matches($fwOut, "(?i)State\s+ON")).Count
            $cl = if ($onCount -gt 0) { "Green" } else { "Red" }
            Write-Host "  Firewall: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$onCount perfil(is) ativo(s)" -ForegroundColor $cl
        } catch {}
    } catch {
        Write-Log "Falha ao coletar info do sistema: $_" "AVISO"
    }
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
    Write-Log "Instalando $Name ($Id)..." "STEP"
    $result = winget install --id $Id -e --accept-source-agreements --accept-package-agreements --silent 2>&1
    if ($LASTEXITCODE -eq 0 -or $result -match "instalado|already installed|No applicable") {
        Write-Log "$Name instalado." "OK"
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
    Write-Log "Preparando instalacao do Office $Version via ODT..." "STEP"
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
    $script:OfficeXML[$Version] | Out-File -FilePath $xmlPath -Encoding UTF8
    Write-Log "Iniciando instalacao Office $Version (pode demorar varios minutos)..." "STEP"
    Start-Process $odtExe -ArgumentList "/configure `"$xmlPath`"" -Wait
    Write-Log "Instalacao do Office $Version concluida." "OK"
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
        Write-Log "Nao foi possivel desativar (pode nao existir neste Windows): $_" "AVISO"
    }
}

function Invoke-TweakDrivers {
    Write-Log "Verificando modulo PSWindowsUpdate..." "STEP"
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log "Instalando PSWindowsUpdate..." "INFO"
        try {
            Install-Module PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
            Write-Log "Modulo instalado." "OK"
        } catch {
            Write-Log "Falha ao instalar modulo: $_" "ERRO"
            return
        }
    }
    Import-Module PSWindowsUpdate -Force
    Write-Log "Buscando atualizacoes de drivers..." "INFO"
    try {
        $updates = Get-WindowsUpdate -Category Drivers -ErrorAction Stop
        if ($updates.Count -eq 0) {
            Write-Log "Nenhuma atualizacao de driver disponivel. Drivers em dia." "OK"
        } else {
            Write-Log "$($updates.Count) driver(s) encontrado(s). Instalando..." "INFO"
            $updates | ForEach-Object { Write-Log " - $($_.Title)" "PLAIN" }
            Install-WindowsUpdate -Category Drivers -AcceptAll -IgnoreReboot -Verbose 2>&1 |
                ForEach-Object { Write-Log $_ "PLAIN" }
            Write-Log "Drivers atualizados. Reinicie para aplicar." "OK"
        }
    } catch {
        Write-Log "Erro ao buscar drivers: $_" "ERRO"
    }
}

# ================================================================
# MANUTENCAO
# ================================================================
function Invoke-OtimizarPC {
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

    $totalRam = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
    $freeRam  = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB
    Write-Log ("RAM: Total {0}MB | Livre {1}MB | Usada {2}MB" -f `
        [math]::Round($totalRam), [math]::Round($freeRam), [math]::Round($totalRam - $freeRam)) "INFO"

    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
        $t = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
        $u = [math]::Round($_.Used / 1GB, 1)
        $f = [math]::Round($_.Free / 1GB, 1)
        Write-Log "Disco $($_.Name): ${t}GB total | ${u}GB usado | ${f}GB livre" "INFO"
    }
    Write-Log "Otimizacao concluida." "OK"
}

function Invoke-Diagnostico {
    Write-Log "=== TOP 10 PROCESSOS POR RAM ===" "STEP"
    Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 | ForEach-Object {
        $mem = [math]::Round($_.WorkingSet64 / 1MB, 1)
        Write-Log (" {0} {1} MB" -f $_.Name.PadRight(28), $mem.ToString().PadLeft(8)) "PLAIN"
    }

    Write-Log "=== PROGRAMAS NA INICIALIZACAO ===" "STEP"
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

    Write-Log "=== SEGURANCA ===" "STEP"
    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        if ($def.AntivirusEnabled) {
            Write-Log "Defender ATIVO | Definicoes: $($def.AntivirusSignatureLastUpdated.ToString('dd/MM/yyyy'))" "OK"
        } else {
            Write-Log "Defender INATIVO" "ERRO"
        }
    } catch { Write-Log "Nao foi possivel verificar o Defender." "AVISO" }

    try {
        $fwOut = netsh advfirewall show allprofiles state 2>&1 | Out-String
        $onCount = ([regex]::Matches($fwOut, "(?i)State\s+ON")).Count
        if ($onCount -gt 0) { Write-Log "Firewall: $onCount perfil(is) ativo(s)" "OK" }
        else { Write-Log "Firewall INATIVO" "ERRO" }
    } catch { Write-Log "Nao foi possivel verificar Firewall." "AVISO" }

    Write-Log "=== ERROS CRITICOS (ULTIMAS 24H) ===" "STEP"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=([datetime]::Now.AddHours(-24))} `
            -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($events -and $events.Count -gt 0) {
            Write-Log "$($events.Count) erro(s) critico(s) nas ultimas 24h:" "AVISO"
            $events | ForEach-Object {
                $first = $_.Message.Split("`n")[0]
                $msg = $first.Substring(0, [Math]::Min(90, $first.Length))
                Write-Log " [$($_.TimeCreated.ToString('HH:mm'))] $($_.ProviderName): $msg" "PLAIN"
            }
        } else {
            Write-Log "Nenhum erro critico nas ultimas 24h." "OK"
        }
    } catch { Write-Log "Nao foi possivel verificar Event Viewer." "AVISO" }

    Write-Log "Diagnostico concluido." "OK"
}

function Invoke-SFCDISM {
    Write-Log "Executando SFC /scannow (aguarde)..." "STEP"
    $sfc = sfc /scannow 2>&1 | Out-String
    if ($sfc -match "encontrou|found") {
        Write-Log "SFC: problemas encontrados e corrigidos." "OK"
    } elseif ($sfc -match "nao encontrou|did not find") {
        Write-Log "SFC: nenhuma violacao de integridade encontrada." "OK"
    } else {
        Write-Log "SFC concluido." "INFO"
    }
    Write-Log "Executando DISM RestoreHealth (aguarde)..." "STEP"
    $dism = DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
    if ($dism -match "concluida|successfully") {
        Write-Log "DISM concluido com sucesso." "OK"
    } else {
        Write-Log "DISM: verifique o resultado manualmente." "AVISO"
    }
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
    Write-Log "Configurando DNS em '$Adapter' -> $DNS1 / $DNS2 ..." "STEP"
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { Write-Log "Adaptador '$Adapter' nao encontrado." "ERRO"; return }
        $dns = @($DNS1); if ($DNS2) { $dns += $DNS2 }
        $r = Invoke-CimMethod -InputObject $cfg -MethodName SetDNSServerSearchOrder `
             -Arguments @{ DNSServerSearchOrder = [string[]]$dns }
        if ($r.ReturnValue -eq 0) { Write-Log "DNS configurado." "OK" }
        else { Write-Log "Concluido com codigo WMI: $($r.ReturnValue)" "AVISO" }
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
        if ($r.ReturnValue -eq 0) { Write-Log "DNS resetado para DHCP automatico." "OK" }
        else { Write-Log "Concluido com codigo WMI: $($r.ReturnValue)" "AVISO" }
    } catch { Write-Log $_ "ERRO" }
}

function Invoke-TestarConectividade {
    foreach ($target in @("8.8.8.8","1.1.1.1","google.com")) {
        $ping = New-Object System.Net.NetworkInformation.Ping
        try {
            $reply = $ping.Send($target, 3000)
            if ($reply.Status -eq "Success") {
                Write-Log "Ping ${target}: $($reply.RoundtripTime)ms" "OK"
            } else {
                Write-Log "Ping ${target}: $($reply.Status)" "ERRO"
            }
        } catch {
            Write-Log "Ping ${target} falhou: $_" "ERRO"
        }
    }
}

function Show-IPConfig {
    Write-Log "IPConfig /all:" "STEP"
    ipconfig /all 2>&1 | ForEach-Object {
        if ($_.Trim()) { Write-Host "   $_" -ForegroundColor Gray }
    }
}

function Invoke-JoinDomain {
    param([string]$Domain,[string]$User,[string]$Pass,[string]$NewName)
    if (-not $Domain -or -not $User -or -not $Pass) {
        Write-Log "Preencha Dominio, Usuario e Senha." "ERRO"
        return
    }
    Write-Log "Ingressando em $Domain como $User..." "STEP"
    try {
        $cred = New-Object PSCredential("$Domain\$User", (ConvertTo-SecureString $Pass -AsPlainText -Force))
        if ($NewName) {
            Add-Computer -DomainName $Domain -Credential $cred -NewName $NewName -Force
            Write-Log "PC renomeado para '$NewName' e ingressado em $Domain." "OK"
        } else {
            Add-Computer -DomainName $Domain -Credential $cred -Force
            Write-Log "Ingressado em $Domain com sucesso." "OK"
        }
        Write-Log "Reinicie o computador para aplicar as alteracoes." "AVISO"
    } catch {
        Write-Log "Falha ao ingressar no dominio: $_" "ERRO"
    }
}

# ================================================================
# MENUS
# ================================================================
function Show-MainMenu {
    Clear-Host
    Write-Host ""
    Write-Host "  _   _           _  _____           _" -ForegroundColor Cyan
    Write-Host " | \ | | _____  _| ||_   _|__   ___ | |" -ForegroundColor Cyan
    Write-Host " |  \| |/ _ \ \/ / __|| |/ _ \ / _ \| |" -ForegroundColor Cyan
    Write-Host " | |\  |  __/>  <| |_ | | (_) | (_) | |" -ForegroundColor Cyan
    Write-Host " |_| \_|\___/_/\_\\__||_|\___/ \___/|_|" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "        Ferramenta de TI da Next  v$script:VERSION" -ForegroundColor DarkCyan
    Show-SysInfo
    Write-Sep
    Write-Host "  [1] Instalacoes de Software" -ForegroundColor White
    Write-Host "  [2] Tweaks do Sistema"       -ForegroundColor White
    Write-Host "  [3] Manutencao"              -ForegroundColor White
    Write-Host "  [4] Rede / Dominio"          -ForegroundColor White
    Write-Host "  [5] Abrir pasta de relatorios ($script:REPORT_DIR)" -ForegroundColor White
    Write-Host "  [0] Sair" -ForegroundColor DarkGray
    Write-Sep
}

function Install-PadraoNext {
    Write-Log "Instalando padrao Next: Chrome, WinRAR, Adobe Reader, AnyDesk, TeamViewer..." "STEP"
    if (-not (Test-Winget)) { Install-Winget }
    Install-WingetApp "Google.Chrome"                "Google Chrome"
    Install-WingetApp "RARLab.WinRAR"                "WinRAR"
    Install-WingetApp "Adobe.Acrobat.Reader.64-bit"  "Adobe Acrobat Reader"
    Install-WingetApp "AnyDesk.AnyDesk"              "AnyDesk"
    Install-WingetApp "TeamViewer.TeamViewer"         "TeamViewer"
    Write-Log "Padrao Next instalado. Escolha a versao do Office separadamente." "OK"
}

function Menu-Instalacoes {
    while ($true) {
        Clear-Host
        Write-Header "INSTALACOES DE SOFTWARE"
        Write-Host "  --- Padrao Next ---" -ForegroundColor DarkCyan
        Write-Host "  [P] Instalar PADRAO NEXT (Chrome + WinRAR + Adobe Reader + AnyDesk + TeamViewer)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  --- Programas individuais ---" -ForegroundColor DarkCyan
        Write-Host "  [1] Google Chrome              (winget)" -ForegroundColor White
        Write-Host "  [2] WinRAR                     (winget)" -ForegroundColor White
        Write-Host "  [3] AnyDesk                    (winget)" -ForegroundColor White
        Write-Host "  [4] TeamViewer                 (winget)" -ForegroundColor White
        Write-Host "  [8] Adobe Acrobat Reader DC    (winget)" -ForegroundColor White
        Write-Host ""
        Write-Host "  --- Microsoft Office (ODT) ---" -ForegroundColor DarkCyan
        Write-Host "  [5] Microsoft 365              (ODT)"    -ForegroundColor White
        Write-Host "  [6] Office 2021 Pro Plus VL    (ODT)"    -ForegroundColor White
        Write-Host "  [7] Office 2016 Pro Plus VL    (ODT)"    -ForegroundColor White
        Write-Sep
        Write-Host "  [0] Voltar" -ForegroundColor DarkGray
        $op = Read-Option
        switch ($op) {
            "P" { Install-PadraoNext; Pause-Enter }
            "p" { Install-PadraoNext; Pause-Enter }
            "1" { if (-not (Test-Winget)) { Install-Winget }; Install-WingetApp "Google.Chrome"         "Google Chrome"; Pause-Enter }
            "2" { if (-not (Test-Winget)) { Install-Winget }; Install-WingetApp "RARLab.WinRAR"         "WinRAR"; Pause-Enter }
            "3" { if (-not (Test-Winget)) { Install-Winget }; Install-WingetApp "AnyDesk.AnyDesk"       "AnyDesk"; Pause-Enter }
            "4" { if (-not (Test-Winget)) { Install-Winget }; Install-WingetApp "TeamViewer.TeamViewer" "TeamViewer"; Pause-Enter }
            "8" { if (-not (Test-Winget)) { Install-Winget }; Install-WingetApp "Adobe.Acrobat.Reader.64-bit" "Adobe Acrobat Reader"; Pause-Enter }
            "5" { Install-Office "365"; Pause-Enter }
            "6" { Install-Office "2021"; Pause-Enter }
            "7" { Install-Office "2016"; Pause-Enter }
            "0" { return }
            default { Write-Log "Opcao invalida." "AVISO"; Start-Sleep -Milliseconds 600 }
        }
    }
}

function Menu-Tweaks {
    while ($true) {
        Clear-Host
        Write-Header "TWEAKS DO SISTEMA"
        Write-Host "  [1] Desativar hibernacao (powercfg -h off)" -ForegroundColor White
        Write-Host "  [2] Desativar Smart App Control (Win11)"    -ForegroundColor White
        Write-Host "  [3] Atualizar drivers via Windows Update"   -ForegroundColor White
        Write-Host "  [4] Aplicar TODOS"                          -ForegroundColor White
        Write-Sep
        Write-Host "  [0] Voltar" -ForegroundColor DarkGray
        $op = Read-Option
        switch ($op) {
            "1" { Invoke-TweakHibernacao; Pause-Enter }
            "2" { Invoke-TweakSmartApp; Pause-Enter }
            "3" { Invoke-TweakDrivers; Pause-Enter }
            "4" { Invoke-TweakHibernacao; Invoke-TweakSmartApp; Invoke-TweakDrivers; Pause-Enter }
            "0" { return }
            default { Write-Log "Opcao invalida." "AVISO"; Start-Sleep -Milliseconds 600 }
        }
    }
}

function Menu-Manutencao {
    while ($true) {
        Clear-Host
        Write-Header "MANUTENCAO"
        Write-Host "  [1] Otimizar PC (temp, lixeira, cleanmgr, DNS)" -ForegroundColor White
        Write-Host "  [2] Diagnostico completo"                       -ForegroundColor White
        Write-Host "  [3] SFC + DISM"                                  -ForegroundColor White
        Write-Host "  [4] Flush DNS"                                   -ForegroundColor White
        Write-Host "  [5] Abrir pasta de relatorios"                   -ForegroundColor White
        Write-Sep
        Write-Host "  [0] Voltar" -ForegroundColor DarkGray
        $op = Read-Option
        switch ($op) {
            "1" { Invoke-OtimizarPC; Pause-Enter }
            "2" { Invoke-Diagnostico; Pause-Enter }
            "3" { Invoke-SFCDISM; Pause-Enter }
            "4" { ipconfig /flushdns | Out-Null; Write-Log "Cache DNS limpo." "OK"; Pause-Enter }
            "5" { Start-Process explorer.exe $script:REPORT_DIR }
            "0" { return }
            default { Write-Log "Opcao invalida." "AVISO"; Start-Sleep -Milliseconds 600 }
        }
    }
}

function Menu-Rede {
    while ($true) {
        Clear-Host
        Write-Header "REDE / DOMINIO"
        Write-Host "  [1] Listar adaptadores"               -ForegroundColor White
        Write-Host "  [2] Configurar DNS manualmente"       -ForegroundColor White
        Write-Host "  [3] Resetar DNS para DHCP"            -ForegroundColor White
        Write-Host "  [4] Testar conectividade (ping)"      -ForegroundColor White
        Write-Host "  [5] Exibir ipconfig /all"             -ForegroundColor White
        Write-Host "  [6] Ingressar em dominio AD"          -ForegroundColor White
        Write-Sep
        Write-Host "  [0] Voltar" -ForegroundColor DarkGray
        $op = Read-Option
        switch ($op) {
            "1" { Show-Adapters; Pause-Enter }
            "2" {
                Show-Adapters
                $ad   = Read-Input "Nome do adaptador (ex: Ethernet, Wi-Fi)"
                $d1   = Read-Input "DNS primario" "8.8.8.8"
                $d2   = Read-Input "DNS secundario (vazio para nenhum)" "8.8.4.4"
                Invoke-SetDNS -Adapter $ad -DNS1 $d1 -DNS2 $d2
                Pause-Enter
            }
            "3" {
                Show-Adapters
                $ad = Read-Input "Nome do adaptador"
                Invoke-ResetDNS -Adapter $ad
                Pause-Enter
            }
            "4" { Invoke-TestarConectividade; Pause-Enter }
            "5" { Show-IPConfig; Pause-Enter }
            "6" {
                Write-Log "Ingresso em dominio - reinicio sera necessario." "AVISO"
                $dom   = Read-Input "Dominio (ex: empresa.local)"
                $user  = Read-Input "Usuario com permissao de join"
                Write-Host "Senha > " -ForegroundColor Yellow -NoNewline
                $secure = Read-Host -AsSecureString
                $pass   = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                          [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
                $name  = Read-Input "Novo nome do PC (vazio = manter atual)"
                if (Confirm-Yes "Confirmar ingresso em '$dom' como '$user'?") {
                    Invoke-JoinDomain -Domain $dom -User $user -Pass $pass -NewName $name
                }
                Pause-Enter
            }
            "0" { return }
            default { Write-Log "Opcao invalida." "AVISO"; Start-Sleep -Milliseconds 600 }
        }
    }
}

# ================================================================
# LOOP PRINCIPAL
# ================================================================
Write-Log "NextTool v$script:VERSION iniciado em $env:COMPUTERNAME (log: $script:LOG_FILE)" "INFO"

while ($true) {
    Show-MainMenu
    $op = Read-Option
    switch ($op) {
        "1" { Menu-Instalacoes }
        "2" { Menu-Tweaks }
        "3" { Menu-Manutencao }
        "4" { Menu-Rede }
        "5" { Start-Process explorer.exe $script:REPORT_DIR }
        "0" { Write-Log "Encerrando NextTool." "INFO"; break }
        default { Write-Log "Opcao invalida." "AVISO"; Start-Sleep -Milliseconds 600 }
    }
    if ($op -eq "0") { break }
}
