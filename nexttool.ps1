#Requires -Version 5.1
# ================================================================
# NextTool v4.1 - Ferramenta de TI da Next (GUI)
# github.com/matheusgabsilva/nexttool
#
# Uso via URL:
#   irm nexttool.matheusgabsilva.digital | iex
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
        Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -Command `"irm nexttool.matheusgabsilva.digital | iex`"" -Verb RunAs
    }
    exit
}

# === OCULTAR JANELA DO CONSOLE (sem relançar - funciona com arquivo e com irm|iex) ===
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
$script:VERSION      = "4.1"
$script:REPORT_DIR   = "C:\Next-Relatorios"
$script:SENHA_PADRAO = "next@2025"          # Senha padrao — altere aqui antes de distribuir
$script:REG_PATH     = "HKLM:\SOFTWARE\NextTool"
$script:MODE         = "USER"               # USER ou ADMIN
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
# SENHA / MODO ADM
# ================================================================
function Get-SenhaHash {
    param([string]$Texto)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Texto)
    $hash  = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ($hash | ForEach-Object { $_.ToString("x2") }) -join ""
}

function Initialize-Senha {
    if (-not (Test-Path $script:REG_PATH)) {
        New-Item -Path $script:REG_PATH -Force | Out-Null
    }
    $existing = (Get-ItemProperty -Path $script:REG_PATH -Name "AdminHash" -ErrorAction SilentlyContinue).AdminHash
    if (-not $existing) {
        $hash = Get-SenhaHash $script:SENHA_PADRAO
        Set-ItemProperty -Path $script:REG_PATH -Name "AdminHash" -Value $hash -Force
    }
}

function Test-SenhaAdm {
    param([string]$Tentativa)
    $stored = (Get-ItemProperty -Path $script:REG_PATH -Name "AdminHash" -ErrorAction SilentlyContinue).AdminHash
    if (-not $stored) { return $false }
    return ((Get-SenhaHash $Tentativa) -eq $stored)
}

function Set-SenhaAdm {
    param([string]$NovaSenha)
    if ($NovaSenha.Length -lt 4) { Write-Log "Senha deve ter ao menos 4 caracteres." "AVISO"; return }
    $hash = Get-SenhaHash $NovaSenha
    if (-not (Test-Path $script:REG_PATH)) { New-Item -Path $script:REG_PATH -Force | Out-Null }
    Set-ItemProperty -Path $script:REG_PATH -Name "AdminHash" -Value $hash -Force
    Write-Log "Senha ADM alterada com sucesso." "OK"
}

function Show-DialogSenha {
    Add-Type -AssemblyName PresentationFramework
    $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Acesso ADM" Height="180" Width="340"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#282C34" FontFamily="Segoe UI" FontSize="13">
  <StackPanel Margin="24,20">
    <TextBlock Text="Digite a senha ADM:" Foreground="#ABB2BF" Margin="0,0,0,10"/>
    <PasswordBox x:Name="PbSenha" Background="#1E2128" Foreground="#ABB2BF"
                 BorderBrush="#3E4451" BorderThickness="1" Padding="8,6" Margin="0,0,0,14"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnOk"      Content="Entrar"   Width="80" Height="30" Margin="0,0,8,0"
              Background="#61AFEF" Foreground="#1E2128" FontWeight="SemiBold"/>
      <Button x:Name="BtnCancel" Content="Cancelar" Width="80" Height="30"
              Background="#4B5263" Foreground="#ABB2BF"/>
    </StackPanel>
  </StackPanel>
</Window>
"@
    [xml]$dlgXml = $dlgXaml
    $dlgReader   = New-Object System.Xml.XmlNodeReader $dlgXml
    $dlg         = [Windows.Markup.XamlReader]::Load($dlgReader)
    $pb          = $dlg.FindName("PbSenha")
    $dlg.FindName("BtnOk").Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $dlg.FindName("BtnCancel").Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $pb.Add_KeyDown({ if ($_.Key -eq "Return") { $dlg.DialogResult = $true; $dlg.Close() } })
    $dlg.ShowDialog() | Out-Null
    if ($dlg.DialogResult -eq $true) { return $pb.Password } else { return $null }
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

function Invoke-TweakSuspender {
    Write-Log "Desativando suspender (sleep)..." "STEP"
    powercfg /change standby-timeout-ac 0 | Out-Null   # plugado
    powercfg /change standby-timeout-dc 0 | Out-Null   # bateria
    Write-Log "Suspender desativado (AC e DC)." "OK"
}

function Invoke-TweakTela {
    Write-Log "Desativando desligamento de tela..." "STEP"
    powercfg /change monitor-timeout-ac 0 | Out-Null   # plugado
    powercfg /change monitor-timeout-dc 0 | Out-Null   # bateria
    Write-Log "Desligamento de tela desativado (AC e DC)." "OK"
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

function Invoke-ReinstalarImpressoras {
    Write-Log "=== REINSTALAR DRIVERS DE IMPRESSORA ===" "STEP"
    $impressoras = Get-Printer -ErrorAction SilentlyContinue
    if (-not $impressoras) { Write-Log "Nenhuma impressora encontrada." "AVISO"; return }
    foreach ($imp in $impressoras) {
        Write-Log "Removendo: $($imp.Name)..." "INFO"
        Remove-Printer -Name $imp.Name -ErrorAction SilentlyContinue
    }
    Write-Log "Impressoras removidas. Redescobindo via PnP..." "INFO"
    pnputil /scan-devices | Out-Null
    Start-Sleep -Seconds 3
    $novas = Get-Printer -ErrorAction SilentlyContinue
    if ($novas) {
        $novas | ForEach-Object { Write-Log "Encontrada: $($_.Name)" "OK" }
    } else {
        Write-Log "Nenhuma impressora redescoberta automaticamente. Reconecte o dispositivo." "AVISO"
    }
    Write-Log "Reinstalacao de impressoras concluida." "OK"
}

# ================================================================
# ALTA PRIORIDADE
# ================================================================
function Invoke-PingVisual {
    param([string]$Destino = "8.8.8.8", [int]$Count = 4)
    Write-Log "=== PING: $Destino ===" "STEP"
    $resultados = Test-Connection -ComputerName $Destino -Count $Count -ErrorAction SilentlyContinue
    if ($resultados) {
        $resultados | ForEach-Object {
            $ms = $_.ResponseTime
            $level = if ($ms -lt 50) {"OK"} elseif ($ms -lt 150) {"AVISO"} else {"ERRO"}
            Write-Log "  Resposta de $($_.Address): ${ms}ms" $level
        }
        $avg = [math]::Round(($resultados | Measure-Object ResponseTime -Average).Average, 1)
        $lost = $Count - $resultados.Count
        Write-Log "Enviados: $Count  |  Perdidos: $lost  |  Latencia media: ${avg}ms" "OK"
    } else {
        Write-Log "Sem resposta de $Destino - host inacessivel ou sem rede." "ERRO"
    }
}

function Invoke-PingContinuo {
    param([string]$Destino = "8.8.8.8", [int]$Count = 20)
    Write-Log "=== PING CONTINUO: $Destino ($Count pacotes) ===" "STEP"
    $ok = 0; $fail = 0
    for ($i = 1; $i -le $Count; $i++) {
        $r = Test-Connection -ComputerName $Destino -Count 1 -ErrorAction SilentlyContinue
        if ($r) {
            $ms = $r.ResponseTime
            $ok++
            $level = if ($ms -lt 50) {"OK"} elseif ($ms -lt 150) {"AVISO"} else {"ERRO"}
            Write-Log "  [$i/$Count] ${ms}ms" $level
        } else {
            $fail++
            Write-Log "  [$i/$Count] Timeout" "ERRO"
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Log "Concluido - OK: $ok  |  Falhas: $fail  |  Perda: $([math]::Round($fail/$Count*100))%" $(if ($fail -eq 0){"OK"} else {"AVISO"})
}

function Invoke-TracertVisual {
    param([string]$Destino = "8.8.8.8")
    Write-Log "=== TRACERT: $Destino ===" "STEP"
    try {
        $saida = & tracert.exe -h 20 -w 1000 $Destino 2>&1
        $saida | ForEach-Object {
            $line = $_.ToString().Trim()
            if ($line) { Write-Log "  $line" "PLAIN" }
        }
        Write-Log "Tracert concluido." "OK"
    } catch { Write-Log "Erro no tracert: $_" "ERRO" }
}

function Export-RelatorioHTML {
    param([string]$ReportDir = "C:\Next-Relatorios")
    Write-Log "=== GERANDO RELATORIO HTML ===" "STEP"
    try {
        if (-not (Test-Path $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null }
        $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
        $file = "$ReportDir\relatorio_$($env:COMPUTERNAME)_$ts.html"

        $cs   = Get-CimInstance Win32_ComputerSystem
        $os   = Get-CimInstance Win32_OperatingSystem
        $cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
        $gpu  = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Caption
        $mb   = Get-CimInstance Win32_BaseBoard
        $bios = Get-CimInstance Win32_BIOS
        $ramT = [math]::Round($cs.TotalPhysicalMemory/1GB,1)
        $boot = $os.LastBootUpTime
        $up   = (Get-Date) - $boot
        $discos = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
            $t = [math]::Round(($_.Used + $_.Free)/1GB,1)
            $f = [math]::Round($_.Free/1GB,1)
            "<tr><td>$($_.Name):</td><td>${t} GB total</td><td>${f} GB livre</td></tr>"
        }
        $nic = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" | Select-Object -First 1
        $ip  = if ($nic.IPAddress) { $nic.IPAddress[0] } else { "N/A" }
        $dns = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder -join ", " } else { "N/A" }
        $gw  = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway[0] } else { "N/A" }

        $defSt = "N/A"
        try { $def = Get-MpComputerStatus -ErrorAction Stop; $defSt = if ($def.AntivirusEnabled) {"ATIVO"} else {"INATIVO"} } catch {}

        $dominio = if ($cs.PartOfDomain) { "Dominio: $($cs.Domain)" } else { "Grupo de trabalho: $($cs.Domain)" }
        $licenca = try { (Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND LicenseStatus=1 AND Name LIKE 'Windows%'" -ErrorAction Stop | Select-Object -First 1).Name } catch { "N/A" }

        $html = @"
<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>NextTool - Relatorio $($env:COMPUTERNAME)</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#1e2128;color:#abb2bf;margin:0;padding:20px}
h1{color:#61afef;border-bottom:2px solid #61afef;padding-bottom:8px}
h2{color:#98c379;margin-top:24px;font-size:14px;text-transform:uppercase;letter-spacing:1px}
table{width:100%;border-collapse:collapse;margin-bottom:16px}
td,th{padding:6px 12px;border:1px solid #3e4451;font-size:13px}
th{background:#21252b;color:#61afef;text-align:left}
tr:nth-child(even){background:#21252b}
.badge{display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;font-weight:bold}
.ok{background:#2d4a3e;color:#98c379}.er{background:#4a2d2d;color:#e06c75}
.info{color:#5c6370;font-size:11px}
</style></head><body>
<h1>NextTool - Relatorio do Sistema</h1>
<p class="info">Gerado em: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")  |  Operador: $env:USERNAME</p>
<h2>Computador</h2>
<table><tr><th>Campo</th><th>Valor</th></tr>
<tr><td>Nome</td><td>$($env:COMPUTERNAME)</td></tr>
<tr><td>Dominio / Grupo</td><td>$dominio</td></tr>
<tr><td>Sistema Operacional</td><td>$($os.Caption) (Build $($os.BuildNumber))</td></tr>
<tr><td>Licenca Windows</td><td>$licenca</td></tr>
<tr><td>Uptime</td><td>$($up.Days)d $($up.Hours)h $($up.Minutes)m</td></tr>
<tr><td>Ultimo boot</td><td>$($boot.ToString('dd/MM/yyyy HH:mm'))</td></tr>
</table>
<h2>Hardware</h2>
<table><tr><th>Campo</th><th>Valor</th></tr>
<tr><td>Processador</td><td>$($cpu.Name.Trim())</td></tr>
<tr><td>Nucleos</td><td>$($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads</td></tr>
<tr><td>RAM Total</td><td>$ramT GB</td></tr>
<tr><td>GPU</td><td>$gpu</td></tr>
<tr><td>Placa-mae</td><td>$($mb.Manufacturer) $($mb.Product)</td></tr>
<tr><td>BIOS</td><td>$($bios.Manufacturer) v$($bios.SMBIOSBIOSVersion)</td></tr>
</table>
<h2>Armazenamento</h2>
<table><tr><th>Drive</th><th>Total</th><th>Livre</th></tr>$($discos -join "")</table>
<h2>Rede</h2>
<table><tr><th>Campo</th><th>Valor</th></tr>
<tr><td>Adaptador</td><td>$($nic.Description)</td></tr>
<tr><td>Endereco IP</td><td>$ip</td></tr>
<tr><td>Gateway</td><td>$gw</td></tr>
<tr><td>DNS</td><td>$dns</td></tr>
</table>
<h2>Seguranca</h2>
<table><tr><th>Item</th><th>Status</th></tr>
<tr><td>Windows Defender</td><td><span class="badge $(if($defSt -eq 'ATIVO'){'ok'}else{'er'})">$defSt</span></td></tr>
</table>
</body></html>
"@
        $html | Out-File -FilePath $file -Encoding UTF8
        Write-Log "Relatorio gerado: $file" "OK"
        Start-Process $file
    } catch { Write-Log "Erro ao gerar relatorio: $_" "ERRO" }
}

function Invoke-VerificarDominio {
    Write-Log "=== STATUS DE DOMINIO ===" "STEP"
    $cs = Get-CimInstance Win32_ComputerSystem
    if ($cs.PartOfDomain) {
        Write-Log "PC no dominio: $($cs.Domain)" "OK"
        try {
            $dc = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().DomainControllers | Select-Object -First 1
            Write-Log "DC encontrado: $($dc.Name)" "OK"
        } catch { Write-Log "Nao foi possivel contatar o DC." "AVISO" }
    } else {
        Write-Log "PC nao esta em dominio. Grupo de trabalho: $($cs.Domain)" "AVISO"
    }
    $site = try { [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name } catch { "N/A" }
    if ($site -ne "N/A") { Write-Log "Site AD: $site" "INFO" }
}

function Invoke-DesbloquearUsuario {
    param([string]$NomeUsuario)
    if (-not $NomeUsuario) { Write-Log "Selecione um usuario na lista primeiro." "AVISO"; return }
    Write-Log "Desbloqueando usuario: $NomeUsuario..." "STEP"
    try {
        Unlock-LocalUser -Name $NomeUsuario -ErrorAction Stop
        Write-Log "Usuario '$NomeUsuario' desbloqueado com sucesso." "OK"
    } catch { Write-Log "Erro ao desbloquear '$NomeUsuario': $_" "ERRO" }
}

function Show-IPInfo {
    Write-Log "=== INFORMACOES DE REDE ===" "STEP"
    Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" | ForEach-Object {
        Write-Log "Adaptador : $($_.Description)" "INFO"
        if ($_.IPAddress)            { Write-Log " IP       : $($_.IPAddress[0])" "PLAIN" }
        if ($_.IPSubnet)             { Write-Log " Mascara  : $($_.IPSubnet[0])" "PLAIN" }
        if ($_.DefaultIPGateway)     { Write-Log " Gateway  : $($_.DefaultIPGateway[0])" "PLAIN" }
        if ($_.DNSServerSearchOrder) { Write-Log " DNS      : $($_.DNSServerSearchOrder -join ' | ')" "PLAIN" }
        if ($_.MACAddress)           { Write-Log " MAC      : $($_.MACAddress)" "PLAIN" }
        Write-Log "" "PLAIN"
    }
}

# ================================================================
# MEDIA PRIORIDADE
# ================================================================
function Invoke-HistoricoLogs {
    param([string]$ReportDir = "C:\Next-Relatorios")
    Write-Log "=== HISTORICO DE SESSOES ===" "STEP"
    $logs = Get-ChildItem -Path $ReportDir -Filter "nexttool_*.log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 10
    if (-not $logs) { Write-Log "Nenhum log anterior encontrado." "INFO"; return }
    Write-Log "$($logs.Count) sessao(oes) encontrada(s):" "OK"
    $logs | ForEach-Object {
        $size = [math]::Round($_.Length/1KB, 1)
        Write-Log "  $($_.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))  |  $($_.Name)  ($size KB)" "PLAIN"
    }
    Start-Process explorer.exe $ReportDir
}

function Invoke-VerificarUpdates {
    Write-Log "=== VERIFICAR ATUALIZACOES PENDENTES ===" "STEP"
    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        Write-Log "Consultando Windows Update (aguarde)..." "INFO"
        $result = $searcher.Search("IsInstalled=0 and Type='Software'")
        if ($result.Updates.Count -eq 0) {
            Write-Log "Nenhuma atualizacao pendente." "OK"
        } else {
            Write-Log "$($result.Updates.Count) atualizacao(oes) pendente(s):" "AVISO"
            $result.Updates | Select-Object -First 15 | ForEach-Object {
                $size = if ($_.MaxDownloadSize -gt 0) { " ($([math]::Round($_.MaxDownloadSize/1MB,1)) MB)" } else { "" }
                Write-Log "  - $($_.Title)$size" "PLAIN"
            }
            if ($result.Updates.Count -gt 15) {
                Write-Log "  ... e mais $($result.Updates.Count - 15) atualizacao(oes)." "PLAIN"
            }
        }
    } catch { Write-Log "Erro ao verificar updates: $_" "ERRO" }
}

function Show-ServiciosCriticos {
    Write-Log "=== STATUS DE SERVICOS CRITICOS ===" "STEP"
    $servicos = @{
        "Spooler"       = "Fila de Impressao"
        "wuauserv"      = "Windows Update"
        "W32Time"       = "Hora do Windows"
        "LanmanServer"  = "Compartilhamento (Server)"
        "Netlogon"      = "Logon de Rede"
        "BITS"          = "BITS (Downloads)"
        "Dnscache"      = "Cache DNS"
        "RpcSs"         = "RPC"
        "EventLog"      = "Log de Eventos"
        "WinDefend"     = "Windows Defender"
    }
    foreach ($svc in $servicos.GetEnumerator()) {
        try {
            $s = Get-Service -Name $svc.Key -ErrorAction Stop
            $status = $s.Status
            $level  = if ($status -eq "Running") {"OK"} else {"AVISO"}
            Write-Log ("  {0,-28} [{1}]" -f $svc.Value, $status) $level
        } catch {
            Write-Log ("  {0,-28} [NAO ENCONTRADO]" -f $svc.Value) "PLAIN"
        }
    }
}

function Invoke-ReiniciarServico {
    param([string]$NomeServico)
    if (-not $NomeServico) { Write-Log "Informe o nome do servico." "AVISO"; return }
    Write-Log "Reiniciando servico: $NomeServico..." "STEP"
    try {
        Restart-Service -Name $NomeServico -Force -ErrorAction Stop
        Write-Log "Servico '$NomeServico' reiniciado com sucesso." "OK"
    } catch { Write-Log "Erro ao reiniciar '$NomeServico': $_" "ERRO" }
}

function Invoke-LimparPerfilCorrompido {
    Write-Log "=== LIMPAR PERFIS CORROMPIDOS ===" "STEP"
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $perfis  = Get-ChildItem $regPath -ErrorAction SilentlyContinue
    $encontrados = 0
    foreach ($p in $perfis) {
        $caminho = (Get-ItemProperty $p.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        $isBak   = $p.PSChildName -match "\.bak$"
        $isTemp  = $caminho -match "TEMP$|\.TEMP$"
        if ($isBak -or $isTemp) {
            Write-Log "Perfil corrompido encontrado: $caminho" "AVISO"
            try {
                Remove-Item $p.PSPath -Recurse -Force -ErrorAction Stop
                Write-Log "Entrada removida do registro: $($p.PSChildName)" "OK"
                $encontrados++
            } catch { Write-Log "Erro ao remover perfil: $_" "ERRO" }
        }
    }
    if ($encontrados -eq 0) {
        Write-Log "Nenhum perfil corrompido ou temporario encontrado." "OK"
    } else {
        Write-Log "$encontrados perfil(is) removido(s). Reinicie o computador." "OK"
    }
}

function Invoke-DesinstalarApp {
    param([string]$NomeApp)
    if (-not $NomeApp) { Write-Log "Informe o nome do aplicativo." "AVISO"; return }
    Write-Log "=== DESINSTALAR: $NomeApp ===" "STEP"
    $wingetExe = $null
    foreach ($p in @("$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
                     "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe")) {
        $found = Resolve-Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { $wingetExe = $found.Path; break }
    }
    if (-not $wingetExe) { Write-Log "winget nao encontrado." "ERRO"; return }
    Write-Log "Procurando '$NomeApp' via winget..." "INFO"
    $proc = Start-Process $wingetExe -ArgumentList "uninstall --name `"$NomeApp`" --silent --accept-source-agreements" `
            -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\wg_out.txt" -ErrorAction SilentlyContinue
    $out = Get-Content "$env:TEMP\wg_out.txt" -ErrorAction SilentlyContinue
    if ($out) { $out | ForEach-Object { if ($_) { Write-Log "  $_" "PLAIN" } } }
    if ($proc.ExitCode -eq 0) {
        Write-Log "Aplicativo '$NomeApp' desinstalado com sucesso." "OK"
    } else {
        Write-Log "Desinstalacao retornou codigo $($proc.ExitCode). Verifique o nome do app." "AVISO"
    }
}

function Invoke-RenovarIP {
    Write-Log "=== RENOVAR ENDERECO IP (DHCP) ===" "STEP"
    Write-Log "Liberando IP atual..." "INFO"
    ipconfig /release 2>&1 | Out-Null
    Write-Log "Solicitando novo IP..." "INFO"
    ipconfig /renew 2>&1 | Out-Null
    ipconfig /flushdns 2>&1 | Out-Null
    $ip = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex (
        Get-NetRoute -DestinationPrefix "0.0.0.0/0" | Sort-Object RouteMetric | Select-Object -First 1 -ExpandProperty InterfaceIndex
    ) -ErrorAction SilentlyContinue).IPAddress
    if ($ip) { Write-Log "Novo IP obtido: $ip" "OK" }
    Write-Log "IP renovado com sucesso." "OK"
}

function Invoke-ResetarProxy {
    Write-Log "=== RESETAR CONFIGURACOES DE PROXY ===" "STEP"
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    Set-ItemProperty -Path $path -Name ProxyEnable      -Value 0 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $path -Name ProxyServer       -Value "" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $path -Name ProxyOverride     -Value "" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $path -Name AutoConfigURL     -Value "" -ErrorAction SilentlyContinue
    netsh winhttp reset proxy 2>&1 | Out-Null
    Write-Log "Proxy do WinInet desativado." "OK"
    Write-Log "Proxy do WinHTTP resetado." "OK"
    Write-Log "Configuracoes de proxy limpas com sucesso." "OK"
}

function Invoke-SincronizarHora {
    Write-Log "=== SINCRONIZAR HORA DO SISTEMA ===" "STEP"
    $svc = Get-Service -Name W32Time -ErrorAction SilentlyContinue
    if (-not $svc) { Write-Log "Servico W32Time nao encontrado." "ERRO"; return }
    if ($svc.Status -ne 'Running') {
        Start-Service W32Time -ErrorAction SilentlyContinue
        Write-Log "Servico W32Time iniciado." "INFO"
    }
    w32tm /resync /force 2>&1 | ForEach-Object { if ($_) { Write-Log $_ "INFO" } }
    $hora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    Write-Log "Hora atual do sistema: $hora" "OK"
    Write-Log "Sincronizacao de hora concluida." "OK"
}

function Invoke-LimparMiniaturas {
    Write-Log "=== LIMPAR CACHE DE MINIATURAS ===" "STEP"
    $dir = "$env:LocalAppData\Microsoft\Windows\Explorer"
    $arquivos = Get-ChildItem -Path $dir -Filter "thumbcache_*.db" -ErrorAction SilentlyContinue
    if (-not $arquivos) { Write-Log "Cache de miniaturas ja estava vazio." "INFO"; return }
    $total = $arquivos.Count
    $bytes = ($arquivos | Measure-Object -Property Length -Sum).Sum
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
    $removidos = 0
    foreach ($f in $arquivos) {
        Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path $f.FullName)) { $removidos++ }
    }
    Start-Process explorer.exe
    $mb = [math]::Round($bytes / 1MB, 2)
    Write-Log "$removidos de $total arquivo(s) removido(s) ($mb MB liberados)." "OK"
    Write-Log "Cache de miniaturas limpo. Explorer reiniciado." "OK"
}

function Invoke-LimparCredenciais {
    Write-Log "=== LIMPAR CREDENCIAIS SALVAS ===" "STEP"
    $creds = cmdkey /list 2>&1 | Select-String "Destino:" | ForEach-Object { ($_ -split "Destino:\s*")[1].Trim() }
    if (-not $creds) {
        # Tentar em ingles
        $creds = cmdkey /list 2>&1 | Select-String "Target:" | ForEach-Object { ($_ -split "Target:\s*")[1].Trim() }
    }
    if (-not $creds) { Write-Log "Nenhuma credencial salva encontrada." "INFO"; return }
    $count = 0
    foreach ($c in $creds) {
        if ($c) {
            cmdkey /delete:$c 2>&1 | Out-Null
            Write-Log "Removida: $c" "INFO"
            $count++
        }
    }
    Write-Log "$count credencial(is) removida(s) com sucesso." "OK"
}

function Invoke-LimparCacheTeamsOffice {
    Write-Log "=== LIMPAR CACHE DO TEAMS E OFFICE ===" "STEP"
    $caminhos = @{
        "Teams (novo)"   = "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
        "Teams (classico)" = "$env:AppData\Microsoft\Teams\Cache"
        "Teams tmp"      = "$env:AppData\Microsoft\Teams\blob_storage"
        "Teams databases"= "$env:AppData\Microsoft\Teams\databases"
        "Office cache"   = "$env:LocalAppData\Microsoft\Office\16.0\OfficeFileCache"
        "Outlook cache"  = "$env:LocalAppData\Microsoft\Outlook\RoamCache"
    }
    $totalMB = 0
    foreach ($item in $caminhos.GetEnumerator()) {
        if (Test-Path $item.Value) {
            $size = (Get-ChildItem $item.Value -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            Remove-Item "$($item.Value)\*" -Recurse -Force -ErrorAction SilentlyContinue
            $mb = [math]::Round($size / 1MB, 2)
            $totalMB += $mb
            Write-Log "$($item.Key): $mb MB liberados." "OK"
        } else {
            Write-Log "$($item.Key): nao encontrado." "INFO"
        }
    }
    Write-Log "Total liberado: $([math]::Round($totalMB, 2)) MB." "OK"
    Write-Log "Cache do Teams/Office limpo. Reinicie os apps se necessario." "OK"
}

function Invoke-LimparSpooler {
    Write-Log "=== LIMPAR FILA DE IMPRESSAO ===" "STEP"
    Write-Log "Parando servico Spooler..." "INFO"
    Stop-Service Spooler -Force -ErrorAction SilentlyContinue
    $dir = "$env:SystemRoot\System32\spool\PRINTERS"
    $arquivos = Get-ChildItem -Path $dir -Recurse -ErrorAction SilentlyContinue
    if ($arquivos) {
        Remove-Item "$dir\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "$($arquivos.Count) arquivo(s) removido(s) da fila." "OK"
    } else {
        Write-Log "Fila ja estava vazia." "INFO"
    }
    Start-Service Spooler -ErrorAction SilentlyContinue
    Write-Log "Servico Spooler reiniciado." "OK"
    Write-Log "Fila de impressao limpa com sucesso." "OK"
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
            # SpecVersion ex: "2.0, 0, 1.59" - pega o primeiro segmento
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
    } catch { Write-Log "Tracert: erro - $_" "ERRO" }
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
# AREA DE TRABALHO
# ================================================================
$script:DesktopIconGuids = [ordered]@{
    "Meu Computador"      = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    "Arquivos do Usuario" = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
    "Rede"                = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
    "Lixeira"             = "{645FF040-5081-101B-9F08-00AA002F954E}"
    "Painel de Controle"  = "{26EE0668-A00A-44D7-9371-BEB064C98683}"
}

function Get-DesktopIconState {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    $result = [ordered]@{}
    foreach ($name in $DesktopIconGuids.Keys) {
        $guid = $DesktopIconGuids[$name]
        $val  = (Get-ItemProperty -Path $regPath -Name $guid -ErrorAction SilentlyContinue).$guid
        $result[$name] = ($val -ne 1)   # 0 ou ausente = visivel
    }
    return $result
}

function Set-DesktopIconState {
    param([string]$IconName, [bool]$Visible)
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    $guid = $DesktopIconGuids[$IconName]
    if (-not $guid) { Write-Log "Icone '$IconName' nao encontrado." "ERRO"; return }
    Set-ItemProperty -Path $regPath -Name $guid -Value ([int](-not $Visible)) -Type DWord -Force
    # Notifica o Shell para atualizar icones imediatamente
    try {
        Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class ShellNotify{[DllImport("shell32.dll")]public static extern void SHChangeNotify(int e,uint f,IntPtr a,IntPtr b);}' -ErrorAction SilentlyContinue
        [ShellNotify]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    } catch {}
    Write-Log "Icone '$IconName': $(if ($Visible){'visivel'}else{'oculto'})." "OK"
}

function Get-InstalledApps {
    $dirs = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    )

    # Padroes de itens a excluir (ferramentas de sistema, docs, desinstaladores)
    $excludePattern = @(
        'desinstalar','uninstall','uninst',
        'manual do console','release notes','notas de instala',
        'what''s new','o que h[aá] de novo',
        'module docs','manuals','documentation','ajuda$','help$',
        'sobre o ','about ',
        'log de telemetria',
        'configurar java','java [\d]+ update',
        '\(x86\)$','\(32.bit\)$'
    ) -join '|'

    # Coleta todos os .lnk
    $raw = foreach ($dir in $dirs) {
        if (-not (Test-Path $dir)) { continue }
        Get-ChildItem -Path $dir -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object {
                [PSCustomObject]@{
                    Nome   = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                    Origem = if ($dir -like "*ProgramData*") {"Sistema"} else {"Usuario"}
                    Caminho= $_.FullName
                }
            }
    }

    # Filtra ruido
    $filtered = $raw | Where-Object { $_.Nome -notmatch $excludePattern }

    # Remove duplicatas: para cada nome base (sem sufixos x64/32-bit),
    # prefere 64-bit sobre 32-bit e, empate, mantém o primeiro
    $seen    = @{}
    $result  = foreach ($app in ($filtered | Sort-Object Nome)) {
        # Chave normalizada: remove sufixos de arquitetura para comparacao
        $key = $app.Nome -replace '\s*\(x64\)$','' `
                         -replace '\s*\(64.bit\)$','' `
                         -replace '\s*64$','' `
                         -replace '\s*ISE$','' |
               ForEach-Object { $_.Trim().ToLower() }

        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $app
        } else {
            # Prefere versao sem sufixo de arquitetura (mais limpa)
            $existingLen = $seen[$key].Nome.Length
            if ($app.Nome.Length -lt $existingLen) { $seen[$key] = $app }
        }
    }
    return @($seen.Values | Sort-Object Nome)
}

function New-DesktopShortcut {
    param([string]$LnkSource)
    $desktop = [Environment]::GetFolderPath("Desktop")
    $nome    = [System.IO.Path]::GetFileName($LnkSource)
    $dest    = Join-Path $desktop $nome
    try {
        Copy-Item -Path $LnkSource -Destination $dest -Force -ErrorAction Stop
        Write-Log "Atalho criado: $([System.IO.Path]::GetFileNameWithoutExtension($nome))" "OK"
    } catch { Write-Log "Falha ao criar atalho '$nome': $_" "ERRO" }
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
    'Invoke-TweakSuspender','Invoke-TweakTela',
    'Invoke-TweakUltimatePerf','Invoke-TweakDarkTheme','Invoke-TweakWidgets','Invoke-TweakVerboseLogon',
    'Invoke-TweakDrivers','Invoke-OtimizarPC','Invoke-SFCDISM','Invoke-CheckDisk',
    'Invoke-ResetWinsock','Invoke-LimparCacheWindowsUpdate','Invoke-LimparSpooler','Invoke-ReinstalarImpressoras',
    'Invoke-RenovarIP','Invoke-ResetarProxy','Invoke-SincronizarHora',
    'Invoke-LimparMiniaturas','Invoke-LimparCredenciais','Invoke-LimparCacheTeamsOffice',
    'Invoke-GpUpdate','Invoke-RestartExplorer',
    'Invoke-PingVisual','Invoke-PingContinuo','Invoke-TracertVisual','Export-RelatorioHTML',
    'Invoke-VerificarDominio','Invoke-DesbloquearUsuario','Show-IPInfo',
    'Invoke-HistoricoLogs','Invoke-VerificarUpdates','Show-ServiciosCriticos',
    'Invoke-ReiniciarServico','Invoke-LimparPerfilCorrompido','Invoke-DesinstalarApp',
    'Invoke-Diagnostico',
    'Get-NicConfig','Show-Adapters','Invoke-SetDNS','Invoke-ResetDNS','Invoke-RenovarDHCP',
    'Invoke-TestarConectividade','Show-IPConfig','Invoke-RenomearPC','Invoke-JoinDomain',
    'Get-LocalUsersInfo','Invoke-CreateUser','Invoke-SetPassword',
    'Invoke-ToggleUser','Invoke-AddToAdmins','Invoke-RemoveUser',
    'Format-Size','Get-FolderSize','Invoke-AnalisarPasta',
    'Get-DesktopIconState','Set-DesktopIconState','Get-InstalledApps','New-DesktopShortcut'
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
    $rs.SessionStateProxy.SetVariable("FolderQueue",       $script:FolderQueue)
    $rs.SessionStateProxy.SetVariable("FileQueue",         $script:FileQueue)
    $rs.SessionStateProxy.SetVariable("DesktopIconGuids",  $script:DesktopIconGuids)
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
    Title="NextTool v4.1"
    Height="740" Width="1020"
    MinHeight="620" MinWidth="860"
    WindowStartupLocation="CenterScreen"
    Background="#282C34"
    FontFamily="Segoe UI"
    FontSize="13">

  <Window.Resources>

    <!-- ── BUTTON ─────────────────────────────────────────────── -->
    <Style TargetType="Button">
      <Setter Property="Background"        Value="#61AFEF"/>
      <Setter Property="Foreground"        Value="#1E2128"/>
      <Setter Property="FontWeight"        Value="SemiBold"/>
      <Setter Property="Padding"           Value="14,7"/>
      <Setter Property="BorderThickness"   Value="0"/>
      <Setter Property="Cursor"            Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="BtnBorder" Background="{TemplateBinding Background}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <Grid>
                <Border x:Name="HoverOverlay" Background="White" CornerRadius="6" Opacity="0"/>
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <EventTrigger RoutedEvent="MouseEnter">
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="HoverOverlay"
                                     Storyboard.TargetProperty="Opacity"
                                     To="0.15" Duration="0:0:0.12"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <EventTrigger RoutedEvent="MouseLeave">
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="HoverOverlay"
                                     Storyboard.TargetProperty="Opacity"
                                     To="0" Duration="0:0:0.12"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="HoverOverlay" Property="Opacity" Value="0.28"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── CHECKBOX ────────────────────────────────────────────── -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#ABB2BF"/>
      <Setter Property="Margin"     Value="0,5"/>
      <Setter Property="Cursor"     Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="ChkBox" Width="16" Height="16" CornerRadius="4"
                      BorderBrush="#3E4451" BorderThickness="1.5"
                      Background="Transparent" VerticalAlignment="Center" Margin="0,0,8,0">
                <Path x:Name="ChkMark" Visibility="Collapsed"
                      Data="M3,8 L7,12 L14,4"
                      Stroke="#1E2128" StrokeThickness="1.8"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <ContentPresenter VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="ChkBox"  Property="Background"  Value="#61AFEF"/>
                <Setter TargetName="ChkBox"  Property="BorderBrush" Value="#61AFEF"/>
                <Setter TargetName="ChkMark" Property="Visibility"  Value="Visible"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ChkBox" Property="BorderBrush" Value="#61AFEF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── TEXTBOX ─────────────────────────────────────────────── -->
    <Style TargetType="TextBox">
      <Setter Property="Background"      Value="#1E2128"/>
      <Setter Property="Foreground"      Value="#ABB2BF"/>
      <Setter Property="CaretBrush"      Value="#ABB2BF"/>
      <Setter Property="BorderBrush"     Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="8,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="TxtBorder"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ScrollViewer x:Name="PART_ContentHost" Focusable="False"
                            HorizontalScrollBarVisibility="Hidden"
                            VerticalScrollBarVisibility="Hidden"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="TxtBorder" Property="BorderBrush" Value="#61AFEF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── PASSWORDBOX ─────────────────────────────────────────── -->
    <Style TargetType="PasswordBox">
      <Setter Property="Background"      Value="#1E2128"/>
      <Setter Property="Foreground"      Value="#ABB2BF"/>
      <Setter Property="BorderBrush"     Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="8,6"/>
    </Style>

    <!-- ── GROUPBOX ────────────────────────────────────────────── -->
    <Style TargetType="GroupBox">
      <Setter Property="Margin"  Value="0,0,0,16"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GroupBox">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E4451" BorderThickness="0,0,0,1">
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                  <Border Width="3" CornerRadius="2" Background="#61AFEF" Margin="0,2,10,2"/>
                  <ContentPresenter ContentSource="Header"
                                    TextElement.Foreground="#ABB2BF"
                                    TextElement.FontWeight="SemiBold"
                                    TextElement.FontSize="12"
                                    VerticalAlignment="Center"/>
                </StackPanel>
              </Border>
              <ContentPresenter Grid.Row="1" Margin="0,10,0,0"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── TABCONTROL / TABITEM ────────────────────────────────── -->
    <Style TargetType="TabControl">
      <Setter Property="Background"      Value="#21252B"/>
      <Setter Property="BorderBrush"     Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="0,1,0,0"/>
      <Setter Property="Padding"         Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="Foreground"      Value="#5C6370"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Padding"         Value="18,11"/>
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
                <Setter Property="Foreground" Value="#ABB2BF"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#ABB2BF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── LABEL / SEPARATOR / PROGRESSBAR ────────────────────── -->
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#5C6370"/>
      <Setter Property="Padding"    Value="0,0,0,2"/>
      <Setter Property="FontSize"   Value="11"/>
    </Style>
    <Style TargetType="Separator">
      <Setter Property="Background" Value="#3E4451"/>
      <Setter Property="Margin"     Value="0,8"/>
    </Style>
    <Style TargetType="ProgressBar">
      <Setter Property="Height"          Value="4"/>
      <Setter Property="Background"      Value="#2D3139"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground"      Value="#61AFEF"/>
    </Style>

    <!-- ── LISTVIEW / LISTVIEWITEM ─────────────────────────────── -->
    <Style TargetType="ListView">
      <Setter Property="Background"      Value="#1E2128"/>
      <Setter Property="Foreground"      Value="#ABB2BF"/>
      <Setter Property="BorderBrush"     Value="#3E4451"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="ListViewItem">
      <Setter Property="Foreground"      Value="#ABB2BF"/>
      <Setter Property="Padding"         Value="4,3"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListViewItem">
            <Border x:Name="ItemBorder" Background="Transparent"
                    BorderThickness="0" Padding="{TemplateBinding Padding}">
              <GridViewRowPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ItemBorder" Property="Background" Value="#2D3139"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="ItemBorder" Property="Background" Value="#3A3F4B"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="Background"  Value="#21252B"/>
      <Setter Property="Foreground"  Value="#61AFEF"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="BorderBrush" Value="#3E4451"/>
      <Setter Property="Padding"     Value="8,5"/>
    </Style>

    <!-- ── SCROLLBAR (slim) ────────────────────────────────────── -->
    <Style TargetType="ScrollBar">
      <Setter Property="Stylus.IsFlicksEnabled" Value="False"/>
      <Setter Property="Width"    Value="8"/>
      <Setter Property="MinWidth" Value="8"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent" Width="8">
              <Track x:Name="PART_Track" IsDirectionReversed="True" Margin="0,2">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.LineUpCommand" Opacity="0" Focusable="False">
                    <RepeatButton.Template>
                      <ControlTemplate TargetType="RepeatButton"><Border/></ControlTemplate>
                    </RepeatButton.Template>
                  </RepeatButton>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.LineDownCommand" Opacity="0" Focusable="False">
                    <RepeatButton.Template>
                      <ControlTemplate TargetType="RepeatButton"><Border/></ControlTemplate>
                    </RepeatButton.Template>
                  </RepeatButton>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb">
                        <Border Background="#4B5263" CornerRadius="4" Margin="2,0"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="54"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="240"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Grid Grid.Row="0" Background="#21252B">
      <Grid.RowDefinitions>
        <RowDefinition Height="2"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <!-- Faixa gradiente no topo -->
      <Border Grid.Row="0">
        <Border.Background>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
            <GradientStop Color="#61AFEF" Offset="0"/>
            <GradientStop Color="#98C379" Offset="0.45"/>
            <GradientStop Color="#C678DD" Offset="1"/>
          </LinearGradientBrush>
        </Border.Background>
      </Border>
      <!-- Conteúdo do header -->
      <Border Grid.Row="1" BorderBrush="#3E4451" BorderThickness="0,0,0,1">
        <Grid Margin="20,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="Next" FontSize="22" FontWeight="Bold" Foreground="#61AFEF"/>
            <TextBlock Text="Tool" FontSize="22" FontWeight="Bold" Foreground="#ABB2BF"/>
            <Border Background="#2D3139" CornerRadius="4" Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
              <TextBlock Text="v4.1" FontSize="10" Foreground="#5C6370" FontWeight="SemiBold"/>
            </Border>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,4,0">
            <TextBlock x:Name="HdrPC"     Foreground="#5C6370" FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
            <TextBlock x:Name="HdrUptime" Foreground="#5C6370" FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
            <!-- Badge de modo -->
            <Border x:Name="BadgeModo" Background="#2D3139" CornerRadius="4"
                    Padding="8,3" Margin="0,0,10,0" VerticalAlignment="Center">
              <TextBlock x:Name="TxtModo" Text="USUARIO" Foreground="#E5C07B"
                         FontSize="10" FontWeight="Bold"/>
            </Border>
            <Button x:Name="BtnModoAdm" Content="Entrar ADM"
                    Background="#E5C07B" Foreground="#1E2128" FontSize="11"
                    Padding="10,5" FontWeight="SemiBold" Margin="0,0,8,0"/>
            <Button x:Name="BtnRelatorio" Content="Relatorios"
                    Background="#2D3139" Foreground="#ABB2BF" FontSize="11"
                    Padding="10,5" FontWeight="Normal"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>

    <!-- PAINEL USUARIO -->
    <Grid x:Name="UserPanel" Grid.Row="1" Visibility="Visible">
      <TabControl x:Name="UserTabs" Margin="0">

        <!-- ===== USER: SISTEMA ===== -->
        <TabItem Header="  Sistema  ">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
            <StackPanel Margin="20,16">
              <GroupBox Header="Informacoes do Computador">
                <Grid Margin="4,4">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,12,0">
                    <TextBlock Text="COMPUTADOR"  Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiPC"     Foreground="#ABB2BF" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,8"/>
                    <TextBlock Text="SISTEMA"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiOS"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                    <TextBlock Text="USUARIO"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiUser"   Foreground="#ABB2BF" FontSize="11"/>
                  </StackPanel>
                  <StackPanel Grid.Column="1" Margin="0,0,12,0">
                    <TextBlock Text="ENDERECO IP" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiIP"     Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                    <TextBlock Text="GATEWAY"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiGW"     Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                    <TextBlock Text="DNS"         Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiDNS"    Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="DOMINIO"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiDomain" Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                    <TextBlock Text="UPTIME"      Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiUptime" Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                    <TextBlock Text="MEMORIA RAM" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                    <TextBlock x:Name="USiRAM"    Foreground="#ABB2BF" FontSize="11"/>
                  </StackPanel>
                </Grid>
              </GroupBox>

              <!-- Diagnostico -->
              <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <Button x:Name="UBtnDiagnostico" Content="Executar Diagnostico Completo"
                        Background="#61AFEF" Foreground="#1E2128" Padding="18,10" FontWeight="SemiBold"/>
              </StackPanel>

              <!-- Ajuda: Relatorio -->
              <GroupBox Header="Como enviar informacoes para o suporte">
                <Grid Margin="4,4">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>

                  <Border Grid.Column="0" Background="#1E2128" CornerRadius="5" Padding="12,10">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <Border Background="#61AFEF" CornerRadius="3" Padding="7,2" Margin="0,0,8,0">
                          <TextBlock Text="EXPORTAR" Foreground="#1E2128" FontWeight="Bold" FontSize="10"/>
                        </Border>
                        <TextBlock Text="Log da sessao atual" Foreground="#ABB2BF" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" Foreground="#E5C07B" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,4"
                                 Text="1. Execute o Diagnostico Completo acima primeiro."/>
                      <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11" Margin="0,0,0,4"
                                 Text="2. Clique em 'Exportar' no rodape da tela. Um arquivo .txt sera salvo em:"/>
                      <Border Background="#21252B" CornerRadius="3" Padding="8,5" Margin="0,4,0,6">
                        <TextBlock Foreground="#98C379" FontSize="11" FontFamily="Consolas" Text="C:\Next-Relatorios\"/>
                      </Border>
                      <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                 Text="3. Encaminhe o arquivo .txt ao suporte por e-mail ou WhatsApp."/>
                    </StackPanel>
                  </Border>

                  <Border Grid.Column="2" Background="#1E2128" CornerRadius="5" Padding="12,10">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <Border Background="#2D3139" CornerRadius="3" Padding="7,2" Margin="0,0,8,0">
                          <TextBlock Text="RELATORIOS" Foreground="#ABB2BF" FontWeight="Bold" FontSize="10"/>
                        </Border>
                        <TextBlock Text="Historico de todas as sessoes" Foreground="#ABB2BF" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" Foreground="#E5C07B" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,4"
                                 Text="1. Execute o Diagnostico Completo acima primeiro."/>
                      <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11" Margin="0,0,0,4"
                                 Text="2. Clique em 'Relatorios' no canto superior direito. A pasta sera aberta:"/>
                      <Border Background="#21252B" CornerRadius="3" Padding="8,5" Margin="0,4,0,6">
                        <TextBlock Foreground="#98C379" FontSize="11" FontFamily="Consolas" Text="C:\Next-Relatorios\"/>
                      </Border>
                      <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                 Text="3. Envie o arquivo mais recente (nome do PC + data/hora) ao suporte."/>
                    </StackPanel>
                  </Border>

                </Grid>
              </GroupBox>

            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <!-- ===== USER: LIMPEZA ===== -->
        <TabItem Header="  Limpeza  ">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
            <StackPanel Margin="20,16">
              <GroupBox Header="Limpeza Rapida">
                <WrapPanel Margin="4,4">
                  <Button x:Name="UBtnOtimizar"    Content="Limpar Temporarios"    Width="185" Height="64" Margin="0,0,10,10"/>
                  <Button x:Name="UBtnTeamsOffice" Width="185" Height="64" Margin="0,0,10,10">
                    <TextBlock Text="Limpar Cache Teams/Office" TextWrapping="Wrap" TextAlignment="Center"/>
                  </Button>
                  <Button x:Name="UBtnSpooler"     Width="185" Height="64" Margin="0,0,10,10">
                    <TextBlock Text="Limpar Fila de Impressao" TextWrapping="Wrap" TextAlignment="Center"/>
                  </Button>
                  <Button x:Name="UBtnFlushDns"    Content="Flush DNS"             Width="185" Height="64" Margin="0,0,10,10"/>
                  <Button x:Name="UBtnRenovarIP"   Content="Renovar IP (DHCP)"     Width="185" Height="64" Margin="0,0,10,10"/>
                  <Button x:Name="UBtnSincHora"    Content="Sincronizar Hora"      Width="185" Height="64" Margin="0,0,10,10"/>
                </WrapPanel>
              </GroupBox>

              <!-- Ajuda: Limpeza -->
              <GroupBox Header="O que faz cada funcao">
                <Grid Margin="4,4">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8" Margin="0,0,0,6">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Limpar Temporarios" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Remove arquivos temporarios do Windows e esvazia a lixeira. Use quando o PC estiver lento ou com pouco espaco em disco."/>
                      </StackPanel>
                    </Border>
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8" Margin="0,0,0,6">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Limpar Cache Teams/Office" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Apaga arquivos temporarios do Teams e Office que acumulam GB ao longo do tempo. Resolve lentidao, travamentos e erros de sincronizacao. Reinicie os apps apos."/>
                      </StackPanel>
                    </Border>
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Limpar Fila de Impressao" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Cancela todos os documentos presos na fila da impressora e reinicia o servico. Use quando a impressora parou de imprimir sem motivo aparente."/>
                      </StackPanel>
                    </Border>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8" Margin="0,0,0,6">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Flush DNS" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Limpa o cache de enderecos de sites guardados pelo Windows. Use quando um site nao abre ou abre com erro, mesmo estando com internet."/>
                      </StackPanel>
                    </Border>
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8" Margin="0,0,0,6">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Renovar IP (DHCP)" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Libera o endereco IP atual e solicita um novo ao roteador. Use quando o PC esta com 'Rede sem internet' ou com IP de conflito (169.254.x.x)."/>
                      </StackPanel>
                    </Border>
                    <Border Background="#1E2128" CornerRadius="5" Padding="12,8">
                      <StackPanel>
                        <TextBlock Foreground="#61AFEF" FontWeight="SemiBold" FontSize="12" Text="Sincronizar Hora" Margin="0,0,0,3"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#5C6370" FontSize="11"
                                   Text="Sincroniza o relogio do PC com o servidor de hora da Microsoft. Use quando aparecem erros de certificado, login no Teams ou problemas de autenticacao."/>
                      </StackPanel>
                    </Border>
                  </StackPanel>
                </Grid>
              </GroupBox>

            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <!-- ===== USER: REDE ===== -->
        <TabItem Header="  Rede  ">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
            <StackPanel Margin="20,16">

              <GroupBox Header="Teste de Rede">
                <StackPanel Margin="4,4">
                  <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBox x:Name="UTxtPingHost" Text="8.8.8.8" Width="200" Margin="0,0,8,0"/>
                    <Button x:Name="UBtnPing"         Content="Ping"           Padding="14,8" Margin="0,0,8,0"/>
                    <Button x:Name="UBtnTracert"      Content="Tracert"        Padding="14,8" Margin="0,0,8,0"
                            Background="#4B5263" Foreground="#ABB2BF"/>
                    <Button x:Name="UBtnTestarConect" Content="Testar Conexao" Padding="14,8"/>
                  </StackPanel>
                  <TextBlock Foreground="#5C6370" FontSize="11" Margin="0,4,0,0" TextWrapping="Wrap"
                             Text="Digite um IP ou endereco (ex: 8.8.8.8, google.com) e clique em Ping ou Tracert. Os resultados aparecem no Log abaixo."/>
                </StackPanel>
              </GroupBox>

              <!-- Guia compacto -->
              <GroupBox Header="Como interpretar os resultados">
                <Grid Margin="4,4">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>

                  <!-- Ping -->
                  <Border Grid.Column="0" Background="#1E2128" CornerRadius="5" Padding="12,10">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <Border Background="#61AFEF" CornerRadius="3" Padding="7,2" Margin="0,0,8,0">
                          <TextBlock Text="PING" Foreground="#1E2128" FontWeight="Bold" FontSize="10"/>
                        </Border>
                        <TextBlock Text="Testa alcance e latencia" Foreground="#ABB2BF" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#98C379" Text="Abaixo de 50ms"/><Run Foreground="#5C6370" Text=" — otimo"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E5C07B" Text="50ms a 150ms"/><Run Foreground="#5C6370" Text=" — aceitavel, video pode travar"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E06C75" Text="Acima de 150ms"/><Run Foreground="#5C6370" Text=" — ruim, contate suporte"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11">
                        <Run Foreground="#E06C75" Text="Sem resposta"/><Run Foreground="#5C6370" Text=" — sem internet, teste 8.8.8.8"/>
                      </TextBlock>
                    </StackPanel>
                  </Border>

                  <!-- Tracert -->
                  <Border Grid.Column="2" Background="#1E2128" CornerRadius="5" Padding="12,10">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <Border Background="#4B5263" CornerRadius="3" Padding="7,2" Margin="0,0,8,0">
                          <TextBlock Text="TRACERT" Foreground="#ABB2BF" FontWeight="Bold" FontSize="10"/>
                        </Border>
                        <TextBlock Text="Mapeia rota ate o destino" Foreground="#ABB2BF" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#ABB2BF" Text="Cada linha = 1 roteador no caminho"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E5C07B" Text="* * *"/><Run Foreground="#5C6370" Text=" — roteador nao responde (normal), trafego passa"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E06C75" Text="Latencia alta num salto"/><Run Foreground="#5C6370" Text=" — gargalo naquele ponto"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11">
                        <Run Foreground="#E06C75" Text="Para no meio"/><Run Foreground="#5C6370" Text=" — bloqueado por firewall"/>
                      </TextBlock>
                    </StackPanel>
                  </Border>

                  <!-- Testar Conexao -->
                  <Border Grid.Column="4" Background="#1E2128" CornerRadius="5" Padding="12,10">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <Border Background="#61AFEF" CornerRadius="3" Padding="7,2" Margin="0,0,8,0">
                          <TextBlock Text="TESTAR" Foreground="#1E2128" FontWeight="Bold" FontSize="10"/>
                        </Border>
                        <TextBlock Text="Internet + DNS juntos" Foreground="#ABB2BF" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E06C75" Text="8.8.8.8 falha"/><Run Foreground="#5C6370" Text=" — sem internet, verifique cabo/Wi-Fi"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,6">
                        <Run Foreground="#E5C07B" Text="8.8.8.8 ok, site falha"/><Run Foreground="#5C6370" Text=" — DNS com problema, use Flush DNS"/>
                      </TextBlock>
                      <TextBlock TextWrapping="Wrap" FontSize="11">
                        <Run Foreground="#98C379" Text="Tudo ok"/><Run Foreground="#5C6370" Text=" — conexao normal"/>
                      </TextBlock>
                    </StackPanel>
                  </Border>

                </Grid>
              </GroupBox>

            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <!-- ===== USER: ARMAZENAMENTO ===== -->
        <TabItem Header="  Armazenamento  ">
          <Grid Background="#21252B">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Discos -->
            <GroupBox Grid.Row="0" Header="Discos" Margin="16,12,16,0">
              <StackPanel x:Name="UDrivePanel" Margin="4,4"/>
            </GroupBox>

            <!-- Analisar pasta -->
            <Grid Grid.Row="1" Margin="16,8,16,12">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#1E2128" CornerRadius="4" Padding="12,8" Margin="0,0,0,8">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="UTxtAnalyzePath" Grid.Column="0" Text="C:\" Margin="0,0,8,0"/>
                  <TextBlock Grid.Column="1" Text="Arquivos acima de" Foreground="#5C6370"
                             VerticalAlignment="Center" Margin="0,0,6,0" FontSize="11"/>
                  <TextBox x:Name="UTxtMinMB" Grid.Column="2" Text="50" Width="50" Margin="0,0,8,0"/>
                  <Button x:Name="UBtnAnalisarPasta" Grid.Column="3" Content="Analisar"
                          Background="#61AFEF" Foreground="#1E2128"/>
                </Grid>
              </Border>
              <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <GroupBox Grid.Column="0" Header="Subpastas  (por tamanho)  — duplo clique para abrir">
                  <ListView x:Name="ULvFolders">
                    <ListView.View>
                      <GridView>
                        <GridViewColumn Header="Pasta"   DisplayMemberBinding="{Binding Nome}"    Width="160"/>
                        <GridViewColumn Header="Tamanho" DisplayMemberBinding="{Binding Tamanho}" Width="80"/>
                        <GridViewColumn Header="Caminho" DisplayMemberBinding="{Binding Caminho}" Width="220"/>
                      </GridView>
                    </ListView.View>
                  </ListView>
                </GroupBox>
                <GroupBox Grid.Column="2" Header="Maiores arquivos  — duplo clique para abrir pasta">
                  <ListView x:Name="ULvFiles">
                    <ListView.View>
                      <GridView>
                        <GridViewColumn Header="Arquivo"  DisplayMemberBinding="{Binding Nome}"    Width="160"/>
                        <GridViewColumn Header="Tamanho"  DisplayMemberBinding="{Binding Tamanho}" Width="80"/>
                        <GridViewColumn Header="Caminho"  DisplayMemberBinding="{Binding Caminho}" Width="220"/>
                      </GridView>
                    </ListView.View>
                  </ListView>
                </GroupBox>
              </Grid>
            </Grid>
          </Grid>
        </TabItem>

      </TabControl>
    </Grid>

    <!-- PAINEL ADM -->
    <Grid x:Name="AdminPanel" Grid.Row="1" Visibility="Collapsed">
    <!-- TABS ADM -->
    <TabControl x:Name="MainTabs" Margin="0">

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
                  <TextBlock Text="DOMINIO / USUARIO" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiUser"    Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                  <TextBlock Text="UPTIME" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiUptime"  Foreground="#ABB2BF" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="0,0,12,0">
                  <TextBlock Text="PROCESSADOR" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiCPU"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="MEMORIA RAM"  Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiRAM"     Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                  <TextBlock Text="GPU"          Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiGPU"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="PLACA-MAE / BIOS" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiMobo"    Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="ARMAZENAMENTO" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiDisk"    Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="SEGURANCA"    Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiSec"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="TPM / SECURE BOOT" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiTpm"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <!-- Rede rapida -->
            <GroupBox Header="Rede">
              <Grid Margin="4,4">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,12,0">
                  <TextBlock Text="ENDERECO IP" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiIP"      Foreground="#ABB2BF" FontSize="11" Margin="0,0,0,8"/>
                  <TextBlock Text="GATEWAY"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiGW"      Foreground="#ABB2BF" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="0,0,12,0">
                  <TextBlock Text="DNS"         Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiDNS"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="ADAPTADOR"   Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiNIC"     Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="DOMINIO / WORKGROUP" Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiDomain"  Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8"/>
                  <TextBlock Text="LICENCA WINDOWS"     Foreground="#5C6370" FontSize="9" FontWeight="Bold" Margin="0,0,0,2"/>
                  <TextBlock x:Name="SiLicenca" Foreground="#ABB2BF" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <!-- Acao -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,14">
              <Button x:Name="BtnDiagnostico" Content="Diagnostico Completo"
                      Background="#61AFEF" Foreground="#1E2128" Padding="18,10" Margin="0,0,10,0"/>
              <Button x:Name="BtnAtualizarDrivers" Content="Atualizar Drivers + winget"
                      Background="#98C379" Foreground="#1E2128" Padding="18,10" Margin="0,0,10,0"/>
              <Button x:Name="BtnExportarRelatorio" Content="Exportar Relatorio HTML"
                      Background="#C678DD" Foreground="#1E2128" Padding="18,10" Margin="0,0,10,0"/>
              <Button x:Name="BtnHistoricoLogs" Content="Historico de Sessoes"
                      Background="#4B5263" Foreground="#ABB2BF" Padding="18,10"/>
            </StackPanel>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ===== MANUTENCAO ===== -->
      <TabItem Header="  Manutencao  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#21252B">
          <StackPanel Margin="20,16">

            <GroupBox Header="Limpeza">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnOtimizar"           Content="Otimizar PC"              Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnSfcDism"            Content="SFC + DISM"               Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnCheckDisk"          Content="Verificar Disco (C:)"     Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnLimparWU"           Content="Limpar Cache WU"          Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnLimparMiniaturas"  Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Limpar Cache Miniaturas" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
                <Button x:Name="BtnLimparCredenciais" Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Limpar Credenciais" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
                <Button x:Name="BtnLimparTeamsOffice" Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Limpar Cache Teams/Office" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
              </WrapPanel>
            </GroupBox>

            <GroupBox Header="Rede">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnFlushDns"        Content="Flush DNS"           Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnResetWinsock"    Content="Reset Winsock/IP"    Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnRenovarIP"       Content="Renovar IP (DHCP)"   Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnResetarProxy"    Content="Resetar Proxy"       Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnSincronizarHora" Content="Sincronizar Hora"    Width="175" Height="64" Margin="0,0,10,10"/>
              </WrapPanel>
            </GroupBox>

            <GroupBox Header="Impressao">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnLimparSpooler"         Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Limpar Fila de Impressao" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
                <Button x:Name="BtnReinstalarImpressoras" Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Reinstalar Impressoras" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
              </WrapPanel>
            </GroupBox>

            <GroupBox Header="Sistema">
              <WrapPanel Margin="4,4">
                <Button x:Name="BtnGpUpdate"          Content="gpupdate /force"       Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnRestartExplorer"   Content="Reiniciar Explorer"    Width="175" Height="64" Margin="0,0,10,10"
                        Background="#E5C07B" Foreground="#1E2128"/>
                <Button x:Name="BtnVerificarUpdates"  Content="Verificar Updates"     Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnServicos"          Content="Status de Servicos"    Width="175" Height="64" Margin="0,0,10,10"/>
                <Button x:Name="BtnPerfilCorrompido"  Width="190" Height="64" Margin="0,0,10,10">
                  <TextBlock Text="Limpar Perfil Corrompido" TextWrapping="Wrap" TextAlignment="Center"/>
                </Button>
              </WrapPanel>
            </GroupBox>

            <GroupBox Header="Reiniciar Servico">
              <StackPanel Margin="4,4" Orientation="Horizontal">
                <TextBox x:Name="TxtServico" Width="200" Margin="0,0,8,0"/>
                <Button x:Name="BtnReiniciarServico" Content="Reiniciar"
                        Background="#E5C07B" Foreground="#1E2128" Padding="14,5"/>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="Desinstalar Aplicativo (winget)">
              <StackPanel Margin="4,4" Orientation="Horizontal">
                <TextBox x:Name="TxtDesinstalarApp" Width="280" Margin="0,0,8,0"/>
                <Button x:Name="BtnDesinstalarApp" Content="Desinstalar"
                        Background="#E06C75" Foreground="#1E2128" Padding="14,5"/>
              </StackPanel>
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
              <TextBlock Text="Selecao rapida:" Foreground="#5C6370" VerticalAlignment="Center" Margin="0,0,12,0" FontSize="11"/>
              <Button x:Name="BtnPresetNext"   Content="Padrao Next" Background="#98C379" Foreground="#1E2128" Padding="14,5" Margin="0,0,8,0"/>
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
                    <CheckBox x:Name="ChkSuspender"        Content="Desativar Suspender (Sleep)"/>
                    <CheckBox x:Name="ChkTela"             Content="Desativar Desligamento de Tela"/>
                    <CheckBox x:Name="ChkSmartApp"         Content="Desativar Smart App Control  (Win11)"/>
                  </StackPanel>
                </GroupBox>
              </StackPanel>
              <StackPanel Grid.Column="2">
                <GroupBox Header="Preferencias  (opcionais)">
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
      <TabItem Header="  Rede / Dominio  ">
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
                  <Button x:Name="BtnListarAdapters"   Content="Listar Adaptadores"   Padding="14,8" Margin="0,0,8,0"/>
                  <Button x:Name="BtnTestarConect"     Content="Testar Conectividade" Padding="14,8" Margin="0,0,8,0"/>
                  <Button x:Name="BtnIPConfig"         Content="IPConfig /all"        Padding="14,8" Margin="0,0,8,0"/>
                  <Button x:Name="BtnVerificarDominio" Content="Status Dominio"       Padding="14,8" Margin="0,0,8,0"/>
                  <Button x:Name="BtnIPInfo"           Content="Info Rede Completa"   Padding="14,8" Margin="0,0,0,0"/>
                </WrapPanel>
              </GroupBox>

              <GroupBox Header="Ping / Tracert">
                <StackPanel Margin="4,4">
                  <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                    <TextBox x:Name="TxtPingHost" Text="8.8.8.8" Width="180" Margin="0,0,8,0"/>
                    <Button x:Name="BtnPing"        Content="Ping"            Margin="0,0,8,0" Padding="12,5"/>
                    <Button x:Name="BtnPingContinuo" Content="Ping Continuo"  Margin="0,0,8,0" Padding="12,5"
                            Background="#E5C07B" Foreground="#1E2128"/>
                    <Button x:Name="BtnTracert"     Content="Tracert"         Padding="12,5"
                            Background="#4B5263" Foreground="#ABB2BF"/>
                  </StackPanel>
                </StackPanel>
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
              <GroupBox Header="Ingressar em Dominio AD">
                <StackPanel Margin="4,4">
                  <Label Content="Domínio  (ex: empresa.local)"/>
                  <TextBox x:Name="TxtDominio" Margin="0,0,0,8"/>
                  <Label Content="Usuário com permissão de join"/>
                  <TextBox x:Name="TxtDomUser" Margin="0,0,0,8"/>
                  <Label Content="Senha"/>
                  <PasswordBox x:Name="TxtDomPass" Margin="0,0,0,8"/>
                  <Label Content="Novo nome do PC  (opcional)"/>
                  <TextBox x:Name="TxtDomName" Margin="0,0,0,14"/>
                  <Button x:Name="BtnJoinDomain" Content="Ingressar no Dominio"
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
      <TabItem Header="  Usuarios  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Botoes de acao -->
          <Border Grid.Row="0" Background="#1E2128" Padding="14,10">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="BtnListarUsers" Content="Atualizar Lista"
                      Background="#4B5263" Foreground="#ABB2BF" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnToggleUser" Content="Ativar/Desativar"
                      Background="#E5C07B" Foreground="#1E2128" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnAddToAdmins" Content="&gt; Administradores"
                      Background="#61AFEF" Foreground="#1E2128" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnRemoveUser" Content="Remover Usuario"
                      Background="#E06C75" Foreground="#1E2128" Padding="12,5" Margin="0,0,8,0"/>
              <Button x:Name="BtnDesbloquearUser" Content="Desbloquear"
                      Background="#98C379" Foreground="#1E2128" Padding="12,5"/>
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
                <TextBlock Text="CRIAR USUARIO" Foreground="#5C6370" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/>
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

      <!-- ===== AREA DE TRABALHO ===== -->
      <TabItem Header="  Area de Trabalho  ">
        <Grid Background="#21252B">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Icones do sistema -->
          <GroupBox Grid.Row="0" Header="Ícones do Sistema" Margin="16,12,16,0">
            <StackPanel Margin="4,6">
              <WrapPanel>
                <CheckBox x:Name="ChkDtComputer" Content="Meu Computador"      Width="210" Margin="0,0,10,6"/>
                <CheckBox x:Name="ChkDtUser"     Content="Arquivos do Usuario" Width="210" Margin="0,0,10,6"/>
                <CheckBox x:Name="ChkDtNetwork"  Content="Rede"                Width="210" Margin="0,0,10,6"/>
                <CheckBox x:Name="ChkDtRecycle"  Content="Lixeira"             Width="210" Margin="0,0,10,6"/>
                <CheckBox x:Name="ChkDtControl"  Content="Painel de Controle"  Width="210" Margin="0,0,10,6"/>
              </WrapPanel>
              <Button x:Name="BtnAplicarIcones" Content="Aplicar Icones"
                      HorizontalAlignment="Left" Background="#61AFEF" Foreground="#1E2128"
                      Padding="14,6" Margin="0,8,0,0"/>
            </StackPanel>
          </GroupBox>

          <!-- Apps instalados -->
          <GroupBox Grid.Row="1" Header="Atalhos de Aplicativos" Margin="16,10,16,0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <!-- Barra superior -->
              <Grid Grid.Row="0" Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="TxtAppSearch" Grid.Column="0" Margin="0,0,8,0"
                         FontSize="12"/>
                <TextBlock Text="?" Grid.Column="0" Foreground="#5C6370" FontSize="13"
                           HorizontalAlignment="Right" VerticalAlignment="Center"
                           Margin="0,0,16,0" IsHitTestVisible="False"/>
                <Button x:Name="BtnCarregarApps"  Grid.Column="1" Content="Carregar"
                        Background="#4B5263" Foreground="#ABB2BF" Padding="10,5" Margin="0,0,6,0"/>
                <Button x:Name="BtnMarcarTodos"   Grid.Column="2" Content="Marcar Todos"
                        Background="#4B5263" Foreground="#ABB2BF" Padding="10,5" Margin="0,0,6,0"/>
                <Button x:Name="BtnDesmarcarTodos" Grid.Column="3" Content="Desmarcar"
                        Background="#4B5263" Foreground="#ABB2BF" Padding="10,5"/>
              </Grid>

              <!-- Grade de checkboxes (igual ao painel de icones do sistema) -->
              <Border Grid.Row="1" Background="#1E2128" BorderBrush="#3E4451" BorderThickness="1" CornerRadius="4">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                  <WrapPanel x:Name="AppCheckPanel" Margin="10,8" Orientation="Horizontal"/>
                </ScrollViewer>
              </Border>
            </Grid>
          </GroupBox>

          <!-- Rodape -->
          <Border Grid.Row="2" Background="#1E2128" Padding="16,10">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="BtnCriarAtalho" Content="+ Criar Atalhos Selecionados"
                      Background="#98C379" Foreground="#1E2128" Padding="14,7" Margin="0,0,14,0"/>
              <TextBlock x:Name="TxtAppCount" Foreground="#5C6370" FontSize="11"
                         VerticalAlignment="Center"/>
            </StackPanel>
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
            <GroupBox Grid.Column="0" Header="Subpastas  (por tamanho)  — duplo clique para abrir">
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
            <GroupBox Grid.Column="2" Header="Maiores arquivos  — duplo clique para abrir pasta">
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
    </Grid><!-- fim AdminPanel -->

    <!-- LOG PANEL -->
    <Grid Grid.Row="2" Background="#1E2128" MinHeight="220">
      <Grid.RowDefinitions>
        <RowDefinition Height="28"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Grid.Row="0" Background="#21252B" BorderBrush="#3E4451" BorderThickness="0,1,0,1">
        <Grid Margin="14,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
            <Border Width="7" Height="7" CornerRadius="4" Background="#3E4451" Margin="0,0,8,0"/>
            <TextBlock Text="LOG DE SAIDA" Foreground="#5C6370" FontSize="10"
                       FontWeight="Bold" VerticalAlignment="Center"/>
          </StackPanel>
          <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
            <Button x:Name="BtnLimparLog" Content="Limpar"
                    Padding="10,3" Margin="0,0,6,0"
                    Background="#3E4451" Foreground="#ABB2BF"
                    FontSize="11" FontWeight="Normal"/>
            <Button x:Name="BtnExportLog" Content="Exportar"
                    Padding="10,3"
                    Background="#61AFEF" Foreground="#1E2128"
                    FontSize="11" FontWeight="SemiBold"/>
          </StackPanel>
        </Grid>
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
# --- Modo ---
$UserPanel          = $Window.FindName("UserPanel")
$AdminPanel         = $Window.FindName("AdminPanel")
$BtnModoAdm         = $Window.FindName("BtnModoAdm")
$TxtModo            = $Window.FindName("TxtModo")
$BadgeModo          = $Window.FindName("BadgeModo")
# --- User mode controls ---
$USiPC              = $Window.FindName("USiPC")
$USiOS              = $Window.FindName("USiOS")
$USiUser            = $Window.FindName("USiUser")
$USiIP              = $Window.FindName("USiIP")
$USiGW              = $Window.FindName("USiGW")
$USiDNS             = $Window.FindName("USiDNS")
$USiDomain          = $Window.FindName("USiDomain")
$USiUptime          = $Window.FindName("USiUptime")
$USiRAM             = $Window.FindName("USiRAM")
$UBtnOtimizar       = $Window.FindName("UBtnOtimizar")
$UBtnTeamsOffice    = $Window.FindName("UBtnTeamsOffice")
$UBtnSpooler        = $Window.FindName("UBtnSpooler")
$UBtnFlushDns       = $Window.FindName("UBtnFlushDns")
$UBtnRenovarIP      = $Window.FindName("UBtnRenovarIP")
$UBtnSincHora       = $Window.FindName("UBtnSincHora")
$UTxtPingHost       = $Window.FindName("UTxtPingHost")
$UBtnPing           = $Window.FindName("UBtnPing")
$UBtnTracert        = $Window.FindName("UBtnTracert")
$UBtnTestarConect   = $Window.FindName("UBtnTestarConect")
$UBtnDiagnostico    = $Window.FindName("UBtnDiagnostico")
$UDrivePanel        = $Window.FindName("UDrivePanel")
$UTxtAnalyzePath    = $Window.FindName("UTxtAnalyzePath")
$UTxtMinMB          = $Window.FindName("UTxtMinMB")
$UBtnAnalisarPasta  = $Window.FindName("UBtnAnalisarPasta")
$ULvFolders         = $Window.FindName("ULvFolders")
$ULvFiles           = $Window.FindName("ULvFiles")
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
$BtnLimparWU              = $Window.FindName("BtnLimparWU")
$BtnLimparSpooler         = $Window.FindName("BtnLimparSpooler")
$BtnReinstalarImpressoras = $Window.FindName("BtnReinstalarImpressoras")
$BtnRenovarIP             = $Window.FindName("BtnRenovarIP")
$BtnResetarProxy          = $Window.FindName("BtnResetarProxy")
$BtnSincronizarHora       = $Window.FindName("BtnSincronizarHora")
$BtnLimparMiniaturas      = $Window.FindName("BtnLimparMiniaturas")
$BtnLimparCredenciais     = $Window.FindName("BtnLimparCredenciais")
$BtnLimparTeamsOffice     = $Window.FindName("BtnLimparTeamsOffice")
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
$ChkSuspender       = $Window.FindName("ChkSuspender")
$ChkTela            = $Window.FindName("ChkTela")
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
$SiIP               = $Window.FindName("SiIP")
$SiGW               = $Window.FindName("SiGW")
$SiDNS              = $Window.FindName("SiDNS")
$SiNIC              = $Window.FindName("SiNIC")
$SiDomain           = $Window.FindName("SiDomain")
$SiLicenca          = $Window.FindName("SiLicenca")
$BtnExportarRelatorio = $Window.FindName("BtnExportarRelatorio")
$BtnHistoricoLogs   = $Window.FindName("BtnHistoricoLogs")
$BtnVerificarUpdates= $Window.FindName("BtnVerificarUpdates")
$BtnServicos        = $Window.FindName("BtnServicos")
$BtnPerfilCorrompido= $Window.FindName("BtnPerfilCorrompido")
$TxtServico         = $Window.FindName("TxtServico")
$BtnReiniciarServico= $Window.FindName("BtnReiniciarServico")
$TxtDesinstalarApp  = $Window.FindName("TxtDesinstalarApp")
$BtnDesinstalarApp  = $Window.FindName("BtnDesinstalarApp")
$TxtPingHost        = $Window.FindName("TxtPingHost")
$BtnPing            = $Window.FindName("BtnPing")
$BtnPingContinuo    = $Window.FindName("BtnPingContinuo")
$BtnTracert         = $Window.FindName("BtnTracert")
$BtnVerificarDominio= $Window.FindName("BtnVerificarDominio")
$BtnIPInfo          = $Window.FindName("BtnIPInfo")
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
$BtnRemoveUser        = $Window.FindName("BtnRemoveUser")
$BtnDesbloquearUser   = $Window.FindName("BtnDesbloquearUser")
$TxtNewUserName     = $Window.FindName("TxtNewUserName")
$TxtNewUserPass     = $Window.FindName("TxtNewUserPass")
$ChkNewUserAdmin    = $Window.FindName("ChkNewUserAdmin")
$BtnCreateUser      = $Window.FindName("BtnCreateUser")
$TxtChgUser         = $Window.FindName("TxtChgUser")
$TxtChgPass         = $Window.FindName("TxtChgPass")
$BtnSetPassword     = $Window.FindName("BtnSetPassword")
$ChkDtComputer      = $Window.FindName("ChkDtComputer")
$ChkDtUser          = $Window.FindName("ChkDtUser")
$ChkDtNetwork       = $Window.FindName("ChkDtNetwork")
$ChkDtRecycle       = $Window.FindName("ChkDtRecycle")
$ChkDtControl       = $Window.FindName("ChkDtControl")
$BtnAplicarIcones   = $Window.FindName("BtnAplicarIcones")
$TxtAppSearch       = $Window.FindName("TxtAppSearch")
$BtnCarregarApps    = $Window.FindName("BtnCarregarApps")
$AppCheckPanel      = $Window.FindName("AppCheckPanel")
$BtnMarcarTodos     = $Window.FindName("BtnMarcarTodos")
$BtnDesmarcarTodos  = $Window.FindName("BtnDesmarcarTodos")
$BtnCriarAtalho     = $Window.FindName("BtnCriarAtalho")
$TxtAppCount        = $Window.FindName("TxtAppCount")
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
# TIMER - drena filas dos runspaces para a UI
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
        $LvFolders.Items.Clear();  $ULvFolders.Items.Clear()
        foreach ($r in $res) { [void]$LvFolders.Items.Add($r); [void]$ULvFolders.Items.Add($r) }
    }
    if ($script:FileQueue.TryDequeue([ref]$res)) {
        $LvFiles.Items.Clear();  $ULvFiles.Items.Clear()
        foreach ($r in $res) { [void]$LvFiles.Items.Add($r); [void]$ULvFiles.Items.Add($r) }
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

    $HdrPC.Text     = $env:COMPUTERNAME
    $HdrUptime.Text = "$($up.Days)d $($up.Hours)h $($up.Minutes)m"

    # Painel usuario
    try {
        $USiPC.Text     = $env:COMPUTERNAME
        $USiOS.Text     = "$($os.Caption -replace 'Microsoft ','')  (Build $($os.BuildNumber))"
        $USiUser.Text   = "$env:USERNAME  @  $($cs.Domain)"
        $USiUptime.Text = "$($up.Days)d $($up.Hours)h $($up.Minutes)m"
        $ramT2 = [math]::Round($cs.TotalPhysicalMemory/1GB,1)
        $ramF2 = [math]::Round($os.FreePhysicalMemory/1MB,1)
        $USiRAM.Text    = "Total: ${ramT2} GB   |   Livre: ${ramF2} GB"
        $USiDomain.Text = if ($cs.PartOfDomain) { "Dominio: $($cs.Domain)" } else { "Workgroup: $($cs.Domain)" }
        $nic2 = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" | Select-Object -First 1
        $USiIP.Text  = if ($nic2.IPAddress)            { $nic2.IPAddress[0] }                     else { "N/A" }
        $USiGW.Text  = if ($nic2.DefaultIPGateway)     { $nic2.DefaultIPGateway[0] }              else { "N/A" }
        $USiDNS.Text = if ($nic2.DNSServerSearchOrder) { $nic2.DNSServerSearchOrder -join " | " } else { "N/A" }
    } catch {}

    # Rede rapida
    try {
        $nic = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" | Select-Object -First 1
        $SiIP.Text  = if ($nic.IPAddress)            { $nic.IPAddress[0] }             else { "N/A" }
        $SiGW.Text  = if ($nic.DefaultIPGateway)     { $nic.DefaultIPGateway[0] }      else { "N/A" }
        $SiDNS.Text = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder -join " | " } else { "N/A" }
        $SiNIC.Text = $nic.Description
    } catch { $SiIP.Text = "N/A" }

    # Dominio
    try {
        $SiDomain.Text = if ($cs.PartOfDomain) { "Dominio: $($cs.Domain)" } else { "Workgroup: $($cs.Domain)" }
    } catch {}

    # Licenca Windows
    try {
        $lic = Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND LicenseStatus=1 AND Name LIKE 'Windows%'" -ErrorAction Stop | Select-Object -First 1
        $SiLicenca.Text = if ($lic) { $lic.Name -replace "Windows ","" } else { "N/A" }
    } catch { $SiLicenca.Text = "N/A" }
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
        $info.Text = "${usedGB}GB / ${totalGB}GB  (${freeGB}GB livre - ${pct}%)"
        $info.FontSize = 11; $info.VerticalAlignment = "Center"
        $info.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#5C6370")
        [System.Windows.Controls.Grid]::SetColumn($info, 2)

        $row.Children.Add($lbl) | Out-Null
        $row.Children.Add($pb)  | Out-Null
        $row.Children.Add($info)| Out-Null
        $DrivePanel.Children.Add($row) | Out-Null

        # Replica para o painel usuario
        $row2 = New-Object System.Windows.Controls.Grid; $row2.Margin = "0,0,0,10"
        $u1 = New-Object System.Windows.Controls.ColumnDefinition; $u1.Width = "60"
        $u2 = New-Object System.Windows.Controls.ColumnDefinition; $u2.Width = "*"
        $u3 = New-Object System.Windows.Controls.ColumnDefinition; $u3.Width = "210"
        $row2.ColumnDefinitions.Add($u1); $row2.ColumnDefinitions.Add($u2); $row2.ColumnDefinitions.Add($u3)
        $lbl2 = New-Object System.Windows.Controls.TextBlock
        $lbl2.Text = "$($_.Name):"; $lbl2.FontWeight = "SemiBold"; $lbl2.VerticalAlignment = "Center"
        $lbl2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ABB2BF")
        [System.Windows.Controls.Grid]::SetColumn($lbl2, 0)
        $pb2 = New-Object System.Windows.Controls.ProgressBar
        $pb2.Value = $pct; $pb2.Maximum = 100; $pb2.VerticalAlignment = "Center"; $pb2.Margin = "0,0,12,0"
        $pb2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($color)
        [System.Windows.Controls.Grid]::SetColumn($pb2, 1)
        $info2 = New-Object System.Windows.Controls.TextBlock
        $info2.Text = "${usedGB}GB / ${totalGB}GB  (${freeGB}GB livre - ${pct}%)"
        $info2.FontSize = 11; $info2.VerticalAlignment = "Center"
        $info2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#5C6370")
        [System.Windows.Controls.Grid]::SetColumn($info2, 2)
        $row2.Children.Add($lbl2) | Out-Null
        $row2.Children.Add($pb2)  | Out-Null
        $row2.Children.Add($info2)| Out-Null
        $UDrivePanel.Children.Add($row2) | Out-Null
    }
} catch {}

# ================================================================
# EVENT HANDLERS
# ================================================================

# ================================================================
# MODO ADM / USUARIO
# ================================================================
Initialize-Senha

function Enter-ModoAdm {
    $senha = Show-DialogSenha
    if ($null -eq $senha) { return }
    if (Test-SenhaAdm $senha) {
        $script:MODE = "ADMIN"
        $UserPanel.Visibility  = "Collapsed"
        $AdminPanel.Visibility = "Visible"
        $TxtModo.Text          = "ADM"
        $TxtModo.Foreground    = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#98C379")
        $BtnModoAdm.Content    = "Sair ADM"
        $BtnModoAdm.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4B5263")
        $BtnModoAdm.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ABB2BF")
        Write-Log "Modo ADM ativado." "OK"
    } else {
        [System.Windows.MessageBox]::Show("Senha incorreta.", "Acesso negado", "OK", "Warning") | Out-Null
        Write-Log "Tentativa de acesso ADM com senha incorreta." "AVISO"
    }
}

function Exit-ModoAdm {
    $script:MODE = "USER"
    $AdminPanel.Visibility = "Collapsed"
    $UserPanel.Visibility  = "Visible"
    $TxtModo.Text          = "USUARIO"
    $TxtModo.Foreground    = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#E5C07B")
    $BtnModoAdm.Content    = "Entrar ADM"
    $BtnModoAdm.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#E5C07B")
    $BtnModoAdm.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1E2128")
    Write-Log "Saiu do modo ADM." "INFO"
}

$BtnModoAdm.Add_Click({
    if ($script:MODE -eq "ADMIN") { Exit-ModoAdm } else { Enter-ModoAdm }
})

# --- Painel Usuario ---
$UBtnOtimizar.Add_Click({    Invoke-Async { Invoke-OtimizarPC } })
$UBtnTeamsOffice.Add_Click({ Invoke-Async { Invoke-LimparCacheTeamsOffice } })
$UBtnSpooler.Add_Click({     Invoke-Async { Invoke-LimparSpooler } })
$UBtnFlushDns.Add_Click({    Invoke-Async { ipconfig /flushdns | Out-Null; Write-Log "Cache DNS limpo." "OK" } })
$UBtnRenovarIP.Add_Click({   Invoke-Async { Invoke-RenovarIP } })
$UBtnSincHora.Add_Click({    Invoke-Async { Invoke-SincronizarHora } })
$UBtnPing.Add_Click({
    $h = $UTxtPingHost.Text.Trim(); if (-not $h) { $h = "8.8.8.8" }
    Invoke-Async { Invoke-PingVisual -Destino $H } -Vars @{ H = $h }
})
$UBtnTracert.Add_Click({
    $h = $UTxtPingHost.Text.Trim(); if (-not $h) { $h = "8.8.8.8" }
    Invoke-Async { Invoke-TracertVisual -Destino $H } -Vars @{ H = $h }
})
$UBtnTestarConect.Add_Click({  Invoke-Async { Invoke-TestarConectividade } })
$UBtnDiagnostico.Add_Click({   Invoke-Async { Invoke-Diagnostico } })

$UBtnAnalisarPasta.Add_Click({
    $path  = $UTxtAnalyzePath.Text.Trim()
    $minMB = try { [int]$UTxtMinMB.Text } catch { 50 }
    if (-not $path) { $path = "C:\" }
    if (-not (Test-Path $path)) { Write-Log "Pasta nao encontrada: $path" "ERRO"; return }
    $ULvFolders.Items.Clear(); $ULvFiles.Items.Clear()
    Invoke-Async { Invoke-AnalisarPasta -Path $P -MinMB $M } -Vars @{ P=$path; M=$minMB }
})

$ULvFolders.Add_MouseDoubleClick({
    $sel = $ULvFolders.SelectedItem
    if ($sel -and $sel.Caminho -and (Test-Path $sel.Caminho)) {
        Start-Process explorer.exe $sel.Caminho
    }
})
$ULvFiles.Add_MouseDoubleClick({
    $sel = $ULvFiles.SelectedItem
    if ($sel -and $sel.Caminho) {
        $pasta = Split-Path $sel.Caminho -Parent
        if (Test-Path $pasta) { Start-Process explorer.exe $pasta }
    }
})

# Redireciona resultado de analise para os dois paineis

# --- Sistema ---
$BtnDiagnostico.Add_Click({       Invoke-Async { Invoke-Diagnostico } })
$BtnAtualizarDrivers.Add_Click({  Invoke-Async { Invoke-TweakDrivers } })
$BtnExportarRelatorio.Add_Click({
    $dir = $script:REPORT_DIR
    Invoke-Async { Export-RelatorioHTML -ReportDir $Dir } -Vars @{ Dir = $dir }
})
$BtnHistoricoLogs.Add_Click({
    $dir = $script:REPORT_DIR
    Invoke-Async { Invoke-HistoricoLogs -ReportDir $Dir } -Vars @{ Dir = $dir }
})
$BtnRelatorio.Add_Click({        Start-Process explorer.exe $script:REPORT_DIR })

# --- Manutencao ---
$BtnOtimizar.Add_Click({         Invoke-Async { Invoke-OtimizarPC } })
$BtnSfcDism.Add_Click({          Invoke-Async { Invoke-SFCDISM } })
$BtnCheckDisk.Add_Click({        Invoke-Async { Invoke-CheckDisk "C:" } })
$BtnLimparWU.Add_Click({              Invoke-Async { Invoke-LimparCacheWindowsUpdate } })
$BtnLimparMiniaturas.Add_Click({      Invoke-Async { Invoke-LimparMiniaturas } })
$BtnLimparCredenciais.Add_Click({     Invoke-Async { Invoke-LimparCredenciais } })
$BtnLimparTeamsOffice.Add_Click({     Invoke-Async { Invoke-LimparCacheTeamsOffice } })
$BtnFlushDns.Add_Click({              Invoke-Async { ipconfig /flushdns | Out-Null; Write-Log "Cache DNS limpo." "OK" } })
$BtnResetWinsock.Add_Click({          Invoke-Async { Invoke-ResetWinsock } })
$BtnRenovarIP.Add_Click({             Invoke-Async { Invoke-RenovarIP } })
$BtnResetarProxy.Add_Click({          Invoke-Async { Invoke-ResetarProxy } })
$BtnSincronizarHora.Add_Click({       Invoke-Async { Invoke-SincronizarHora } })
$BtnLimparSpooler.Add_Click({         Invoke-Async { Invoke-LimparSpooler } })
$BtnReinstalarImpressoras.Add_Click({ Invoke-Async { Invoke-ReinstalarImpressoras } })
$BtnGpUpdate.Add_Click({          Invoke-Async { Invoke-GpUpdate } })
$BtnRestartExplorer.Add_Click({   Invoke-Async { Invoke-RestartExplorer } })
$BtnVerificarUpdates.Add_Click({  Invoke-Async { Invoke-VerificarUpdates } })
$BtnServicos.Add_Click({          Invoke-Async { Show-ServiciosCriticos } })
$BtnPerfilCorrompido.Add_Click({  Invoke-Async { Invoke-LimparPerfilCorrompido } })
$BtnReiniciarServico.Add_Click({
    $svc = $TxtServico.Text.Trim()
    Invoke-Async { Invoke-ReiniciarServico -NomeServico $Svc } -Vars @{ Svc = $svc }
})
$BtnDesinstalarApp.Add_Click({
    $app = $TxtDesinstalarApp.Text.Trim()
    Invoke-Async { Invoke-DesinstalarApp -NomeApp $App } -Vars @{ App = $app }
})

# --- Tweaks ---
$script:AllTweakChks = @(
    $ChkTelemetria,$ChkActivityHistory,$ChkLocationTracking,
    $ChkFileExtensions,$ChkHiddenFiles,$ChkNumLock,
    $ChkEndTask,$ChkServices,$ChkHibernacao,$ChkSmartApp,
    $ChkUltimatePerf,$ChkDarkTheme,$ChkWidgets,$ChkVerboseLogon,
    $ChkSuspender,$ChkTela
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
    $ChkSuspender.IsChecked        = $true
    $ChkTela.IsChecked             = $true
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
        Sus=$ChkSuspender.IsChecked; Tla=$ChkTela.IsChecked
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
        if ($V.Sus)  { Invoke-TweakSuspender }
        if ($V.Tla)  { Invoke-TweakTela }
        Write-Log "Todos os tweaks aplicados." "OK"
    } -Vars @{ V = $v }
})

# --- Rede ---
$BtnListarAdapters.Add_Click({   Invoke-Async { Show-Adapters } })
$BtnVerificarDominio.Add_Click({ Invoke-Async { Invoke-VerificarDominio } })
$BtnIPInfo.Add_Click({           Invoke-Async { Show-IPInfo } })
$BtnPing.Add_Click({
    $h = $TxtPingHost.Text.Trim(); if (-not $h) { $h = "8.8.8.8" }
    Invoke-Async { Invoke-PingVisual -Destino $H } -Vars @{ H = $h }
})
$BtnPingContinuo.Add_Click({
    $h = $TxtPingHost.Text.Trim(); if (-not $h) { $h = "8.8.8.8" }
    Invoke-Async { Invoke-PingContinuo -Destino $H } -Vars @{ H = $h }
})
$BtnTracert.Add_Click({
    $h = $TxtPingHost.Text.Trim(); if (-not $h) { $h = "8.8.8.8" }
    Invoke-Async { Invoke-TracertVisual -Destino $H } -Vars @{ H = $h }
})
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
$BtnDesbloquearUser.Add_Click({
    $u = if ($LvUsers.SelectedItem) { $LvUsers.SelectedItem.Nome } else { "" }
    Invoke-Async { Invoke-DesbloquearUsuario -NomeUsuario $U } -Vars @{ U = $u }
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

# --- Area de Trabalho ---

# Carrega estado atual dos icones do sistema ao abrir
$script:AllApps = @()
try {
    $iconState = Get-DesktopIconState
    $ChkDtComputer.IsChecked = $iconState["Meu Computador"]
    $ChkDtUser.IsChecked     = $iconState["Arquivos do Usuario"]
    $ChkDtNetwork.IsChecked  = $iconState["Rede"]
    $ChkDtRecycle.IsChecked  = $iconState["Lixeira"]
    $ChkDtControl.IsChecked  = $iconState["Painel de Controle"]
} catch {}

$BtnAplicarIcones.Add_Click({
    Invoke-Async {
        Set-DesktopIconState "Meu Computador"      $States["Computer"]
        Set-DesktopIconState "Arquivos do Usuario" $States["User"]
        Set-DesktopIconState "Rede"                $States["Network"]
        Set-DesktopIconState "Lixeira"             $States["Recycle"]
        Set-DesktopIconState "Painel de Controle"  $States["Control"]
        Write-Log "Icones da area de trabalho atualizados." "OK"
    } -Vars @{ States = @{
        Computer = [bool]$ChkDtComputer.IsChecked
        User     = [bool]$ChkDtUser.IsChecked
        Network  = [bool]$ChkDtNetwork.IsChecked
        Recycle  = [bool]$ChkDtRecycle.IsChecked
        Control  = [bool]$ChkDtControl.IsChecked
    }}
})

# Helper: reconstroi checkboxes no painel conforme filtro
function Update-AppCheckPanel {
    param([string]$Filter = "")
    $AppCheckPanel.Children.Clear()
    $filtered = if ($Filter) {
        $script:AllApps | Where-Object { $_.Nome -like "*$Filter*" }
    } else { $script:AllApps }
    foreach ($app in $filtered) {
        $chk = New-Object System.Windows.Controls.CheckBox
        $chk.Content   = $app.Nome
        $chk.Tag       = $app.Caminho
        $chk.Width     = 220
        $chk.Margin    = [System.Windows.Thickness]::new(0, 0, 10, 6)
        $chk.Foreground= [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ABB2BF")
        $chk.Cursor    = [System.Windows.Input.Cursors]::Hand
        $chk.ToolTip   = $app.Caminho
        [void]$AppCheckPanel.Children.Add($chk)
    }
    $TxtAppCount.Text = "$($AppCheckPanel.Children.Count) de $($script:AllApps.Count) app(s)"
}

# Carregar apps instalados
$BtnCarregarApps.Add_Click({
    $TxtAppCount.Text = "Carregando..."
    $script:AllApps   = Get-InstalledApps
    Update-AppCheckPanel
    Write-Log "$($script:AllApps.Count) aplicativo(s) listado(s)." "OK"
})

# Filtro em tempo real
$TxtAppSearch.Add_TextChanged({ Update-AppCheckPanel -Filter $TxtAppSearch.Text.Trim() })

# Marcar / Desmarcar todos
$BtnMarcarTodos.Add_Click({
    $AppCheckPanel.Children | ForEach-Object { $_.IsChecked = $true }
})
$BtnDesmarcarTodos.Add_Click({
    $AppCheckPanel.Children | ForEach-Object { $_.IsChecked = $false }
})

# Criar atalhos marcados
$BtnCriarAtalho.Add_Click({
    $caminhos = @($AppCheckPanel.Children | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag })
    if ($caminhos.Count -eq 0) { Write-Log "Marque ao menos um aplicativo." "AVISO"; return }
    Invoke-Async {
        foreach ($c in $Caminhos) { New-DesktopShortcut -LnkSource $c }
        Write-Log "$($Caminhos.Count) atalho(s) criado(s) na area de trabalho." "OK"
    } -Vars @{ Caminhos = $caminhos }
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

# --- Armazenamento: duplo clique para abrir ---
$LvFolders.Add_MouseDoubleClick({
    $sel = $LvFolders.SelectedItem
    if ($sel -and $sel.Caminho -and (Test-Path $sel.Caminho)) {
        Start-Process explorer.exe $sel.Caminho
    }
})
$LvFiles.Add_MouseDoubleClick({
    $sel = $LvFiles.SelectedItem
    if ($sel -and $sel.Caminho) {
        $pasta = Split-Path $sel.Caminho -Parent
        if (Test-Path $pasta) { Start-Process explorer.exe $pasta }
    }
})

# --- Log ---
$BtnLimparLog.Add_Click({ $LogBox.Items.Clear() })
$BtnExportLog.Add_Click({
    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $dest = "$script:REPORT_DIR\Next_Suporte_$($env:COMPUTERNAME)_$ts.txt"
    $LogBox.Items | ForEach-Object { $_.Text } | Out-File -FilePath $dest -Encoding UTF8
    Write-Log "Log exportado: $dest" "OK"
})

# ================================================================
# INICIAR
# ================================================================
Write-Log "NextTool v$script:VERSION iniciado em $env:COMPUTERNAME" "INFO"
[void]$Window.ShowDialog()
$LogTimer.Stop()
