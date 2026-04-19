#Requires -Version 5.1
# ================================================================
# NextTool v1.0 - Ferramenta de TI da Next
# github.com/matheusgabsilva/nexttool
#
# Uso via URL:
#   irm "https://raw.githubusercontent.com/matheusgabsilva/nexttool/master/nexttool.ps1" | iex
# ================================================================

Set-StrictMode -Off
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# === ELEVACAO DE PRIVILEGIO ===
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell "-NoProfile -ExecutionPolicy Bypass -Command `"irm 'https://raw.githubusercontent.com/matheusgabsilva/nexttool/master/nexttool.ps1' | iex`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ================================================================
# XAML - INTERFACE GRAFICA
# ================================================================
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="NextTool - Ferramenta de TI"
    Height="700" Width="920"
    MinHeight="600" MinWidth="800"
    WindowStartupLocation="CenterScreen"
    Background="#1A1A2E">

  <Window.Resources>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="#1A1A2E"/>
      <Setter Property="BorderBrush" Value="#3A3A5C"/>
      <Setter Property="BorderThickness" Value="0,1,0,0"/>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#22223A"/>
      <Setter Property="Foreground" Value="#9090B0"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="border" Background="{TemplateBinding Background}"
                    Padding="{TemplateBinding Padding}" BorderThickness="0,0,0,3"
                    BorderBrush="Transparent">
              <ContentPresenter ContentSource="Header"
                                HorizontalAlignment="Center" VerticalAlignment="Center"
                                TextElement.Foreground="{TemplateBinding Foreground}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="border" Property="Background" Value="#1A1A2E"/>
                <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Background" Value="#252540"/>
                <Setter Property="Foreground" Value="#D0D0F0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#D0D0E8"/>
      <Setter Property="Margin" Value="4,5"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="#D0D0E8"/>
      <Setter Property="Margin" Value="4,5"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="GroupBox">
      <Setter Property="Foreground" Value="#8080A8"/>
      <Setter Property="BorderBrush" Value="#3A3A5C"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Margin" Value="0,6"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background" Value="#0078D4"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Padding" Value="18,8"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1188E0"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#005FA3"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Background" Value="#3A3A5C"/>
                <Setter Property="Foreground" Value="#606070"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnSecondary" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#3A3A5C"/>
      <Setter Property="Padding" Value="14,7"/>
    </Style>

    <Style x:Key="BtnDanger" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#C42B1C"/>
    </Style>

    <Style x:Key="InputBox" TargetType="TextBox">
      <Setter Property="Background" Value="#252540"/>
      <Setter Property="Foreground" Value="#E0E0F0"/>
      <Setter Property="BorderBrush" Value="#3A3A5C"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="CaretBrush" Value="White"/>
    </Style>

    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#8080A8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="ScrollViewer">
      <Setter Property="Background" Value="Transparent"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="65"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="155"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#12122A">
      <Grid Margin="20,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Next" FontSize="26" FontWeight="Black" Foreground="#0078D4"/>
          <TextBlock Text="Tool" FontSize="26" FontWeight="Black" Foreground="#FFFFFF"/>
          <TextBlock Text=" — Ferramenta de TI" FontSize="13"
                     Foreground="#60607A" VerticalAlignment="Bottom" Margin="6,0,0,3"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <TextBlock x:Name="txtComputador" Foreground="#505068" FontSize="11" VerticalAlignment="Center" Margin="0,0,15,0"/>
          <TextBlock Text="v1.0" Foreground="#404058" FontSize="11" VerticalAlignment="Center"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- TABS -->
    <TabControl Grid.Row="1" x:Name="tabMain" Margin="8,0,8,0">

      <!-- ======================================================
           ABA 0: SISTEMA
      ====================================================== -->
      <TabItem Header="  Sistema  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
          <StackPanel Margin="22,16,22,16">

            <TextBlock Text="Informações do Sistema" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>

            <!-- Cards de info rapida -->
            <Grid Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <Border Grid.Column="0" Background="#12122A" CornerRadius="6" Padding="14,10">
                <StackPanel>
                  <TextBlock Text="COMPUTADOR" FontSize="10" Foreground="#606080" FontWeight="Bold" Margin="0,0,0,4"/>
                  <TextBlock x:Name="cardPC" Text="—" FontSize="16" FontWeight="Bold" Foreground="#FFFFFF" TextWrapping="Wrap"/>
                  <TextBlock x:Name="cardOS" Text="—" FontSize="11" Foreground="#8080A8" Margin="0,3,0,0" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>

              <Border Grid.Column="2" Background="#12122A" CornerRadius="6" Padding="14,10">
                <StackPanel>
                  <TextBlock Text="PROCESSADOR" FontSize="10" Foreground="#606080" FontWeight="Bold" Margin="0,0,0,4"/>
                  <TextBlock x:Name="cardCPU" Text="—" FontSize="13" FontWeight="SemiBold" Foreground="#FFFFFF" TextWrapping="Wrap"/>
                  <TextBlock x:Name="cardRAM" Text="—" FontSize="11" Foreground="#8080A8" Margin="0,3,0,0"/>
                </StackPanel>
              </Border>

              <Border Grid.Column="4" Background="#12122A" CornerRadius="6" Padding="14,10">
                <StackPanel>
                  <TextBlock Text="ARMAZENAMENTO" FontSize="10" Foreground="#606080" FontWeight="Bold" Margin="0,0,0,4"/>
                  <TextBlock x:Name="cardDisk" Text="—" FontSize="13" FontWeight="SemiBold" Foreground="#FFFFFF" TextWrapping="Wrap"/>
                  <TextBlock x:Name="cardUser" Text="—" FontSize="11" Foreground="#8080A8" Margin="0,3,0,0"/>
                </StackPanel>
              </Border>
            </Grid>

            <!-- Detalhes completos -->
            <GroupBox Header="Detalhes">
              <StackPanel>
                <TextBlock x:Name="txtSysInfo" Foreground="#C0C0E0" FontFamily="Consolas"
                           FontSize="12" TextWrapping="Wrap" LineHeight="22"/>
              </StackPanel>
            </GroupBox>

            <!-- Seguranca -->
            <GroupBox Header="Segurança">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" x:Name="cardDefender" Background="#1A2A1A" CornerRadius="4" Padding="12,8">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="cardDefenderIcon" Text="●" Foreground="#50D090" FontSize="14" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <StackPanel>
                      <TextBlock Text="Windows Defender" FontSize="12" FontWeight="Bold" Foreground="#FFFFFF"/>
                      <TextBlock x:Name="cardDefenderTxt" Text="Verificando..." FontSize="11" Foreground="#8080A8"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
                <Border Grid.Column="2" x:Name="cardFW" Background="#1A2A1A" CornerRadius="4" Padding="12,8">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="cardFWIcon" Text="●" Foreground="#50D090" FontSize="14" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <StackPanel>
                      <TextBlock Text="Firewall" FontSize="12" FontWeight="Bold" Foreground="#FFFFFF"/>
                      <TextBlock x:Name="cardFWTxt" Text="Verificando..." FontSize="11" Foreground="#8080A8"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
              </Grid>
            </GroupBox>

            <Button x:Name="btnSysInfo" Content="↻  Atualizar" Style="{StaticResource BtnSecondary}"
                    HorizontalAlignment="Left" Margin="0,4,0,0"/>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ======================================================
           ABA 1: INSTALACOES
      ====================================================== -->
      <TabItem Header="  Instalações  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
          <StackPanel Margin="22,16,22,16">

            <TextBlock Text="Instalação de Softwares" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>

            <GroupBox Header="Utilitários">
              <StackPanel>
                <CheckBox x:Name="chkAcrobat"    Content="Adobe Acrobat Reader DC"/>
                <CheckBox x:Name="chkWinRAR"     Content="WinRAR"/>
                <CheckBox x:Name="chkAnyDesk"    Content="AnyDesk"/>
                <CheckBox x:Name="chkTeamViewer" Content="TeamViewer"/>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="Microsoft Office  —  selecione a versão">
              <StackPanel>
                <RadioButton x:Name="rdoOfficeNone"  Content="Não instalar Office"                      GroupName="Office" IsChecked="True"/>
                <RadioButton x:Name="rdoOffice365"   Content="Microsoft 365  (requer assinatura ativa)" GroupName="Office"/>
                <RadioButton x:Name="rdoOffice2021"  Content="Office 2021 Professional Plus"            GroupName="Office"/>
                <RadioButton x:Name="rdoOffice2016"  Content="Office 2016 Professional Plus"            GroupName="Office"/>
                <TextBlock Margin="22,6,0,0" FontSize="11" Foreground="#FFA040" TextWrapping="Wrap"
                           Text="Office 2021/2016 é instalado via ODT (Office Deployment Tool). Requer licença de volume (KMS/MAK) ou ativação manual posterior."/>
              </StackPanel>
            </GroupBox>

            <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
              <Button x:Name="btnInstall"   Content="▶  Instalar Selecionados" Style="{StaticResource BtnPrimary}" Margin="0,0,10,0"/>
              <Button x:Name="btnSelAll"    Content="Marcar Todos"             Style="{StaticResource BtnSecondary}" Margin="0,0,6,0"/>
              <Button x:Name="btnDeselAll"  Content="Desmarcar Todos"          Style="{StaticResource BtnSecondary}"/>
            </StackPanel>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ======================================================
           ABA 2: TWEAKS
      ====================================================== -->
      <TabItem Header="  Tweaks  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
          <StackPanel Margin="22,16,22,16">

            <TextBlock Text="Ajustes do Sistema" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>

            <GroupBox Header="Energia e Performance">
              <StackPanel>
                <CheckBox x:Name="chkHibernacao" Content="Desativar Hibernação"
                          ToolTip="powercfg -h off — libera espaço em disco e melhora boot"/>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="Segurança (Windows 11)">
              <StackPanel>
                <CheckBox x:Name="chkSmartApp" Content="Desativar Controle Inteligente de Aplicativos (Smart App Control)"
                          ToolTip="Recomendado em ambientes corporativos com antivírus gerenciado. Exige reinicialização."/>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="Drivers">
              <StackPanel>
                <CheckBox x:Name="chkDrivers" Content="Atualizar Drivers via Windows Update (PSWindowsUpdate)"
                          ToolTip="Instala o módulo PSWindowsUpdate e busca atualizações de drivers disponíveis"/>
              </StackPanel>
            </GroupBox>

            <Button x:Name="btnTweaks" Content="▶  Aplicar Tweaks Selecionados"
                    Style="{StaticResource BtnPrimary}" HorizontalAlignment="Left" Margin="0,14,0,0"/>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ======================================================
           ABA 3: MANUTENCAO
      ====================================================== -->
      <TabItem Header="  Manutenção  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
          <StackPanel Margin="22,16,22,16">

            <TextBlock Text="Manutenção e Diagnóstico" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>

            <GroupBox Header="Ações Rápidas">
              <WrapPanel>
                <Button x:Name="btnOtimizar"    Content="🧹  Otimizar PC"         Style="{StaticResource BtnPrimary}"   Margin="0,4,8,4"/>
                <Button x:Name="btnDiagnostico" Content="🔍  Diagnóstico"          Style="{StaticResource BtnPrimary}"   Margin="0,4,8,4"/>
                <Button x:Name="btnSFCDISM"     Content="🛠  SFC + DISM"           Style="{StaticResource BtnSecondary}" Margin="0,4,8,4"/>
                <Button x:Name="btnFlushDNS"    Content="🌐  Flush DNS"            Style="{StaticResource BtnSecondary}" Margin="0,4,8,4"/>
                <Button x:Name="btnRelatorio"   Content="📄  Abrir Relatórios"     Style="{StaticResource BtnSecondary}" Margin="0,4,8,4"/>
              </WrapPanel>
            </GroupBox>


          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ======================================================
           ABA 4: REDE / DOMINIO
      ====================================================== -->
      <TabItem Header="  Rede / Domínio  ">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
          <StackPanel Margin="22,16,22,16">

            <TextBlock Text="Rede e Domínio" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>

            <!-- DNS -->
            <GroupBox Header="Configurar DNS">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="140"/>
                  <ColumnDefinition Width="*" MaxWidth="300"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Label   Grid.Row="0" Grid.Column="0" Content="Adaptador de Rede:"/>
                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtAdapter"
                         Style="{StaticResource InputBox}" Text="Ethernet" Margin="0,3"/>

                <Label   Grid.Row="1" Grid.Column="0" Content="DNS Primário:"/>
                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtDNS1"
                         Style="{StaticResource InputBox}" Text="8.8.8.8" Margin="0,3"/>

                <Label   Grid.Row="2" Grid.Column="0" Content="DNS Secundário:"/>
                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtDNS2"
                         Style="{StaticResource InputBox}" Text="8.8.4.4" Margin="0,3"/>

                <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,8,0,0">
                  <Button x:Name="btnSetDNS"   Content="Aplicar DNS"    Style="{StaticResource BtnPrimary}"   Margin="0,0,8,0"/>
                  <Button x:Name="btnResetDNS" Content="Resetar (DHCP)" Style="{StaticResource BtnSecondary}"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <!-- Diagnostico de Rede -->
            <GroupBox Header="Diagnóstico de Rede">
              <WrapPanel>
                <Button x:Name="btnPing"     Content="Testar Conectividade" Style="{StaticResource BtnSecondary}" Margin="0,4,8,4"/>
                <Button x:Name="btnIPConfig" Content="Ver IP / Config"       Style="{StaticResource BtnSecondary}" Margin="0,4,8,4"/>
              </WrapPanel>
            </GroupBox>

            <!-- Dominio AD -->
            <GroupBox Header="Ingresso em Domínio AD  —  opcional">
              <StackPanel>
                <TextBlock Foreground="#8080A8" FontSize="11" Margin="0,0,0,10" TextWrapping="Wrap"
                           Text="Preencha apenas se necessário ingressar esta máquina no domínio. Campos em branco serão ignorados."/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="140"/>
                    <ColumnDefinition Width="*" MaxWidth="300"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <Label   Grid.Row="0" Grid.Column="0" Content="Domínio:"/>
                  <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtDomain"
                           Style="{StaticResource InputBox}" Margin="0,3"/>

                  <Label   Grid.Row="1" Grid.Column="0" Content="Usuário AD:"/>
                  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtDomainUser"
                           Style="{StaticResource InputBox}" Margin="0,3"/>

                  <Label   Grid.Row="2" Grid.Column="0" Content="Senha AD:"/>
                  <PasswordBox Grid.Row="2" Grid.Column="1" x:Name="txtDomainPass"
                               Background="#252540" Foreground="#E0E0F0" BorderBrush="#3A3A5C"
                               BorderThickness="1" Padding="8,6" FontSize="13" Margin="0,3"/>

                  <Label   Grid.Row="3" Grid.Column="0" Content="Renomear PC para:"/>
                  <TextBox Grid.Row="3" Grid.Column="1" x:Name="txtNewName"
                           Style="{StaticResource InputBox}" Margin="0,3"/>

                  <Button  Grid.Row="4" Grid.Column="1" x:Name="btnJoinDomain"
                           Content="Ingressar no Domínio"
                           Style="{StaticResource BtnPrimary}" HorizontalAlignment="Left" Margin="0,10,0,0"/>
                </Grid>
              </StackPanel>
            </GroupBox>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ======================================================
           ABA 5: ESPACO EM DISCO
      ====================================================== -->
      <TabItem Header="  Espaço em Disco  ">
        <Grid Background="#1A1A2E">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- Controles de scan -->
          <StackPanel Grid.Row="0" Margin="22,16,22,8">
            <TextBlock Text="Análise de Espaço em Disco" FontSize="15" FontWeight="Bold"
                       Foreground="#0078D4" Margin="0,0,0,12"/>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Label Content="Caminho:" Margin="0,0,6,0"/>
              <TextBox x:Name="txtScanPath" Style="{StaticResource InputBox}"
                       Text="C:\" Width="220" Margin="0,0,10,0"/>
              <Button x:Name="btnScan" Content="🔍  Escanear" Style="{StaticResource BtnPrimary}" Margin="0,0,10,0"/>
              <Button x:Name="btnScanClear" Content="Limpar" Style="{StaticResource BtnSecondary}"/>
            </StackPanel>
            <TextBlock Text="Exibe pastas de primeiro nível com tamanho recursivo + maiores arquivos individuais (acima de 50 MB)."
                       Foreground="#606080" FontSize="11" Margin="0,8,0,0"/>
          </StackPanel>

          <!-- Resultados -->
          <Grid Grid.Row="1" Margin="22,4,22,12">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="10"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Grid.Column="0" Header="Pastas mais pesadas">
              <ListBox x:Name="lstFolders" Background="#0F0F1E" Foreground="#C0C0E0"
                       BorderThickness="0" FontFamily="Consolas" FontSize="11"
                       ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
            </GroupBox>

            <GroupBox Grid.Column="2" Header="Maiores arquivos  (acima de 50 MB)">
              <ListBox x:Name="lstFiles" Background="#0F0F1E" Foreground="#C0C0E0"
                       BorderThickness="0" FontFamily="Consolas" FontSize="11"
                       ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
            </GroupBox>
          </Grid>
        </Grid>
      </TabItem>

    </TabControl>

    <!-- LOG PANEL -->
    <Border Grid.Row="2" Background="#0F0F1E" BorderBrush="#3A3A5C" BorderThickness="0,1,0,0" Margin="8,0,8,8">
      <Grid Margin="10,7,10,7">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,4">
          <TextBlock Text="LOG" FontSize="10" FontWeight="Bold" Foreground="#404058"
                     VerticalAlignment="Center" Margin="0,0,8,0"/>
          <Button x:Name="btnClearLog" Content="Limpar" FontSize="10" Padding="6,2"
                  Background="#252540" Foreground="#606080" BorderThickness="0" Cursor="Hand"/>
        </StackPanel>
        <ScrollViewer Grid.Row="1" x:Name="logScroll" VerticalScrollBarVisibility="Auto">
          <TextBlock x:Name="txtLog" Foreground="#50D090" FontFamily="Consolas"
                     FontSize="11" TextWrapping="Wrap"/>
        </ScrollViewer>
      </Grid>
    </Border>

  </Grid>
</Window>
"@

# ================================================================
# PARSE XAML E REFERENCIAS
# ================================================================
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

function Get-Control { param($name) $window.FindName($name) }

$txtLog         = Get-Control "txtLog"
$logScroll      = Get-Control "logScroll"
$txtComputador  = Get-Control "txtComputador"
$txtSysInfo     = Get-Control "txtSysInfo"
$cardPC         = Get-Control "cardPC"
$cardOS         = Get-Control "cardOS"
$cardCPU        = Get-Control "cardCPU"
$cardRAM        = Get-Control "cardRAM"
$cardDisk       = Get-Control "cardDisk"
$cardUser       = Get-Control "cardUser"
$cardDefenderIcon = Get-Control "cardDefenderIcon"
$cardDefenderTxt  = Get-Control "cardDefenderTxt"
$cardDefender     = Get-Control "cardDefender"
$cardFWIcon     = Get-Control "cardFWIcon"
$cardFWTxt      = Get-Control "cardFWTxt"
$cardFW         = Get-Control "cardFW"

# Instalacoes
$chkAcrobat     = Get-Control "chkAcrobat"
$chkWinRAR      = Get-Control "chkWinRAR"
$chkAnyDesk     = Get-Control "chkAnyDesk"
$chkTeamViewer  = Get-Control "chkTeamViewer"
$rdoOfficeNone  = Get-Control "rdoOfficeNone"
$rdoOffice365   = Get-Control "rdoOffice365"
$rdoOffice2021  = Get-Control "rdoOffice2021"
$rdoOffice2016  = Get-Control "rdoOffice2016"
$btnInstall     = Get-Control "btnInstall"
$btnSelAll      = Get-Control "btnSelAll"
$btnDeselAll    = Get-Control "btnDeselAll"

# Tweaks
$chkHibernacao  = Get-Control "chkHibernacao"
$chkSmartApp    = Get-Control "chkSmartApp"
$chkDrivers     = Get-Control "chkDrivers"
$btnTweaks      = Get-Control "btnTweaks"

# Manutencao
$btnOtimizar    = Get-Control "btnOtimizar"
$btnDiagnostico = Get-Control "btnDiagnostico"
$btnSFCDISM     = Get-Control "btnSFCDISM"
$btnFlushDNS    = Get-Control "btnFlushDNS"
$btnRelatorio   = Get-Control "btnRelatorio"
$btnSysInfo     = Get-Control "btnSysInfo"

# Rede / Dominio
$txtAdapter     = Get-Control "txtAdapter"
$txtDNS1        = Get-Control "txtDNS1"
$txtDNS2        = Get-Control "txtDNS2"
$btnSetDNS      = Get-Control "btnSetDNS"
$btnResetDNS    = Get-Control "btnResetDNS"
$btnPing        = Get-Control "btnPing"
$btnIPConfig    = Get-Control "btnIPConfig"
$txtDomain      = Get-Control "txtDomain"
$txtDomainUser  = Get-Control "txtDomainUser"
$txtDomainPass  = Get-Control "txtDomainPass"
$txtNewName     = Get-Control "txtNewName"
$btnJoinDomain  = Get-Control "btnJoinDomain"
$btnClearLog    = Get-Control "btnClearLog"

# Espaco em Disco
$txtScanPath    = Get-Control "txtScanPath"
$btnScan        = Get-Control "btnScan"
$btnScanClear   = Get-Control "btnScanClear"
$lstFolders     = Get-Control "lstFolders"
$lstFiles       = Get-Control "lstFiles"
$global:syncHash.LstFolders = $lstFolders
$global:syncHash.LstFiles   = $lstFiles

# ================================================================
# SISTEMA DE LOG + RUNSPACE
# ================================================================
$global:syncHash = [hashtable]::Synchronized(@{
    Log      = $txtLog
    Scroll   = $logScroll
    Running  = $false
    Queue    = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
})

$global:relatorio = [System.Collections.Generic.List[string]]::new()
$global:dataExec  = Get-Date -Format "yyyy-MM-dd_HH-mm"
$global:nomePC    = $env:COMPUTERNAME

function Write-GuiLog {
    param([string]$msg, [string]$color = "White")
    $ts   = [datetime]::Now.ToString("HH:mm:ss")
    $line = "[$ts]  $msg"
    $global:relatorio.Add($line)
    $capturedLine = $line
    $sh = $global:syncHash
    $sb = [System.Action]{ $sh.Log.Text += "$capturedLine`n"; $sh.Scroll.ScrollToBottom() }
    if ($sh.Log.Dispatcher.CheckAccess()) { $sh.Log.Text += "$capturedLine`n"; $sh.Scroll.ScrollToBottom() }
    else { $sh.Log.Dispatcher.Invoke($sb) }
}

function QLog { param($m) if ($syncHash) { $syncHash.Queue.Enqueue($m) } }

function Invoke-BackgroundTask {
    param(
        [scriptblock]$Task,
        [string]$Label = "Tarefa",
        [hashtable]$Vars = @{},
        [scriptblock]$OnDone = $null
    )
    $global:syncHash.OnDoneCallback = $OnDone

    if ($global:syncHash.Running) {
        Write-GuiLog "Aguarde a operacao atual terminar..." "Orange"
        return
    }

    $global:syncHash.Running = $true
    $global:syncHash.Queue   = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    Write-GuiLog "Iniciando: $Label" "Cyan"

    # Serializa todas as funcoes customizadas para o runspace
    $fnNames = @(
        "Format-Size","Invoke-EspacoDisco",
        "Install-Winget","Install-WingetApp","Install-Office",
        "Invoke-TweakHibernacao","Invoke-TweakSmartApp","Invoke-TweakDrivers",
        "Invoke-OtimizarPC","Invoke-Diagnostico","Invoke-SFCDISM",
        "Get-NicConfig","Invoke-SetDNS","Invoke-ResetDNS","Invoke-TestarConectividade","Invoke-JoinDomain"
    )
    $fnDefs = ($fnNames | ForEach-Object {
        $fn = Get-Item "function:$_" -ErrorAction SilentlyContinue
        if ($fn) { "function $_ {`n$($fn.ScriptBlock)`n}" }
    }) -join "`n`n"

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("syncHash",    $global:syncHash)
    $rs.SessionStateProxy.SetVariable("task",        $Task)
    $rs.SessionStateProxy.SetVariable("fnDefs",      $fnDefs)
    $rs.SessionStateProxy.SetVariable("rsOfficeXML", $script:OfficeXML)
    foreach ($kv in $Vars.GetEnumerator()) {
        $rs.SessionStateProxy.SetVariable($kv.Key, $kv.Value)
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript({
        # QLog global: visivel em qualquer modulo/escopo do runspace
        function global:QLog { param($m) if ($syncHash) { $syncHash.Queue.Enqueue($m) } }
        Invoke-Expression $fnDefs
        $script:OfficeXML = $rsOfficeXML
        try { & $task } catch { $syncHash.Queue.Enqueue("[ERRO] $_") }
        $syncHash.Queue.Enqueue("__DONE__")
    })

    $handle = $ps.BeginInvoke()

    # Armazena referencias no syncHash para o tick acessar sem depender de closure
    $global:syncHash.ActiveTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:syncHash.ActiveTimer.Interval = [timespan]::FromMilliseconds(250)
    $global:syncHash.ActivePS = $ps
    $global:syncHash.ActiveRS = $rs

    $global:syncHash.ActiveTimer.Add_Tick({
        $msg = ""
        while ($global:syncHash.Queue.TryDequeue([ref]$msg)) {
            $ts = [datetime]::Now.ToString("HH:mm:ss")
            if ($msg -eq "__DONE__") {
                $global:syncHash.ActiveTimer.Stop()
                $global:syncHash.Running = $false
                $entry = "[$ts]  Concluido."
                $global:relatorio.Add($entry)
                $global:syncHash.Log.Text  += "$entry`n"
                $global:syncHash.Scroll.ScrollToBottom()
                try { $global:syncHash.ActivePS.Dispose(); $global:syncHash.ActiveRS.Dispose() } catch {}
                if ($global:syncHash.OnDoneCallback) {
                    try { & $global:syncHash.OnDoneCallback } catch {}
                    $global:syncHash.OnDoneCallback = $null
                }
                return
            }
            $display = $msg -replace "^\[\w+\]\s*",""
            $entry   = "[$ts]  $display"
            $global:relatorio.Add($entry)
            $global:syncHash.Log.Text  += "$entry`n"
            $global:syncHash.Scroll.ScrollToBottom()
        }
    })

    $global:syncHash.ActiveTimer.Start()
}

# ================================================================
# WINGET - VERIFICACAO E INSTALACAO
# ================================================================
function Test-Winget {
    return ($null -ne (Get-Command winget -ErrorAction SilentlyContinue))
}

function Install-Winget {
    QLog "[INFO] Instalando winget (App Installer)..."
    try {
        $url = "https://aka.ms/getwinget"
        $tmp = "$env:TEMP\AppInstaller.msixbundle"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
        Add-AppxPackage -Path $tmp
        QLog "[OK] winget instalado."
    } catch {
        QLog "[ERRO] Falha ao instalar winget: $_"
    }
}

function Install-WingetApp {
    param([string]$Id, [string]$Name)
    QLog "[INFO] Instalando $Name..."
    $result = winget install --id $Id -e --accept-source-agreements --accept-package-agreements --silent 2>&1
    if ($LASTEXITCODE -eq 0 -or $result -match "instalado|already installed|No applicable") {
        QLog "[OK] $Name instalado com sucesso."
    } else {
        QLog "[ERRO] Falha ao instalar $Name. Codigo: $LASTEXITCODE"
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
    QLog "[INFO] Preparando instalacao do Office $Version via ODT..."

    $odtDir = "$env:TEMP\NextODT"
    New-Item -Path $odtDir -ItemType Directory -Force | Out-Null

    $odtExe = "$odtDir\setup.exe"
    QLog "[INFO] Baixando Office Deployment Tool..."
    try {
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17928-20114.exe" `
            -OutFile "$odtDir\odt.exe" -UseBasicParsing
        Start-Process "$odtDir\odt.exe" -ArgumentList "/quiet /extract:$odtDir" -Wait
        QLog "[OK] ODT extraido."
    } catch {
        QLog "[ERRO] Falha ao baixar ODT: $_"
        return
    }

    $xmlPath = "$odtDir\config_$Version.xml"
    $script:OfficeXML[$Version] | Out-File -FilePath $xmlPath -Encoding UTF8
    QLog "[INFO] Iniciando instalacao Office $Version (pode demorar varios minutos)..."
    Start-Process $odtExe -ArgumentList "/configure `"$xmlPath`"" -Wait
    QLog "[OK] Instalacao do Office $Version concluida."
}

# ================================================================
# TWEAKS
# ================================================================
function Invoke-TweakHibernacao {
    QLog "[INFO] Desativando hibernacao..."
    powercfg -h off 2>&1 | Out-Null
    QLog "[OK] Hibernacao desativada."
}

function Invoke-TweakSmartApp {
    QLog "[INFO] Desativando Smart App Control..."
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
    try {
        Set-ItemProperty -Path $path -Name "VerifiedAndReputablePolicyState" -Value 0 -Type DWord -Force
        QLog "[OK] Smart App Control desativado. Reinicie para aplicar."
    } catch {
        QLog "[AVISO] Nao foi possivel desativar (pode nao existir neste Windows): $_"
    }
}

function Invoke-TweakDrivers {
    QLog "[INFO] Verificando modulo PSWindowsUpdate..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        QLog "[INFO] Instalando PSWindowsUpdate..."
        try {
            Install-Module PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
            QLog "[OK] Modulo instalado."
        } catch {
            QLog "[ERRO] Falha ao instalar modulo: $_"
            return
        }
    }
    Import-Module PSWindowsUpdate -Force
    QLog "[INFO] Buscando atualizacoes de drivers..."
    try {
        $updates = Get-WindowsUpdate -Category Drivers -ErrorAction Stop
        if ($updates.Count -eq 0) {
            QLog "[OK] Nenhuma atualizacao de driver disponivel. Drivers em dia!"
        } else {
            QLog "[INFO] $($updates.Count) driver(s) encontrado(s). Instalando..."
            $updates | ForEach-Object { QLog "[INFO]  - $($_.Title)" }
            Install-WindowsUpdate -Category Drivers -AcceptAll -IgnoreReboot -Verbose 2>&1 |
                ForEach-Object { QLog "[INFO] $_" }
            QLog "[OK] Drivers atualizados. Reinicie para aplicar."
        }
    } catch {
        QLog "[ERRO] Erro ao buscar drivers: $_"
    }
}

# ================================================================
# MANUTENCAO (NextTI-Manutencao integrado)
# ================================================================
function Invoke-OtimizarPC {
    QLog "[INFO] Limpando arquivos temporarios..."
    $paths = @($env:TEMP, "C:\Windows\Temp", "C:\Windows\Prefetch")
    $total = 0
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $files = Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue
            $total += $files.Count
            Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    QLog "[OK] $total arquivo(s) temporario(s) removido(s)."

    QLog "[INFO] Esvaziando lixeira..."
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    QLog "[OK] Lixeira esvaziada."

    QLog "[INFO] Executando Limpeza de Disco..."
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    @("Temporary Files","Internet Cache Files","Recycle Bin","Thumbnail Cache",
      "Downloaded Program Files","Memory Dump Files","Old ChkDsk Files",
      "Setup Log Files","System error memory dump files","Update Cleanup") | ForEach-Object {
        $p = "$regPath\$_"
        if (Test-Path $p) { Set-ItemProperty -Path $p -Name StateFlags0064 -Value 2 -ErrorAction SilentlyContinue }
    }
    $cg = Start-Process cleanmgr -ArgumentList "/sagerun:64" -PassThru -WindowStyle Hidden
    $done = $cg.WaitForExit(60000)
    if (-not $done) { try { $cg.Kill() } catch {} ; QLog "[INFO] Limpeza de disco encerrada (timeout 60s)." }
    else { QLog "[OK] Limpeza de disco concluida." }

    QLog "[INFO] Limpando cache DNS..."
    ipconfig /flushdns | Out-Null
    QLog "[OK] Cache DNS limpo."

    $totalRam = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
    $freeRam  = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB
    QLog "[INFO] RAM: Total $([math]::Round($totalRam))MB | Livre $([math]::Round($freeRam))MB | Usada $([math]::Round($totalRam - $freeRam))MB"

    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object {
        $total = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
        $used  = [math]::Round($_.Used / 1GB, 1)
        $free  = [math]::Round($_.Free / 1GB, 1)
        QLog "[INFO] Disco $($_.Name): ${total}GB total | ${used}GB usado | ${free}GB livre"
    }
    QLog "[OK] Otimizacao concluida!"
}

function Invoke-Diagnostico {
    QLog "[INFO] === TOP 10 PROCESSOS POR RAM ==="
    Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 | ForEach-Object {
        $mem = [math]::Round($_.WorkingSet64 / 1MB, 1)
        QLog "[INFO]  $($_.Name.PadRight(28)) $($mem.ToString().PadLeft(8)) MB"
    }

    QLog "[INFO] === PROGRAMAS NA INICIALIZACAO ==="
    @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
      "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run") | ForEach-Object {
        $scope = if ($_ -like "*HKLM*") { "Sistema" } else { "Usuario" }
        if (Test-Path $_) {
            Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue |
                Get-Member -MemberType NoteProperty |
                Where-Object { $_.Name -notlike "PS*" } |
                ForEach-Object { QLog "[INFO]  [$scope] $($_.Name)" }
        }
    }

    QLog "[INFO] === SEGURANCA ==="
    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        $est = if ($def.AntivirusEnabled) { "[OK] ATIVO" } else { "[ERRO] INATIVO" }
        QLog "$est  Windows Defender | Definicoes: $($def.AntivirusSignatureLastUpdated.ToString('dd/MM/yyyy'))"
    } catch { QLog "[AVISO] Nao foi possivel verificar o Defender." }

    try {
        $fwOut = netsh advfirewall show allprofiles state 2>&1 | Out-String
        $onCount = ([regex]::Matches($fwOut, "(?i)State\s+ON")).Count
        $s = if ($onCount -gt 0) { "[OK]" } else { "[ERRO]" }
        QLog "$s  Firewall: $onCount perfil(is) ativo(s)"
    } catch { QLog "[AVISO] Nao foi possivel verificar Firewall." }

    QLog "[INFO] === ERROS CRITICOS (ULTIMAS 24H) ==="
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=([datetime]::Now.AddHours(-24))} `
            -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($events -and $events.Count -gt 0) {
            QLog "[AVISO] $($events.Count) erro(s) critico(s) nas ultimas 24h:"
            $events | ForEach-Object {
                $msg = ($_.Message.Split("`n")[0]).Substring(0,[Math]::Min(90,$_.Message.Split("`n")[0].Length))
                QLog "[AVISO]  [$($_.TimeCreated.ToString('HH:mm'))] $($_.ProviderName): $msg"
            }
        } else {
            QLog "[OK] Nenhum erro critico nas ultimas 24h."
        }
    } catch { QLog "[AVISO] Nao foi possivel verificar Event Viewer." }

    QLog "[OK] Diagnostico concluido!"
}

function Invoke-SFCDISM {
    QLog "[INFO] Executando SFC /scannow (aguarde)..."
    $sfc = sfc /scannow 2>&1 | Out-String
    if ($sfc -match "encontrou" -or $sfc -match "found") {
        QLog "[OK] SFC: problemas encontrados e corrigidos."
    } elseif ($sfc -match "nao encontrou" -or $sfc -match "did not find") {
        QLog "[OK] SFC: nenhuma violacao de integridade encontrada."
    } else {
        QLog "[INFO] SFC concluido."
    }

    QLog "[INFO] Executando DISM RestoreHealth (aguarde)..."
    $dism = DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
    if ($dism -match "concluida|successfully") {
        QLog "[OK] DISM concluido com sucesso."
    } else {
        QLog "[AVISO] DISM: verifique o resultado manualmente."
    }
}

function Update-SysInfo {
    $bc = [System.Windows.Media.BrushConverter]::new()
    $brushGreen  = $bc.ConvertFromString("#50D090")
    $brushRed    = $bc.ConvertFromString("#C42B1C")
    $brushBgGood = $bc.ConvertFromString("#1A2A1A")
    $brushBgBad  = $bc.ConvertFromString("#2A1A1A")

    try {
        $os    = Get-CimInstance Win32_OperatingSystem
        $cs    = Get-CimInstance Win32_ComputerSystem
        $cpu   = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
        $ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        $freeGB= [math]::Round($os.FreePhysicalMemory / 1MB / 1024, 1)
        $usedGB= [math]::Round($ramGB - $freeGB, 1)
        $disks = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 }
        $diskSummary = ($disks | ForEach-Object { "$($_.Name): $([math]::Round($_.Free/1GB,1))GB livre" }) -join "  |  "
        $diskDetail  = ($disks | ForEach-Object {
            $t = [math]::Round(($_.Used+$_.Free)/1GB,1)
            $u = [math]::Round($_.Used/1GB,1)
            $f = [math]::Round($_.Free/1GB,1)
            "$($_.Name):  ${t}GB total  |  ${u}GB usado  |  ${f}GB livre"
        }) -join "`n"

        $cardPC.Text   = $env:COMPUTERNAME
        $cardOS.Text   = "$($os.Caption) $($os.OSArchitecture)"
        $cardCPU.Text  = ($cpu -replace "\s{2,}"," ")
        $cardRAM.Text  = "RAM: ${ramGB}GB total  |  ${usedGB}GB usada  |  ${freeGB}GB livre"
        $cardDisk.Text = $diskSummary
        $cardUser.Text = "Usuário: $($env:USERNAME)"

        $txtSysInfo.Text = "Computador  : $($env:COMPUTERNAME)`nUsuário     : $($env:USERNAME)`nDomínio     : $($cs.Domain)`nSO          : $($os.Caption) $($os.OSArchitecture)`nVersão SO   : $($os.Version)`nCPU         : $($cpu -replace '\s{2,}',' ')`nRAM         : ${ramGB}GB total  |  ${usedGB}GB usada  |  ${freeGB}GB livre`n$diskDetail`nData/Hora   : $([datetime]::Now.ToString('dd/MM/yyyy HH:mm'))"
    } catch {}

    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        if ($def.AntivirusEnabled) {
            $cardDefenderIcon.Foreground = $brushGreen
            $cardDefender.Background    = $brushBgGood
            $cardDefenderTxt.Text = "Ativo  |  Def: $($def.AntivirusSignatureLastUpdated.ToString('dd/MM/yy'))"
        } else {
            $cardDefenderIcon.Foreground = $brushRed
            $cardDefender.Background    = $brushBgBad
            $cardDefenderTxt.Text = "INATIVO"
        }
    } catch { $cardDefenderTxt.Text = "Indisponível" }

    try {
        $fwOut = netsh advfirewall show allprofiles state 2>&1 | Out-String
        $onCount = ([regex]::Matches($fwOut, "(?i)State\s+ON")).Count
        if ($onCount -gt 0) {
            $cardFWIcon.Foreground = $brushGreen
            $cardFW.Background    = $brushBgGood
            $cardFWTxt.Text = "$onCount perfil(is) ativo(s)"
        } else {
            $cardFWIcon.Foreground = $brushRed
            $cardFW.Background    = $brushBgBad
            $cardFWTxt.Text = "INATIVO"
        }
    } catch { $cardFWTxt.Text = "Indisponível" }
}

# ================================================================
# ESPACO EM DISCO
# ================================================================
function Format-Size {
    param([long]$bytes)
    if ($bytes -ge 1GB) { return "$([math]::Round($bytes/1GB,2).ToString('0.00')) GB" }
    if ($bytes -ge 1MB) { return "$([math]::Round($bytes/1MB,1).ToString('0.0')) MB " }
    if ($bytes -ge 1KB) { return "$([math]::Round($bytes/1KB,0)) KB " }
    return "$bytes B  "
}

function Invoke-EspacoDisco {
    $path = $syncHash.ScanPath
    QLog "[INFO] Calculando tamanho das pastas em $path ..."

    $dirs = Get-ChildItem $path -Directory -ErrorAction SilentlyContinue
    $folderData = [System.Collections.Generic.List[object]]::new()
    $idx = 0
    foreach ($d in $dirs) {
        $idx++
        QLog "[INFO] [$idx/$($dirs.Count)] $($d.Name)"
        try {
            $sz = (Get-ChildItem $d.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object Length -Sum -ErrorAction SilentlyContinue).Sum
            if (-not $sz) { $sz = 0 }
            $folderData.Add([pscustomobject]@{ Bytes=$sz; Linha="$(Format-Size $sz)  $($d.FullName)" })
        } catch {}
    }
    $syncHash.ScanFolderList = @(
        $folderData | Sort-Object Bytes -Descending | Select-Object -First 25 | ForEach-Object { $_.Linha }
    )

    QLog "[INFO] Buscando arquivos acima de 50 MB..."
    $syncHash.ScanFileList = @(
        Get-ChildItem $path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -ge 50MB } |
            Sort-Object Length -Descending | Select-Object -First 30 |
            ForEach-Object { "$(Format-Size $_.Length)  $($_.FullName)" }
    )
    QLog "[OK] Scan concluido."
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

function Invoke-SetDNS {
    param([string]$Adapter, [string]$DNS1, [string]$DNS2)
    QLog "[INFO] Configurando DNS em '$Adapter' -> $DNS1 / $DNS2 ..."
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { QLog "[ERRO] Adaptador '$Adapter' nao encontrado."; return }
        $r = Invoke-CimMethod -InputObject $cfg -MethodName SetDNSServerSearchOrder `
             -Arguments @{ DNSServerSearchOrder = [string[]]@($DNS1, $DNS2) }
        if ($r.ReturnValue -eq 0) { QLog "[OK] DNS configurado: $DNS1 / $DNS2." }
        else { QLog "[AVISO] Concluido com codigo WMI: $($r.ReturnValue)" }
    } catch { QLog "[ERRO] $_" }
}

function Invoke-ResetDNS {
    param([string]$Adapter)
    QLog "[INFO] Resetando DNS de '$Adapter' para DHCP..."
    try {
        $cfg = Get-NicConfig -Adapter $Adapter
        if (-not $cfg) { QLog "[ERRO] Adaptador '$Adapter' nao encontrado."; return }
        $r = Invoke-CimMethod -InputObject $cfg -MethodName SetDNSServerSearchOrder `
             -Arguments @{ DNSServerSearchOrder = [string[]]$null }
        if ($r.ReturnValue -eq 0) { QLog "[OK] DNS resetado para DHCP automatico." }
        else { QLog "[AVISO] Concluido com codigo WMI: $($r.ReturnValue)" }
    } catch { QLog "[ERRO] $_" }
}

function Invoke-TestarConectividade {
    foreach ($target in @("8.8.8.8","1.1.1.1","google.com")) {
        $ping = New-Object System.Net.NetworkInformation.Ping
        try {
            $reply = $ping.Send($target, 3000)
            if ($reply.Status -eq "Success") {
                QLog "[OK] Ping ${target}: $($reply.RoundtripTime)ms"
            } else {
                QLog "[ERRO] Ping ${target}: $($reply.Status)"
            }
        } catch {
            QLog "[ERRO] Ping ${target} falhou: $_"
        }
    }
}

function Invoke-JoinDomain {
    param([string]$Domain,[string]$User,[string]$Pass,[string]$NewName)
    if (-not $Domain -or -not $User -or -not $Pass) {
        QLog "[ERRO] Preencha Dominio, Usuario e Senha."
        return
    }
    QLog "[INFO] Ingressando em $Domain como $User..."
    try {
        $cred = New-Object PSCredential("$Domain\$User", (ConvertTo-SecureString $Pass -AsPlainText -Force))
        if ($NewName) {
            Add-Computer -DomainName $Domain -Credential $cred -NewName $NewName -Force
            QLog "[OK] PC renomeado para '$NewName' e ingressado em $Domain."
        } else {
            Add-Computer -DomainName $Domain -Credential $cred -Force
            QLog "[OK] Ingressado em $Domain com sucesso."
        }
        QLog "[AVISO] Reinicie o computador para aplicar as alteracoes."
    } catch {
        QLog "[ERRO] Falha ao ingressar no dominio: $_"
    }
}

# ================================================================
# EVENT HANDLERS
# ================================================================

# --- Header: nome do PC ---
$txtComputador.Text = "$env:COMPUTERNAME  |  $env:USERNAME"

# --- Sys info ao abrir ---
$window.Add_Loaded({
    Update-SysInfo
})

# --- Limpar log ---
$btnClearLog.Add_Click({
    $txtLog.Text = ""
})

# --- Selecionar / Desselecionar todos ---
$btnSelAll.Add_Click({
    $chkAcrobat.IsChecked = $chkWinRAR.IsChecked = $chkAnyDesk.IsChecked = $chkTeamViewer.IsChecked = $true
})
$btnDeselAll.Add_Click({
    $chkAcrobat.IsChecked = $chkWinRAR.IsChecked = $chkAnyDesk.IsChecked = $chkTeamViewer.IsChecked = $false
    $rdoOfficeNone.IsChecked = $true
})

# --- Instalar ---
$btnInstall.Add_Click({
    $toInstall   = [System.Collections.Generic.List[object]]::new()
    $officeVer   = ""

    if ($chkAcrobat.IsChecked)    { $toInstall.Add(@{Id="Adobe.Acrobat.Reader.64-bit"; Name="Adobe Acrobat Reader"}) }
    if ($chkWinRAR.IsChecked)     { $toInstall.Add(@{Id="RARLab.WinRAR";               Name="WinRAR"}) }
    if ($chkAnyDesk.IsChecked)    { $toInstall.Add(@{Id="AnyDesk.AnyDesk";             Name="AnyDesk"}) }
    if ($chkTeamViewer.IsChecked) { $toInstall.Add(@{Id="TeamViewer.TeamViewer";        Name="TeamViewer"}) }
    if ($rdoOffice365.IsChecked)  { $officeVer = "365" }
    if ($rdoOffice2021.IsChecked) { $officeVer = "2021" }
    if ($rdoOffice2016.IsChecked) { $officeVer = "2016" }

    if ($toInstall.Count -eq 0 -and -not $officeVer) {
        Write-GuiLog "Selecione ao menos um software para instalar." "Orange"
        return
    }

    $global:syncHash.installList   = @($toInstall)
    $global:syncHash.installOffice = $officeVer

    Invoke-BackgroundTask -Label "Instalacao de Softwares" -Task {
        $list   = $syncHash.installList
        $office = $syncHash.installOffice
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { Install-Winget }
        foreach ($app in $list) { Install-WingetApp -Id $app.Id -Name $app.Name }
        if ($office) { Install-Office -Version $office }
    }
})

# --- Tweaks ---
$btnTweaks.Add_Click({
    $doHib     = $chkHibernacao.IsChecked
    $doSmart   = $chkSmartApp.IsChecked
    $doDrivers = $chkDrivers.IsChecked

    if (-not ($doHib -or $doSmart -or $doDrivers)) {
        Write-GuiLog "Selecione ao menos um tweak." "Orange"
        return
    }

    $global:syncHash.tweakHib     = [bool]$doHib
    $global:syncHash.tweakSmart   = [bool]$doSmart
    $global:syncHash.tweakDrivers = [bool]$doDrivers

    Invoke-BackgroundTask -Label "Aplicar Tweaks" -Task {
        if ($syncHash.tweakHib)     { Invoke-TweakHibernacao }
        if ($syncHash.tweakSmart)   { Invoke-TweakSmartApp }
        if ($syncHash.tweakDrivers) { Invoke-TweakDrivers }
    }
})

# --- Manutencao ---
$btnOtimizar.Add_Click({
    Invoke-BackgroundTask -Label "Otimizacao de PC" -Task { Invoke-OtimizarPC }
})

$btnDiagnostico.Add_Click({
    Invoke-BackgroundTask -Label "Diagnostico do Sistema" -Task { Invoke-Diagnostico }
})

$btnSFCDISM.Add_Click({
    Invoke-BackgroundTask -Label "SFC + DISM" -Task { Invoke-SFCDISM }
})

$btnFlushDNS.Add_Click({
    ipconfig /flushdns | Out-Null
    Write-GuiLog "Cache DNS limpo." "Green"
})

$btnRelatorio.Add_Click({
    $pasta = "C:\Next-Relatorios"
    if (-not (Test-Path $pasta)) { New-Item $pasta -ItemType Directory | Out-Null }
    Start-Process explorer.exe $pasta
})

$btnSysInfo.Add_Click({
    Update-SysInfo
    Write-GuiLog "Informacoes do sistema atualizadas." "Cyan"
})

# --- Rede / DNS ---
$btnSetDNS.Add_Click({
    $global:syncHash.dnsAdapter = $txtAdapter.Text.Trim()
    $global:syncHash.dns1 = $txtDNS1.Text.Trim()
    $global:syncHash.dns2 = $txtDNS2.Text.Trim()
    Invoke-BackgroundTask -Label "Configurar DNS" -Task {
        Invoke-SetDNS -Adapter $syncHash.dnsAdapter -DNS1 $syncHash.dns1 -DNS2 $syncHash.dns2
    }
})

$btnResetDNS.Add_Click({
    $global:syncHash.dnsAdapter = $txtAdapter.Text.Trim()
    Invoke-BackgroundTask -Label "Resetar DNS" -Task { Invoke-ResetDNS -Adapter $syncHash.dnsAdapter }
})

$btnPing.Add_Click({
    Invoke-BackgroundTask -Label "Teste de Conectividade" -Task { Invoke-TestarConectividade }
})

$btnIPConfig.Add_Click({
    Invoke-BackgroundTask -Label "IPConfig" -Task {
        $info = ipconfig /all 2>&1 | Out-String
        $info.Split("`n") | ForEach-Object { if ($_.Trim()) { QLog "[INFO] $_" } }
    }
})

# --- Dominio ---
$btnJoinDomain.Add_Click({
    $dom   = $txtDomain.Text.Trim()
    $user  = $txtDomainUser.Text.Trim()
    $pass  = $txtDomainPass.Password
    $name  = $txtNewName.Text.Trim()

    if (-not $dom) { Write-GuiLog "Preencha o campo Dominio." "Orange"; return }

    $result = [System.Windows.MessageBox]::Show(
        "Ingressar em '$dom' como '$user'?`n`nO computador precisara ser reiniciado.",
        "Confirmar Ingresso no Dominio",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($result -eq "Yes") {
        $global:syncHash.adDomain = $dom
        $global:syncHash.adUser   = $user
        $global:syncHash.adPass   = $pass
        $global:syncHash.adName   = $name
        Invoke-BackgroundTask -Label "Ingresso no Dominio" -Task {
            Invoke-JoinDomain -Domain $syncHash.adDomain -User $syncHash.adUser `
                              -Pass $syncHash.adPass -NewName $syncHash.adName
        }
    }
})

# --- Espaco em Disco ---
$btnScan.Add_Click({
    $global:syncHash.ScanPath = $txtScanPath.Text.Trim()
    if (-not (Test-Path $global:syncHash.ScanPath)) {
        Write-GuiLog "Caminho nao encontrado: $($global:syncHash.ScanPath)" "Orange"
        return
    }
    $lstFolders.ItemsSource = $null
    $lstFiles.ItemsSource   = $null

    Invoke-BackgroundTask -Label "Scan de Disco: $($global:syncHash.ScanPath)" -Task { Invoke-EspacoDisco } -OnDone {
        $global:syncHash.LstFolders.ItemsSource = $global:syncHash.ScanFolderList
        $global:syncHash.LstFiles.ItemsSource   = $global:syncHash.ScanFileList
    }
})

$btnScanClear.Add_Click({
    $lstFolders.ItemsSource = $null
    $lstFiles.ItemsSource   = $null
})

# ================================================================
# EXIBIR JANELA
# ================================================================
Write-GuiLog "NextTool v1.0 iniciado em $env:COMPUTERNAME." "Cyan"
[void]$window.ShowDialog()
