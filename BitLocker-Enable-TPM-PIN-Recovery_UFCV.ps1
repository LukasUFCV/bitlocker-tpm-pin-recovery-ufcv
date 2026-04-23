# ==========================================================
#  Script d'activation BitLocker avec interface graphique
#  Version : UFCV GUI 1.0 (+ Progress UI)
#  Nom de fichier : BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1
#  Fonction : Activation du chiffrement BitLocker avec TPM + PIN + Recovery,
#              compatible GPO Network Unlock.
#  Auteurs : Lukas Mauffré & Olivier Marchoud
#  Structure : UFCV – DSI Pantin
#  Date : 03/11/2025
# ==========================================================

# ==========================================================
# Encodage & Culture - Contexte France
# Force l'encodage UTF-8 avec BOM pour la console et les sorties
# Configure la culture française pour les formats de date et nombre
# ==========================================================

# Forcer l'encodage UTF-8 avec BOM pour la console (accents, emoji)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($true)

# Forcer l'encodage par défaut de PowerShell (utile pour les fichiers)
$OutputEncoding = [System.Text.UTF8Encoding]::new($true)

# Définir la culture française (France)
Set-Culture fr-FR
Set-WinSystemLocale fr-FR
Set-WinUILanguageOverride fr-FR
[Threading.Thread]::CurrentThread.CurrentCulture = 'fr-FR'
[Threading.Thread]::CurrentThread.CurrentUICulture = 'fr-FR'

Write-Host "[INFO] Encodage UTF-8 (BOM) et culture fr-FR appliqués." -ForegroundColor Cyan

# Vérification des prérequis (avertissement si non-SYSTEM / LocalSystem)
$SystemSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
if ([System.Security.Principal.WindowsIdentity]::GetCurrent().User -ne $SystemSid) {
    Write-Warning "Script conçu pour le contexte SYSTEM (LocalSystem). Exécutez-le en tant que SYSTEM si nécessaire."
}

# Charger les assemblies WPF correctement
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ==========================================================
# Vérification (lecture seule) de la configuration BitLocker (clé FVE)
# Compare le registre avec les valeurs attendues (aucune correction)
# Affichage forcé (Out-Host) même si tu stockes le résultat dans une variable
# ==========================================================

$FveSubKey   = "SOFTWARE\Policies\Microsoft\FVE"

# Valeurs attendues (selon tes nouvelles GPO)
$RequiredKeys = @{
    "NetworkUnlockProvider"                  = "C:\Windows\System32\nkpprov.dll"
    "OSManageNKP"                            = 1
    "TPMAutoReseal"                          = 1
    "EncryptionMethodWithXtsOs"              = 7
    "EncryptionMethodWithXtsFdv"             = 7
    "EncryptionMethodWithXtsRdv"             = 4
    "OSEnablePrebootInputProtectorsOnSlates" = 1
    "OSEncryptionType"                       = 2
    "OSRecovery"                             = 1
    "OSManageDRA"                            = 1
    "OSRecoveryPassword"                     = 2
    "OSRecoveryKey"                          = 2
    "OSHideRecoveryPage"                     = 1
    "OSActiveDirectoryBackup"                = 1
    "OSActiveDirectoryInfoToStore"           = 1
    "OSRequireActiveDirectoryBackup"         = 1
    "ActiveDirectoryBackup"                  = 1
    "RequireActiveDirectoryBackup"           = 1
    "ActiveDirectoryInfoToStore"             = 1
    "UseRecoveryPassword"                    = 1
    "UseRecoveryDrive"                       = 1
    "UseAdvancedStartup"                     = 1
    "EnableBDEWithNoTPM"                     = 0
    "UseTPM"                                 = 0
    "UseTPMPIN"                              = 1
    "UseTPMKey"                              = 0
    "UseTPMKeyPIN"                           = 0
}

function Get-ExpectedRegistryKind($value) {
    if ($value -is [string]) { return [Microsoft.Win32.RegistryValueKind]::String }
    if ($value -is [int] -or $value -is [int32]) { return [Microsoft.Win32.RegistryValueKind]::DWord }
    return [Microsoft.Win32.RegistryValueKind]::Unknown
}

function Convert-ValueForCompare($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [string]) { return $v.Trim() }
    return $v
}

function Test-ValueEquality($current, $expected) {
    if ($expected -is [string]) {
        return ([string]$current).Trim().ToLowerInvariant() -eq $expected.Trim().ToLowerInvariant()
    }
    return $current -eq $expected
}

function Write-Section($title) {
    Write-Host ""
    Write-Host "══════════════════════════════════════════════════════════════" -ForegroundColor DarkCyan
    Write-Host "  $title" -ForegroundColor Cyan
    Write-Host "══════════════════════════════════════════════════════════════" -ForegroundColor DarkCyan
    Write-Host ""
}

Write-Section "BitLocker FVE Policy - Comparaison Registre"

# Ouvre la clé registre (lecture)
$rk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($FveSubKey, $false)

if ($null -eq $rk) {
    Write-Warning "Clé FVE absente : HKLM:\$FveSubKey (tout sera MISSING)"
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($name in ($RequiredKeys.Keys | Sort-Object)) {
    $expected     = $RequiredKeys[$name]
    $expectedKind = Get-ExpectedRegistryKind $expected

    $current     = $null
    $currentKind = $null
    $exists      = $false

    if ($null -ne $rk) {
        try {
            $current = $rk.GetValue($name, $null)
            if ($null -ne $current) {
                $exists = $true
                try { $currentKind = $rk.GetValueKind($name) } catch { $currentKind = $null }
            }
        } catch {
            $exists = $false
        }
    }

    $currentNorm  = Convert-ValueForCompare $current
    $expectedNorm = Convert-ValueForCompare $expected

    # Type OK ?
    $typeOk = $true
    if ($exists -and $null -ne $currentKind -and $expectedKind -ne [Microsoft.Win32.RegistryValueKind]::Unknown) {
        if ($expectedKind -eq [Microsoft.Win32.RegistryValueKind]::String) {
            # accepter ExpandString pour les chemins
            $typeOk = @([Microsoft.Win32.RegistryValueKind]::String, [Microsoft.Win32.RegistryValueKind]::ExpandString) -contains $currentKind
        } else {
            $typeOk = ($currentKind -eq $expectedKind)
        }
    }

    # Valeur OK ?
    $valueOk = $exists -and (Test-ValueEquality $currentNorm $expectedNorm)

    $status =
        if (-not $exists) { "MISSING" }
        elseif (-not $typeOk) { "TYPE_MISMATCH" }
        elseif (-not $valueOk) { "DIFF" }
        else { "OK" }

    $results.Add([pscustomobject]@{
        Name         = $name
        Status       = $status
        Expected     = $expected
        Current      = $current
        ExpectedType = $expectedKind.ToString()
        CurrentType  = if ($null -eq $currentKind) { $null } else { $currentKind.ToString() }
    }) | Out-Null
}

if ($null -ne $rk) { $rk.Close() }

# -------------------------
# Affichage (comme ton script Check-FVEPolicy) + contrôle d'éligibilité
# -------------------------

$okItems   = @($results | Where-Object { $_.Status -eq "OK" })
$diffItems = @($results | Where-Object { $_.Status -in @("DIFF","TYPE_MISMATCH") })
$missItems = @($results | Where-Object { $_.Status -eq "MISSING" })
$nonOk     = @($results | Where-Object { $_.Status -ne "OK" })

$okCount   = $okItems.Count
$diffCount = $diffItems.Count
$missCount = $missItems.Count

Write-Host "Résumé :" -ForegroundColor Cyan
Write-Host "  OK        : $okCount" -ForegroundColor Green
Write-Host "  DIFF/TYPE : $diffCount" -ForegroundColor Yellow
Write-Host "  MISSING   : $missCount" -ForegroundColor Red
Write-Host ""

Write-Host "Détails (hors OK) :" -ForegroundColor Cyan
if ($nonOk.Count -gt 0) {
    $nonOk | Format-Table -AutoSize Name, Status, ExpectedType, CurrentType, Expected, Current | Out-Host
} else {
    Write-Host "(Aucun écart)" -ForegroundColor DarkGray
}

$results | Format-Table -AutoSize Name, Status, Expected, Current, ExpectedType, CurrentType | Out-Host

Write-Host ""
Write-Host "Terminé." -ForegroundColor DarkCyan

# Bloquer si la GPO BitLocker attendue n'est pas appliquée
if ($diffCount -gt 0 -or $missCount -gt 0) {

    $msgUi = "Ce poste n'est pas éligible au déploiement BitLocker pour le moment.`n`n" +
             "La configuration attendue (GPO BitLocker) n'est pas appliquée.`n`n" +
             "OK : $okCount / DIFF/TYPE : $diffCount / MISSING : $missCount`n`n" +
             "Veuillez contacter la DSI (UFCV)."

    Write-Warning "Poste non éligible : configuration attendue (GPO BitLocker) non appliquée."

    [System.Windows.MessageBox]::Show($msgUi, "BitLocker - Poste non éligible", "OK", "Error") | Out-Null
    exit 1
}

# ==========================================================
# Vérification réseau UFCV (domaine + contrôleur de domaine)
# Attendu : domaine = ufcvfr.lan
# Exemple DC : SrvDC1.ufcvfr.lan (ou autre DC du domaine)
# ==========================================================

$ExpectedDomain = "ufcvfr.lan".Trim().ToLowerInvariant()

try {
    # 1) Domaine AD "officiel" (le plus fiable)
    $domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $currentDomain = ($domainObj.Name).ToLowerInvariant()

    # 2) Trouver un DC joignable (LAN ou VPN)
    $dcObj = $domainObj.FindDomainController()
    $dcName = ($dcObj.Name).ToLowerInvariant()

    # Vérifs strictes
    if ($currentDomain -ne $ExpectedDomain) {
        throw "Domaine détecté : $currentDomain (attendu : $ExpectedDomain)."
    }

    if (-not $dcName.EndsWith("." + $ExpectedDomain)) {
        throw "Contrôleur de domaine détecté : $dcName (hors domaine $ExpectedDomain)."
    }

    Write-Host "[OK] Réseau UFCV validé : domaine=$currentDomain ; DC=$dcName" -ForegroundColor Green
}
catch {
    $detail = $_.Exception.Message

    Write-Warning "Réseau UFCV non détecté : $detail"

    [System.Windows.MessageBox]::Show(
        "Ce poste n'est pas connecté au réseau UFCV (LAN/VPN) ou n'est pas sur le bon domaine.`n`n" +
        "Domaine attendu : $ExpectedDomain`n" +
        "Détail : $detail`n`n" +
        "Veuillez vous connecter au réseau interne ou au VPN UFCV puis relancer.",
        "BitLocker - Réseau requis",
        "OK",
        "Error"
    ) | Out-Null

    exit 1
}

# ==========================================================
# Gestion du compteur de reports (max 99 fois)
# ==========================================================
$CounterPath = "$env:ProgramData\BitLockerActivation\PostponeCount.txt"
$MaxPostpones = 99

$CounterDir = Split-Path $CounterPath -Parent
if (-not (Test-Path $CounterDir)) {
    New-Item -ItemType Directory -Path $CounterDir -Force | Out-Null
}

if (Test-Path $CounterPath) {
    $CurrentPostponeCount = [int](Get-Content $CounterPath -ErrorAction SilentlyContinue)
} else {
    $CurrentPostponeCount = 0
}

Write-Output "Reports restants : $($MaxPostpones - $CurrentPostponeCount)"

if ($CurrentPostponeCount -ge $MaxPostpones) {
    Write-Warning "Limite de reports atteinte. Activation BitLocker obligatoire."
}

# ==========================================================
# Vérification préalable de l'état BitLocker avant affichage GUI
# ==========================================================
$blv = Get-BitLockerVolume -MountPoint "C:"

switch ($blv.VolumeStatus) {
    'EncryptionInProgress' {
        [System.Windows.MessageBox]::Show(
            "Un chiffrement BitLocker est déjà en cours sur ce poste. Patientez jusqu'à la fin avant de relancer.",
            "Information", "OK", "Information"
        ) | Out-Null
        exit
    }
    'DecryptionInProgress' {
        [System.Windows.MessageBox]::Show(
            "Un déchiffrement BitLocker est actuellement en cours. Attendez qu'il soit terminé avant de relancer.",
            "Information", "OK", "Information"
        ) | Out-Null
        exit
    }
    'FullyEncrypted' {
        if ($blv.ProtectionStatus -eq 'On') {
            [System.Windows.MessageBox]::Show(
                "BitLocker est déjà activé sur ce poste. Aucune action n'est nécessaire.",
                "Information", "OK", "Information"
            ) | Out-Null
            exit
        }
    }
}

# ==========================================================
# XAML (UI PIN + UI Progression)
# ==========================================================
$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="UFCV - Protection BitLocker"
    Height="620" Width="920"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    ShowInTaskbar="True"
    Topmost="True">
    <Window.Resources>
        <Storyboard x:Key="WindowFadeIn">
            <DoubleAnimation Storyboard.TargetProperty="Opacity" From="0" To="1" Duration="0:0:0.4">
                <DoubleAnimation.EasingFunction>
                    <CubicEase EasingMode="EaseOut"/>
                </DoubleAnimation.EasingFunction>
            </DoubleAnimation>
        </Storyboard>

        <SolidColorBrush x:Key="WindowBrush" Color="#F3F6F9"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="SurfaceAltBrush" Color="#F8FBFD"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D9E4ED"/>
        <SolidColorBrush x:Key="InputBorderBrush" Color="#C7D8E6"/>
        <SolidColorBrush x:Key="UfcvBlueBrush" Color="#1696D2"/>
        <SolidColorBrush x:Key="UfcvBlueDarkBrush" Color="#0F7CB5"/>
        <SolidColorBrush x:Key="UfcvBlueSoftBrush" Color="#E8F5FB"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#24364A"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#5F7285"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#7F92A4"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#2F855A"/>
        <SolidColorBrush x:Key="SuccessSoftBrush" Color="#EAF7F0"/>
        <SolidColorBrush x:Key="WarningBrush" Color="#D48627"/>
        <SolidColorBrush x:Key="WarningSoftBrush" Color="#FFF4E5"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#C44F4B"/>
        <SolidColorBrush x:Key="ErrorSoftBrush" Color="#FCEEEE"/>
        <SolidColorBrush x:Key="InfoSoftBrush" Color="#EDF7FD"/>

        <Style x:Key="BaseButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="42"/>
            <Setter Property="MinWidth" Value="138"/>
            <Setter Property="Padding" Value="20,0"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonRoot"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="10"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonRoot" Property="Opacity" Value="0.55"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="ButtonRoot" Property="BorderBrush" Value="{StaticResource UfcvBlueDarkBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{StaticResource UfcvBlueBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource UfcvBlueBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource UfcvBlueDarkBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource UfcvBlueDarkBrush}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#0B6898"/>
                    <Setter Property="BorderBrush" Value="#0B6898"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource UfcvBlueSoftBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource UfcvBlueBrush}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#D8EEF9"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="CloseButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
            <Setter Property="Width" Value="32"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="CloseRoot"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CloseRoot" Property="Background" Value="{StaticResource UfcvBlueSoftBrush}"/>
                                <Setter TargetName="CloseRoot" Property="BorderBrush" Value="{StaticResource UfcvBlueBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource UfcvBlueDarkBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CloseRoot" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PasswordFieldStyle" TargetType="PasswordBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="14,0"/>
        </Style>

        <Style x:Key="ProgressBarStyle" TargetType="ProgressBar">
            <Setter Property="Height" Value="12"/>
            <Setter Property="Foreground" Value="{StaticResource UfcvBlueBrush}"/>
            <Setter Property="Background" Value="#D8E6F0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Grid MinHeight="12" SnapsToDevicePixels="True">
                            <Border x:Name="PART_Track"
                                    CornerRadius="6"
                                    Background="{TemplateBinding Background}"/>
                            <Border x:Name="PART_Indicator"
                                    HorizontalAlignment="Left"
                                    CornerRadius="6"
                                    Background="{TemplateBinding Foreground}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <ContentPresenter/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Background="Transparent" Margin="12">
        <Border.Effect>
            <DropShadowEffect Color="#22000000" BlurRadius="20" ShadowDepth="3" Opacity="0.45"/>
        </Border.Effect>

        <Border Background="{StaticResource WindowBrush}" CornerRadius="24" BorderThickness="1" BorderBrush="#E3EBF2">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="22,14,18,10" Panel.ZIndex="40">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Grid Width="120" Height="40">
                                <Image Name="LogoImage"
                                       Stretch="Uniform"
                                       HorizontalAlignment="Left"
                                       VerticalAlignment="Center"/>
                                <TextBlock Name="LogoFallback"
                                           Visibility="Collapsed"
                                           Text="UFCV"
                                           FontFamily="Bahnschrift SemiCondensed"
                                           FontSize="26"
                                           FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}"
                                           VerticalAlignment="Center"/>
                            </Grid>

                            <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
                                <TextBlock Text="Protection des postes UFCV"
                                           FontSize="17"
                                           FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>
                                <TextBlock Text="Déploiement BitLocker TPM + PIN"
                                           Margin="0,1,0,0"
                                           FontSize="11"
                                           Foreground="{StaticResource TextSecondaryBrush}"/>
                            </StackPanel>
                        </StackPanel>

                        <StackPanel Grid.Column="1"
                                    Orientation="Horizontal"
                                    VerticalAlignment="Center">
                            <Border Background="{StaticResource UfcvBlueSoftBrush}"
                                    BorderBrush="#B7DDF1"
                                    BorderThickness="1"
                                    CornerRadius="11"
                                    Padding="10,6"
                                    Margin="0,0,10,0"
                                    VerticalAlignment="Center">
                                <TextBlock Text="DSI - sécurisation du poste"
                                           FontSize="10.5"
                                           FontWeight="SemiBold"
                                           Foreground="{StaticResource UfcvBlueDarkBrush}"/>
                            </Border>

                            <Button Name="CloseButton"
                                    Content="×"
                                    Style="{StaticResource CloseButton}"
                                    Panel.ZIndex="50"
                                    VerticalAlignment="Center"/>
                        </StackPanel>
                    </Grid>

                    <Border Grid.Row="1" Margin="20,0,20,0" CornerRadius="14" Padding="18,12,18,12">
                        <Border.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                                <GradientStop Color="#0F88C7" Offset="0"/>
                                <GradientStop Color="#189AD8" Offset="0.55"/>
                                <GradientStop Color="#46AEE3" Offset="1"/>
                            </LinearGradientBrush>
                        </Border.Background>

                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="2.8*"/>
                                <ColumnDefinition Width="1.2*"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Text="Activation BitLocker encadrée"
                                           FontFamily="Bahnschrift SemiCondensed"
                                           FontSize="22"
                                           FontWeight="Bold"
                                           Foreground="White"/>
                                <TextBlock Text="Le poste UFCV sera protégé en deux temps : choix du code PIN, puis configuration et redémarrage."
                                           Margin="0,5,0,0"
                                           FontSize="11.5"
                                           Foreground="#EFF8FD"
                                           TextWrapping="Wrap"
                                           LineHeight="16"/>
                            </StackPanel>

                            <Border Grid.Column="1"
                                    Background="#1FFFFFFF"
                                    BorderBrush="#2EFFFFFF"
                                    BorderThickness="1"
                                    CornerRadius="12"
                                    Padding="12,9"
                                    Margin="16,0,0,0"
                                    VerticalAlignment="Center">
                                <StackPanel>
                                    <TextBlock Text="Parcours utilisateur"
                                               FontFamily="Bahnschrift SemiCondensed"
                                               FontSize="14"
                                               FontWeight="Bold"
                                               Foreground="White"/>
                                    <TextBlock Text="1. PIN  2. Configuration  3. Redémarrage"
                                               Margin="0,5,0,0"
                                               FontSize="10.5"
                                               Foreground="#F7FBFE"
                                               TextWrapping="Wrap"/>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </Border>
                </Grid>

                <Grid Grid.Row="1"
                      Margin="20,12,20,18"
                      VerticalAlignment="Stretch">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="2.22*"/>
                            <ColumnDefinition Width="0.92*"/>
                        </Grid.ColumnDefinitions>

                        <Border Grid.Column="0"
                                VerticalAlignment="Stretch"
                                Background="{StaticResource SurfaceBrush}"
                                BorderBrush="{StaticResource BorderBrush}"
                                BorderThickness="1"
                                CornerRadius="14"
                                Padding="18,15,18,15">
                            <Grid>
                                <Grid Name="PinView">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <Border Grid.Row="0"
                                        HorizontalAlignment="Left"
                                        Background="{StaticResource UfcvBlueSoftBrush}"
                                        BorderBrush="#C4E5F5"
                                        BorderThickness="1"
                                        CornerRadius="10"
                                        Padding="9,4">
                                    <TextBlock Text="Étape 1 sur 2"
                                               FontSize="10"
                                               FontWeight="SemiBold"
                                               Foreground="{StaticResource UfcvBlueDarkBrush}"/>
                                </Border>

                                <StackPanel Grid.Row="1" Margin="0,10,0,0">
                                    <TextBlock Text="Choisissez votre code PIN de démarrage"
                                               FontFamily="Bahnschrift SemiCondensed"
                                               FontSize="22"
                                               FontWeight="Bold"
                                               Foreground="{StaticResource TextPrimaryBrush}"/>
                                    <TextBlock Text="Ce code vous sera demandé au démarrage pour confirmer votre identité avant l'ouverture de Windows."
                                               Margin="0,5,0,0"
                                               FontSize="11"
                                               Foreground="{StaticResource TextSecondaryBrush}"
                                               TextWrapping="Wrap"
                                               LineHeight="15"/>
                                </StackPanel>

                                <TextBlock Grid.Row="2"
                                           Margin="0,10,0,0"
                                           Text="Choisissez un code personnel de 6 à 20 chiffres, différent des suites simples."
                                           FontSize="11"
                                           Foreground="{StaticResource TextSecondaryBrush}"
                                           TextWrapping="Wrap"
                                           LineHeight="15"/>

                                <Grid Grid.Row="3" Margin="0,12,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="14"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <TextBlock Grid.Column="0"
                                               Grid.Row="0"
                                               Text="Code PIN"
                                               Margin="0,0,0,6"
                                               FontSize="11.5"
                                               FontWeight="SemiBold"
                                               Foreground="{StaticResource TextPrimaryBrush}"/>

                                    <Border Name="PinInputBorder"
                                            Grid.Column="0"
                                            Grid.Row="1"
                                            Height="42"
                                            CornerRadius="10"
                                            BorderThickness="2"
                                            BorderBrush="{StaticResource InputBorderBrush}"
                                            Background="{StaticResource SurfaceAltBrush}">
                                        <PasswordBox Name="PinInput"
                                                     Style="{StaticResource PasswordFieldStyle}"/>
                                    </Border>

                                    <TextBlock Grid.Column="2"
                                               Grid.Row="0"
                                               Text="Confirmation du PIN"
                                               Margin="0,0,0,6"
                                               FontSize="11.5"
                                               FontWeight="SemiBold"
                                               Foreground="{StaticResource TextPrimaryBrush}"/>

                                    <Border Name="PinConfirmBorder"
                                            Grid.Column="2"
                                            Grid.Row="1"
                                            Height="42"
                                            CornerRadius="10"
                                            BorderThickness="2"
                                            BorderBrush="{StaticResource InputBorderBrush}"
                                            Background="{StaticResource SurfaceAltBrush}">
                                        <PasswordBox Name="PinConfirm"
                                                     Style="{StaticResource PasswordFieldStyle}"/>
                                    </Border>
                                </Grid>

                                <Border Grid.Row="4"
                                        Margin="0,10,0,0"
                                        CornerRadius="10"
                                        BorderThickness="1"
                                        BorderBrush="{StaticResource BorderBrush}"
                                        Background="{StaticResource SurfaceAltBrush}"
                                        Padding="12,8">
                                    <TextBlock Name="PinStatusText"
                                               Text="Utilisez un code PIN personnel de 6 à 20 chiffres. Les suites simples comme 123456 ou 654321 ne sont pas autorisées."
                                               FontSize="10.5"
                                               Foreground="{StaticResource TextSecondaryBrush}"
                                               TextWrapping="Wrap"
                                               LineHeight="14"/>
                                </Border>

                                <Border Grid.Row="6"
                                        Margin="0,10,0,0"
                                        CornerRadius="12"
                                        Background="{StaticResource UfcvBlueSoftBrush}"
                                        BorderBrush="#C6E5F4"
                                        BorderThickness="1"
                                        Padding="14,10">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>

                                        <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                            <TextBlock Text="Conseil UFCV"
                                                       FontFamily="Bahnschrift SemiCondensed"
                                                       FontSize="15"
                                                       FontWeight="Bold"
                                                       Foreground="{StaticResource UfcvBlueDarkBrush}"/>
                                            <TextBlock Text="Votre code PIN reste nécessaire au démarrage du poste."
                                                       Margin="0,3,0,0"
                                                       FontSize="10.5"
                                                       Foreground="{StaticResource TextPrimaryBrush}"
                                                       TextWrapping="Wrap"
                                                       LineHeight="14"/>
                                        </StackPanel>

                                        <StackPanel Grid.Column="1"
                                                    Orientation="Horizontal"
                                                    HorizontalAlignment="Right"
                                                    VerticalAlignment="Center"
                                                    Margin="14,0,0,0">
                                            <Button Name="PostponeButton"
                                                    Content="Reporter"
                                                    Width="126"
                                                    Style="{StaticResource SecondaryButton}"/>
                                            <Button Name="ValidateButton"
                                                    Content="Lancer l'activation"
                                                    Width="152"
                                                    Margin="8,0,0,0"
                                                    Style="{StaticResource PrimaryButton}"/>
                                        </StackPanel>
                                    </Grid>
                                </Border>
                                </Grid>

                                <Grid Name="ProgressView" Visibility="Collapsed">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <Border Grid.Row="0"
                                        HorizontalAlignment="Left"
                                        Background="{StaticResource UfcvBlueSoftBrush}"
                                        BorderBrush="#C4E5F5"
                                        BorderThickness="1"
                                        CornerRadius="11"
                                        Padding="9,4">
                                    <TextBlock Text="Étape 2 sur 2"
                                               FontSize="10"
                                               FontWeight="SemiBold"
                                               Foreground="{StaticResource UfcvBlueDarkBrush}"/>
                                </Border>

                                <StackPanel Grid.Row="1" Margin="0,10,0,0">
                                    <TextBlock Text="Configuration BitLocker en cours"
                                               FontFamily="Bahnschrift SemiCondensed"
                                               FontSize="22"
                                               FontWeight="Bold"
                                               Foreground="{StaticResource TextPrimaryBrush}"/>
                                    <TextBlock Text="La sécurisation du poste est en cours. Merci de laisser cette fenêtre ouverte jusqu'à la fin de l'opération."
                                               Margin="0,5,0,0"
                                               FontSize="11"
                                               Foreground="{StaticResource TextSecondaryBrush}"
                                               TextWrapping="Wrap"
                                               LineHeight="15"/>
                                </StackPanel>

                                <Border Grid.Row="2"
                                        Margin="0,10,0,0"
                                        Background="{StaticResource SurfaceAltBrush}"
                                        BorderBrush="{StaticResource BorderBrush}"
                                        BorderThickness="1"
                                        CornerRadius="14"
                                        Padding="14,11">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>

                                        <TextBlock Grid.Row="0"
                                                   Grid.Column="0"
                                                   Text="Suivi du provisioning"
                                                   FontSize="12"
                                                   FontWeight="SemiBold"
                                                   Foreground="{StaticResource TextPrimaryBrush}"/>

                                        <TextBlock Name="ProgressPercent"
                                                   Grid.Row="0"
                                                   Grid.Column="1"
                                                   Text="0%"
                                                   FontSize="14"
                                                   FontWeight="SemiBold"
                                                   Foreground="{StaticResource UfcvBlueDarkBrush}"
                                                   VerticalAlignment="Center"/>

                                        <StackPanel Grid.Row="1"
                                                    Grid.ColumnSpan="2"
                                                    Margin="0,10,0,0">
                                            <ProgressBar Name="ProgressBar"
                                                         Minimum="0"
                                                         Maximum="100"
                                                         Value="0"
                                                         Style="{StaticResource ProgressBarStyle}"/>
                                            <TextBlock Name="ProgressStatus"
                                                       Margin="0,8,0,0"
                                                       Text="Préparation..."
                                                       FontSize="11"
                                                       FontWeight="SemiBold"
                                                       Foreground="{StaticResource TextPrimaryBrush}"
                                                       TextWrapping="Wrap"
                                                       LineHeight="15"/>
                                        </StackPanel>
                                    </Grid>
                                </Border>

                                <Border Grid.Row="3"
                                        Margin="0,10,0,0"
                                        Background="{StaticResource SurfaceAltBrush}"
                                        BorderBrush="{StaticResource BorderBrush}"
                                        BorderThickness="1"
                                        CornerRadius="14"
                                        Padding="12">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                        </Grid.RowDefinitions>

                                        <TextBlock Grid.Row="0"
                                                   Text="Détail des opérations"
                                                   FontSize="12"
                                                   FontWeight="SemiBold"
                                                   Foreground="{StaticResource TextPrimaryBrush}"/>

                                        <ListBox Name="ProgressSteps"
                                                 Grid.Row="1"
                                                 Margin="0,8,0,0"
                                                 ScrollViewer.VerticalScrollBarVisibility="Hidden"
                                                 ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                                                 ScrollViewer.CanContentScroll="False"/>
                                    </Grid>
                                </Border>

                                <StackPanel Grid.Row="4"
                                            Orientation="Horizontal"
                                            HorizontalAlignment="Right"
                                            Margin="0,10,0,0">
                                    <Button Name="FinishButton"
                                            Content="Fermer"
                                            Width="142"
                                            Visibility="Collapsed"
                                            Style="{StaticResource PrimaryButton}"/>
                                </StackPanel>
                                </Grid>
                            </Grid>
                        </Border>

                        <Border Grid.Column="1"
                                VerticalAlignment="Stretch"
                                Margin="12,0,0,0"
                                Background="{StaticResource SurfaceBrush}"
                                BorderBrush="{StaticResource BorderBrush}"
                                BorderThickness="1"
                                CornerRadius="14"
                                Padding="15,13">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Grid.Row="0"
                                           Text="À quoi ça sert ?"
                                           FontFamily="Bahnschrift SemiCondensed"
                                           FontSize="17"
                                           FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>

                                <TextBlock Grid.Row="1"
                                           Margin="0,6,0,0"
                                           Text="Cette opération sert à protéger le poste UFCV. BitLocker chiffre les données de l'ordinateur pour mieux sécuriser les informations en cas de perte, de vol ou d'accès non autorisé. Le code PIN demandé renforce cette protection au démarrage."
                                           FontSize="10.5"
                                           Foreground="{StaticResource TextSecondaryBrush}"
                                           TextWrapping="Wrap"
                                           LineHeight="14"/>

                                <Border Grid.Row="2"
                                        Margin="0,10,0,0"
                                        Background="{StaticResource UfcvBlueSoftBrush}"
                                        BorderBrush="#C6E5F4"
                                        BorderThickness="1"
                                        CornerRadius="11"
                                        Padding="12,10">
                                    <StackPanel>
                                        <TextBlock Text="Pendant l'opération"
                                                   FontSize="11"
                                                   FontWeight="SemiBold"
                                                   Foreground="{StaticResource UfcvBlueDarkBrush}"/>
                                        <TextBlock Text="Laissez la fenêtre ouverte pendant la configuration."
                                                   Margin="0,6,0,0"
                                                   FontSize="10.5"
                                                   Foreground="{StaticResource TextPrimaryBrush}"
                                                   TextWrapping="Wrap"
                                                   LineHeight="14"/>
                                        <TextBlock Text="Un redémarrage peut être requis pour finaliser l'activation."
                                                   Margin="0,4,0,0"
                                                   FontSize="10.5"
                                                   Foreground="{StaticResource TextPrimaryBrush}"
                                                   TextWrapping="Wrap"
                                                   LineHeight="14"/>
                                    </StackPanel>
                                </Border>

                                <Border Grid.Row="4"
                                        Margin="0,10,0,0"
                                        Height="1"
                                        Background="{StaticResource BorderBrush}"/>

                                <TextBlock Grid.Row="5"
                                           Margin="0,10,0,0"
                                           Text="Disponibilité"
                                           FontFamily="Bahnschrift SemiCondensed"
                                           FontSize="17"
                                           FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>

                                <StackPanel Grid.Row="6" Margin="0,6,0,0">
                                    <TextBlock Name="PostponeCounter"
                                               Text="Reports restants : 99/99"
                                               FontSize="16"
                                               FontWeight="SemiBold"
                                               Foreground="{StaticResource UfcvBlueDarkBrush}"
                                               TextWrapping="Wrap"/>
                                    <TextBlock Text="Si nécessaire, l'opération peut être reportée dans la limite autorisée."
                                               Margin="0,6,0,0"
                                               FontSize="10.5"
                                               Foreground="{StaticResource TextSecondaryBrush}"
                                               TextWrapping="Wrap"
                                               LineHeight="14"/>
                                    <TextBlock Text="La fermeture est bloquée dès que la configuration BitLocker démarre."
                                               Margin="0,8,0,0"
                                               FontSize="10.5"
                                               Foreground="{StaticResource TextMutedBrush}"
                                               TextWrapping="Wrap"
                                               LineHeight="14"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                </Grid>
            </Grid>
        </Border>
    </Border>
</Window>
"@

# ==========================================================
# Parser XAML / créer fenêtre
# ==========================================================
try {
    $XamlBytes  = [System.Text.Encoding]::UTF8.GetBytes($Xaml)
    $XamlString = [System.Text.Encoding]::UTF8.GetString($XamlBytes)
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$XamlString)
    $Window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Output "XAML chargé avec succès."

    $Window.Opacity = 0
    $fadeInStoryboard = $Window.Resources["WindowFadeIn"]
    $fadeInStoryboard.Begin($Window)

} catch {
    Write-Error "Erreur lors du parsing du XAML : $($_.Exception.Message). Vérifiez l'encodage du fichier (UTF-8 BOM recommandé)."
    exit 1
}

# ==========================================================
# Récupérer contrôles
# ==========================================================
$PinInput         = $Window.FindName("PinInput")
$PinConfirm       = $Window.FindName("PinConfirm")
$PinInputBorder   = $Window.FindName("PinInputBorder")
$PinConfirmBorder = $Window.FindName("PinConfirmBorder")
$PinStatusText    = $Window.FindName("PinStatusText")
$ValidateButton   = $Window.FindName("ValidateButton")
$PostponeButton   = $Window.FindName("PostponeButton")
$PostponeCounter  = $Window.FindName("PostponeCounter")
$CloseButton      = $Window.FindName("CloseButton")
$LogoImage        = $Window.FindName("LogoImage")
$LogoFallback     = $Window.FindName("LogoFallback")

# Progress UI controls
$PinView         = $Window.FindName("PinView")
$ProgressView    = $Window.FindName("ProgressView")
$ProgressBar     = $Window.FindName("ProgressBar")
$ProgressPercent = $Window.FindName("ProgressPercent")
$ProgressSteps   = $Window.FindName("ProgressSteps")
$ProgressStatus  = $Window.FindName("ProgressStatus")
$FinishButton    = $Window.FindName("FinishButton")

if (-not $PinInput -or -not $PinConfirm -or -not $PinInputBorder -or -not $PinConfirmBorder -or -not $PinStatusText -or -not $ValidateButton -or -not $PostponeButton -or -not $PostponeCounter -or -not $CloseButton -or -not $LogoImage -or -not $LogoFallback `
    -or -not $PinView -or -not $ProgressView -or -not $ProgressBar -or -not $ProgressPercent -or -not $ProgressSteps -or -not $ProgressStatus -or -not $FinishButton) {
    Write-Error "Échec de récupération des contrôles XAML. Le XAML peut être corrompu."
    exit 1
}

# Ressources UI réutilisées par les handlers et les états visuels
$UiBrushes = @{
    InputBorder   = [System.Windows.Media.Brush]$Window.FindResource("InputBorderBrush")
    TextPrimary   = [System.Windows.Media.Brush]$Window.FindResource("TextPrimaryBrush")
    TextSecondary = [System.Windows.Media.Brush]$Window.FindResource("TextSecondaryBrush")
    TextMuted     = [System.Windows.Media.Brush]$Window.FindResource("TextMutedBrush")
    Info          = [System.Windows.Media.Brush]$Window.FindResource("UfcvBlueDarkBrush")
    InfoSoft      = [System.Windows.Media.Brush]$Window.FindResource("InfoSoftBrush")
    Success       = [System.Windows.Media.Brush]$Window.FindResource("SuccessBrush")
    SuccessSoft   = [System.Windows.Media.Brush]$Window.FindResource("SuccessSoftBrush")
    Warning       = [System.Windows.Media.Brush]$Window.FindResource("WarningBrush")
    WarningSoft   = [System.Windows.Media.Brush]$Window.FindResource("WarningSoftBrush")
    Error         = [System.Windows.Media.Brush]$Window.FindResource("ErrorBrush")
    ErrorSoft     = [System.Windows.Media.Brush]$Window.FindResource("ErrorSoftBrush")
}

# Charger le logo UFCV depuis le dépôt ; fallback textuel si absent.
$UiRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$LogoPath = Join-Path -Path $UiRoot -ChildPath "Assets\UFCV-logo.png"

if (Test-Path -LiteralPath $LogoPath) {
    try {
        $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bitmap.BeginInit()
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.UriSource = [System.Uri]::new((Resolve-Path -LiteralPath $LogoPath).Path, [System.UriKind]::Absolute)
        $bitmap.EndInit()
        $bitmap.Freeze()
        $LogoImage.Source = $bitmap
        $LogoImage.Visibility = "Visible"
        $LogoFallback.Visibility = "Collapsed"
    } catch {
        $LogoImage.Visibility = "Collapsed"
        $LogoFallback.Visibility = "Visible"
    }
} else {
    $LogoImage.Visibility = "Collapsed"
    $LogoFallback.Visibility = "Visible"
}

# ==========================================================
# Init UI / variables
# ==========================================================
$PinInput.MaxLength   = 20
$PinConfirm.MaxLength = 20

$RemainingPostpones = $MaxPostpones - $CurrentPostponeCount
$PostponeCounter.Text = "Reports restants : $RemainingPostpones/$MaxPostpones"

$script:UserAction = $null
$script:Pin = $null
$script:IsProvisioning = $false
$script:RestartPromptShown = $false

# Couleur compteur selon urgence
if ($RemainingPostpones -le 1) {
    $PostponeCounter.Foreground = $UiBrushes.Error
} elseif ($RemainingPostpones -le 2) {
    $PostponeCounter.Foreground = $UiBrushes.Warning
} else {
    $PostponeCounter.Foreground = $UiBrushes.Success
}

# Désactiver "Plus tard" si limite atteinte
if ($CurrentPostponeCount -ge $MaxPostpones) {
    $PostponeButton.IsEnabled = $false
    $PostponeButton.Content = "Limite atteinte"
    $PostponeButton.Opacity = 0.5

    $CloseButton.IsEnabled = $false
    $CloseButton.Opacity = 0.3

    $PostponeCounter.Foreground = $UiBrushes.Error
    $PostponeCounter.Text = "Limite de reports atteinte (0/$MaxPostpones)"
}

# Désactiver Valider au démarrage
$ValidateButton.IsEnabled = $false
$ValidateButton.Opacity = 0.5
$PinStatusText.Text = "Utilisez un code PIN personnel de 6 à 20 chiffres. Les suites simples comme 123456 ou 654321 ne sont pas autorisées."
$PinStatusText.Foreground = $UiBrushes.TextSecondary

# Bloquer caractères non numériques
$PinInput.AddHandler([System.Windows.Input.TextCompositionManager]::PreviewTextInputEvent,
    [System.Windows.Input.TextCompositionEventHandler] {
        param($src, $e)
        if ($e.Text -notmatch "^\d$") { $e.Handled = $true }
    })

$PinConfirm.AddHandler([System.Windows.Input.TextCompositionManager]::PreviewTextInputEvent,
    [System.Windows.Input.TextCompositionEventHandler] {
        param($src, $e)
        if ($e.Text -notmatch "^\d$") { $e.Handled = $true }
    })

# ==========================================================
# Validation PIN + UI
# ==========================================================
function Test-Pin {
    param($Pin)

    if ([string]::IsNullOrEmpty($Pin) -or $Pin.Length -lt 6 -or $Pin.Length -gt 20 -or $Pin -notmatch "^\d+$") {
        return $false, "PIN invalide : 6 à 20 chiffres requis."
    }

    $isAscending = $true
    for ($i = 0; $i -lt $Pin.Length - 1; $i++) {
        $current = [int]::Parse($Pin[$i].ToString())
        $next    = [int]::Parse($Pin[$i + 1].ToString())
        if ($next -ne ($current + 1)) { $isAscending = $false; break }
    }
    if ($isAscending) {
        return $false, "PIN invalide : les chiffres ne doivent pas être en ordre croissant (ex : 123456)."
    }

    $isDescending = $true
    for ($i = 0; $i -lt $Pin.Length - 1; $i++) {
        $current = [int]::Parse($Pin[$i].ToString())
        $next    = [int]::Parse($Pin[$i + 1].ToString())
        if ($next -ne ($current - 1)) { $isDescending = $false; break }
    }
    if ($isDescending) {
        return $false, "PIN invalide : les chiffres ne doivent pas être en ordre décroissant (ex : 654321)."
    }

    return $true, "OK"
}

function Set-PinStatus {
    param(
        [string]$Text,
        [System.Windows.Media.Brush]$Foreground
    )

    $PinStatusText.Text = $Text
    $PinStatusText.Foreground = $Foreground
}

function Update-WindowViewport {
    $workArea = [System.Windows.SystemParameters]::WorkArea
    $reservedWidth = 36
    $reservedHeight = 32

    $maxWidth = [Math]::Max(320, $workArea.Width - $reservedWidth)
    $maxHeight = [Math]::Max(320, $workArea.Height - $reservedHeight)

    $Window.MaxWidth = $maxWidth
    $Window.MaxHeight = $maxHeight

    if ($Window.Width -gt $maxWidth) {
        $Window.Width = $maxWidth
    }
    if ($Window.Height -gt $maxHeight) {
        $Window.Height = $maxHeight
    }

    $Window.Left = $workArea.Left + [Math]::Max(0, ($workArea.Width - $Window.Width) / 2)
    $Window.Top = $workArea.Top + [Math]::Max(0, ($workArea.Height - $Window.Height) / 2)
}

Update-WindowViewport
$Window.Add_Loaded({
    Update-WindowViewport
})

function Show-RestartPrompt {
    if ($script:RestartPromptShown) {
        return "later"
    }

    $script:RestartPromptShown = $true

    $restartXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="UFCV - Redémarrage requis"
    Width="520"
    SizeToContent="Height"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    ShowInTaskbar="False"
    Topmost="True">
    <Border Background="Transparent" Margin="10">
        <Border.Effect>
            <DropShadowEffect Color="#22000000" BlurRadius="18" ShadowDepth="3" Opacity="0.45"/>
        </Border.Effect>

        <Border Background="#F3F6F9" BorderBrush="#D9E4ED" BorderThickness="1" CornerRadius="18" Padding="18">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel>
                        <TextBlock Text="Redémarrage nécessaire"
                                   FontFamily="Bahnschrift SemiCondensed"
                                   FontSize="24"
                                   FontWeight="Bold"
                                   Foreground="#24364A"/>
                        <TextBlock Text="Finalisation de la protection BitLocker"
                                   Margin="0,3,0,0"
                                   FontSize="11.5"
                                   Foreground="#5F7285"/>
                    </StackPanel>

                    <Button Name="RestartCloseButton"
                            Grid.Column="1"
                            Width="32"
                            Height="32"
                            Margin="12,0,0,0"
                            Background="White"
                            BorderBrush="#D9E4ED"
                            BorderThickness="1"
                            Foreground="#5F7285"
                            FontSize="16"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="×"/>
                </Grid>

                <Border Grid.Row="1"
                        Background="#1696D2"
                        CornerRadius="14"
                        Padding="16,14">
                    <TextBlock Text="La configuration du poste est terminée. Un redémarrage est recommandé maintenant pour finaliser la mise en protection."
                               FontSize="12.5"
                               Foreground="White"
                               TextWrapping="Wrap"
                               LineHeight="18"/>
                </Border>

                <Border Grid.Row="2"
                        Margin="0,12,0,0"
                        Background="#E8F5FB"
                        BorderBrush="#C6E5F4"
                        BorderThickness="1"
                        CornerRadius="12"
                        Padding="14,10">
                    <StackPanel>
                        <TextBlock Text="Vous pourrez aussi redémarrer plus tard si vous devez terminer une tâche en cours."
                                   FontSize="11.5"
                                   Foreground="#24364A"
                                   TextWrapping="Wrap"
                                   LineHeight="17"/>
                        <TextBlock Name="RestartErrorText"
                                   Visibility="Collapsed"
                                   Margin="0,8,0,0"
                                   FontSize="11"
                                   FontWeight="SemiBold"
                                   Foreground="#C44F4B"
                                   TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

                <StackPanel Grid.Row="3"
                            Orientation="Horizontal"
                            HorizontalAlignment="Right"
                            Margin="0,14,0,0">
                    <Button Name="RestartLaterButton"
                            Width="140"
                            Height="42"
                            Margin="0,0,8,0"
                            Background="White"
                            BorderBrush="#D9E4ED"
                            BorderThickness="1"
                            Foreground="#24364A"
                            FontSize="13"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="Plus tard"/>
                    <Button Name="RestartNowButton"
                            Width="178"
                            Height="42"
                            Background="#1696D2"
                            BorderBrush="#1696D2"
                            BorderThickness="1"
                            Foreground="White"
                            FontSize="13"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="Redémarrer maintenant"/>
                </StackPanel>
            </Grid>
        </Border>
    </Border>
</Window>
"@

    try {
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$restartXaml)
        $restartWindow = [Windows.Markup.XamlReader]::Load($reader)
    } catch {
        $script:RestartPromptShown = $false
        Write-Warning "Impossible de charger la fenêtre de redémarrage : $($_.Exception.Message)"
        return "later"
    }

    if ($Window -and $Window.IsVisible) {
        try { $restartWindow.Owner = $Window } catch {}
    } else {
        $restartWindow.WindowStartupLocation = "CenterScreen"
    }

    $restartNowButton   = $restartWindow.FindName("RestartNowButton")
    $restartLaterButton = $restartWindow.FindName("RestartLaterButton")
    $restartCloseButton = $restartWindow.FindName("RestartCloseButton")
    $restartErrorText   = $restartWindow.FindName("RestartErrorText")

    $state = [pscustomobject]@{
        Choice = "later"
    }

    $restartLaterAction = {
        $state.Choice = "later"
        $restartWindow.Close()
    }

    $restartLaterButton.Add_Click($restartLaterAction)
    $restartCloseButton.Add_Click($restartLaterAction)

    $restartNowButton.Add_Click({
        $restartNowButton.IsEnabled = $false
        $restartLaterButton.IsEnabled = $false
        $restartCloseButton.IsEnabled = $false

        try {
            Restart-Computer -Force -ErrorAction Stop
            $state.Choice = "restart"
            $restartWindow.Close()
        } catch {
            $restartErrorText.Text = "Le redémarrage automatique n'a pas pu être lancé. Vous pouvez réessayer ou choisir Plus tard."
            $restartErrorText.Visibility = "Visible"
            $restartNowButton.IsEnabled = $true
            $restartLaterButton.IsEnabled = $true
            $restartCloseButton.IsEnabled = $true
        }
    })

    try {
        [void]$restartWindow.ShowDialog()
    } catch {
        Write-Warning "Impossible d'afficher la fenêtre de redémarrage : $($_.Exception.Message)"
        $state.Choice = "later"
    }

    return $state.Choice
}

function Show-AdBackupFailurePrompt {
    param(
        [string]$TechnicalDetail
    )

    $dialogXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="UFCV - Sauvegarde AD requise"
    Width="560"
    SizeToContent="Height"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    ShowInTaskbar="False"
    Topmost="True">
    <Border Background="Transparent" Margin="10">
        <Border.Effect>
            <DropShadowEffect Color="#22000000" BlurRadius="18" ShadowDepth="3" Opacity="0.45"/>
        </Border.Effect>

        <Border Background="#F3F6F9" BorderBrush="#D9E4ED" BorderThickness="1" CornerRadius="18" Padding="18">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel>
                        <TextBlock Text="Sauvegarde de la clé non confirmée"
                                   FontFamily="Bahnschrift SemiCondensed"
                                   FontSize="24"
                                   FontWeight="Bold"
                                   Foreground="#24364A"/>
                        <TextBlock Text="Finalisation de la protection BitLocker"
                                   Margin="0,3,0,0"
                                   FontSize="11.5"
                                   Foreground="#5F7285"/>
                    </StackPanel>

                    <Button Name="AdErrorCloseButton"
                            Grid.Column="1"
                            Width="32"
                            Height="32"
                            Margin="12,0,0,0"
                            Background="White"
                            BorderBrush="#D9E4ED"
                            BorderThickness="1"
                            Foreground="#5F7285"
                            FontSize="16"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="×"/>
                </Grid>

                <Border Grid.Row="1"
                        Background="#FCEEEE"
                        BorderBrush="#E8B7B4"
                        BorderThickness="1"
                        CornerRadius="14"
                        Padding="16,14">
                    <TextBlock Text="L'opération ne peut pas être finalisée tant que la sauvegarde de la clé BitLocker dans l'annuaire UFCV n'a pas pu être validée."
                               FontSize="12.5"
                               Foreground="#24364A"
                               TextWrapping="Wrap"
                               LineHeight="18"/>
                </Border>

                <Border Grid.Row="2"
                        Margin="0,12,0,0"
                        Background="#E8F5FB"
                        BorderBrush="#C6E5F4"
                        BorderThickness="1"
                        CornerRadius="12"
                        Padding="14,10">
                    <StackPanel>
                        <TextBlock Text="Causes probables"
                                   FontSize="11.5"
                                   FontWeight="SemiBold"
                                   Foreground="#0F7CB5"/>
                        <TextBlock Margin="0,6,0,0"
                                   Text="Le poste n'est peut-être plus correctement connecté au réseau UFCV, le contrôleur de domaine n'est pas joignable, ou la sauvegarde dans l'AD n'a pas pu être confirmée."
                                   FontSize="11"
                                   Foreground="#24364A"
                                   TextWrapping="Wrap"
                                   LineHeight="17"/>
                        <TextBlock Margin="0,6,0,0"
                                   Text="Reconnectez-vous au réseau UFCV ou au VPN, puis choisissez Réessayer."
                                   FontSize="11"
                                   Foreground="#24364A"
                                   TextWrapping="Wrap"
                                   LineHeight="17"/>
                        <TextBlock Name="AdErrorTechnicalText"
                                   Visibility="Collapsed"
                                   Margin="0,8,0,0"
                                   FontSize="10.5"
                                   Foreground="#5F7285"
                                   TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

                <TextBlock Grid.Row="3"
                           Name="AdErrorReporterHint"
                           Visibility="Collapsed"
                           Margin="0,10,0,0"
                           FontSize="10.5"
                           Foreground="#C44F4B"
                           Text="Le report n'est plus disponible sur ce poste."/>

                <StackPanel Grid.Row="4"
                            Orientation="Horizontal"
                            HorizontalAlignment="Right"
                            Margin="0,14,0,0">
                    <Button Name="AdErrorPostponeButton"
                            Width="132"
                            Height="42"
                            Margin="0,0,8,0"
                            Background="White"
                            BorderBrush="#D9E4ED"
                            BorderThickness="1"
                            Foreground="#24364A"
                            FontSize="13"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="Reporter"/>
                    <Button Name="AdErrorRetryButton"
                            Width="150"
                            Height="42"
                            Background="#1696D2"
                            BorderBrush="#1696D2"
                            BorderThickness="1"
                            Foreground="White"
                            FontSize="13"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            Content="Réessayer"/>
                </StackPanel>
            </Grid>
        </Border>
    </Border>
</Window>
"@

    try {
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($reader)
    } catch {
        Write-Warning "Impossible de charger la fenêtre d'erreur AD : $($_.Exception.Message)"
        return "dismiss"
    }

    if ($Window -and $Window.IsVisible) {
        try { $dialog.Owner = $Window } catch {}
    } else {
        $dialog.WindowStartupLocation = "CenterScreen"
    }

    $closeButton    = $dialog.FindName("AdErrorCloseButton")
    $retryButton    = $dialog.FindName("AdErrorRetryButton")
    $postponeButton = $dialog.FindName("AdErrorPostponeButton")
    $hintText       = $dialog.FindName("AdErrorReporterHint")
    $technicalText  = $dialog.FindName("AdErrorTechnicalText")

    if (-not [string]::IsNullOrWhiteSpace($TechnicalDetail)) {
        $technicalText.Text = "Détail technique : $TechnicalDetail"
        $technicalText.Visibility = "Visible"
    }

    if ($CurrentPostponeCount -ge $MaxPostpones) {
        $postponeButton.IsEnabled = $false
        $hintText.Visibility = "Visible"
    }

    $state = [pscustomobject]@{
        Choice = "dismiss"
    }

    $closeAction = {
        $state.Choice = "dismiss"
        $dialog.Close()
    }

    $closeButton.Add_Click($closeAction)

    $retryButton.Add_Click({
        $state.Choice = "retry"
        $dialog.Close()
    })

    $postponeButton.Add_Click({
        $state.Choice = "postpone"
        $dialog.Close()
    })

    try {
        [void]$dialog.ShowDialog()
    } catch {
        Write-Warning "Impossible d'afficher la fenêtre d'erreur AD : $($_.Exception.Message)"
        $state.Choice = "dismiss"
    }

    return $state.Choice
}

function Update-PinBorderColors {
    $pin        = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    if ([string]::IsNullOrEmpty($pin) -and [string]::IsNullOrEmpty($pinConfirm)) {
        $PinInputBorder.BorderBrush   = $UiBrushes.InputBorder
        $PinConfirmBorder.BorderBrush = $UiBrushes.InputBorder
        Set-PinStatus -Text "Utilisez un code PIN personnel de 6 à 20 chiffres. Les suites simples comme 123456 ou 654321 ne sont pas autorisées." -Foreground $UiBrushes.TextSecondary
        return
    }

    $validationResult  = Test-Pin -Pin $pin
    $isValidPin        = $validationResult[0]
    $validationResult2 = Test-Pin -Pin $pinConfirm
    $isValidPinConfirm = $validationResult2[0]

    $PinInputBorder.BorderBrush = $UiBrushes.InputBorder
    $PinConfirmBorder.BorderBrush = $UiBrushes.InputBorder

    if (-not [string]::IsNullOrEmpty($pin) -and -not $isValidPin) {
        $PinInputBorder.BorderBrush = $UiBrushes.Warning
        Set-PinStatus -Text $validationResult[1] -Foreground $UiBrushes.Warning
        return
    }

    if (-not [string]::IsNullOrEmpty($pinConfirm) -and -not $isValidPinConfirm) {
        $PinInputBorder.BorderBrush = if ($isValidPin) { $UiBrushes.Info } else { $UiBrushes.Warning }
        $PinConfirmBorder.BorderBrush = $UiBrushes.Warning
        Set-PinStatus -Text $validationResult2[1] -Foreground $UiBrushes.Warning
        return
    }

    if (-not [string]::IsNullOrEmpty($pin) -and [string]::IsNullOrEmpty($pinConfirm) -and $isValidPin) {
        $PinInputBorder.BorderBrush = $UiBrushes.Info
        Set-PinStatus -Text "Le code PIN respecte les critères. Confirmez-le pour poursuivre." -Foreground $UiBrushes.Info
        return
    }

    if (-not [string]::IsNullOrEmpty($pin) -and -not [string]::IsNullOrEmpty($pinConfirm) -and $pin -ne $pinConfirm) {
        $PinInputBorder.BorderBrush = $UiBrushes.Error
        $PinConfirmBorder.BorderBrush = $UiBrushes.Error
        Set-PinStatus -Text "Les deux saisies doivent être identiques pour lancer l'activation." -Foreground $UiBrushes.Error
        return
    }

    if (-not [string]::IsNullOrEmpty($pin) -and -not [string]::IsNullOrEmpty($pinConfirm) -and $pin -eq $pinConfirm -and $isValidPin -and $isValidPinConfirm) {
        $PinInputBorder.BorderBrush = $UiBrushes.Success
        $PinConfirmBorder.BorderBrush = $UiBrushes.Success
        Set-PinStatus -Text "Les deux saisies sont cohérentes. Vous pouvez lancer l'activation BitLocker." -Foreground $UiBrushes.Success
        return
    }
}

function Update-ValidateButtonState {
    $pin        = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    if ([string]::IsNullOrEmpty($pin) -or [string]::IsNullOrEmpty($pinConfirm)) {
        $ValidateButton.IsEnabled = $false
        $ValidateButton.Opacity = 0.5
        return
    }

    $validationResult1 = Test-Pin -Pin $pin
    $validationResult2 = Test-Pin -Pin $pinConfirm
    $isValidPin        = $validationResult1[0]
    $isValidPinConfirm = $validationResult2[0]
    $pinsMatch         = $pin -eq $pinConfirm

    if ($isValidPin -and $isValidPinConfirm -and $pinsMatch) {
        $ValidateButton.IsEnabled = $true
        $ValidateButton.Opacity = 1.0
    } else {
        $ValidateButton.IsEnabled = $false
        $ValidateButton.Opacity = 0.5
    }
}

$PinInput.Add_PasswordChanged({
    Update-PinBorderColors
    Update-ValidateButtonState
})
$PinConfirm.Add_PasswordChanged({
    Update-PinBorderColors
    Update-ValidateButtonState
})

# ==========================================================
# UI helpers (thread-safe) + Progress
# ==========================================================
function Invoke-Ui([scriptblock]$sb) {
    if ($Window.Dispatcher.CheckAccess()) { & $sb }
    else { $Window.Dispatcher.Invoke($sb) }
}

function Show-ProgressUi {
    Invoke-Ui {
        $PinView.Visibility      = "Collapsed"
        $ProgressView.Visibility = "Visible"

        # Bloquer actions pendant provisioning
        $CloseButton.IsEnabled    = $false
        $CloseButton.Opacity      = 0.3
        $PostponeButton.IsEnabled = $false
        $PostponeButton.Opacity   = 0.5
        $ValidateButton.IsEnabled = $false
        $ValidateButton.Opacity   = 0.5

        $ProgressBar.Value = 0
        $ProgressPercent.Text = "0%"
        $ProgressPercent.Foreground = $UiBrushes.Info
        $ProgressSteps.Items.Clear()
        $ProgressStatus.Text = "Démarrage de la configuration..."
        $ProgressStatus.Foreground = $UiBrushes.TextPrimary
        $FinishButton.Visibility = "Collapsed"
    }
}

function Find-VisualChildByType {
    param(
        [Parameter(Mandatory)] [System.Windows.DependencyObject]$Parent,
        [Parameter(Mandatory)] [Type]$TargetType
    )

    $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
    for ($i = 0; $i -lt $childCount; $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
        if ($null -eq $child) {
            continue
        }

        if ($TargetType.IsAssignableFrom($child.GetType())) {
            return $child
        }

        $match = Find-VisualChildByType -Parent $child -TargetType $TargetType
        if ($null -ne $match) {
            return $match
        }
    }

    return $null
}

function Move-ProgressStepsToLatest {
    if (-not $ProgressSteps -or $ProgressSteps.Items.Count -eq 0) {
        return
    }

    $latestItem = $ProgressSteps.Items[$ProgressSteps.Items.Count - 1]
    if ($null -eq $latestItem) {
        return
    }

    $scrollAction = [System.Action]{
        try {
            $ProgressSteps.UpdateLayout()
            $ProgressSteps.ScrollIntoView($latestItem)

            $container = $ProgressSteps.ItemContainerGenerator.ContainerFromItem($latestItem)
            if ($container) {
                $container.BringIntoView()
            }

            $scrollViewer = Find-VisualChildByType -Parent $ProgressSteps -TargetType ([System.Windows.Controls.ScrollViewer])
            if ($scrollViewer) {
                $scrollViewer.UpdateLayout()
                $scrollViewer.ScrollToVerticalOffset($scrollViewer.ScrollableHeight)
            }
        } catch {
        }
    }

    if ($Window.Dispatcher.CheckAccess()) {
        $null = $Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Render, $scrollAction)
    } else {
        Invoke-Ui {
            $null = $Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Render, $scrollAction)
        }
    }
}

function Get-ProgressStateBrush([string]$tag) {
    switch ($tag) {
        "ok"    { return $UiBrushes.Success }
        "done"  { return $UiBrushes.Success }
        "warn"  { return $UiBrushes.Warning }
        "error" { return $UiBrushes.Error }
        default { return $UiBrushes.Info }
    }
}

function Get-StepPalette([string]$tag) {
    switch ($tag) {
        "ok" {
            return @{
                Label      = "Succès"
                Background = $UiBrushes.SuccessSoft
                Border     = $UiBrushes.Success
                Foreground = $UiBrushes.Success
            }
        }
        "done" {
            return @{
                Label      = "Terminé"
                Background = $UiBrushes.SuccessSoft
                Border     = $UiBrushes.Success
                Foreground = $UiBrushes.Success
            }
        }
        "warn" {
            return @{
                Label      = "Attention"
                Background = $UiBrushes.WarningSoft
                Border     = $UiBrushes.Warning
                Foreground = $UiBrushes.Warning
            }
        }
        "error" {
            return @{
                Label      = "Erreur"
                Background = $UiBrushes.ErrorSoft
                Border     = $UiBrushes.Error
                Foreground = $UiBrushes.Error
            }
        }
        default {
            return @{
                Label      = "Information"
                Background = $UiBrushes.InfoSoft
                Border     = $UiBrushes.Info
                Foreground = $UiBrushes.Info
            }
        }
    }
}

function Add-StepLine([string]$text, [string]$tag = "info") {
    Invoke-Ui {
        $palette = Get-StepPalette -tag $tag

        $itemBorder = New-Object System.Windows.Controls.Border
        $itemBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $itemBorder.BorderThickness = [System.Windows.Thickness]::new(1)
        $itemBorder.Padding = [System.Windows.Thickness]::new(8, 5, 8, 5)
        $itemBorder.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
        $itemBorder.Background = $palette.Background
        $itemBorder.BorderBrush = $palette.Border

        $panel = New-Object System.Windows.Controls.DockPanel
        $panel.LastChildFill = $true

        $marker = New-Object System.Windows.Shapes.Ellipse
        $marker.Width = 8
        $marker.Height = 8
        $marker.Fill = $palette.Foreground
        $marker.Margin = [System.Windows.Thickness]::new(0, 4, 8, 0)
        $marker.VerticalAlignment = [System.Windows.VerticalAlignment]::Top

        $message = New-Object System.Windows.Controls.TextBlock
        $message.Text = $text
        $message.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $message.FontSize = 11
        $message.LineHeight = 15
        $message.Foreground = $UiBrushes.TextPrimary
        $message.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        [System.Windows.Controls.DockPanel]::SetDock($marker, [System.Windows.Controls.Dock]::Left)
        [void]$panel.Children.Add($marker)
        [void]$panel.Children.Add($message)
        $itemBorder.Child = $panel

        [void]$ProgressSteps.Items.Add($itemBorder)
        Move-ProgressStepsToLatest
    }
}

function Set-Progress([int]$percent, [string]$status, [string]$tag = "info") {
    if ($percent -lt 0) { $percent = 0 }
    if ($percent -gt 100) { $percent = 100 }
    Invoke-Ui {
        $ProgressBar.Value = $percent
        $ProgressPercent.Text = "$percent%"
        $ProgressPercent.Foreground = Get-ProgressStateBrush -tag $tag
        if ($status) {
            $ProgressStatus.Text = $status
            $ProgressStatus.Foreground = Get-ProgressStateBrush -tag $tag
        }
    }
}

function Complete-Ui([string]$finalStatus, [bool]$isError = $false, [string]$state = $null) {
    if ([string]::IsNullOrWhiteSpace($state)) {
        $state = if ($isError) { "error" } else { "done" }
    }

    Invoke-Ui {
        $ProgressStatus.Text = $finalStatus
        $ProgressStatus.Foreground = Get-ProgressStateBrush -tag $state
        $ProgressPercent.Foreground = Get-ProgressStateBrush -tag $state
        $script:IsProvisioning = $false

        # Marquer l'écran comme terminé (permet fermeture sans incrément de report)
        if ($script:UserAction -eq "Provisioning") {
            $script:UserAction = "Completed"
        }

        # Autoriser fermeture + bouton
        $FinishButton.Visibility = "Visible"
        $FinishButton.IsEnabled = $true

        $CloseButton.IsEnabled = $true
        $CloseButton.Opacity = 1.0
    }

    Add-StepLine -text $finalStatus -tag $state
}

# ==========================================================
# Provisioning BitLocker en asynchrone (runspace) pour UI fluide
# ==========================================================
function Start-BitLockerProvisioningAsync {
    param([Parameter(Mandatory)] [string]$PlainPin)

    $script:IsProvisioning = $true
    Show-ProgressUi
    Add-StepLine -text "Initialisation de la configuration BitLocker." -tag "info"

    # ----------------------------
    # Script exécuté dans runspace
    # ----------------------------
    $provisionScript = {
        param([string]$Pin)

        function Emit([int]$percent, [string]$text, [string]$tag = "info") {
            [pscustomobject]@{ kind="progress"; percent=$percent; text=$text; tag=$tag }
        }
        function Result([string]$status, [string]$message, [string]$detail = $null) {
            [pscustomobject]@{ kind="result"; status=$status; message=$message; detail=$detail }
        }

        try { Import-Module BitLocker -ErrorAction Stop } catch { }

        $MountPoint       = "C:"
        $EncryptionMethod = "XtsAes256"

        Emit 5 "Vérification de l'état BitLocker..." | Write-Output
        $blv = Get-BitLockerVolume -MountPoint $MountPoint

        if ($blv.VolumeStatus -eq 'EncryptionInProgress') { return Result "already" "Un chiffrement BitLocker est déjà en cours. Patientez puis relancez." }
        if ($blv.VolumeStatus -eq 'DecryptionInProgress') { return Result "already" "Un déchiffrement BitLocker est en cours. Patientez puis relancez." }
        if ($blv.VolumeStatus -eq 'FullyEncrypted' -and $blv.ProtectionStatus -eq 'On') { return Result "already" "BitLocker est déjà activé sur ce poste. Aucune action nécessaire." }

        function Get-Protector([string]$mp, [string]$type) {
            (Get-BitLockerVolume -MountPoint $mp).KeyProtector | Where-Object { $_.KeyProtectorType -eq $type }
        }
        function Get-FirstProtectorId([string]$mp, [string]$type) {
            Get-Protector -mp $mp -type $type | Select-Object -ExpandProperty KeyProtectorId -First 1
        }

        # 1) RecoveryPassword
        Emit 20 "Étape 1/3 : vérification / création du RecoveryPassword..." | Write-Output
        $recId = Get-FirstProtectorId -mp $MountPoint -type "RecoveryPassword"

        if (-not $recId) {
            Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop | Out-Null
            $recId = Get-FirstProtectorId -mp $MountPoint -type "RecoveryPassword"
            if (-not $recId) { throw "Impossible de récupérer l'ID du RecoveryPassword après création." }
            Emit 30 "RecoveryPassword ajouté." "ok" | Write-Output
        } else {
            Emit 30 "RecoveryPassword déjà présent (réutilisation)." "ok" | Write-Output
        }

        # 2) Backup AD
        Emit 45 "Étape 2/3 : sauvegarde du RecoveryPassword dans AD DS..." | Write-Output
        try {
            Backup-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recId -ErrorAction Stop | Out-Null
            Emit 55 "Sauvegarde AD effectuée." "ok" | Write-Output
        }
        catch {
            $detail = $_.Exception.Message
            Emit 55 "Impossible de confirmer la sauvegarde de la clé dans l'annuaire UFCV." "error" | Write-Output
            return Result "ad_backup_failed" "La sauvegarde de la clé BitLocker dans l'annuaire UFCV n'a pas pu être validée." $detail
        }

        # 3) Enable-BitLocker
        Emit 65 "Étape 3/3 : activation BitLocker (Used Space Only, TPM + PIN)..." | Write-Output
        $UserPin = ConvertTo-SecureString $Pin -AsPlainText -Force

        $existingTpmPins = @(Get-Protector -mp $MountPoint -type "TpmPin")
        if ($existingTpmPins.Count -gt 0) {
            Emit 70 "Un protecteur TPM+PIN existe déjà : suppression avant recréation..." "warn" | Write-Output
            foreach ($kp in $existingTpmPins) {
                Remove-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop
            }
            Emit 72 "Protecteur(s) TPM+PIN supprimé(s)." "ok" | Write-Output
        }

        try {
            Enable-BitLocker -MountPoint $MountPoint `
                -EncryptionMethod $EncryptionMethod `
                -UsedSpaceOnly `
                -TpmAndPinProtector `
                -Pin $UserPin `
                -ErrorAction Stop | Out-Null

            Emit 85 "Enable-BitLocker lancé." "ok" | Write-Output
        }
        catch {
            $msg = $_.Exception.Message
            $hr  = $_.Exception.HResult

            if ($hr -eq -2144272384 -or $msg -match "0x80310060") {
                New-Item -ItemType File -Path "$env:ProgramData\BitLockerActivation\PendingReboot.flag" -Force | Out-Null
                return Result "policy_pending" "La stratégie BitLocker n'autorise pas encore le PIN (0x80310060). Redémarrez puis relancez le script."
            }

            throw "Échec Enable-BitLocker : $msg"
        }

        # Succès -> suppression compteur reports
        $CounterPath = "$env:ProgramData\BitLockerActivation\PostponeCount.txt"
        if (Test-Path $CounterPath) { Remove-Item $CounterPath -Force }

        Emit 100 "Configuration terminée. Redémarrage requis pour finaliser et démarrer le chiffrement." "done" | Write-Output
        return Result "success" "BitLocker configuré. Un redémarrage est requis."
    }

    # ----------------------------
    # Runspace + async
    # ----------------------------
    $script:__BL_RS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:__BL_RS.ApartmentState = "MTA"
    $script:__BL_RS.ThreadOptions  = "ReuseThread"
    $script:__BL_RS.Open()

    $script:__BL_PS = [System.Management.Automation.PowerShell]::Create()
    $script:__BL_PS.Runspace = $script:__BL_RS

    # IMPORTANT : AddScript attend une string => on passe le contenu du scriptblock
    [void]$script:__BL_PS.AddScript($provisionScript.ToString()).AddArgument($PlainPin)

    $script:__BL_Output = New-Object System.Management.Automation.PSDataCollection[psobject]
    $script:__BL_Input  = New-Object System.Management.Automation.PSDataCollection[psobject]
    $script:__BL_LastIndex = 0

    try {
        $script:__BL_Async = $script:__BL_PS.BeginInvoke($script:__BL_Input, $script:__BL_Output)
    } catch {
        Complete-Ui -finalStatus ("Erreur lancement asynchrone : " + $_.Exception.Message) -isError $true
        return
    }

    # ----------------------------
    # Timer WPF : lecture output + UI
    # ----------------------------
    if ($script:__BL_Timer) {
        try { $script:__BL_Timer.Stop() } catch {}
        $script:__BL_Timer = $null
    }

    $script:__BL_Timer = New-Object System.Windows.Threading.DispatcherTimer
    $script:__BL_Timer.Interval = [TimeSpan]::FromMilliseconds(200)

    $script:__BL_Timer.Add_Tick({
        $showRestartPrompt = $false
        $adBackupPrompt = $null

        # Consommer les nouveaux éléments
        while ($script:__BL_LastIndex -lt $script:__BL_Output.Count) {
            $item = $script:__BL_Output[$script:__BL_LastIndex]
            $script:__BL_LastIndex++

            if ($null -eq $item) { continue }

            if ($item.kind -eq "progress") {
                $p = [int]$item.percent
                $t = [string]$item.text
                $tag = [string]$item.tag
                Set-Progress -percent $p -status $t -tag $tag
                Add-StepLine -text $t -tag $tag
                continue
            }
        }

        # Fin async ?
        if ($script:__BL_Async -and $script:__BL_Async.IsCompleted) {
            $script:__BL_Timer.Stop()

            try {
                # EndInvoke peut throw si erreur non gérée
                $null = $script:__BL_PS.EndInvoke($script:__BL_Async)
            } catch {
                Complete-Ui -finalStatus ("Erreur : " + $_.Exception.Message) -isError $true
            }

            # Erreurs PowerShell stream ?
            if ($script:__BL_PS.Streams.Error.Count -gt 0) {
                $first = $script:__BL_PS.Streams.Error[0]
                Complete-Ui -finalStatus ("Erreur : " + $first.Exception.Message) -isError $true
            } else {
                # Récupérer result
                $res = $null
                foreach ($o in $script:__BL_Output) {
                    if ($o -and $o.kind -eq "result") { $res = $o }
                }

                if ($res -and $res.status -eq "already") {
                    Set-Progress 100 $res.message "warn"
                    Complete-Ui -finalStatus $res.message -isError $false -state "info"
                }
                elseif ($res -and $res.status -eq "ad_backup_failed") {
                    Set-Progress 100 $res.message "error"
                    Complete-Ui -finalStatus $res.message -isError $true -state "error"
                    if ($res.detail) {
                        Add-StepLine -text ("Détail technique : " + $res.detail) -tag "error"
                    }
                    $adBackupPrompt = $res
                }
                elseif ($res -and $res.status -eq "policy_pending") {
                    Set-Progress 100 $res.message "warn"
                    Complete-Ui -finalStatus $res.message -isError $false -state "warn"
                }
                elseif ($res -and $res.status -eq "success") {
                    Set-Progress 100 $res.message "done"
                    Complete-Ui -finalStatus $res.message -isError $false -state "done"
                    $showRestartPrompt = $true
                }
                else {
                    Complete-Ui -finalStatus "Terminé." -isError $false -state "done"
                }
            }

            # Cleanup
            try { $script:__BL_PS.Dispose() } catch {}
            try { $script:__BL_RS.Close(); $script:__BL_RS.Dispose() } catch {}
            $script:__BL_PS = $null
            $script:__BL_RS = $null
            $script:__BL_Async = $null

            if ($adBackupPrompt) {
                $adChoice = Show-AdBackupFailurePrompt -TechnicalDetail $adBackupPrompt.detail

                if ($adChoice -eq "retry" -and -not [string]::IsNullOrWhiteSpace($script:Pin)) {
                    $script:UserAction = "Provisioning"
                    Start-BitLockerProvisioningAsync -PlainPin $script:Pin
                    return
                }

                if ($adChoice -eq "postpone") {
                    $script:UserAction = "Postponed"
                    if ($Window.IsVisible) {
                        $Window.DialogResult = $false
                        $Window.Close()
                    }
                    return
                }
            }

            if ($showRestartPrompt) {
                $restartChoice = Show-RestartPrompt
                if ($restartChoice -eq "restart") {
                    $script:UserAction = "Validated"
                    if ($Window.IsVisible) {
                        $Window.DialogResult = $true
                        $Window.Close()
                    }
                }
            }
        }
    })

    $script:__BL_Timer.Start()
}

# ==========================================================
# Events boutons
# ==========================================================
$ValidateButton.Add_Click({
    $pin = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    $validationResult = Test-Pin -Pin $pin
    $isValid = $validationResult[0]
    $errorMessage = $validationResult[1]

    if (-not $isValid) {
        [System.Windows.MessageBox]::Show($errorMessage, "Erreur", "OK", "Error") | Out-Null
        return
    }

    if ($pin -ne $pinConfirm) {
        [System.Windows.MessageBox]::Show("Les deux codes PIN ne correspondent pas. Veuillez réessayer.", "Erreur", "OK", "Error") | Out-Null
        $PinConfirm.Clear()
        return
    }

    $script:UserAction = "Provisioning"
    $script:Pin = $pin

    try {
        Start-BitLockerProvisioningAsync -PlainPin $pin
    } catch {
        Complete-Ui -finalStatus ("Erreur interne : " + $_.Exception.Message) -isError $true
    }
})

$PostponeButton.Add_Click({
    Write-Output "Activation reportée via le bouton. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
    $script:UserAction = "Postponed"
    $Window.DialogResult = $false
    $Window.Close()
})

# CloseButton : si provisioning terminé => fermer sans report, sinon comportement "Plus tard"
$CloseButton.Add_Click({
    if ($script:IsProvisioning) {
        return
    }

    if ($script:UserAction -eq "Completed") {
        $script:UserAction = "Validated"
        $Window.DialogResult = $true
        $Window.Close()
        return
    }

    Write-Output "Activation reportée via le bouton X. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
    $script:UserAction = "Postponed"
    $Window.DialogResult = $false
    $Window.Close()
})

$FinishButton.Add_Click({
    $script:UserAction = "Validated"
    $Window.DialogResult = $true
    $Window.Close()
})

# ==========================================================
# Bloquer fermeture pendant provisioning + gérer limite reports
# ==========================================================
$Window.Add_Closing({
    param($src, $e)

    # Provisioning en cours => on bloque
    if ($script:IsProvisioning) {
        $e.Cancel = $true
        Add-StepLine -text "Merci de patienter : la configuration BitLocker est toujours en cours." -tag "warn"
        return
    }

    # Limite de reports atteinte => on bloque toute fermeture tant que pas terminé
    if ($CurrentPostponeCount -ge $MaxPostpones -and $script:UserAction -notin @("Validated","Completed")) {
        $e.Cancel = $true
        [System.Windows.MessageBox]::Show(
            "Limite de reports atteinte. L'activation BitLocker est obligatoire.",
            "BitLocker", "OK", "Warning"
        ) | Out-Null
        return
    }

    # Si terminé => laisser fermer sans incrément
    if ($script:UserAction -in @("Validated","Completed")) {
        return
    }

    # Fermeture sans action => considéré comme "Plus tard"
    if ([string]::IsNullOrEmpty($script:UserAction)) {
        Write-Output "Activation reportée via fermeture de la fenêtre. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
        $script:UserAction = "Postponed"
        return
    }
})

# ==========================================================
# Afficher fenêtre
# ==========================================================
try {
    $dialogResult = $Window.ShowDialog()
    Write-Output "DialogResult : $dialogResult"
} catch {
    Write-Error "Erreur d'affichage WPF : $($_.Exception.Message). Essayez sans AllowsTransparency si le problème persiste."
    exit 1
}

# ==========================================================
# Incrémenter compteur si report
# ==========================================================
if ($script:UserAction -eq "Postponed") {
    $CurrentPostponeCount++
    $CurrentPostponeCount | Set-Content $CounterPath -Force
    Write-Output "Compteur incrémenté. Nouvelle valeur : $CurrentPostponeCount"
} else {
    Write-Output "Aucun report."
}

# ==========================================================
# Nettoyage sécurisé
# ==========================================================
$script:Pin = $null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
