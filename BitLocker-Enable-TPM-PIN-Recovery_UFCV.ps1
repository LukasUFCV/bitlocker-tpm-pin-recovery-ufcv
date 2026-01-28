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
    Title="Activation BitLocker - Saisir le PIN"
    Height="670" Width="680"
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

        <Style TargetType="Button">
            <Setter Property="Background">
                <Setter.Value>
                    <SolidColorBrush Color="#40FFFFFF"/>
                </Setter.Value>
            </Setter>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Padding" Value="24,12"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderBrush" Value="#60FFFFFF"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="10"
                                Padding="{TemplateBinding Padding}"
                                Name="border">
                            <Border.Effect>
                                <DropShadowEffect Color="#30000000" BlurRadius="12" ShadowDepth="2" Opacity="0.5"/>
                            </Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background">
                                    <Setter.Value>
                                        <SolidColorBrush Color="#60FFFFFF"/>
                                    </Setter.Value>
                                </Setter>
                                <Setter TargetName="border" Property="BorderBrush" Value="#80FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background">
                <Setter.Value>
                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                        <GradientStop Color="#5E8FD9" Offset="0"/>
                        <GradientStop Color="#4A7AC2" Offset="1"/>
                    </LinearGradientBrush>
                </Setter.Value>
            </Setter>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#70FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1"
                                CornerRadius="10"
                                Padding="{TemplateBinding Padding}"
                                Name="border">
                            <Border.Effect>
                                <DropShadowEffect Color="#50000000" BlurRadius="16" ShadowDepth="3" Opacity="0.6"/>
                            </Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background">
                                    <Setter.Value>
                                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                                            <GradientStop Color="#6EA0E8" Offset="0"/>
                                            <GradientStop Color="#5A8AD1" Offset="1"/>
                                        </LinearGradientBrush>
                                    </Setter.Value>
                                </Setter>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect Color="#60000000" BlurRadius="20" ShadowDepth="4" Opacity="0.7"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="CloseButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#B0B0B0"/>
            <Setter Property="Width" Value="36"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="FontSize" Value="18"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="18">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background">
                                    <Setter.Value>
                                        <SolidColorBrush Color="#20FFFFFF"/>
                                    </Setter.Value>
                                </Setter>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <LinearGradientBrush x:Key="WindowBackground" StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#252540" Offset="0"/>
            <GradientStop Color="#2A2A45" Offset="1"/>
        </LinearGradientBrush>
    </Window.Resources>

    <Border Background="{StaticResource WindowBackground}" CornerRadius="16" Margin="10">
        <Border CornerRadius="14" Margin="3" BorderThickness="1">
            <Border.BorderBrush>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#30FFFFFF" Offset="0"/>
                    <GradientStop Color="#10FFFFFF" Offset="1"/>
                </LinearGradientBrush>
            </Border.BorderBrush>
            <Border.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1" Opacity="0.95">
                    <GradientStop Color="#2A2A45" Offset="0"/>
                    <GradientStop Color="#2F2F4A" Offset="1"/>
                </LinearGradientBrush>
            </Border.Background>

            <!-- ========================================= -->
            <!-- CONTENU PRINCIPAL : PIN + PROGRESS         -->
            <!-- ========================================= -->
            <Grid Margin="45,40,45,45">

                <!-- ========================= -->
                <!-- VUE 1 : Saisie du PIN     -->
                <!-- ========================= -->
                <Grid Name="PinView">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Button Name="CloseButton"
                            Grid.Row="0"
                            HorizontalAlignment="Right"
                            VerticalAlignment="Top"
                            Margin="0,-15,-15,0"
                            Style="{StaticResource CloseButton}"
                            Content="×"/>

                    <Viewbox Grid.Row="0" Width="64" Height="64" Margin="0,5,0,28">
                        <Canvas Width="24" Height="24">
                            <Path Fill="#5E8FD9" Data="M12,1L3,5V11C3,16.55 6.84,21.74 12,23C17.16,21.74 21,16.55 21,11V5L12,1M12,7C13.4,7 14.8,8.1 14.8,9.5V11C15.4,11 16,11.6 16,12.3V15.8C16,16.4 15.4,17 14.7,17H9.2C8.6,17 8,16.4 8,15.7V12.2C8,11.6 8.6,11 9.2,11V9.5C9.2,8.1 10.6,7 12,7M12,8.2C11.2,8.2 10.5,8.7 10.5,9.5V11H13.5V9.5C13.5,8.7 12.8,8.2 12,8.2Z">
                                <Path.Effect>
                                    <DropShadowEffect Color="#305E8FD9" BlurRadius="15" ShadowDepth="0" Opacity="0.6"/>
                                </Path.Effect>
                            </Path>
                        </Canvas>
                    </Viewbox>

                    <TextBlock Grid.Row="1"
                               HorizontalAlignment="Center"
                               Text="Protection des données"
                               FontSize="28"
                               FontWeight="SemiBold"
                               Margin="0,0,0,22"
                               Foreground="#FFFFFF">
                        <TextBlock.Effect>
                            <DropShadowEffect Color="#30000000" BlurRadius="8" ShadowDepth="2" Opacity="0.4"/>
                        </TextBlock.Effect>
                    </TextBlock>

                    <TextBlock Grid.Row="2"
                               HorizontalAlignment="Left"
                               Text="La sécurité de vos données est essentielle pour l'UFCV. BitLocker chiffre le contenu de votre disque afin de rendre les informations illisibles en cas de vol ou d'accès non autorisé. Vos fichiers restent ainsi protégés, même si l'ordinateur quitte les locaux de l'UFCV."
                               FontSize="12"
                               Margin="0,0,0,16"
                               Foreground="#D0D0D0"
                               TextWrapping="Wrap"
                               LineHeight="19"/>

                    <TextBlock Grid.Row="3"
                               HorizontalAlignment="Left"
                               Text="Il vous suffit maintenant de choisir un code PIN de 6 à 20 chiffres, de préférence le même que celui utilisé pour ouvrir votre session, afin de sécuriser l'accès à votre poste au démarrage."
                               FontSize="12"
                               Margin="0,0,0,32"
                               Foreground="#D0D0D0"
                               TextWrapping="Wrap"
                               LineHeight="19"/>

                    <Grid Grid.Row="4" Margin="0,0,0,24">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="16"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Grid.Column="0" Grid.Row="0"
                                   HorizontalAlignment="Left"
                                   Text="Code PIN"
                                   FontSize="13"
                                   Margin="0,0,0,10"
                                   Foreground="#E0E0E0"
                                   FontWeight="SemiBold"/>

                        <Border Name="PinInputBorder"
                                Grid.Column="0" Grid.Row="1"
                                CornerRadius="12"
                                BorderThickness="2"
                                Height="54">
                            <Border.BorderBrush>
                                <SolidColorBrush Color="#40FFFFFF"/>
                            </Border.BorderBrush>
                            <Border.Background>
                                <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                                    <GradientStop Color="#20FFFFFF" Offset="0"/>
                                    <GradientStop Color="#15FFFFFF" Offset="1"/>
                                </LinearGradientBrush>
                            </Border.Background>
                            <Border.Effect>
                                <DropShadowEffect Color="#20000000" BlurRadius="10" ShadowDepth="2" Opacity="0.4"/>
                            </Border.Effect>

                            <PasswordBox Name="PinInput"
                                         FontSize="15"
                                         VerticalContentAlignment="Center"
                                         Padding="18,0"
                                         Background="Transparent"
                                         Foreground="#FFFFFF"
                                         BorderThickness="0"
                                         FontWeight="Normal"/>
                        </Border>

                        <TextBlock Grid.Column="2" Grid.Row="0"
                                   HorizontalAlignment="Left"
                                   Text="Confirmation"
                                   FontSize="13"
                                   Margin="0,0,0,10"
                                   Foreground="#E0E0E0"
                                   FontWeight="SemiBold"/>

                        <Border Name="PinConfirmBorder"
                                Grid.Column="2" Grid.Row="1"
                                CornerRadius="12"
                                BorderThickness="2"
                                Height="54">
                            <Border.BorderBrush>
                                <SolidColorBrush Color="#40FFFFFF"/>
                            </Border.BorderBrush>
                            <Border.Background>
                                <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                                    <GradientStop Color="#20FFFFFF" Offset="0"/>
                                    <GradientStop Color="#15FFFFFF" Offset="1"/>
                                </LinearGradientBrush>
                            </Border.Background>
                            <Border.Effect>
                                <DropShadowEffect Color="#20000000" BlurRadius="10" ShadowDepth="2" Opacity="0.4"/>
                            </Border.Effect>

                            <PasswordBox Name="PinConfirm"
                                         FontSize="15"
                                         VerticalContentAlignment="Center"
                                         Padding="18,0"
                                         Background="Transparent"
                                         Foreground="#FFFFFF"
                                         BorderThickness="0"
                                         FontWeight="Normal"/>
                        </Border>
                    </Grid>

                    <TextBlock Grid.Row="5"
                               Name="PostponeCounter"
                               HorizontalAlignment="Center"
                               Text="Reports restants : 99/99"
                               FontSize="11"
                               Margin="0,0,0,24"
                               Foreground="#A0A0A0"
                               FontWeight="SemiBold"/>

                    <StackPanel Grid.Row="6"
                                Orientation="Horizontal"
                                HorizontalAlignment="Center">
                        <Button Name="ValidateButton"
                                Content="Valider"
                                Width="140"
                                Style="{StaticResource PrimaryButton}"/>
                        <Button Name="PostponeButton"
                                Content="Plus tard"
                                Width="140"/>
                    </StackPanel>
                </Grid>

                <!-- ========================= -->
                <!-- VUE 2 : Progression       -->
                <!-- ========================= -->
                <Grid Name="ProgressView" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0"
                               Text="Configuration BitLocker"
                               FontSize="22"
                               FontWeight="SemiBold"
                               Foreground="#FFFFFF"
                               HorizontalAlignment="Center"
                               Margin="0,0,0,12"/>

                    <TextBlock Grid.Row="1"
                               Name="ProgressStatus"
                               Text="Préparation..."
                               FontSize="12"
                               Foreground="#D0D0D0"
                               TextWrapping="Wrap"
                               HorizontalAlignment="Center"
                               Margin="0,0,0,14"/>

                    <Grid Grid.Row="2" Margin="0,0,0,18">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <ProgressBar Name="ProgressBar"
                                     Grid.Column="0"
                                     Height="16"
                                     Minimum="0"
                                     Maximum="100"
                                     Value="0"
                                     Margin="0,0,12,0"/>

                        <TextBlock Name="ProgressPercent"
                                   Grid.Column="1"
                                   Text="0%"
                                   Foreground="#A0A0A0"
                                   FontSize="12"
                                   VerticalAlignment="Center"/>
                    </Grid>

                    <Border Grid.Row="3" CornerRadius="12" BorderThickness="1" BorderBrush="#30FFFFFF">
                        <Border.Background>
                            <SolidColorBrush Color="#1AFFFFFF"/>
                        </Border.Background>

                        <ListBox Name="ProgressSteps"
                                 Background="Transparent"
                                 BorderThickness="0"
                                 Foreground="#E0E0E0"
                                 FontSize="12"
                                 Margin="10"
                                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                    </Border>

                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,18,0,0">
                        <Button Name="FinishButton"
                                Content="Fermer"
                                Width="140"
                                Visibility="Collapsed"
                                Style="{StaticResource PrimaryButton}"/>
                    </StackPanel>
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
$ValidateButton   = $Window.FindName("ValidateButton")
$PostponeButton   = $Window.FindName("PostponeButton")
$PostponeCounter  = $Window.FindName("PostponeCounter")
$CloseButton      = $Window.FindName("CloseButton")

# Progress UI controls
$PinView         = $Window.FindName("PinView")
$ProgressView    = $Window.FindName("ProgressView")
$ProgressBar     = $Window.FindName("ProgressBar")
$ProgressPercent = $Window.FindName("ProgressPercent")
$ProgressSteps   = $Window.FindName("ProgressSteps")
$ProgressStatus  = $Window.FindName("ProgressStatus")
$FinishButton    = $Window.FindName("FinishButton")

if (-not $PinInput -or -not $PinConfirm -or -not $PinInputBorder -or -not $PinConfirmBorder -or -not $ValidateButton -or -not $PostponeButton -or -not $PostponeCounter -or -not $CloseButton `
    -or -not $PinView -or -not $ProgressView -or -not $ProgressBar -or -not $ProgressPercent -or -not $ProgressSteps -or -not $ProgressStatus -or -not $FinishButton) {
    Write-Error "Échec de récupération des contrôles XAML. Le XAML peut être corrompu."
    exit 1
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

# Couleur compteur selon urgence
if ($RemainingPostpones -le 1) {
    $PostponeCounter.Foreground = "#EF5350"
} elseif ($RemainingPostpones -le 2) {
    $PostponeCounter.Foreground = "#FFB74D"
} else {
    $PostponeCounter.Foreground = "#66BB6A"
}

# Désactiver "Plus tard" si limite atteinte
if ($CurrentPostponeCount -ge $MaxPostpones) {
    $PostponeButton.IsEnabled = $false
    $PostponeButton.Content = "Limite atteinte"
    $PostponeButton.Opacity = 0.5

    $CloseButton.IsEnabled = $false
    $CloseButton.Opacity = 0.3

    $PostponeCounter.Foreground = "#EF5350"
    $PostponeCounter.Text = "Limite de reports atteinte (0/$MaxPostpones)"
}

# Désactiver Valider au démarrage
$ValidateButton.IsEnabled = $false
$ValidateButton.Opacity = 0.5

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

function Update-PinBorderColors {
    $pin        = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    if ([string]::IsNullOrEmpty($pin) -and [string]::IsNullOrEmpty($pinConfirm)) {
        $PinInputBorder.BorderBrush   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x30, 0xFF, 0xFF, 0xFF))
        $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x30, 0xFF, 0xFF, 0xFF))
        return
    }

    $validationResult  = Test-Pin -Pin $pin
    $isValidPin        = $validationResult[0]
    $validationResult2 = Test-Pin -Pin $pinConfirm
    $isValidPinConfirm = $validationResult2[0]

    if (-not [string]::IsNullOrEmpty($pin)) {
        if (-not $isValidPin) {
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xFF, 0xB7, 0x4D))
        } elseif (-not [string]::IsNullOrEmpty($pinConfirm) -and $pin -eq $pinConfirm) {
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0x66, 0xBB, 0x6A))
        } elseif (-not [string]::IsNullOrEmpty($pinConfirm) -and $pin -ne $pinConfirm) {
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xEF, 0x53, 0x50))
        } else {
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x40, 0xFF, 0xFF, 0xFF))
        }
    }

    if (-not [string]::IsNullOrEmpty($pinConfirm)) {
        if (-not $isValidPinConfirm) {
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xFF, 0xB7, 0x4D))
        } elseif (-not [string]::IsNullOrEmpty($pin) -and $pin -eq $pinConfirm -and $isValidPin) {
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0x66, 0xBB, 0x6A))
        } elseif (-not [string]::IsNullOrEmpty($pin) -and $pin -ne $pinConfirm) {
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xEF, 0x53, 0x50))
        } else {
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x40, 0xFF, 0xFF, 0xFF))
        }
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
        $ProgressSteps.Items.Clear()
        $ProgressStatus.Text = "Démarrage de la configuration..."
        $FinishButton.Visibility = "Collapsed"
    }
}

function Add-StepLine([string]$text) {
    Invoke-Ui {
        [void]$ProgressSteps.Items.Add($text)
        $ProgressSteps.ScrollIntoView($ProgressSteps.Items[$ProgressSteps.Items.Count - 1])
    }
}

function Set-Progress([int]$percent, [string]$status) {
    if ($percent -lt 0) { $percent = 0 }
    if ($percent -gt 100) { $percent = 100 }
    Invoke-Ui {
        $ProgressBar.Value = $percent
        $ProgressPercent.Text = "$percent%"
        if ($status) { $ProgressStatus.Text = $status }
    }
}

function Complete-Ui([string]$finalStatus, [bool]$isError = $false) {
    Invoke-Ui {
        $ProgressStatus.Text = $finalStatus
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

    if ($isError) {
        Add-StepLine "❌ Échec : $finalStatus"
    } else {
        Add-StepLine "✅ Terminé : $finalStatus"
    }
}

# ==========================================================
# Provisioning BitLocker en asynchrone (runspace) pour UI fluide
# ==========================================================
function Start-BitLockerProvisioningAsync {
    param([Parameter(Mandatory)] [string]$PlainPin)

    $script:IsProvisioning = $true
    Show-ProgressUi
    Add-StepLine "⏳ Initialisation..."

    # ----------------------------
    # Script exécuté dans runspace
    # ----------------------------
    $provisionScript = {
        param([string]$Pin)

        function Emit([int]$percent, [string]$text, [string]$tag = "info") {
            [pscustomobject]@{ kind="progress"; percent=$percent; text=$text; tag=$tag }
        }
        function Result([string]$status, [string]$message) {
            [pscustomobject]@{ kind="result"; status=$status; message=$message }
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
        Backup-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recId -ErrorAction Stop | Out-Null
        Emit 55 "Sauvegarde AD effectuée." "ok" | Write-Output

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
        # Consommer les nouveaux éléments
        while ($script:__BL_LastIndex -lt $script:__BL_Output.Count) {
            $item = $script:__BL_Output[$script:__BL_LastIndex]
            $script:__BL_LastIndex++

            if ($null -eq $item) { continue }

            if ($item.kind -eq "progress") {
                $p = [int]$item.percent
                $t = [string]$item.text
                Set-Progress -percent $p -status $t

                switch ($item.tag) {
                    "ok"   { Add-StepLine "✅ $t" }
                    "warn" { Add-StepLine "⚠️ $t" }
                    "done" { Add-StepLine "✅ $t" }
                    default { Add-StepLine "⏳ $t" }
                }
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
                    Set-Progress 100 $res.message
                    Complete-Ui -finalStatus $res.message -isError $false
                }
                elseif ($res -and $res.status -eq "policy_pending") {
                    Set-Progress 100 $res.message
                    Complete-Ui -finalStatus $res.message -isError $false
                }
                elseif ($res -and $res.status -eq "success") {
                    Set-Progress 100 $res.message
                    Complete-Ui -finalStatus $res.message -isError $false
                }
                else {
                    Complete-Ui -finalStatus "Terminé." -isError $false
                }
            }

            # Cleanup
            try { $script:__BL_PS.Dispose() } catch {}
            try { $script:__BL_RS.Close(); $script:__BL_RS.Dispose() } catch {}
            $script:__BL_PS = $null
            $script:__BL_RS = $null
            $script:__BL_Async = $null
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
        Add-StepLine "⛔ Merci de patienter : configuration en cours..."
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