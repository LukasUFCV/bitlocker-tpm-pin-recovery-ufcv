# ==========================================================
#  Script d'activation BitLocker avec interface graphique
#  Version : UFCV GUI 1.0
#  Nom de fichier : BitLocker-Enable-TPM-PIN-Recovery_UFCV.ps1
#  Fonction : Activation du chiffrement BitLocker avec TPM + PIN + Recovery,
#              compatible GPO Network Unlock.
#  Auteurs : Lukas Mauffré & Olivier Marchoud
#  Structure : UFCV – DSI Pantin
#  Date : 03/11/2025
# ==========================================================

# Forcer encodage UTF-8 pour console (corrige les accents)
# $OutputEncoding = [System.Text.Encoding]::UTF8

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

# Vérification des prérequis (avertissement si non-System)
if (-not ([System.Security.Principal.WindowsIdentity]::GetCurrent().User -eq 'SYSTEM')) {
    Write-Warning "Script conçu pour contexte Système ; adaptez si nécessaire."
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

$FveRegPath  = "HKLM:\SOFTWARE\Policies\Microsoft\FVE"
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

function Normalize-ValueForCompare($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [string]) { return $v.Trim() }
    return $v
}

function Values-AreEqual($current, $expected) {
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

    if ($rk -ne $null) {
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

    $currentNorm  = Normalize-ValueForCompare $current
    $expectedNorm = Normalize-ValueForCompare $expected

    # Type OK ?
    $typeOk = $true
    if ($exists -and $currentKind -ne $null -and $expectedKind -ne [Microsoft.Win32.RegistryValueKind]::Unknown) {
        if ($expectedKind -eq [Microsoft.Win32.RegistryValueKind]::String) {
            # accepter ExpandString pour les chemins
            $typeOk = @([Microsoft.Win32.RegistryValueKind]::String, [Microsoft.Win32.RegistryValueKind]::ExpandString) -contains $currentKind
        } else {
            $typeOk = ($currentKind -eq $expectedKind)
        }
    }

    # Valeur OK ?
    $valueOk = $exists -and (Values-AreEqual $currentNorm $expectedNorm)

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

# -------------------------
# Affichage (comme ton script Check-FVEPolicy)
# -------------------------
$okCount    = ($results | Where-Object { $_.Status -eq "OK" }).Count
$diffCount  = ($results | Where-Object { $_.Status -in @("DIFF","TYPE_MISMATCH") }).Count
$missCount  = ($results | Where-Object { $_.Status -eq "MISSING" }).Count

Write-Host "Résumé :" -ForegroundColor Cyan
Write-Host "  OK        : $okCount" -ForegroundColor Green
Write-Host "  DIFF/TYPE : $diffCount" -ForegroundColor Yellow
Write-Host "  MISSING   : $missCount" -ForegroundColor Red
Write-Host ""

Write-Host "Détails (hors OK) :" -ForegroundColor Cyan
$nonOk = $results | Where-Object { $_.Status -ne "OK" }

if ($nonOk.Count -gt 0) {
    $nonOk | Format-Table -AutoSize Name, Status, ExpectedType, CurrentType, Expected, Current | Out-Host
} else {
    Write-Host "(Aucun écart)" -ForegroundColor DarkGray
}

# Afficher TOUT (comme ton script)
$results | Format-Table -AutoSize Name, Status, Expected, Current, ExpectedType, CurrentType | Out-Host

Write-Host ""
Write-Host "Terminé." -ForegroundColor DarkCyan

# (Optionnel) si tu veux garder le résultat pour plus tard :
$FveAudit = $results

# Gestion du compteur de reports (max 99 fois)
$CounterPath = "$env:ProgramData\BitLockerActivation\PostponeCount.txt"
$MaxPostpones = 99

# Créer le dossier si nécessaire
$CounterDir = Split-Path $CounterPath -Parent
if (-not (Test-Path $CounterDir)) {
    New-Item -ItemType Directory -Path $CounterDir -Force | Out-Null
}

# Lire le compteur actuel
if (Test-Path $CounterPath) {
    $CurrentPostponeCount = [int](Get-Content $CounterPath -ErrorAction SilentlyContinue)
} else {
    $CurrentPostponeCount = 0
}

Write-Output "Reports restants : $($MaxPostpones - $CurrentPostponeCount)"

# Si limite atteinte, afficher un avertissement
if ($CurrentPostponeCount -ge $MaxPostpones) {
    Write-Warning "Limite de reports atteinte. Activation BitLocker obligatoire."
}

# Vérification préalable de l'état BitLocker avant affichage GUI
$blv = Get-BitLockerVolume -MountPoint "C:"

switch ($blv.VolumeStatus) {
    'EncryptionInProgress' {
        [System.Windows.MessageBox]::Show(
            "Un chiffrement BitLocker est déjà en cours sur ce poste. Patientez jusqu'à la fin avant de relancer.",
            "Information", "OK", "Information"
        )
        exit
    }
    'DecryptionInProgress' {
        [System.Windows.MessageBox]::Show(
            "Un déchiffrement BitLocker est actuellement en cours. Attendez qu'il soit terminé avant de relancer.",
            "Information", "OK", "Information"
        )
        exit
    }
    'FullyEncrypted' {
        if ($blv.ProtectionStatus -eq 'On') {
            [System.Windows.MessageBox]::Show(
                "BitLocker est déjà activé sur ce poste. Aucune action n'est nécessaire.",
                "Information", "OK", "Information"
            )
            exit
        }
    }
}

# XAML avec design Glassmorphism moderne et sobre
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
        <!-- Animation fade-in pour la fenêtre -->
        <Storyboard x:Key="WindowFadeIn">
            <DoubleAnimation Storyboard.TargetProperty="Opacity" From="0" To="1" Duration="0:0:0.4">
                <DoubleAnimation.EasingFunction>
                    <CubicEase EasingMode="EaseOut"/>
                </DoubleAnimation.EasingFunction>
            </DoubleAnimation>
        </Storyboard>
        
        <!-- Style moderne pour les boutons avec effet glassmorphism -->
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
        
        <!-- Style pour le bouton principal (Valider) avec dégradé vibrant -->
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
        
        <!-- Style pour le bouton de fermeture -->
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
        
        <!-- Dégradé de fond avec couleurs plus riches -->
        <LinearGradientBrush x:Key="WindowBackground" StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#252540" Offset="0"/>
            <GradientStop Color="#2A2A45" Offset="1"/>
        </LinearGradientBrush>
    </Window.Resources>
    
    <!-- Fond dégradé externe -->
    <Border Background="{StaticResource WindowBackground}" CornerRadius="16" Margin="10">
        
        <!-- Carte glassmorphism principale -->
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

            <Grid Margin="45,40,45,45">
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
                
                <!-- Bouton de fermeture X -->
                <Button Name="CloseButton" 
                        Grid.Row="0"
                        HorizontalAlignment="Right" 
                        VerticalAlignment="Top"
                        Margin="0,-15,-15,0"
                        Style="{StaticResource CloseButton}"
                        Content="×"/>
                
                <!-- Icône décorative avec couleur vibrante -->
                <Viewbox Grid.Row="0" Width="64" Height="64" Margin="0,5,0,28">
                    <Canvas Width="24" Height="24">
                        <Path Fill="#5E8FD9" Data="M12,1L3,5V11C3,16.55 6.84,21.74 12,23C17.16,21.74 21,16.55 21,11V5L12,1M12,7C13.4,7 14.8,8.1 14.8,9.5V11C15.4,11 16,11.6 16,12.3V15.8C16,16.4 15.4,17 14.7,17H9.2C8.6,17 8,16.4 8,15.7V12.2C8,11.6 8.6,11 9.2,11V9.5C9.2,8.1 10.6,7 12,7M12,8.2C11.2,8.2 10.5,8.7 10.5,9.5V11H13.5V9.5C13.5,8.7 12.8,8.2 12,8.2Z">
                            <Path.Effect>
                                <DropShadowEffect Color="#305E8FD9" BlurRadius="15" ShadowDepth="0" Opacity="0.6"/>
                            </Path.Effect>
                        </Path>
                    </Canvas>
                </Viewbox>
                
                <!-- Titre principal avec meilleur contraste -->
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
                
                <!-- Texte explicatif - Paragraphe 1 avec meilleur contraste -->
                <TextBlock Grid.Row="2" 
                           HorizontalAlignment="Left" 
                           Text="La sécurité de vos données est essentielle pour l'UFCV. BitLocker chiffre le contenu de votre disque afin de rendre les informations illisibles en cas de vol ou d'accès non autorisé. Vos fichiers restent ainsi protégés, même si l'ordinateur quitte les locaux de l'UFCV." 
                           FontSize="12" 
                           Margin="0,0,0,16" 
                           Foreground="#D0D0D0" 
                           TextWrapping="Wrap"
                           LineHeight="19"/>

                <!-- Texte explicatif - Paragraphe 2 avec meilleur contraste -->
                <TextBlock Grid.Row="3" 
                           HorizontalAlignment="Left" 
                           Text="Il vous suffit maintenant de choisir un code PIN à 6 chiffres, de préférence le même que celui utilisé pour ouvrir votre session, afin de sécuriser l'accès à votre poste au démarrage." 
                           FontSize="12" 
                           Margin="0,0,0,32" 
                           Foreground="#D0D0D0" 
                           TextWrapping="Wrap"
                           LineHeight="19"/>
                
                <!-- Grid pour les deux champs côte à côte -->
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
                    
                    <!-- Colonne 1 - PIN -->
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
                    
                    <!-- Colonne 2 - Confirmation -->
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
                
                <!-- Compteur de reports avec style amélioré -->
                <TextBlock Grid.Row="5" 
                           Name="PostponeCounter"
                           HorizontalAlignment="Center" 
                           Text="Reports restants : 99/99" 
                           FontSize="11" 
                           Margin="0,0,0,24" 
                           Foreground="#A0A0A0" 
                           FontWeight="SemiBold"/>
                
                <!-- Boutons -->
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
        </Border>
    </Border>
</Window>
"@

# Parser le XAML et créer la fenêtre
try {
    $XamlBytes = [System.Text.Encoding]::UTF8.GetBytes($Xaml)
    $XamlString = [System.Text.Encoding]::UTF8.GetString($XamlBytes)
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$XamlString)
    $Window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Output "XAML chargé avec succès."
    
    # Démarrer l'animation de fade-in
    $Window.Opacity = 0
    $fadeInStoryboard = $Window.Resources["WindowFadeIn"]
    $fadeInStoryboard.Begin($Window)
    
} catch {
    Write-Error "Erreur lors du parsing du XAML : $($_.Exception.Message). Vérifiez l'encodage du fichier (UTF-8 BOM recommandé)."
    exit 1
}

# Récupérer les contrôles (vérification null)
$PinInput = $Window.FindName("PinInput")
$PinConfirm = $Window.FindName("PinConfirm")
$PinInputBorder = $Window.FindName("PinInputBorder")
$PinConfirmBorder = $Window.FindName("PinConfirmBorder")
$ValidateButton = $Window.FindName("ValidateButton")
$PostponeButton = $Window.FindName("PostponeButton")
$PostponeCounter = $Window.FindName("PostponeCounter")
$CloseButton = $Window.FindName("CloseButton")

if (-not $PinInput -or -not $PinConfirm -or -not $PinInputBorder -or -not $PinConfirmBorder -or -not $ValidateButton -or -not $PostponeButton -or -not $PostponeCounter -or -not $CloseButton) {
    Write-Error "Échec de récupération des contrôles XAML. Le XAML peut être corrompu."
    exit 1
}

# Mettre à jour le compteur visuel
$RemainingPostpones = $MaxPostpones - $CurrentPostponeCount
$PostponeCounter.Text = "Reports restants : $RemainingPostpones/$MaxPostpones"

# Variable pour suivre l'action de l'utilisateur
$script:UserAction = $null

# Changer la couleur selon l'urgence avec des couleurs plus vibrantes
if ($RemainingPostpones -le 1) {
    $PostponeCounter.Foreground = "#EF5350"
} elseif ($RemainingPostpones -le 2) {
    $PostponeCounter.Foreground = "#FFB74D"
} else {
    $PostponeCounter.Foreground = "#66BB6A"
}

# Désactiver le bouton « Remettre à plus tard » si limite atteinte
if ($CurrentPostponeCount -ge $MaxPostpones) {
    $PostponeButton.IsEnabled = $false
    $PostponeButton.Content = "Limite atteinte"
    $PostponeButton.Opacity = 0.5
    $CloseButton.IsEnabled = $false
    $CloseButton.Opacity = 0.3
    $PostponeCounter.Foreground = "#EF5350"
    $PostponeCounter.Text = "Limite de reports atteinte (0/$MaxPostpones)"
}

# Désactiver le bouton Valider au démarrage
$ValidateButton.IsEnabled = $false
$ValidateButton.Opacity = 0.5

# Ajouter un gestionnaire pour bloquer les caractères non numériques
$PinInput.AddHandler([System.Windows.Input.TextCompositionManager]::PreviewTextInputEvent, 
    [System.Windows.Input.TextCompositionEventHandler] {
        param($sender, $e)
        if ($e.Text -notmatch "^\d$") {
            $e.Handled = $true
        }
    })

$PinConfirm.AddHandler([System.Windows.Input.TextCompositionManager]::PreviewTextInputEvent, 
    [System.Windows.Input.TextCompositionEventHandler] {
        param($sender, $e)
        if ($e.Text -notmatch "^\d$") {
            $e.Handled = $true
        }
    })

# Ajouter des gestionnaires pour la validation en temps réel
$PinInput.Add_PasswordChanged({
    Update-PinBorderColors
    Update-ValidateButtonState
})

$PinConfirm.Add_PasswordChanged({
    Update-PinBorderColors
    Update-ValidateButtonState
})

# Fonction de validation du PIN
function Validate-PIN {
    param($Pin)
    
    # Vérifier la longueur et que ce sont uniquement des chiffres
    if ($Pin.Length -lt 6 -or $Pin.Length -gt 20 -or $Pin -notmatch "^\d+$") {
        return $false, "PIN invalide : 6 à 20 chiffres requis."
    }
    
    # Vérifier les séquences croissantes
    $isAscending = $true
    for ($i = 0; $i -lt $Pin.Length - 1; $i++) {
        $current = [int]::Parse($Pin[$i].ToString())
        $next = [int]::Parse($Pin[$i + 1].ToString())
        if ($next -ne ($current + 1)) {
            $isAscending = $false
            break
        }
    }
    
    if ($isAscending) {
        return $false, "PIN invalide : les chiffres ne doivent pas être en ordre croissant (ex : 123456)."
    }
    
    # Vérifier les séquences décroissantes
    $isDescending = $true
    for ($i = 0; $i -lt $Pin.Length - 1; $i++) {
        $current = [int]::Parse($Pin[$i].ToString())
        $next = [int]::Parse($Pin[$i + 1].ToString())
        if ($next -ne ($current - 1)) {
            $isDescending = $false
            break
        }
    }
    
    if ($isDescending) {
        return $false, "PIN invalide : les chiffres ne doivent pas être en ordre décroissant (ex : 654321)."
    }
    
    return $true, "OK"
}

# Fonction pour mettre à jour les couleurs des bordures en temps réel
function Update-PinBorderColors {
    $pin = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    # Si les champs sont vides, remettre la couleur par défaut
    if ([string]::IsNullOrEmpty($pin) -and [string]::IsNullOrEmpty($pinConfirm)) {
        $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x30, 0xFF, 0xFF, 0xFF))
        $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x30, 0xFF, 0xFF, 0xFF))
        return
    }

    # Valider le premier PIN
    $validationResult = Validate-PIN -Pin $pin
    $isValidPin = $validationResult[0]

    # Valider le second PIN
    $validationResult2 = Validate-PIN -Pin $pinConfirm
    $isValidPinConfirm = $validationResult2[0]

    # Déterminer les couleurs
    if (-not [string]::IsNullOrEmpty($pin)) {
        if (-not $isValidPin) {
            # Orange si non conforme
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xFF, 0xB7, 0x4D))
        } elseif (-not [string]::IsNullOrEmpty($pinConfirm) -and $pin -eq $pinConfirm) {
            # Vert lumineux si conforme et égaux
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0x66, 0xBB, 0x6A))
        } elseif (-not [string]::IsNullOrEmpty($pinConfirm) -and $pin -ne $pinConfirm) {
            # Rouge si différents
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xEF, 0x53, 0x50))
        } else {
            # Couleur par défaut si confirmation vide
            $PinInputBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x40, 0xFF, 0xFF, 0xFF))
        }
    }

    if (-not [string]::IsNullOrEmpty($pinConfirm)) {
        if (-not $isValidPinConfirm) {
            # Orange si non conforme
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xFF, 0xB7, 0x4D))
        } elseif (-not [string]::IsNullOrEmpty($pin) -and $pin -eq $pinConfirm -and $isValidPin) {
            # Vert lumineux si conforme et égaux
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0x66, 0xBB, 0x6A))
        } elseif (-not [string]::IsNullOrEmpty($pin) -and $pin -ne $pinConfirm) {
            # Rouge si différents
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0xFF, 0xEF, 0x53, 0x50))
        } else {
            # Couleur par défaut si PIN vide
            $PinConfirmBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(0x40, 0xFF, 0xFF, 0xFF))
        }
    }
}

# Fonction pour mettre à jour l'état du bouton Valider
function Update-ValidateButtonState {
    $pin = $PinInput.Password
    $pinConfirm = $PinConfirm.Password

    # Vérifier que les deux champs sont remplis
    if ([string]::IsNullOrEmpty($pin) -or [string]::IsNullOrEmpty($pinConfirm)) {
        $ValidateButton.IsEnabled = $false
        $ValidateButton.Opacity = 0.5
        return
    }

    # Vérifier que les deux PINs sont valides
    $validationResult1 = Validate-PIN -Pin $pin
    $validationResult2 = Validate-PIN -Pin $pinConfirm
    $isValidPin = $validationResult1[0]
    $isValidPinConfirm = $validationResult2[0]

    # Vérifier que les deux PINs correspondent
    $pinsMatch = $pin -eq $pinConfirm

    # Activer le bouton seulement si tout est conforme
    if ($isValidPin -and $isValidPinConfirm -and $pinsMatch) {
        $ValidateButton.IsEnabled = $true
        $ValidateButton.Opacity = 1.0
    } else {
        $ValidateButton.IsEnabled = $false
        $ValidateButton.Opacity = 0.5
    }
}

# Événement Valider : valider et fermer
$ValidateButton.Add_Click({
    $pin = $PinInput.Password
    $pinConfirm = $PinConfirm.Password
    
    # Vérifier que le PIN est valide
    $validationResult = Validate-PIN -Pin $pin
    $isValid = $validationResult[0]
    $errorMessage = $validationResult[1]
    
    if (-not $isValid) {
        [System.Windows.MessageBox]::Show($errorMessage, "Erreur", "OK", "Error")
        return
    }
    
    # Vérifier que les deux PINs correspondent
    if ($pin -ne $pinConfirm) {
        [System.Windows.MessageBox]::Show("Les deux codes PIN ne correspondent pas. Veuillez réessayer.", "Erreur", "OK", "Error")
        $PinConfirm.Clear()
        return
    }
    
    # Si tout est OK, marquer comme valide et continuer
    $script:UserAction = "Validated"
    $Window.DialogResult = $true
    $script:Pin = $pin
    $Window.Close()
})

# Événement Remettre à plus tard
$PostponeButton.Add_Click({
    Write-Output "Activation reportée via le bouton. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
    $script:UserAction = "Postponed"
    $Window.DialogResult = $false
    $Window.Close()
})

# Événement Bouton de fermeture X (même comportement que Remettre à plus tard)
$CloseButton.Add_Click({
    Write-Output "Activation reportée via le bouton X. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
    $script:UserAction = "Postponed"
    $Window.DialogResult = $false
    $Window.Close()
})

# Gérer la fermeture par la barre des tâches ou Alt+F4
$Window.Add_Closing({
    param($sender, $e)
    
    # Si l'utilisateur a validé, ne rien faire de spécial
    if ($script:UserAction -eq "Validated") {
        return
    }
    
    # Si l'utilisateur n'a pas cliqué sur un bouton (fermeture par la barre des tâches)
    if ([string]::IsNullOrEmpty($script:UserAction)) {
        Write-Output "Activation reportée via fermeture de la fenêtre. Reports restants : $($MaxPostpones - $CurrentPostponeCount - 1)"
        $script:UserAction = "Postponed"
    }
})

# Afficher la boîte de dialogue (avec try-catch pour erreur d'affichage)
try {
    $dialogResult = $Window.ShowDialog()
    Write-Output "DialogResult : $dialogResult"
} catch {
    Write-Error "Erreur d'affichage WPF : $($_.Exception.Message). Essayez sans AllowsTransparency si le problème persiste."
    exit 1
}

# Incrémenter le compteur si l'utilisateur a reporté
if ($script:UserAction -eq "Postponed") {
    $CurrentPostponeCount++
    $CurrentPostponeCount | Set-Content $CounterPath -Force
    Write-Output "Compteur incrémenté. Nouvelle valeur : $CurrentPostponeCount"
}

# Vérification préalable de l'état BitLocker
$blv = Get-BitLockerVolume -MountPoint "C:"
switch ($blv.VolumeStatus) {
    'EncryptionInProgress' {
        [System.Windows.MessageBox]::Show("Un chiffrement BitLocker est déjà en cours sur ce poste. Patientez jusqu'à la fin.", "Information", "OK", "Information")
        return
    }
    'DecryptionInProgress' {
        [System.Windows.MessageBox]::Show("Un déchiffrement BitLocker est actuellement en cours. Attendez qu'il soit terminé avant de relancer.", "Information", "OK", "Information")
        return
    }
    'FullyEncrypted' {
        if ($blv.ProtectionStatus -eq 'On') {
            [System.Windows.MessageBox]::Show("BitLocker est déjà activé sur ce poste.", "Information", "OK", "Information")
            return
        }
    }
}

# Traitement post-saisie (workflow RecoveryPassword + Backup AD + Enable-BitLocker TPM+PIN)
if ($dialogResult -eq $true -and $script:Pin) {

    $MountPoint       = "C:"
    $EncryptionMethod = "XtsAes256"
    $ComputerName     = $env:COMPUTERNAME

    # Helpers (localisés ici pour être autonome)
    function Get-Protector([string]$mp, [string]$type) {
        (Get-BitLockerVolume -MountPoint $mp).KeyProtector | Where-Object { $_.KeyProtectorType -eq $type }
    }
    function Get-FirstProtectorId([string]$mp, [string]$type) {
        Get-Protector -mp $mp -type $type | Select-Object -ExpandProperty KeyProtectorId -First 1
    }

    # Vérification état BitLocker
    $blv = Get-BitLockerVolume -MountPoint $MountPoint
    if ($blv.VolumeStatus -eq "EncryptionInProgress" -or $blv.VolumeStatus -eq "FullyEncrypted" -or $blv.ProtectionStatus -eq "On") {
        Write-Host "BitLocker est déjà activé ou en cours sur $MountPoint. Aucune action requise." -ForegroundColor Yellow
        [System.Windows.MessageBox]::Show(
            "BitLocker est déjà activé ou en cours de chiffrement sur ce poste.",
            "Information", "OK", "Information"
        )
        return
    }

    try {
        Write-Host "=== Activation BitLocker sur $MountPoint ($ComputerName) ===" -ForegroundColor Cyan

        # 1) RecoveryPassword (création ou réutilisation)
        Write-Host "Étape 1/3 : vérification du RecoveryPassword..." -ForegroundColor Cyan
        $recId = Get-FirstProtectorId -mp $MountPoint -type "RecoveryPassword"

        if (-not $recId) {
            Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop | Out-Null
            $recId = Get-FirstProtectorId -mp $MountPoint -type "RecoveryPassword"

            if (-not $recId) {
                throw "Impossible de récupérer l'ID du RecoveryPassword après création."
            }

            Write-Host "[OK] RecoveryPassword ajouté." -ForegroundColor Green
        } else {
            Write-Host "[OK] RecoveryPassword déjà présent (réutilisation)." -ForegroundColor Green
        }

        # 2) Backup AD (obligatoire si GPO l'exige)
        Write-Host "Étape 2/3 : sauvegarde du RecoveryPassword dans AD DS..." -ForegroundColor Cyan
        Backup-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recId -ErrorAction Stop | Out-Null
        Write-Host "[OK] Sauvegarde AD effectuée (Backup-BitLockerKeyProtector)." -ForegroundColor Green

        # 3) Enable-BitLocker (Used Space Only + XtsAes256 + TPM+PIN)
        Write-Host "Étape 3/3 : activation BitLocker (Used Space Only, TPM+PIN)..." -ForegroundColor Cyan

        # Conversion du PIN (GUI) en SecureString
        $UserPin = ConvertTo-SecureString $script:Pin -AsPlainText -Force

        # Si un TPM+PIN existait déjà, on le supprime pour éviter doublons
        $existingTpmPins = @(Get-Protector -mp $MountPoint -type "TpmPin")
        if ($existingTpmPins.Count -gt 0) {
            Write-Host "[WARN] Un protecteur TPM+PIN existe déjà : suppression avant recréation." -ForegroundColor DarkYellow
            foreach ($kp in $existingTpmPins) {
                Remove-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop
            }
            Write-Host "[OK] Protecteur(s) TPM+PIN supprimé(s)." -ForegroundColor Green
        }

        # Lancement de l'activation (pas de barre de progression)
        try {
            Enable-BitLocker -MountPoint $MountPoint `
                -EncryptionMethod $EncryptionMethod `
                -UsedSpaceOnly `
                -TpmAndPinProtector `
                -Pin $UserPin `
                -ErrorAction Stop | Out-Null

            Write-Host "[OK] Enable-BitLocker lancé." -ForegroundColor Green
        }
        catch {
            $msg = $_.Exception.Message
            $hr  = $_.Exception.HResult

            # GPO PIN pas encore appliquée
            if ($hr -eq -2144272384 -or $msg -match "0x80310060") {
                Write-Warning "La stratégie ne permet pas encore le PIN au démarrage (0x80310060)."
                [System.Windows.MessageBox]::Show(
                    "La stratégie BitLocker n'autorise pas encore le PIN au démarrage.`n`n" +
                    "Redémarrez l'ordinateur puis relancez le script.",
                    "Redémarrage requis", "OK", "Warning"
                )
                New-Item -ItemType File -Path "$env:ProgramData\BitLockerActivation\PendingReboot.flag" -Force | Out-Null
                exit 0
            }

            throw "Échec Enable-BitLocker : $msg"
        }

        # Succès -> suppression compteur de reports
        if (Test-Path $CounterPath) {
            Remove-Item $CounterPath -Force
        }

        [System.Windows.MessageBox]::Show(
            "BitLocker a été configuré. Un redémarrage est requis pour finaliser l'initialisation et démarrer le chiffrement.",
            "BitLocker", "OK", "Information"
        )

        Write-Host "[OK] Configuration terminée. Redémarrage requis." -ForegroundColor Green
    }
    catch {
        Write-Host "[ERR] $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show(
            "Erreur : $($_.Exception.Message)",
            "Erreur", "OK", "Error"
        )
    }

} else {
    Write-Output "Activation reportée ou annulée par l'utilisateur."
}

# Nettoyage sécurisé
$script:Pin = $null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()