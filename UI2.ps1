#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '1.12.0' }

Add-Type -AssemblyName PresentationFramework

param (
    [string]$DefaultSiteUrl = "https://google.com"
)

# === CONFIGURATION ===
$clientId = "<YOUR_CLIENT_ID>"
$tenantId = "<YOUR_TENANT_ID>"
$certPath = "<PATH_TO_PFX>"
$certPassword = ConvertTo-SecureString "<YOUR_CERT_PASSWORD>" -AsPlainText -Force
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath, $certPassword)

function Connect-PnPToSite {
    param($siteUrl)
    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Certificate $cert
}

function Get-SharePointLibraries {
    $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -notmatch "^(SiteAssets|Style Library|Form Templates|Site Pages|SiteCollectionDocuments|SiteCollectionImages)$" }
    return $lists
}

function Get-LibraryProperties {
    param($libraryTitle)
    return Get-PnPList -Identity $libraryTitle
}

function Update-LibraryProperties {
    param($libraryTitle, $syncEnabled, $offlineEnabled)

    $params = @{}
    if ($syncEnabled -ne $null) { $params.Add("EnableSync", $syncEnabled) }
    if ($offlineEnabled -ne $null) { $params.Add("EnableOfflineClientAccess", $offlineEnabled) }

    if ($params.Count -gt 0) {
        Set-PnPList -Identity $libraryTitle @params
    }
}

function Update-ToggleFromLibrary {
    param ($libraryTitle)

    try {
        $props = Get-LibraryProperties -libraryTitle $libraryTitle

        $rbSyncYes.IsChecked = $false
        $rbSyncNo.IsChecked = $false
        $rbOfflineYes.IsChecked = $false
        $rbOfflineNo.IsChecked = $false

        if ($props.EnableSync -eq $true) { $rbSyncYes.IsChecked = $true }
        elseif ($props.EnableSync -eq $false) { $rbSyncNo.IsChecked = $true }

        if ($props.EnableOfflineClientAccess -eq $true) { $rbOfflineYes.IsChecked = $true }
        elseif ($props.EnableOfflineClientAccess -eq $false) { $rbOfflineNo.IsChecked = $true }

        Log-Output "Properties loaded for '$libraryTitle'"
    } catch {
        Log-Output "Error loading properties for '$libraryTitle': $_"
    }
}

function Update-SiteSettings {
    param (
        [bool]$disableSyncApp,
        [bool]$clientApp,
        [bool]$offlineAvailability
    )

    if ($disableSyncApp -eq $true) {
        Add-PnPCustomAction -Name "DisableSync" -Title "DisableSync" -Location "ScriptLink" -Sequence 10000 -ScriptSrc "~sitecollection/Style Library/DisableSync.js" -Scope Site -Force
    } elseif ($disableSyncApp -eq $false) {
        Remove-PnPCustomAction -Name "DisableSync" -Scope Site -Force -ErrorAction SilentlyContinue
    }

    if ($clientApp -eq $true) {
        Enable-PnPFeature -Identity "8c0c38f5-6eeb-4c05-865a-0c865b367f9e" -Scope Site -Force
    } elseif ($clientApp -eq $false) {
        Disable-PnPFeature -Identity "8c0c38f5-6eeb-4c05-865a-0c865b367f9e" -Scope Site -Force
    }

    Set-PnPSite -AllowDownloadingNonWebViewableFiles $offlineAvailability
}

function Update-SiteToggles {
    try {
        $customAction = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq "DisableSync" }
        $rbDisableSyncYes.IsChecked = $customAction -ne $null
        $rbDisableSyncNo.IsChecked = -not $rbDisableSyncYes.IsChecked

        $feature = Get-PnPFeature -Scope Site | Where-Object { $_.DefinitionId -eq "8c0c38f5-6eeb-4c05-865a-0c865b367f9e" }
        $rbClientYes.IsChecked = $feature -ne $null
        $rbClientNo.IsChecked = -not $rbClientYes.IsChecked

        $siteProps = Get-PnPSite
        $rbSiteOfflineYes.IsChecked = $siteProps.AllowDownloadingNonWebViewableFiles
        $rbSiteOfflineNo.IsChecked = -not $rbSiteOfflineYes.IsChecked

        Log-Output "Site toggle values loaded."
    } catch {
        Log-Output "Error loading site toggle values: $_"
    }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="SharePoint Library Manager" Height="600" Width="800">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
            <Label Content="SharePoint Site URL:" Width="150" VerticalAlignment="Center"/>
            <TextBox Name="txtSiteUrl" Width="500"/>
            <Button Name="btnGo" Content="Go" Width="80" Margin="10,0,0,0"/>
        </StackPanel>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <ListBox Name="lstLibraries" SelectionMode="Extended" Grid.Column="0"/>

            <StackPanel Grid.Column="1" Margin="10,0,0,0">
                <GroupBox Header="Library Settings">
                    <StackPanel>
                        <Label Content="Sync with OneDrive:"/>
                        <StackPanel Orientation="Horizontal">
                            <RadioButton Name="rbSyncYes" Content="Yes" GroupName="sync"/>
                            <RadioButton Name="rbSyncNo" Content="No" GroupName="sync" Margin="10,0,0,0"/>
                        </StackPanel>
                        <Label Content="Access Without Internet:" Margin="0,10,0,0"/>
                        <StackPanel Orientation="Horizontal">
                            <RadioButton Name="rbOfflineYes" Content="Yes" GroupName="offline"/>
                            <RadioButton Name="rbOfflineNo" Content="No" GroupName="offline" Margin="10,0,0,0"/>
                        </StackPanel>
                        <Button Name="btnApplyLibrary" Content="Apply Library Changes" Margin="0,10,0,0"/>
                    </StackPanel>
                </GroupBox>

                <GroupBox Header="Site Settings" Margin="0,20,0,0">
                    <StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <Label Content="Disable Sync App:" Width="180"/>
                            <RadioButton Name="rbDisableSyncYes" Content="Yes" GroupName="syncApp"/>
                            <RadioButton Name="rbDisableSyncNo" Content="No" GroupName="syncApp" Margin="10,0,0,0"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                            <Label Content="Default Client App:" Width="180"/>
                            <RadioButton Name="rbClientYes" Content="Yes" GroupName="clientApp"/>
                            <RadioButton Name="rbClientNo" Content="No" GroupName="clientApp" Margin="10,0,0,0"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                            <Label Content="Site Offline Availability:" Width="180"/>
                            <RadioButton Name="rbSiteOfflineYes" Content="Yes" GroupName="offlineAvail"/>
                            <RadioButton Name="rbSiteOfflineNo" Content="No" GroupName="offlineAvail" Margin="10,0,0,0"/>
                        </StackPanel>
                        <Button Name="btnApplySite" Content="Apply Site Changes" Margin="0,10,0,0"/>
                    </StackPanel>
                </GroupBox>
            </StackPanel>
        </Grid>

        <DockPanel Grid.Row="2" Margin="0,10,0,0">
            <TextBox Name="txtLog" Height="80" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True" DockPanel.Dock="Top"/>
            <Button Name="btnExit" Content="Exit" Width="80" Height="30" HorizontalAlignment="Right" Margin="0,5,0,0" DockPanel.Dock="Bottom"/>
        </DockPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Link elements
$txtSiteUrl = $window.FindName("txtSiteUrl")
$btnGo = $window.FindName("btnGo")
$lstLibraries = $window.FindName("lstLibraries")
$btnApplyLibrary = $window.FindName("btnApplyLibrary")
$btnApplySite = $window.FindName("btnApplySite")
$rbSyncYes = $window.FindName("rbSyncYes")
$rbSyncNo = $window.FindName("rbSyncNo")
$rbOfflineYes = $window.FindName("rbOfflineYes")
$rbOfflineNo = $window.FindName("rbOfflineNo")
$rbDisableSyncYes = $window.FindName("rbDisableSyncYes")
$rbDisableSyncNo = $window.FindName("rbDisableSyncNo")
$rbClientYes = $window.FindName("rbClientYes")
$rbClientNo = $window.FindName("rbClientNo")
$rbSiteOfflineYes = $window.FindName("rbSiteOfflineYes")
$rbSiteOfflineNo = $window.FindName("rbSiteOfflineNo")
$txtLog = $window.FindName("txtLog")
$btnExit = $window.FindName("btnExit")

# Initialize default URL
$txtSiteUrl.Text = $DefaultSiteUrl

function Log-Output {
    param([string]$message)
    $txtLog.Dispatcher.Invoke([action]{ $txtLog.AppendText("$message`n"); $txtLog.ScrollToEnd() })
}

$btnExit.Add_Click({
    Log-Output "Exiting tool..."
    $window.Close()
})

$global:librariesMap = @{}

$btnGo.Add_Click({
    $siteUrl = $txtSiteUrl.Text
    if (-not $siteUrl) { Log-Output "Please enter a SharePoint URL."; return }

    try {
        Connect-PnPToSite -siteUrl $siteUrl
        $lstLibraries.Items.Clear()
        $libraries = Get-SharePointLibraries
        $global:librariesMap.Clear()
        foreach ($lib in $libraries) {
            $lstLibraries.Items.Add($lib.Title)
            $global:librariesMap[$lib.Title] = $lib.Title
        }
        Update-SiteToggles
        Log-Output "Libraries and site settings loaded. Now select a library to proceed."
    } catch {
        Log-Output "Error: $_"
    }
})

$lstLibraries.Add_SelectionChanged({
    if ($lstLibraries.SelectedItems.Count -eq 1) {
        $selectedLib = $lstLibraries.SelectedItem
        Update-ToggleFromLibrary -libraryTitle $selectedLib
    }
})

$btnApplyLibrary.Add_Click({
    $sync = if ($rbSyncYes.IsChecked) { $true } elseif ($rbSyncNo.IsChecked) { $false } else { $null }
    $offline = if ($rbOfflineYes.IsChecked) { $true } elseif ($rbOfflineNo.IsChecked) { $false } else { $null }
    $selected = $lstLibraries.SelectedItems
    if ($selected.Count -eq 0) { Log-Output "Please select at least one library."; return }
    foreach ($libName in $selected) {
        Update-LibraryProperties -libraryTitle $libName -syncEnabled $sync -offlineEnabled $offline
        Log-Output "Updated: $libName"
    }
    Log-Output "Library settings updated."
})

$btnApplySite.Add_Click({
    $disableSync = if ($rbDisableSyncYes.IsChecked) { $true } elseif ($rbDisableSyncNo.IsChecked) { $false } else { $null }
    $clientApp = if ($rbClientYes.IsChecked) { $true } elseif ($rbClientNo.IsChecked) { $false } else { $null }
    $offlineSite = if ($rbSiteOfflineYes.IsChecked) { $true } elseif ($rbSiteOfflineNo.IsChecked) { $false } else { $null }

    Update-SiteSettings -disableSyncApp $disableSync -clientApp $clientApp -offlineAvailability $offlineSite
    Log-Output "Site settings updated."
})

$window.ShowDialog() | Out-Null
