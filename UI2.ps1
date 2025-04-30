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

[xml]$xaml = @" ... "@

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
