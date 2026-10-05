#Requires -Version 5.1
<#
.SYNOPSIS
    Deaktiviert Company-wide Sharing Links auf SharePoint-Sites (Tenant bioland365).
.DESCRIPTION
    Ohne -Apply wird nur ein Trockenlauf durchgeführt. Es wird KEIN Backup des Ist-Zustands
    erstellt, nur ein Protokoll (CSV) mit dem Ergebnis je Site. Das Protokoll dient auch als
    Grundlage für Restore-CompanyWideLinks.ps1 (Sites mit Status 'OK').

    Ausgenommen werden: Redirect-Sites, App-Katalog, Suchcenter, PWA, Document Center,
    My-Site-Host, Kanal-Sites (TEAMCHANNEL#1, außer mit -IncludeChannelSites), Sites, deren
    LockState nicht 'Unlock' ist, die Hubs aus -ExcludeHubUrl samt allen verbundenen Sites
    sowie die Sites aus -ExcludeUrl.
.EXAMPLE
    .\Set-CompanyWideLinks.ps1                                   # Trockenlauf
    .\Set-CompanyWideLinks.ps1 -OnlyUrl https://bioland365.sharepoint.com/sites/Test -Apply   # Pilot
    .\Set-CompanyWideLinks.ps1 -Apply                            # produktiv
#>
[CmdletBinding()]
param(
    [string]$AdminUrl = 'https://bioland365-admin.sharepoint.com',
    [string]$OutDir   = 'C:\Users\mschmitt\Documents',
    [string[]]$OnlyUrl,                 # nur diese Sites bearbeiten (Pilot/Stichprobe)
    [string[]]$ExcludeHubUrl = @('https://bioland365.sharepoint.com/sites/Projekthub'),   # Hub + alle verbundenen Sites ausnehmen
    [string[]]$ExcludeUrl    = @('https://bioland365.sharepoint.com/sites/BiolandQMS'),   # einzelne Sites ausnehmen
    [switch]$IncludeChannelSites,       # TEAMCHANNEL#1 einbeziehen (Standard: ausgenommen)
    [switch]$Apply                      # ohne -Apply = Trockenlauf
)

$ErrorActionPreference = 'Stop'

# Templates, die nie angefasst werden (Vergleich ist nicht case-sensitiv)
$excludedTemplates = @('APPCATALOG#0', 'SPSMSITEHOST#0', 'SRCHCEN#0', 'PWA#0', 'BDR#0')

function Format-Url([string]$u) { $u.Trim().TrimEnd('/').ToLowerInvariant() }

# Modul laden (unter PowerShell 7 nur über Windows-PowerShell-Kompatibilität)
if ($PSVersionTable.PSEdition -eq 'Core') {
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
} else {
    Import-Module Microsoft.Online.SharePoint.PowerShell
}

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$logPath = Join-Path $OutDir ("SPO_CompanyWideLinks_Log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))

Connect-SPOService -Url $AdminUrl

if (-not (Get-Command Set-SPOSite).Parameters.ContainsKey('DisableCompanyWideSharingLinks')) {
    throw 'Parameter DisableCompanyWideSharingLinks fehlt - Modul aktualisieren (Update-Module Microsoft.Online.SharePoint.PowerShell).'
}

# Hubs auflösen: ID(s) der auszunehmenden Hubs ermitteln. Schlägt das fehl, bricht das Skript ab,
# damit keine Site des Hubs versehentlich geändert wird.
$hubIds = @()
foreach ($hubUrl in $ExcludeHubUrl) {
    try {
        $hub = Get-SPOHubSite -Identity $hubUrl -ErrorAction Stop
    } catch {
        throw "Hub '$hubUrl' nicht gefunden oder nicht lesbar: $($_.Exception.Message)"
    }
    $hubIds += @($hub.ID, $hub.SiteId) | Where-Object { $_ -and $_ -ne [guid]::Empty }
}
$hubIds = @($hubIds | ForEach-Object { "$_".ToLowerInvariant() } | Select-Object -Unique)
if ($ExcludeHubUrl -and -not $hubIds) { throw 'Hub-ID konnte nicht ermittelt werden - Abbruch.' }

$excludeUrlSet = @($ExcludeUrl + $ExcludeHubUrl | Where-Object { $_ } | ForEach-Object { Format-Url $_ })

function Test-HubMember($site) { $hubIds -contains "$($site.HubSiteId)".ToLowerInvariant() }

$sites = Get-SPOSite -Limit All

# Plausibilitätsprüfung: Es muss mindestens eine Site dem Hub zugeordnet sein (der Hub selbst zählt mit).
# Sonst liefert die Listenabfrage HubSiteId nicht mit - dann mit -Detailed neu lesen.
if ($hubIds -and -not ($sites | Where-Object { Test-HubMember $_ })) {
    Write-Warning 'HubSiteId in der Listenabfrage nicht auswertbar - lese erneut mit -Detailed (langsamer).'
    $sites = Get-SPOSite -Limit All -Detailed
    if (-not ($sites | Where-Object { Test-HubMember $_ })) {
        throw 'Keine Hub-Zugehörigkeit auswertbar - Abbruch, damit keine Hub-Sites geändert werden.'
    }
}

$hubMembers = @($sites | Where-Object { Test-HubMember $_ })
Write-Host "Hub-Sites (ausgenommen): $($hubMembers.Count)"
$hubMembers | Select-Object -ExpandProperty Url | ForEach-Object { Write-Host "  $_" }

if ($OnlyUrl) { $sites = $sites | Where-Object { $_.Url -in $OnlyUrl } }
if (-not $sites) { throw 'Keine Sites im Geltungsbereich gefunden.' }

# Ändern bzw. Trockenlauf
$i = 0
$log = foreach ($s in $sites) {
    $i++
    Write-Progress -Activity 'Company-wide Links deaktivieren' -Status $s.Url -PercentComplete ($i / $sites.Count * 100)

    if ($excludeUrlSet -contains (Format-Url $s.Url)) {
        [pscustomobject]@{ Url = $s.Url; Status = 'Übersprungen (ausgenommene Site)' }; continue
    }
    if (Test-HubMember $s) {
        [pscustomobject]@{ Url = $s.Url; Status = 'Übersprungen (Hub-Site)' }; continue
    }
    if ($s.Template -like 'RedirectSite*' -or $s.Template -in $excludedTemplates) {
        [pscustomobject]@{ Url = $s.Url; Status = "Übersprungen (Template $($s.Template))" }; continue
    }
    if ($s.Template -eq 'TEAMCHANNEL#1' -and -not $IncludeChannelSites) {
        [pscustomobject]@{ Url = $s.Url; Status = 'Übersprungen (Kanal-Site)' }; continue
    }
    if ($s.LockState -ne 'Unlock') {
        [pscustomobject]@{ Url = $s.Url; Status = "Übersprungen (LockState $($s.LockState))" }; continue
    }
    if ($s.DisableCompanyWideSharingLinks -eq 'Disabled') {
        [pscustomobject]@{ Url = $s.Url; Status = 'Übersprungen (bereits Disabled)' }; continue
    }
    if (-not $Apply) {
        [pscustomobject]@{ Url = $s.Url; Status = 'Trockenlauf' }; continue
    }

    $status = $null
    foreach ($try in 1..3) {
        try {
            Set-SPOSite -Identity $s.Url -DisableCompanyWideSharingLinks Disabled -ErrorAction Stop
            $status = 'OK'; break
        } catch {
            $status = "Fehler: $($_.Exception.Message)"
            if ($try -lt 3) { Start-Sleep -Seconds (5 * $try) }
        }
    }
    [pscustomobject]@{ Url = $s.Url; Status = $status }
}
Write-Progress -Activity 'Company-wide Links deaktivieren' -Completed

$log | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Protokoll: $logPath"
$log | Group-Object { ($_.Status -split ' ')[0] } | Select-Object Count, Name | Format-Table -AutoSize
$log | Where-Object { $_.Status -like 'Fehler*' }
