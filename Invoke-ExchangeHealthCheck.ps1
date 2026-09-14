#Requires -Version 5.1
<#
.SYNOPSIS
    Health check multi-server per Exchange on-premises con alerting via e-mail.

.DESCRIPTION
    Esegue una serie di controlli sull'intera organizzazione Exchange:
      - raggiungibilita' dei server (WinRM/CIM) e uptime
      - servizi Exchange (Win32_Service + Test-ServiceHealth)
      - spazio disco incluse mount point (Win32_Volume)
      - memoria / CPU / reboot pending
      - ServerComponentState (server rimasti in maintenance mode)
      - Managed Availability (Get-HealthReport)
      - DAG: stato gruppo, witness, copy status, copy/replay queue, content index
      - Test-ReplicationHealth
      - Database: mount state, backup age, copia attiva non su preference 1
      - Code di trasporto (Get-Queue) + back pressure
      - Certificati in scadenza
      - Test-MapiConnectivity

    Le anomalie vengono deduplicate su file di stato (niente mail storm), con
    cooldown configurabile, notifica di escalation e mail di rientro (recovery).

.PARAMETER ConfigPath
    Percorso del file JSON di configurazione. Default: .\ExchangeHealthCheck.config.json

.PARAMETER Server
    Limita l'esecuzione ai server indicati (wildcard ammesse). Sovrascrive la config.

.PARAMETER Check
    Esegue solo i check indicati. Valori: Os, Disk, Services, Components, Health,
    Dag, Replication, Databases, Queues, BackPressure, Certificates, Mapi.

.PARAMETER NoMail
    Esegue i controlli e scrive log/report ma non invia e-mail.

.PARAMETER ForceMail
    Ignora il cooldown e invia comunque il riepilogo.

.PARAMETER TestMail
    Invia solo una mail di test per validare la configurazione SMTP ed esce.

.PARAMETER PassThru
    Ritorna in pipeline tutti i finding.

.EXAMPLE
    .\Invoke-ExchangeHealthCheck.ps1

.EXAMPLE
    .\Invoke-ExchangeHealthCheck.ps1 -Check Disk,Queues -NoMail -Verbose

.NOTES
    Eseguire da Exchange Management Shell oppure lasciare che lo script carichi
    lo snap-in / apra una remote session verso un server Exchange (vedi config).
    Account richiesto: View-Only Organization Management + amministratore locale
    sui server (necessario per WinRM/CIM remoto).
#>

[CmdletBinding()]
param(
    [string]   $ConfigPath,
    [string[]] $Server,
    [ValidateSet('Os','Disk','Services','Components','Health','Dag','Replication','Databases','Queues','BackPressure','Certificates','Mapi')]
    [string[]] $Check,
    [switch]   $NoMail,
    [switch]   $ForceMail,
    [switch]   $TestMail,
    [switch]   $PassThru
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
$script:StartTime      = Get-Date
$script:ScriptRoot     = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:Findings       = New-Object System.Collections.Generic.List[object]
$script:QueueSummary   = New-Object System.Collections.Generic.List[object]
$script:LogFile        = $null
$script:ExSession      = $null
$script:CheckFilter    = $Check

#region ---------------------------------------------------------------- CONFIG

$DefaultConfigJson = @'
{
  "Organization": {
    "Name": "Exchange",
    "ConnectTo": "",
    "ViewEntireForest": true,
    "AutoDetectDomainController": true,
    "PreferredDomainController": "",
    "PreferredGlobalCatalog": ""
  },
  "Servers": {
    "Include": ["*"],
    "Exclude": [],
    "IncludeEdge": false,
    "UseFqdnForRemoting": true,
    "SkipExchangeChecksWhenOffline": true,
    "SiteFilter": []
  },
  "Checks": {
    "Os": true,
    "Disk": true,
    "Services": true,
    "Components": true,
    "Health": true,
    "Dag": true,
    "Replication": true,
    "Databases": true,
    "Queues": true,
    "BackPressure": true,
    "Certificates": true,
    "Mapi": false
  },
  "Thresholds": {
    "DiskMode": "And",
    "DiskFreePercentWarning": 20,
    "DiskFreePercentCritical": 10,
    "DiskFreeGBWarning": 60,
    "DiskFreeGBCritical": 25,
    "MinimumVolumeSizeGB": 5,
    "MemoryFreePercentWarning": 6,
    "MemoryFreePercentCritical": 3,
    "CpuLoadPercentWarning": 95,
    "UptimeMinutesMinimum": 15,
    "QueueWarning": 100,
    "QueueCritical": 500,
    "SubmissionQueueWarning": 50,
    "SubmissionQueueCritical": 250,
    "PoisonQueueWarning": 1,
    "RetryQueueWarning": 20,
    "CopyQueueWarning": 10,
    "CopyQueueCritical": 50,
    "ReplayQueueWarning": 20,
    "ReplayQueueCritical": 100,
    "BackupAgeHoursWarning": 36,
    "BackupAgeHoursCritical": 72,
    "CertExpiryDaysWarning": 30,
    "CertExpiryDaysCritical": 7,
    "ThrottleLimit": 24,
    "RemoteOpenTimeoutSeconds": 20,
    "RemoteOperationTimeoutSeconds": 120
  },
  "VolumeOverrides": [],
  "HealthReport": {
    "IncludeFailingMonitors": true,
    "MaxMonitorsPerHealthSet": 5
  },
  "Queues": {
    "ResolveNextHopHostnames": true,
    "ReverseDnsTimeoutMs": 1000,
    "TryNetBiosFallback": true,
    "NetBiosTimeoutMs": 1500,
    "ResolveSendConnectorName": true
  },
  "Ignore": {
    "Services": ["MSExchangePOP3", "MSExchangePOP3BE", "MSExchangeIMAP4", "MSExchangeIMAP4BE", "MSExchangeEdgeSync"],
    "ServerComponents": ["ForwardSyncDaemon", "ProvisioningRps"],
    "HealthSets": ["FfoQuarantine", "Monitoring", "OutsideInHealth", "Imap", "Pop"],
    "Volumes": [],
    "Databases": [],
    "Keys": []
  },
  "ExtraServices": ["W3SVC", "WinRM", "RemoteRegistry"],
  "Alerting": {
    "CooldownMinutes": 120,
    "SendRecovery": true,
    "HeartbeatHours": 24,
    "NotifyOnUnknown": true,
    "MinimumSeverityToMail": "Warning"
  },
  "Mail": {
    "Enabled": true,
    "SmtpServers": ["smtp.contoso.local"],
    "Port": 25,
    "UseSsl": false,
    "UseDefaultCredentials": false,
    "CredentialFile": "",
    "From": "exchange-monitor@contoso.com",
    "FromDisplayName": "Exchange Health Monitor",
    "To": ["messaging-team@contoso.com"],
    "Cc": [],
    "SubjectPrefix": "[Exchange Health]",
    "AttachCsv": true,
    "IncludeQueueSummary": true
  },
  "Console": {
    "ShowQueueSummary": true
  },
  "Paths": {
    "LogDirectory": "Logs",
    "ReportDirectory": "Reports",
    "StateFile": "State\\alert-state.json",
    "RetentionDays": 30
  }
}
'@

function Merge-HcConfig {
    param($Default, $Override)
    if ($null -eq $Override) { return $Default }
    if (-not ($Default -is [System.Management.Automation.PSCustomObject]) -or
        -not ($Override -is [System.Management.Automation.PSCustomObject])) { return $Override }
    $result = $Default.PSObject.Copy()
    foreach ($prop in $Override.PSObject.Properties) {
        if ($result.PSObject.Properties.Name -contains $prop.Name) {
            $result.$($prop.Name) = Merge-HcConfig -Default $result.$($prop.Name) -Override $prop.Value
        }
        else {
            Add-Member -InputObject $result -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
        }
    }
    return $result
}

function Resolve-HcPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $script:ScriptRoot $Path)
}

#endregion

#region --------------------------------------------------------------- LOGGING

function Write-HcLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('DEBUG','INFO','WARN','ERROR')][string]$Level = 'INFO',
        # Colore esplicito per la console. Serve ai finding, che hanno una severita
        # propria: senza, Critical e Warning finirebbero entrambi in giallo e con
        # -Verbose si confonderebbero con le righe di diagnostica.
        [string]$Color,
        # Riga di dettaglio: a video solo con -Verbose, ma sempre nel file di log.
        [switch]$VerboseOnly
    )
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    $show = -not ($VerboseOnly -and $VerbosePreference -eq 'SilentlyContinue')
    if ($show) {
        if ($Color) {
            Write-Host $line -ForegroundColor $Color
        }
        else {
            switch ($Level) {
                'ERROR' { Write-Host $line -ForegroundColor Red }
                'WARN'  { Write-Host $line -ForegroundColor Yellow }
                'DEBUG' { Write-Verbose $line }
                default { Write-Host $line }
            }
        }
    }

    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
    }
}

function Remove-HcOldFiles {
    param([string]$Directory, [int]$RetentionDays)
    if (-not $Directory -or -not (Test-Path -LiteralPath $Directory)) { return }
    if ($RetentionDays -le 0) { return }
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit } |
        ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { } }
}

#endregion

#region -------------------------------------------------------------- FINDINGS

function Get-SeverityRank {
    param([string]$Severity)
    switch ($Severity) {
        'Critical' { 4 }
        'Unknown'  { 3 }
        'Warning'  { 2 }
        'Info'     { 1 }
        default    { 0 }
    }
}

# Colore console per severita, condiviso tra i finding e il riepilogo code: cosi
# i due output restano visivamente coerenti (rosso = Critical, verde = OK, ...).
function Get-HcConsoleColor {
    param([string]$Severity)
    switch ($Severity) {
        'Critical' { 'Red' }
        'Unknown'  { 'Magenta' }
        'Warning'  { 'Yellow' }
        'Info'     { 'Cyan' }
        default    { 'Green' }
    }
}

function Add-Finding {
    param(
        [Parameter(Mandatory)][string]$Category,
        [string]$Server = '-',
        [string]$Item   = '-',
        [ValidateSet('OK','Info','Warning','Critical','Unknown')][string]$Severity = 'OK',
        [string]$Message,
        [object]$Value
    )
    $key = '{0}|{1}|{2}' -f $Category, $Server, $Item
    $ignoreKeys = @($script:Config.Ignore.Keys)
    foreach ($pattern in $ignoreKeys) {
        if ($pattern -and $key -like $pattern) {
            Write-HcLog "Finding ignorato da configurazione: $key" -Level DEBUG
            return
        }
    }
    $finding = [pscustomobject]@{
        Timestamp = Get-Date
        Category  = $Category
        Server    = $Server
        Item      = $Item
        Severity  = $Severity
        Message   = $Message
        Value     = if ($null -eq $Value) { '' } else { [string]$Value }
        Key       = $key
    }
    $script:Findings.Add($finding) | Out-Null

    # Colore per severita: rosso i Critical, verde gli OK. I finding non allarmanti
    # restano a video solo con -Verbose, ma finiscono comunque nel file di log.
    $color = Get-HcConsoleColor $Severity
    $level = switch ($Severity) {
        'Critical' { 'ERROR' }
        'Unknown'  { 'WARN'  }
        'Warning'  { 'WARN'  }
        default    { 'INFO'  }
    }
    $detailOnly = ((Get-SeverityRank $Severity) -lt 2)
    Write-HcLog ('{0,-8} {1,-22} {2,-18} {3}' -f $Severity.ToUpperInvariant(), $Category, $Server, $Message) `
        -Level $level -Color $color -VerboseOnly:$detailOnly
}

function Test-CheckEnabled {
    param([Parameter(Mandatory)][string]$Name)
    if ($script:CheckFilter) { return ($script:CheckFilter -contains $Name) }
    $value = $script:Config.Checks.$Name
    if ($null -eq $value) { return $true }
    return [bool]$value
}

function Test-HcCommand {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

# Restituisce il primo valore non vuoto tra le proprieta indicate. Serve dove
# cmdlet diversi (o versioni diverse di Exchange) espongono lo stesso dato con
# nomi differenti: meglio leggere entrambe le forme che stampare campi vuoti.
function Get-HcFirstValue {
    param([object]$InputObject, [string[]]$PropertyNames)
    if ($null -eq $InputObject) { return $null }
    foreach ($name in $PropertyNames) {
        if ($InputObject.PSObject.Properties.Name -contains $name) {
            $value = $InputObject.$name
            if ($null -ne $value -and [string]$value -ne '') { return $value }
        }
    }
    return $null
}

# Esegue un blocco di check isolando le eccezioni: un check che esplode non deve
# fermare l'intero giro, diventa un finding "Unknown".
function Invoke-HcCheck {
    param(
        [Parameter(Mandatory)][string]$Category,
        [string]$ServerName = '-',
        [Parameter(Mandatory)][scriptblock]$Body
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        & $Body
    }
    catch {
        Add-Finding -Category $Category -Server $ServerName -Item 'CheckExecution' -Severity 'Unknown' `
            -Message ('Check non eseguibile: {0}' -f $_.Exception.Message)
    }
    finally {
        $sw.Stop()
        Write-HcLog ('Check {0} [{1}] in {2:N1}s' -f $Category, $ServerName, $sw.Elapsed.TotalSeconds) -Level DEBUG
    }
}

#endregion

#region ------------------------------------------------------- EXCHANGE ACCESS

function Connect-HcExchange {
    if (Test-HcCommand 'Get-ExchangeServer') {
        Write-HcLog 'Cmdlet Exchange gia disponibili nella sessione corrente.'
    }
    else {
        $snapin = Get-PSSnapin -Registered -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn' -ErrorAction SilentlyContinue
        if ($snapin) {
            Write-HcLog 'Carico lo snap-in Exchange locale.'
            Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn' -ErrorAction Stop
        }
        else {
            $target = $script:Config.Organization.ConnectTo
            if ([string]::IsNullOrWhiteSpace($target)) {
                throw 'Cmdlet Exchange non disponibili e Organization.ConnectTo non valorizzato nella configurazione.'
            }
            Write-HcLog "Apro una remote PowerShell session verso $target."
            $uri = 'http://{0}/PowerShell/' -f $target
            $script:ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri `
                -Authentication Kerberos -ErrorAction Stop
            Import-PSSession -Session $script:ExSession -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
        }
    }

    # Contesto AD della sessione. In foreste multi-dominio ViewEntireForest puo far
    # risolvere gli oggetti su un DC di un altro dominio (anche solo trusted), con
    # errori tipo "object '*\SERVER' could not be found on <DC>". In quel caso si
    # mette ViewEntireForest a false oppure si fissa il DC del dominio corretto.
    if (Test-HcCommand 'Set-ADServerSettings') {
        # Il valore va IMPOSTATO sempre, anche quando e false: limitarsi ad attivarlo
        # quando e true lascia la sessione con l'impostazione che aveva gia (la EMS
        # puo averla a true da profilo o da un'esecuzione precedente), e in
        # configurazione si legge false mentre in realta e attivo.
        $adSettings = @{ ViewEntireForest = [bool]$script:Config.Organization.ViewEntireForest }

        $preferredDc = [string]$script:Config.Organization.PreferredDomainController
        if ($preferredDc) { $adSettings['PreferredServer'] = $preferredDc }

        $preferredGc = [string]$script:Config.Organization.PreferredGlobalCatalog
        if ($preferredGc) { $adSettings['PreferredGlobalCatalog'] = $preferredGc }

        if ($adSettings.Count -gt 0) {
            try {
                Set-ADServerSettings @adSettings -ErrorAction Stop
                Write-HcLog ('Contesto AD: {0}' -f (($adSettings.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join ', '))
            }
            catch {
                Write-HcLog ('Set-ADServerSettings fallito: {0}' -f $_.Exception.Message) -Level WARN
            }
        }

        # Cosa sta davvero usando la sessione: gli oggetti server di Exchange stanno
        # nel Configuration NC, replicato a tutti i DC della STESSA foresta. Un DC di
        # una foresta trusted non li vedra mai, e l'errore e "could not be found".
        if (Test-HcCommand 'Get-ADServerSettings') {
            try {
                $effective = Get-ADServerSettings -ErrorAction Stop
                Write-HcLog ('AD in uso -> ViewEntireForest={0} | DomainController={1} | GlobalCatalog={2}' -f `
                    $effective.ViewEntireForest,
                    (@($effective.PreferredDomainControllers) -join ','),
                    $effective.DefaultGlobalCatalog)
            }
            catch {
                Write-HcLog ('Get-ADServerSettings non disponibile: {0}' -f $_.Exception.Message) -Level DEBUG
            }
        }
    }
}

# Individua un domain controller VIVO nel dominio indicato. FindOne() fa una vera
# chiamata al locator, quindi restituisce un DC che risponde adesso: la ricerca si
# ripete a ogni esecuzione, senza fissare per sempre un DC che domani potrebbe
# essere spento. Usa System.DirectoryServices.ActiveDirectory, presente su ogni
# Windows: nessuna dipendenza da RSAT o dal modulo ActiveDirectory.
function Resolve-HcDomainController {
    param([Parameter(Mandatory)][string]$DomainName)
    try {
        $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName)
        $controller = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($context)
        return [string]$controller.Name
    }
    catch {
        Write-HcLog ('Nessun domain controller trovato per "{0}": {1}' -f $DomainName, $_.Exception.Message) -Level WARN
        return $null
    }
}

# Il dominio dei server Exchange si ricava dal loro FQDN: si prende il suffisso
# piu diffuso, cosi un singolo server anomalo non sposta la scelta.
function Get-HcServerDomain {
    param([Parameter(Mandatory)][object[]]$Targets)
    $suffixes = @($Targets |
        Where-Object { $_.Fqdn -and $_.Fqdn.Contains('.') } |
        ForEach-Object { ($_.Fqdn -split '\.', 2)[1] })
    if ($suffixes.Count -eq 0) { return $null }
    return ($suffixes | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
}

# Aggancia la sessione a un DC del dominio in cui vivono i server Exchange.
# Chiamata dopo aver risolto il perimetro, quando gli FQDN sono noti.
function Set-HcAutoDomainController {
    param([Parameter(Mandatory)][object[]]$Targets)

    if (-not (Test-HcCommand 'Set-ADServerSettings')) { return }
    if ([string]$script:Config.Organization.PreferredDomainController) { return }  # scelta esplicita: si rispetta
    if (-not $script:Config.Organization.AutoDetectDomainController) { return }

    $domain = Get-HcServerDomain -Targets $Targets
    if (-not $domain) {
        Write-HcLog 'Dominio dei server Exchange non deducibile dagli FQDN: lascio la scelta del DC a Exchange.' -Level DEBUG
        return
    }

    $controller = Resolve-HcDomainController -DomainName $domain
    if (-not $controller) { return }

    try {
        Set-ADServerSettings -PreferredServer $controller -ErrorAction Stop
        Write-HcLog ('Domain controller rilevato automaticamente per il dominio "{0}": {1}' -f $domain, $controller)
    }
    catch {
        Write-HcLog ('Impossibile agganciare il DC {0}: {1}' -f $controller, $_.Exception.Message) -Level WARN
    }
}

function Disconnect-HcExchange {
    if ($script:ExSession) {
        try { Remove-PSSession -Session $script:ExSession -ErrorAction Stop } catch { }
        $script:ExSession = $null
    }
}

function Get-HcTargetServer {
    $include = if ($Server) { $Server } else { @($script:Config.Servers.Include) }
    if (-not $include -or $include.Count -eq 0) { $include = @('*') }
    $exclude = @($script:Config.Servers.Exclude)
    $sites   = @($script:Config.Servers.SiteFilter)

    $all = @(Get-ExchangeServer -ErrorAction Stop)
    Write-HcLog ('Server Exchange trovati in organizzazione: {0}' -f $all.Count)

    $selected = foreach ($srv in $all) {
        $role = [string]$srv.ServerRole
        if (-not $script:Config.Servers.IncludeEdge -and $role -match 'Edge') { continue }

        $matched = $false
        foreach ($pattern in $include) { if ($srv.Name -like $pattern -or $srv.Fqdn -like $pattern) { $matched = $true; break } }
        if (-not $matched) { continue }

        $skip = $false
        foreach ($pattern in $exclude) { if ($pattern -and ($srv.Name -like $pattern -or $srv.Fqdn -like $pattern)) { $skip = $true; break } }
        if ($skip) { continue }

        if ($sites.Count -gt 0) {
            $siteName = ''
            if ($srv.Site) { $siteName = ($srv.Site.ToString() -split '/')[-1] }
            $inSite = $false
            foreach ($pattern in $sites) { if ($siteName -like $pattern) { $inSite = $true; break } }
            if (-not $inSite) { continue }
        }

        $isMailbox = $role -match 'Mailbox'
        [pscustomobject]@{
            Name      = [string]$srv.Name
            Fqdn      = [string]$srv.Fqdn
            Role      = $role
            Version   = [string]$srv.AdminDisplayVersion
            Site      = if ($srv.Site) { ($srv.Site.ToString() -split '/')[-1] } else { '' }
            IsMailbox = $isMailbox
            IsEdge    = ($role -match 'Edge')
            Dag       = ''
            Online    = $false
        }
    }

    $selected = @($selected)
    if ($selected.Count -eq 0) { throw 'Nessun server Exchange selezionato: verificare i filtri in Servers.Include/Exclude.' }

    # Appartenenza al DAG (serve ai check di replica)
    if (Test-HcCommand 'Get-MailboxServer') {
        foreach ($srv in $selected | Where-Object { $_.IsMailbox }) {
            try {
                $mbx = Get-MailboxServer -Identity $srv.Name -ErrorAction Stop
                if ($mbx.DatabaseAvailabilityGroup) { $srv.Dag = [string]$mbx.DatabaseAvailabilityGroup }
            }
            catch {
                Write-HcLog ('Get-MailboxServer {0} fallito: {1}' -f $srv.Name, $_.Exception.Message) -Level DEBUG
            }
        }
    }

    Write-HcLog ('Server selezionati per il check: {0} ({1})' -f $selected.Count, (($selected | Select-Object -ExpandProperty Name) -join ', '))
    return $selected
}

#endregion

#region --------------------------------------------------- RACCOLTA DATI OS

# Un solo fan-out remoto per tutti i server: WinRM parallelizza fino a
# ThrottleLimit. Raccoglie OS, volumi, servizi, memoria, reboot pending.
function Get-HcRemoteData {
    param([Parameter(Mandatory)][object[]]$Targets)

    # Per il remoting si preferisce l'FQDN: il nome corto dipende dal suffisso DNS
    # dell'host che esegue lo script e con Kerberos fallisce come
    # "Cannot find the computer ..." anche se il server e perfettamente attivo.
    # Servers.UseFqdnForRemoting = false per tornare al nome corto.
    $useFqdn = $true
    if ($null -ne $script:Config.Servers.UseFqdnForRemoting) {
        $useFqdn = [bool]$script:Config.Servers.UseFqdnForRemoting
    }
    $names = @($Targets | ForEach-Object {
        if ($useFqdn -and $_.Fqdn) { $_.Fqdn } else { $_.Name }
    })
    $minVol  = [double]$script:Config.Thresholds.MinimumVolumeSizeGB
    $extra   = @($script:Config.ExtraServices)
    $throttle = [int]$script:Config.Thresholds.ThrottleLimit
    if ($throttle -le 0) { $throttle = 24 }

    # Senza OpenTimeout, un server spento tiene occupato uno slot del fan-out per
    # il timeout WinRM di default (~45s). Con decine di server significa minuti
    # persi per sapere una cosa che si sa gia: non risponde.
    $openTimeout = [int]$script:Config.Thresholds.RemoteOpenTimeoutSeconds
    if ($openTimeout -le 0) { $openTimeout = 20 }
    $operationTimeout = [int]$script:Config.Thresholds.RemoteOperationTimeoutSeconds
    if ($operationTimeout -le 0) { $operationTimeout = 120 }
    $sessionOption = New-PSSessionOption -OpenTimeout ($openTimeout * 1000) -OperationTimeout ($operationTimeout * 1000)

    $scriptBlock = {
        param($MinVolumeGB, $ExtraServices)
        $ErrorActionPreference = 'SilentlyContinue'

        $os  = Get-CimInstance -ClassName Win32_OperatingSystem
        $cpu = Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average

        $volumes = @()
        foreach ($vol in (Get-CimInstance -ClassName Win32_Volume -Filter 'DriveType = 3')) {
            if ($null -eq $vol.Capacity -or $vol.Capacity -lt ($MinVolumeGB * 1GB)) { continue }
            $volumes += [pscustomobject]@{
                Name        = $vol.Name
                Label       = $vol.Label
                DriveLetter = $vol.DriveLetter
                CapacityGB  = [math]::Round($vol.Capacity / 1GB, 2)
                FreeGB      = [math]::Round($vol.FreeSpace / 1GB, 2)
                FreePercent = [math]::Round((100 * $vol.FreeSpace / $vol.Capacity), 2)
            }
        }

        $services = @()
        foreach ($svc in (Get-CimInstance -ClassName Win32_Service)) {
            $keep = $false
            if ($svc.Name -like 'MSExchange*') { $keep = $true }
            elseif ($ExtraServices -contains $svc.Name) { $keep = $true }
            if (-not $keep) { continue }
            $services += [pscustomobject]@{
                Name        = $svc.Name
                DisplayName = $svc.DisplayName
                State       = $svc.State
                StartMode   = $svc.StartMode
            }
        }

        $rebootPending = $false
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $rebootPending = $true }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $rebootPending = $true }
        $pfro = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
        if ($pfro) { $rebootPending = $true }

        $totalMemMB = [math]::Round($os.TotalVisibleMemorySize / 1KB, 0)
        $freeMemMB  = [math]::Round($os.FreePhysicalMemory / 1KB, 0)

        [pscustomobject]@{
            ComputerName    = $env:COMPUTERNAME
            LastBootUpTime  = $os.LastBootUpTime
            OsCaption       = $os.Caption
            TotalMemoryMB   = $totalMemMB
            FreeMemoryMB    = $freeMemMB
            FreeMemoryPct   = if ($totalMemMB -gt 0) { [math]::Round(100 * $freeMemMB / $totalMemMB, 2) } else { 0 }
            CpuLoadPercent  = if ($null -ne $cpu.Average) { [math]::Round($cpu.Average, 0) } else { -1 }
            Volumes         = $volumes
            Services        = $services
            RebootPending   = $rebootPending
            LocalTime       = Get-Date
        }
    }

    # La macchina locale si interroga direttamente, senza passare da WinRM: una
    # connessione di loopback verso se stessi richiede comunque listener,
    # autenticazione Kerberos e SPN corretti, e fallisce per motivi che non hanno
    # nulla a che vedere con la salute del server. Girando lo script su un
    # Exchange server, e proprio quel server a fallire per primo.
    $localName  = $env:COMPUTERNAME
    $localNames = @($Targets | Where-Object { $_.Name -ieq $localName } |
        ForEach-Object { if ($useFqdn -and $_.Fqdn) { $_.Fqdn } else { $_.Name } })
    $remoteNames = @($names | Where-Object { $localNames -notcontains $_ })

    $results = @()
    $errors  = $null

    if ($localNames.Count -gt 0) {
        Write-HcLog ('Raccolta dati in locale su {0} (senza WinRM).' -f $localName)
        try {
            $results += & $scriptBlock $minVol $extra
        }
        catch {
            Write-HcLog ('Raccolta locale su {0} fallita: {1}' -f $localName, $_.Exception.Message) -Level WARN
        }
    }

    if ($remoteNames.Count -gt 0) {
        Write-HcLog ('Raccolta dati OS su {0} server remoti (throttle {1}, timeout connessione {2}s)...' -f $remoteNames.Count, $throttle, $openTimeout)
        $results += @(Invoke-Command -ComputerName $remoteNames -ScriptBlock $scriptBlock `
            -ArgumentList $minVol, $extra -ThrottleLimit $throttle -SessionOption $sessionOption `
            -ErrorAction SilentlyContinue -ErrorVariable errors)
    }

    # Indicizzato sia sul nome corto sia su quello restituito dalla sessione: il
    # chiamante cerca per Name, ma qui si e connessi con l'FQDN.
    $map = @{}
    foreach ($res in $results) {
        $keys = @()
        if ($res.PSComputerName) {
            $keys += [string]$res.PSComputerName
            $keys += ([string]$res.PSComputerName -split '\.')[0]
        }
        if ($res.ComputerName) { $keys += [string]$res.ComputerName }
        foreach ($key in ($keys | Select-Object -Unique)) {
            if ($key) { $map[$key.ToUpperInvariant()] = $res }
        }
    }

    foreach ($err in @($errors)) {
        $failed = [string]$err.TargetObject
        if (-not $failed) { $failed = 'server non identificato' }
        Write-HcLog ('Errore remoto su {0}: {1}' -f $failed, $err.Exception.Message) -Level WARN
    }

    return $map
}

#endregion

#region --------------------------------------------------------------- CHECK OS

function Invoke-HcOsCheck {
    param([object]$Target, [object]$Data)

    $t = $script:Config.Thresholds

    if ($null -eq $Data) {
        Add-Finding -Category 'Connectivity' -Server $Target.Name -Item 'WinRM' -Severity 'Critical' `
            -Message 'Server non raggiungibile via WinRM/CIM (o accesso negato): nessun dato OS raccolto.'
        return
    }

    Add-Finding -Category 'Connectivity' -Server $Target.Name -Item 'WinRM' -Severity 'OK' -Message 'Server raggiungibile.'

    # --- Uptime
    if ($Data.LastBootUpTime) {
        $uptime = (Get-Date) - [datetime]$Data.LastBootUpTime
        if ($uptime.TotalMinutes -lt [double]$t.UptimeMinutesMinimum) {
            Add-Finding -Category 'Uptime' -Server $Target.Name -Item 'LastBoot' -Severity 'Warning' `
                -Message ('Server riavviato di recente ({0:N0} minuti fa, boot {1:yyyy-MM-dd HH:mm}).' -f $uptime.TotalMinutes, [datetime]$Data.LastBootUpTime) `
                -Value ('{0:N0} min' -f $uptime.TotalMinutes)
        }
        else {
            Add-Finding -Category 'Uptime' -Server $Target.Name -Item 'LastBoot' -Severity 'OK' `
                -Message ('Uptime {0:N1} giorni.' -f $uptime.TotalDays) -Value ('{0:N1} d' -f $uptime.TotalDays)
        }
    }

    # --- Memoria
    $freePct = [double]$Data.FreeMemoryPct
    $sev = 'OK'
    if ($freePct -lt [double]$t.MemoryFreePercentCritical) { $sev = 'Critical' }
    elseif ($freePct -lt [double]$t.MemoryFreePercentWarning) { $sev = 'Warning' }
    Add-Finding -Category 'Memory' -Server $Target.Name -Item 'PhysicalMemory' -Severity $sev `
        -Message ('Memoria libera {0:N1}% ({1:N0} MB su {2:N0} MB).' -f $freePct, $Data.FreeMemoryMB, $Data.TotalMemoryMB) `
        -Value ('{0:N1}%' -f $freePct)

    # --- CPU (campione istantaneo, solo warning)
    if ([int]$Data.CpuLoadPercent -ge 0) {
        $cpuSev = 'OK'
        if ([double]$Data.CpuLoadPercent -ge [double]$t.CpuLoadPercentWarning) { $cpuSev = 'Warning' }
        Add-Finding -Category 'Cpu' -Server $Target.Name -Item 'LoadPercent' -Severity $cpuSev `
            -Message ('Carico CPU {0}%.' -f $Data.CpuLoadPercent) -Value ('{0}%' -f $Data.CpuLoadPercent)
    }

    # --- Reboot pending
    if ($Data.RebootPending) {
        Add-Finding -Category 'Os' -Server $Target.Name -Item 'RebootPending' -Severity 'Info' `
            -Message 'Riavvio in sospeso (patch/servicing in attesa di reboot).'
    }
}

#endregion

#region ------------------------------------------------------------- CHECK DISK

function Get-HcVolumeThreshold {
    param([string]$ServerName, [object]$Volume)

    $t = $script:Config.Thresholds
    $result = @{
        FreePercentWarning  = [double]$t.DiskFreePercentWarning
        FreePercentCritical = [double]$t.DiskFreePercentCritical
        FreeGBWarning       = [double]$t.DiskFreeGBWarning
        FreeGBCritical      = [double]$t.DiskFreeGBCritical
        Mode                = [string]$t.DiskMode
    }

    foreach ($ovr in @($script:Config.VolumeOverrides)) {
        $serverPattern = if ($ovr.ServerPattern) { $ovr.ServerPattern } else { '*' }
        $volumePattern = if ($ovr.VolumePattern) { $ovr.VolumePattern } else { '*' }
        $volumeId = '{0} {1} {2}' -f $Volume.Name, $Volume.DriveLetter, $Volume.Label
        if ($ServerName -like $serverPattern -and $volumeId -like $volumePattern) {
            if ($null -ne $ovr.FreePercentWarning)  { $result.FreePercentWarning  = [double]$ovr.FreePercentWarning }
            if ($null -ne $ovr.FreePercentCritical) { $result.FreePercentCritical = [double]$ovr.FreePercentCritical }
            if ($null -ne $ovr.FreeGBWarning)       { $result.FreeGBWarning       = [double]$ovr.FreeGBWarning }
            if ($null -ne $ovr.FreeGBCritical)      { $result.FreeGBCritical      = [double]$ovr.FreeGBCritical }
            if ($ovr.Mode)                          { $result.Mode                = [string]$ovr.Mode }
            break
        }
    }
    return $result
}

function Invoke-HcDiskCheck {
    param([object]$Target, [object]$Data)

    if ($null -eq $Data) { return }
    $ignoreVolumes = @($script:Config.Ignore.Volumes)

    foreach ($vol in @($Data.Volumes)) {
        $volumeId = if ($vol.DriveLetter) { [string]$vol.DriveLetter } else { [string]$vol.Name }

        $skip = $false
        foreach ($pattern in $ignoreVolumes) {
            if ($pattern -and ($vol.Name -like $pattern -or $volumeId -like $pattern -or $vol.Label -like $pattern)) { $skip = $true; break }
        }
        if ($skip) { continue }

        $th = Get-HcVolumeThreshold -ServerName $Target.Name -Volume $vol

        $pctCrit = ([double]$vol.FreePercent -lt $th.FreePercentCritical)
        $gbCrit  = ([double]$vol.FreeGB      -lt $th.FreeGBCritical)
        $pctWarn = ([double]$vol.FreePercent -lt $th.FreePercentWarning)
        $gbWarn  = ([double]$vol.FreeGB      -lt $th.FreeGBWarning)

        if ($th.Mode -eq 'Or') {
            $isCritical = ($pctCrit -or $gbCrit)
            $isWarning  = ($pctWarn -or $gbWarn)
        }
        else {
            # "And": evita falsi positivi sui volumi molto grandi (10% di 4 TB = 400 GB)
            $isCritical = ($pctCrit -and $gbCrit)
            $isWarning  = ($pctWarn -and $gbWarn)
        }

        $sev = 'OK'
        if ($isCritical) { $sev = 'Critical' } elseif ($isWarning) { $sev = 'Warning' }

        $label = if ($vol.Label) { ' [' + $vol.Label + ']' } else { '' }
        Add-Finding -Category 'Disk' -Server $Target.Name -Item $volumeId -Severity $sev `
            -Message ('{0}{1} liberi {2:N1} GB su {3:N1} GB ({4:N1}%).' -f $vol.Name, $label, $vol.FreeGB, $vol.CapacityGB, $vol.FreePercent) `
            -Value ('{0:N1} GB / {1:N1}%' -f $vol.FreeGB, $vol.FreePercent)
    }
}

#endregion

#region --------------------------------------------------------- CHECK SERVIZI

function Invoke-HcServiceCheck {
    param([object]$Target, [object]$Data)

    $ignore = @($script:Config.Ignore.Services)

    # --- Servizi impostati in Auto ma non in esecuzione
    if ($null -ne $Data) {
        $stopped = @($Data.Services | Where-Object {
            $_.StartMode -eq 'Auto' -and $_.State -ne 'Running' -and ($ignore -notcontains $_.Name)
        })
        if ($stopped.Count -gt 0) {
            foreach ($svc in $stopped) {
                Add-Finding -Category 'Service' -Server $Target.Name -Item $svc.Name -Severity 'Critical' `
                    -Message ('Servizio "{0}" ({1}) impostato Automatico ma in stato {2}.' -f $svc.DisplayName, $svc.Name, $svc.State) `
                    -Value $svc.State
            }
        }
        else {
            Add-Finding -Category 'Service' -Server $Target.Name -Item 'AutoServices' -Severity 'OK' `
                -Message ('Tutti i servizi automatici monitorati sono in esecuzione ({0} verificati).' -f @($Data.Services).Count)
        }
    }

    # --- Test-ServiceHealth: verifica i servizi richiesti in base al ruolo.
    # Saltato se il server non ha risposto alla raccolta dati: usa RPC e su un
    # host irraggiungibile resterebbe appeso fino al timeout senza dire nulla di
    # nuovo (l'irraggiungibilita e gia segnalata dal check Connectivity).
    if ((Test-HcCommand 'Test-ServiceHealth') -and $null -ne $Data) {
        try {
            $health = @(Test-ServiceHealth -Server $Target.Name -ErrorAction Stop)
            foreach ($role in $health) {
                if (-not $role.RequiredServicesRunning) {
                    $missing = @($role.ServicesNotRunning) | Where-Object { $ignore -notcontains $_ }
                    if ($missing.Count -gt 0) {
                        Add-Finding -Category 'ServiceHealth' -Server $Target.Name -Item ([string]$role.Role) -Severity 'Critical' `
                            -Message ('Ruolo {0}: servizi richiesti non attivi -> {1}' -f $role.Role, ($missing -join ', ')) `
                            -Value ($missing -join ', ')
                    }
                }
                else {
                    Add-Finding -Category 'ServiceHealth' -Server $Target.Name -Item ([string]$role.Role) -Severity 'OK' `
                        -Message ('Ruolo {0}: tutti i servizi richiesti sono attivi.' -f $role.Role)
                }
            }
        }
        catch {
            Add-Finding -Category 'ServiceHealth' -Server $Target.Name -Item 'Test-ServiceHealth' -Severity 'Unknown' `
                -Message ('Test-ServiceHealth fallito: {0}' -f $_.Exception.Message)
        }
    }
}

#endregion

#region ------------------------------------------- CHECK COMPONENT / MANAGED AV

function Invoke-HcComponentCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-ServerComponentState')) { return }

    # Alcuni componenti sono Inactive per progetto su on-premises: esistono nel
    # codice ma li usa solo il datacenter Microsoft (ForwardSyncDaemon,
    # ProvisioningRps). Segnalarli significherebbe generare lo stesso falso
    # positivo su ogni server, per sempre. Lista estendibile da configurazione.
    $ignoreComponents = @($script:Config.Ignore.ServerComponents)

    $states = @(Get-ServerComponentState -Identity $Target.Name -ErrorAction Stop)
    $inactive = @($states | Where-Object {
        if ($_.State -eq 'Active') { return $false }
        $name = [string]$_.Component
        foreach ($pattern in $ignoreComponents) {
            if ($pattern -and $name -like $pattern) { return $false }
        }
        return $true
    })

    if ($inactive.Count -eq 0) {
        Add-Finding -Category 'ComponentState' -Server $Target.Name -Item 'AllComponents' -Severity 'OK' `
            -Message ('Nessun componente inattivo da segnalare ({0} verificati).' -f $states.Count)
        return
    }

    foreach ($comp in $inactive) {
        # ServerWideOffline inattivo = server in maintenance mode: quasi sempre
        # un residuo dimenticato dopo il patching.
        $sev = if ($comp.Component -eq 'ServerWideOffline') { 'Critical' } else { 'Warning' }
        $requester = ''
        try {
            $req = @($comp.LocalStates | Sort-Object TimeStamp -Descending | Select-Object -First 1)
            if ($req) { $requester = [string]$req.Requester }
        } catch { }
        Add-Finding -Category 'ComponentState' -Server $Target.Name -Item ([string]$comp.Component) -Severity $sev `
            -Message ('Componente "{0}" in stato {1}{2}.' -f $comp.Component, $comp.State, $(if ($requester) { " (requester: $requester)" } else { '' })) `
            -Value ([string]$comp.State)
    }
}

# Get-HealthReport si ferma al livello dell'health set: dice CHE cosa non va, non
# PERCHE. Il dettaglio sta un livello sotto, nei monitor, e si ottiene con
# Get-ServerHealth. Viene interrogato solo per gli health set gia risultati non
# sani, quindi il costo e proporzionale ai problemi, non al numero di health set.
function Get-HcHealthSetDetail {
    param(
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$HealthSetName
    )

    if (-not $script:Config.HealthReport.IncludeFailingMonitors) { return '' }
    if (-not (Test-HcCommand 'Get-ServerHealth')) { return '' }
    if ($HealthSetName -eq 'HealthSet sconosciuto') { return '' }

    $maximum = [int]$script:Config.HealthReport.MaxMonitorsPerHealthSet
    if ($maximum -le 0) { $maximum = 5 }

    try {
        $monitors = @(Get-ServerHealth -Identity $ServerName -HealthSet $HealthSetName -ErrorAction Stop |
            Where-Object { $_.AlertValue -and $_.AlertValue -ne 'Healthy' -and $_.AlertValue -ne 'Disabled' })
    }
    catch {
        Write-HcLog ('Dettaglio monitor di {0} su {1} non disponibile: {2}' -f $HealthSetName, $ServerName, $_.Exception.Message) -Level DEBUG
        return ''
    }

    if ($monitors.Count -eq 0) { return '' }

    $described = foreach ($monitor in ($monitors | Select-Object -First $maximum)) {
        $name = [string](Get-HcFirstValue -InputObject $monitor -PropertyNames @('Name', 'MonitorIdentity'))
        $target = [string]$monitor.TargetResource
        if ($target) { '{0} [{1}] = {2}' -f $name, $target, $monitor.AlertValue }
        else { '{0} = {1}' -f $name, $monitor.AlertValue }
    }

    $text = ' Monitor coinvolti: {0}' -f (($described) -join '; ')
    if ($monitors.Count -gt $maximum) {
        $text += ' (e altri {0}).' -f ($monitors.Count - $maximum)
    }
    return $text
}

function Invoke-HcHealthCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-HealthReport')) { return }

    $ignore = @($script:Config.Ignore.HealthSets)
    $report = @(Get-HealthReport -Identity $Target.Name -ErrorAction Stop)

    # Get-HealthReport espone il nome dell'health set in "Name" e l'orario in
    # "LastTransitionTime"; "HealthSetName" e "FirstAlertObservedTime" sono invece
    # di Get-ServerHealth. Leggendo solo le seconde i campi restavano vuoti e, cosa
    # peggiore, la chiave di deduplica diventava identica per tutti gli health set
    # (un solo alert al posto di uno per health set) e Ignore.HealthSets non
    # matchava mai. Si leggono entrambe le forme.
    $schemaLogged = $false
    $entries = foreach ($hs in $report) {
        $name = [string](Get-HcFirstValue -InputObject $hs `
            -PropertyNames @('HealthSetName', 'Name', 'HealthSet', 'HealthSetIdentity', 'Identity'))

        # Se nessuno dei nomi noti restituisce un valore, invece di mostrare un
        # segnaposto inutile si logga lo schema reale dell'oggetto: cosi si vede
        # subito quale proprieta va letta su questa versione di Exchange.
        if (-not $name) {
            if (-not $schemaLogged) {
                $available = ($hs.PSObject.Properties |
                    Where-Object { $_.Value -ne $null -and [string]$_.Value -ne '' } |
                    ForEach-Object { '{0}={1}' -f $_.Name, $_.Value }) -join ' | '
                Write-HcLog ('Nome health set non riconosciuto su {0}. Proprieta disponibili: {1}' -f $Target.Name, $available) -Level WARN
                $schemaLogged = $true
            }
            $name = 'HealthSet sconosciuto'
        }
        [pscustomobject]@{
            Name       = $name
            AlertValue = [string]$hs.AlertValue
            Since      = Get-HcFirstValue -InputObject $hs -PropertyNames @('LastTransitionTime', 'FirstAlertObservedTime')
        }
    }

    # L'esclusione accetta wildcard: "MSExchange*" oltre al nome esatto.
    $monitored = @($entries | Where-Object {
        $current = $_.Name
        $skip = $false
        foreach ($pattern in $ignore) {
            if ($pattern -and $current -like $pattern) { $skip = $true; break }
        }
        -not $skip
    })

    $bad      = @($monitored | Where-Object { $_.AlertValue -eq 'Unhealthy' })
    $degraded = @($monitored | Where-Object { $_.AlertValue -eq 'Degraded' })

    foreach ($hs in $bad) {
        $since = ''
        if ($hs.Since) { $since = ' (dal {0:yyyy-MM-dd HH:mm})' -f [datetime]$hs.Since }
        $detail = Get-HcHealthSetDetail -ServerName $Target.Name -HealthSetName $hs.Name
        Add-Finding -Category 'ManagedAvailability' -Server $Target.Name -Item $hs.Name -Severity 'Critical' `
            -Message ('Health set "{0}" Unhealthy{1}.{2}' -f $hs.Name, $since, $detail) -Value 'Unhealthy'
    }
    foreach ($hs in $degraded) {
        $since = ''
        if ($hs.Since) { $since = ' (dal {0:yyyy-MM-dd HH:mm})' -f [datetime]$hs.Since }
        $detail = Get-HcHealthSetDetail -ServerName $Target.Name -HealthSetName $hs.Name
        Add-Finding -Category 'ManagedAvailability' -Server $Target.Name -Item $hs.Name -Severity 'Warning' `
            -Message ('Health set "{0}" Degraded{1}.{2}' -f $hs.Name, $since, $detail) -Value 'Degraded'
    }
    if ($bad.Count -eq 0 -and $degraded.Count -eq 0) {
        Add-Finding -Category 'ManagedAvailability' -Server $Target.Name -Item 'AllHealthSets' -Severity 'OK' `
            -Message ('Tutti gli health set monitorati sono Healthy ({0} valutati).' -f $report.Count)
    }
}

#endregion

#region ------------------------------------------------------------- CHECK DAG

# Il cluster di un DAG si raggiunge per nome solo se il DAG ha un Administrative
# Access Point. Senza (default dei DAG moderni) il nome non e risolvibile e serve
# passare da un nodo membro: si prova prima il nome del DAG, poi i membri.
function Get-HcClusterNode {
    param(
        [Parameter(Mandatory)][string]$ClusterName,
        [string[]]$FallbackHosts
    )

    $attempts = @($ClusterName) + @($FallbackHosts)
    $failures = @()

    foreach ($endpoint in ($attempts | Where-Object { $_ } | Select-Object -Unique)) {
        try {
            $nodes = @(Get-ClusterNode -Cluster $endpoint -ErrorAction Stop)
            if ($nodes.Count -gt 0) {
                return [pscustomobject]@{ Nodes = $nodes; Endpoint = $endpoint; Error = $null }
            }
        }
        catch {
            $failures += ('{0}: {1}' -f $endpoint, $_.Exception.Message)
        }
    }

    return [pscustomobject]@{ Nodes = @(); Endpoint = $null; Error = ($failures -join ' | ') }
}

function Invoke-HcDagCheck {
    param([object[]]$Targets)

    if (-not (Test-HcCommand 'Get-DatabaseAvailabilityGroup')) { return }

    $dagNames = @($Targets | Where-Object { $_.Dag } | Select-Object -ExpandProperty Dag -Unique)
    if ($dagNames.Count -eq 0) {
        Write-HcLog 'Nessun DAG rilevato sui server selezionati.'
        return
    }

    foreach ($dagName in $dagNames) {
        Invoke-HcCheck -Category 'Dag' -ServerName $dagName -Body {
            $dag = Get-DatabaseAvailabilityGroup -Identity $dagName -Status -ErrorAction Stop

            $stopped = @($dag.StoppedMailboxServers)
            if ($stopped.Count -gt 0) {
                Add-Finding -Category 'Dag' -Server $dagName -Item 'StoppedMembers' -Severity 'Critical' `
                    -Message ('Membri DAG fermi: {0}' -f ($stopped -join ', ')) -Value ($stopped -join ', ')
            }
            else {
                Add-Finding -Category 'Dag' -Server $dagName -Item 'StoppedMembers' -Severity 'OK' `
                    -Message ('Tutti i {0} membri del DAG risultano avviati.' -f @($dag.StartedMailboxServers).Count)
            }

            $witnessState = [string]$dag.WitnessShareInUse
            if ($witnessState -eq 'InvalidConfiguration') {
                Add-Finding -Category 'Dag' -Server $dagName -Item 'Witness' -Severity 'Critical' `
                    -Message ('Witness in configurazione non valida (witness server: {0}).' -f $dag.WitnessServer) -Value $witnessState
            }
            elseif ($witnessState -eq 'Alternate') {
                Add-Finding -Category 'Dag' -Server $dagName -Item 'Witness' -Severity 'Warning' `
                    -Message ('Il DAG sta usando l''alternate witness ({0}).' -f $dag.AlternateWitnessServer) -Value $witnessState
            }
            else {
                Add-Finding -Category 'Dag' -Server $dagName -Item 'Witness' -Severity 'OK' `
                    -Message ('Witness OK ({0}, {1}).' -f $dag.WitnessServer, $witnessState) -Value $witnessState
            }

            # Nodi del cluster sottostante. Un DAG creato senza Administrative
            # Access Point (il default dei DAG moderni) non ha un IP ne un nome
            # risolvibile: "error opening cluster DAG1" e quello, non un guasto.
            # In quel caso si interroga il cluster passando da un nodo membro.
            if (Test-HcCommand 'Get-ClusterNode') {
                $members = @($Targets | Where-Object { $_.Dag -eq $dagName } | Select-Object -ExpandProperty Name)
                if ($members.Count -eq 0) {
                    $members = @($dag.StartedMailboxServers | ForEach-Object { ([string]$_ -split '\.')[0] })
                }

                $cluster = Get-HcClusterNode -ClusterName $dagName -FallbackHosts (@($members) | Select-Object -First 2)

                if ($cluster.Nodes.Count -gt 0) {
                    if ($cluster.Endpoint -ne $dagName) {
                        Write-HcLog ('Cluster {0} raggiunto tramite il nodo {1} (nome del cluster non risolvibile: DAG senza access point).' -f $dagName, $cluster.Endpoint)
                    }
                    foreach ($node in $cluster.Nodes) {
                        if ([string]$node.State -ne 'Up') {
                            Add-Finding -Category 'Cluster' -Server $dagName -Item ([string]$node.Name) -Severity 'Critical' `
                                -Message ('Nodo cluster {0} in stato {1}.' -f $node.Name, $node.State) -Value ([string]$node.State)
                        }
                    }
                    if (@($cluster.Nodes | Where-Object { [string]$_.State -ne 'Up' }).Count -eq 0) {
                        Add-Finding -Category 'Cluster' -Server $dagName -Item 'Nodes' -Severity 'OK' `
                            -Message ('Tutti i {0} nodi cluster sono Up.' -f $cluster.Nodes.Count)
                    }
                }
                else {
                    # Nessun allarme: lo stato dei membri del DAG e gia verificato
                    # dai cmdlet Exchange, questo controllo e solo una conferma.
                    Add-Finding -Category 'Cluster' -Server $dagName -Item 'Nodes' -Severity 'Info' `
                        -Message ('Stato dei nodi cluster non interrogabile per {0} (verificare RSAT Failover Clustering e i permessi sul cluster). Lo stato dei membri del DAG resta verificato via Exchange.' -f $dagName)
                    Write-HcLog ('Get-ClusterNode non disponibile per {0}: {1}' -f $dagName, $cluster.Error) -Level DEBUG
                }
            }
        }
    }
}

function Invoke-HcCopyStatusCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-MailboxDatabaseCopyStatus')) { return }

    $t = $script:Config.Thresholds
    $ignoreDb = @($script:Config.Ignore.Databases)

    # Con -Server, Exchange cerca l'identity "*\<server>". Se su quel server non
    # esiste alcuna copia di database risponde "could not be found" invece di un
    # risultato vuoto: e un'assenza di dati, non un guasto. Stesso messaggio anche
    # quando l'oggetto esiste ma e fuori dallo scope RBAC dell'account.
    try {
        $copies = @(Get-MailboxDatabaseCopyStatus -Server $Target.Name -ErrorAction Stop)
    }
    catch {
        if ($_.Exception.Message -match 'could not be found|non e stato trovato|couldn''t be found') {
            Add-Finding -Category 'DatabaseCopy' -Server $Target.Name -Item 'Copies' -Severity 'Info' `
                -Message ('Nessuna copia di database trovata su {0}: il server non ne ospita, oppure non sono visibili con lo scope RBAC dell''account corrente.' -f $Target.Name)
            return
        }
        throw
    }

    if ($copies.Count -eq 0) {
        Add-Finding -Category 'DatabaseCopy' -Server $Target.Name -Item 'Copies' -Severity 'Info' `
            -Message 'Nessuna copia di database presente su questo server.'
        return
    }

    foreach ($copy in $copies) {
        $dbName = ([string]$copy.Name -split '\\')[0]

        $skip = $false
        foreach ($pattern in $ignoreDb) { if ($pattern -and $dbName -like $pattern) { $skip = $true; break } }
        if ($skip) { continue }

        $status = [string]$copy.Status

        switch -Regex ($status) {
            '^(Mounted|Healthy)$'                    { $sev = 'OK' }
            '^(Seeding|SeedingSource|Initializing|Resynchronizing|SinglePageRestore|DisconnectedAndHealthy|Mounting|Dismounting)$' { $sev = 'Warning' }
            '^(Failed|FailedAndSuspended|Suspended|ServiceDown|Dismounted|DisconnectedAndResynchronizing|Misconfigured|NotKnown)$' { $sev = 'Critical' }
            default                                  { $sev = 'Warning' }
        }

        Add-Finding -Category 'DatabaseCopy' -Server $Target.Name -Item ([string]$copy.Name) -Severity $sev `
            -Message ('Copia {0}: stato {1}{2}.' -f $copy.Name, $status, $(if ($copy.ErrorMessage) { ' - ' + $copy.ErrorMessage } else { '' })) `
            -Value $status

        # --- Copy queue / Replay queue
        $copyQ   = 0
        $replayQ = 0
        if ($null -ne $copy.CopyQueueLength)   { $copyQ   = [int64]$copy.CopyQueueLength }
        if ($null -ne $copy.ReplayQueueLength) { $replayQ = [int64]$copy.ReplayQueueLength }

        if ($status -ne 'Mounted') {
            $qSev = 'OK'
            if ($copyQ -ge [int64]$t.CopyQueueCritical) { $qSev = 'Critical' }
            elseif ($copyQ -ge [int64]$t.CopyQueueWarning) { $qSev = 'Warning' }
            Add-Finding -Category 'CopyQueue' -Server $Target.Name -Item ([string]$copy.Name) -Severity $qSev `
                -Message ('Copy queue {0} log per {1}.' -f $copyQ, $copy.Name) -Value $copyQ

            $rSev = 'OK'
            if ($replayQ -ge [int64]$t.ReplayQueueCritical) { $rSev = 'Critical' }
            elseif ($replayQ -ge [int64]$t.ReplayQueueWarning) { $rSev = 'Warning' }
            Add-Finding -Category 'ReplayQueue' -Server $Target.Name -Item ([string]$copy.Name) -Severity $rSev `
                -Message ('Replay queue {0} log per {1}.' -f $replayQ, $copy.Name) -Value $replayQ
        }

        # --- Content index
        $ci = [string]$copy.ContentIndexState
        if ($ci -and $ci -ne 'Healthy') {
            $ciSev = if ($ci -match '^(Crawling|Seeding|Suspended)$') { 'Warning' } else { 'Critical' }
            Add-Finding -Category 'ContentIndex' -Server $Target.Name -Item ([string]$copy.Name) -Severity $ciSev `
                -Message ('Content index di {0} in stato {1}{2}.' -f $copy.Name, $ci, $(if ($copy.ContentIndexErrorMessage) { ' - ' + $copy.ContentIndexErrorMessage } else { '' })) `
                -Value $ci
        }
        elseif ($ci) {
            Add-Finding -Category 'ContentIndex' -Server $Target.Name -Item ([string]$copy.Name) -Severity 'OK' `
                -Message ('Content index di {0} Healthy.' -f $copy.Name) -Value $ci
        }
    }
}

function Invoke-HcReplicationCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Test-ReplicationHealth')) { return }
    if (-not $Target.Dag) { return }

    $results = @(Test-ReplicationHealth -Identity $Target.Name -ErrorAction Stop)
    $failed  = @($results | Where-Object { [string]$_.Result -eq 'Failed' })
    $warned  = @($results | Where-Object { [string]$_.Result -eq 'Warning' })

    foreach ($res in $failed) {
        Add-Finding -Category 'Replication' -Server $Target.Name -Item ([string]$res.Check) -Severity 'Critical' `
            -Message ('Test-ReplicationHealth "{0}" FAILED: {1}' -f $res.Check, $res.Error) -Value 'Failed'
    }
    foreach ($res in $warned) {
        Add-Finding -Category 'Replication' -Server $Target.Name -Item ([string]$res.Check) -Severity 'Warning' `
            -Message ('Test-ReplicationHealth "{0}" WARNING: {1}' -f $res.Check, $res.Error) -Value 'Warning'
    }
    if ($failed.Count -eq 0 -and $warned.Count -eq 0) {
        Add-Finding -Category 'Replication' -Server $Target.Name -Item 'AllChecks' -Severity 'OK' `
            -Message ('Test-ReplicationHealth: {0} controlli superati.' -f $results.Count)
    }
}

#endregion

#region -------------------------------------------------------- CHECK DATABASE

function Invoke-HcDatabaseCheck {
    param([object[]]$Targets)

    if (-not (Test-HcCommand 'Get-MailboxDatabase')) { return }

    $t = $script:Config.Thresholds
    $ignoreDb   = @($script:Config.Ignore.Databases)
    $serverList = @($Targets | Select-Object -ExpandProperty Name)

    # Interrogazione per server, non a livello di organizzazione: "Get-MailboxDatabase
    # -Status" senza -Server enumera TUTTI i database della foresta e, per ricavare
    # lo stato di mount, contatta ogni server proprietario. Basta un server lento o
    # irraggiungibile fuori perimetro per far sembrare lo script bloccato, perche i
    # cmdlet Exchange non hanno timeout.
    $collectionRan = @($Targets | Where-Object { $_.Online }).Count -gt 0
    $databases = @()
    $seen = @{}

    foreach ($target in $Targets) {
        if ($collectionRan -and -not $target.Online) {
            Write-HcLog ('{0} non raggiungibile: salto l''elenco database.' -f $target.Name) -Level DEBUG
            continue
        }

        Write-HcLog ('Recupero database di {0}...' -f $target.Name) -Level DEBUG
        try {
            $serverDatabases = @(Get-MailboxDatabase -Server $target.Name -Status -ErrorAction Stop)
        }
        catch {
            Add-Finding -Category 'Database' -Server $target.Name -Item 'Get-MailboxDatabase' -Severity 'Unknown' `
                -Message ('Elenco database non recuperabile: {0}' -f $_.Exception.Message)
            continue
        }

        # In un DAG lo stesso database torna una volta per copia: si tiene la prima.
        foreach ($db in $serverDatabases) {
            $id = [string]$db.Guid
            if (-not $id) { $id = [string]$db.Name }
            if ($seen.ContainsKey($id)) { continue }
            $seen[$id] = $true
            $databases += $db
        }
    }

    Write-HcLog ('Database rilevati: {0}' -f $databases.Count)

    foreach ($db in $databases) {
        $dbName = [string]$db.Name

        $skip = $false
        foreach ($pattern in $ignoreDb) { if ($pattern -and $dbName -like $pattern) { $skip = $true; break } }
        if ($skip) { continue }

        $activeServer = [string]$db.Server
        # Ignora i DB montati su server esclusi dal perimetro
        if ($serverList -notcontains ($activeServer -split '\.')[0]) { continue }

        # --- Stato di mount
        if ($db.Mounted -eq $false) {
            Add-Finding -Category 'Database' -Server $activeServer -Item $dbName -Severity 'Critical' `
                -Message ('Database "{0}" NON montato.' -f $dbName) -Value 'Dismounted'
        }
        elseif ($null -eq $db.Mounted) {
            Add-Finding -Category 'Database' -Server $activeServer -Item $dbName -Severity 'Unknown' `
                -Message ('Stato di mount del database "{0}" non determinabile.' -f $dbName)
        }
        else {
            Add-Finding -Category 'Database' -Server $activeServer -Item $dbName -Severity 'OK' `
                -Message ('Database "{0}" montato su {1}.' -f $dbName, $activeServer) -Value 'Mounted'
        }

        # --- Backup
        $lastFull = $db.LastFullBackup
        $lastIncr = $db.LastIncrementalBackup
        $lastBackup = $null
        if ($lastFull) { $lastBackup = [datetime]$lastFull }
        if ($lastIncr -and (-not $lastBackup -or [datetime]$lastIncr -gt $lastBackup)) { $lastBackup = [datetime]$lastIncr }

        if (-not $lastBackup) {
            Add-Finding -Category 'Backup' -Server $activeServer -Item $dbName -Severity 'Critical' `
                -Message ('Nessun backup registrato per il database "{0}".' -f $dbName) -Value 'mai'
        }
        else {
            $ageHours = ((Get-Date) - $lastBackup).TotalHours
            $sev = 'OK'
            if ($ageHours -ge [double]$t.BackupAgeHoursCritical) { $sev = 'Critical' }
            elseif ($ageHours -ge [double]$t.BackupAgeHoursWarning) { $sev = 'Warning' }
            Add-Finding -Category 'Backup' -Server $activeServer -Item $dbName -Severity $sev `
                -Message ('Ultimo backup di "{0}": {1:yyyy-MM-dd HH:mm} ({2:N1} ore fa).' -f $dbName, $lastBackup, $ageHours) `
                -Value ('{0:N1} h' -f $ageHours)
        }

        # --- Copia attiva sulla preferenza 1 (bilanciamento del DAG)
        try {
            $prefs = @($db.ActivationPreference)
            if ($prefs.Count -gt 1) {
                $preferred = $null
                foreach ($pref in $prefs) {
                    if ([int]$pref.Value -eq 1) { $preferred = ([string]$pref.Key -split '\.')[0]; break }
                }
                if ($preferred -and $preferred -ne ($activeServer -split '\.')[0]) {
                    Add-Finding -Category 'Activation' -Server $activeServer -Item $dbName -Severity 'Info' `
                        -Message ('Database "{0}" attivo su {1} anziche sulla copia preferita {2}.' -f $dbName, $activeServer, $preferred) `
                        -Value $preferred
                }
            }
        }
        catch {
            Write-HcLog ('Verifica activation preference di {0} fallita: {1}' -f $dbName, $_.Exception.Message) -Level DEBUG
        }
    }
}

#endregion

#region ----------------------------------------------------------- CHECK QUEUE

# Risoluzione PTR con cache (per IP, per l'intera esecuzione) e timeout: senza un
# timeout esplicito una risoluzione DNS che non risponde puo restare appesa a
# lungo, e per un giro con decine di server con code verso decine di smart host
# significherebbe rallentare l'intero check per un singolo DNS lento o assente.
function Resolve-HcReverseDns {
    param([Parameter(Mandatory)][string]$IpAddress)

    if (-not $script:ReverseDnsCache) { $script:ReverseDnsCache = @{} }
    if ($script:ReverseDnsCache.ContainsKey($IpAddress)) { return $script:ReverseDnsCache[$IpAddress] }

    $timeoutMs = [int]$script:Config.Queues.ReverseDnsTimeoutMs
    if ($timeoutMs -le 0) { $timeoutMs = 1000 }

    $resolved = $null
    try {
        $async = [System.Net.Dns]::BeginGetHostEntry($IpAddress, $null, $null)
        if ($async.AsyncWaitHandle.WaitOne($timeoutMs)) {
            $resolved = [System.Net.Dns]::EndGetHostEntry($async).HostName
        }
        else {
            Write-HcLog ('Risoluzione PTR di {0} oltre il timeout di {1} ms: resta il solo IP.' -f $IpAddress, $timeoutMs) -Level DEBUG
        }
    }
    catch {
        Write-HcLog ('Risoluzione PTR di {0} fallita: {1}' -f $IpAddress, $_.Exception.Message) -Level DEBUG
    }

    $script:ReverseDnsCache[$IpAddress] = $resolved
    return $resolved
}

# Fallback quando manca un record PTR in DNS. In molte reti aziendali, in
# particolare quelle piu datate, le zone di reverse lookup non sono tenute
# aggiornate quanto quelle dirette: "ping -a" su Windows spesso risolve
# comunque il nome perche il resolver di sistema, oltre al DNS, prova anche il
# NetBIOS name query (porta UDP 137) quando l'host e sullo stesso segmento di
# rete. [System.Net.Dns] invece interroga SOLO il DNS: se manca il PTR, fallisce
# anche quando "ping -a" mostra un nome. nbtstat -A replica quella stessa
# interrogazione NetBIOS. Il processo esterno viene vincolato a un timeout
# esplicito (Process.WaitForExit), coerente con l'approccio usato per il DNS:
# nbtstat non ha un timeout nativo e su un host fuori subnet puo attendere a
# lungo prima di arrendersi da solo.
function Resolve-HcNetBiosName {
    param([Parameter(Mandatory)][string]$IpAddress)

    $timeoutMs = [int]$script:Config.Queues.NetBiosTimeoutMs
    if ($timeoutMs -le 0) { $timeoutMs = 1500 }

    $process = $null
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = 'nbtstat.exe'
        $psi.Arguments              = '-A {0}' -f $IpAddress
        $psi.RedirectStandardOutput = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true

        $process = [System.Diagnostics.Process]::Start($psi)
        if (-not $process.WaitForExit($timeoutMs)) {
            try { $process.Kill() } catch { }
            Write-HcLog ('nbtstat su {0} oltre il timeout di {1} ms.' -f $IpAddress, $timeoutMs) -Level DEBUG
            return $null
        }

        $output = $process.StandardOutput.ReadToEnd()
        # Riga tipica: "SERVERNAME     <00>  UNIQUE      Registered"
        $match = [regex]::Match($output, '^\s*([A-Za-z0-9_\-]+)\s*<00>\s*UNIQUE', 'Multiline')
        if ($match.Success) { return $match.Groups[1].Value.Trim() }
        return $null
    }
    catch {
        Write-HcLog ('nbtstat su {0} non eseguibile: {1}' -f $IpAddress, $_.Exception.Message) -Level DEBUG
        return $null
    }
    finally {
        if ($process) { $process.Dispose() }
    }
}

# NextHopDomain a volte e un IP invece di un nome: tipico di uno smart host o di
# un send connector configurato con l'indirizzo anziche l'FQDN. Se e un IP,
# prima si tenta il DNS (PTR) e poi, se non risolve, il NetBIOS: cosi il
# risultato si avvicina a quanto mostrerebbe "ping -a" sulla stessa rete. Il
# nome viene affiancato senza mai nascondere l'IP. Se non e un IP (dominio SMTP
# o nome di un database) resta invariato.
function Get-HcNextHopLabel {
    param([string]$NextHopDomain)

    if ([string]::IsNullOrWhiteSpace($NextHopDomain)) { return $NextHopDomain }
    if (-not $script:Config.Queues.ResolveNextHopHostnames) { return $NextHopDomain }

    $candidate = $NextHopDomain
    # Alcuni next hop sono "IP:porta": si isola la sola parte IP per il parsing.
    if ($candidate -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') { $candidate = $Matches[1] }

    $parsedIp = $null
    if (-not [System.Net.IPAddress]::TryParse($candidate, [ref]$parsedIp)) { return $NextHopDomain }

    $hostName = Resolve-HcReverseDns -IpAddress $candidate
    if (-not $hostName -and $script:Config.Queues.TryNetBiosFallback) {
        $hostName = Resolve-HcNetBiosName -IpAddress $candidate
    }
    if ($hostName) { return '{0} [{1}]' -f $NextHopDomain, $hostName }
    return $NextHopDomain
}

# Get-Queue espone gia il GUID del send connector usato per quell'hop
# (NextHopConnector): non serve indovinarlo dal dominio o dagli AddressSpaces.
# Per una coda di consegna interna (via DAG/database) quel GUID non corrisponde
# a nessun Send Connector: Get-SendConnector fallisce, e in quel caso non si
# mostra nulla in piu, invece di inventare un nome. Cache per GUID: lo stesso
# connector serve piu code sullo stesso server.
function Get-HcSendConnectorName {
    param([string]$ConnectorId)

    if ([string]::IsNullOrWhiteSpace($ConnectorId)) { return $null }
    if (-not $script:Config.Queues.ResolveSendConnectorName) { return $null }
    if (-not (Test-HcCommand 'Get-SendConnector')) { return $null }

    if (-not $script:SendConnectorCache) { $script:SendConnectorCache = @{} }
    if ($script:SendConnectorCache.ContainsKey($ConnectorId)) { return $script:SendConnectorCache[$ConnectorId] }

    $name = $null
    try {
        $connector = Get-SendConnector -Identity $ConnectorId -ErrorAction Stop
        if ($connector) { $name = [string]$connector.Name }
    }
    catch {
        Write-HcLog ('NextHopConnector {0} non corrisponde a un send connector (probabile consegna interna): {1}' -f $ConnectorId, $_.Exception.Message) -Level DEBUG
    }

    $script:SendConnectorCache[$ConnectorId] = $name
    return $name
}

# Etichetta completa di una coda: destinazione (con hostname risolto se e un IP)
# piu, se disponibile, il send connector che la sta instradando.
function Get-HcQueueDestinationLabel {
    param([Parameter(Mandatory)][object]$Queue)

    $label = Get-HcNextHopLabel -NextHopDomain ([string]$Queue.NextHopDomain)

    $connectorName = Get-HcSendConnectorName -ConnectorId ([string]$Queue.NextHopConnector)
    if ($connectorName) { $label = '{0} via connector "{1}"' -f $label, $connectorName }

    return $label
    return $NextHopDomain
}

function Invoke-HcQueueCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-Queue')) { return }

    $t = $script:Config.Thresholds
    $queues = @(Get-Queue -Server $Target.Name -ErrorAction Stop)

    # Le shadow queue trattengono messaggi per progetto: non sono un'anomalia.
    $realQueues = @($queues | Where-Object { [string]$_.DeliveryType -ne 'ShadowRedundancy' })

    $total = 0
    foreach ($q in $realQueues) { $total += [int64]$q.MessageCount }

    $sev = 'OK'
    if ($total -ge [int64]$t.QueueCritical) { $sev = 'Critical' }
    elseif ($total -ge [int64]$t.QueueWarning) { $sev = 'Warning' }
    Add-Finding -Category 'Queue' -Server $Target.Name -Item 'TotalMessages' -Severity $sev `
        -Message ('Totale messaggi in coda: {0} su {1} code attive.' -f $total, $realQueues.Count) -Value $total

    # Fotografia delle code per la vista aggregata in mail: viene raccolta sempre,
    # anche quando nulla supera soglia, perche il quadro del mail flow e utile di
    # per se, non solo quando c'e un allarme.
    $submissionQueue = @($realQueues | Where-Object { [string]$_.Identity -match 'Submission' })
    $poisonQueue     = @($queues     | Where-Object { [string]$_.Identity -match 'Poison' })
    $retryQueues     = @($realQueues | Where-Object { [string]$_.Status -eq 'Retry' })
    $shadowQueues    = @($queues     | Where-Object { [string]$_.DeliveryType -eq 'ShadowRedundancy' })

    $deliveryQueues = @($realQueues | Where-Object {
        [string]$_.Identity -notmatch 'Submission|Poison'
    } | Sort-Object { [int64]$_.MessageCount } -Descending)

    $largest = $deliveryQueues | Select-Object -First 1
    $shadowTotal = 0
    foreach ($q in $shadowQueues) { $shadowTotal += [int64]$q.MessageCount }

    $script:QueueSummary.Add([pscustomobject]@{
        Server       = $Target.Name
        Total        = $total
        QueueCount   = $realQueues.Count
        Submission   = if ($submissionQueue) { [int64]($submissionQueue | Measure-Object -Property MessageCount -Sum).Sum } else { 0 }
        Poison       = if ($poisonQueue)     { [int64]($poisonQueue     | Measure-Object -Property MessageCount -Sum).Sum } else { 0 }
        RetryCount   = $retryQueues.Count
        ShadowTotal  = $shadowTotal
        LargestName  = if ($largest) { Get-HcQueueDestinationLabel -Queue $largest } else { '' }
        LargestCount = if ($largest) { [int64]$largest.MessageCount } else { 0 }
        Severity     = $sev
    }) | Out-Null

    foreach ($q in $realQueues) {
        $identity = [string]$q.Identity
        $count    = [int64]$q.MessageCount
        $status   = [string]$q.Status
        $type     = [string]$q.DeliveryType

        if ($type -eq 'Undefined' -and $identity -match 'Submission') {
            $subSev = 'OK'
            if ($count -ge [int64]$t.SubmissionQueueCritical) { $subSev = 'Critical' }
            elseif ($count -ge [int64]$t.SubmissionQueueWarning) { $subSev = 'Warning' }
            if ($subSev -ne 'OK') {
                Add-Finding -Category 'Queue' -Server $Target.Name -Item 'Submission' -Severity $subSev `
                    -Message ('Submission queue con {0} messaggi (categorizer in difficolta).' -f $count) -Value $count
            }
            continue
        }

        if ($identity -match 'Poison') {
            if ($count -ge [int64]$t.PoisonQueueWarning) {
                Add-Finding -Category 'Queue' -Server $Target.Name -Item 'Poison' -Severity 'Warning' `
                    -Message ('Poison queue con {0} messaggi.' -f $count) -Value $count
            }
            continue
        }

        $qSev = 'OK'
        if ($count -ge [int64]$t.QueueCritical) { $qSev = 'Critical' }
        elseif ($count -ge [int64]$t.QueueWarning) { $qSev = 'Warning' }

        if ($status -eq 'Retry' -and $count -ge [int64]$t.RetryQueueWarning -and (Get-SeverityRank $qSev) -lt 2) {
            $qSev = 'Warning'
        }

        if ($qSev -ne 'OK') {
            # Item resta l'IP/dominio grezzo: e la chiave di deduplica degli alert e
            # non deve dipendere da una risoluzione DNS che puo cambiare da un giro
            # all'altro. Solo il messaggio, destinato a un umano, mostra l'hostname.
            $destination = Get-HcQueueDestinationLabel -Queue $q
            Add-Finding -Category 'Queue' -Server $Target.Name -Item ([string]$q.NextHopDomain) -Severity $qSev `
                -Message ('Coda "{0}" verso {1}: {2} messaggi, stato {3}{4}.' -f $identity, $destination, $count, $status, $(if ($q.LastError) { ' - ' + $q.LastError } else { '' })) `
                -Value $count
        }
    }
}

function Invoke-HcBackPressureCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-ExchangeDiagnosticInfo')) { return }

    $raw = Get-ExchangeDiagnosticInfo -Server $Target.Name -Process EdgeTransport -Component ResourceThrottling -ErrorAction Stop
    if (-not $raw) { return }

    $xml = [xml]$raw
    $meters = @($xml.SelectNodes('//ResourceMeter'))
    if ($meters.Count -eq 0) { return }

    # Attenzione ai campi: "CurrentResourceUse" e lo STATO (Low / Medium / High),
    # "Pressure" e la misura numerica della risorsa. Low e lo stato SANO: e il
    # valore normale su un server che sta benissimo. Confrontare Pressure con
    # 'Normal' faceva scattare l'allarme su ogni risorsa a ogni esecuzione.
    $pressured = @()
    foreach ($meter in $meters) {
        $resource = [string]$meter.Resource

        $level = [string]$meter.CurrentResourceUse
        if ($level -notmatch '^\s*(Low|Medium|High)\s*$') {
            $alternative = [string]$meter.Pressure
            if ($alternative -match '^\s*(Low|Medium|High)\s*$') { $level = $alternative }
        }
        $level = $level.Trim()

        if ($level -match '^(Medium|High)$') {
            $detail = '{0} = {1}' -f $resource, $level
            $measure = [string]$meter.Pressure
            if ($measure -and $measure.Trim() -ne $level) { $detail += ' (valore {0})' -f $measure.Trim() }
            $pressured += [pscustomobject]@{ Text = $detail; Level = $level }
        }
    }

    if ($pressured.Count -gt 0) {
        $sev = 'Warning'
        if (@($pressured | Where-Object { $_.Level -eq 'High' }).Count -gt 0) { $sev = 'Critical' }
        $text = ($pressured | ForEach-Object { $_.Text }) -join '; '
        Add-Finding -Category 'BackPressure' -Server $Target.Name -Item 'Transport' -Severity $sev `
            -Message ('Back pressure sul transport: {0}' -f $text) -Value $text
    }
    else {
        Add-Finding -Category 'BackPressure' -Server $Target.Name -Item 'Transport' -Severity 'OK' `
            -Message ('Nessuna back pressure ({0} risorse monitorate).' -f $meters.Count)
    }
}

#endregion

#region --------------------------------------------------- CHECK CERT E MAPI

function Invoke-HcCertificateCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Get-ExchangeCertificate')) { return }

    $t = $script:Config.Thresholds
    $certs = @(Get-ExchangeCertificate -Server $Target.Name -ErrorAction Stop)
    $relevant = @($certs | Where-Object { [string]$_.Services -match 'IIS|SMTP|IMAP|POP' })

    foreach ($cert in $relevant) {
        $days = ([datetime]$cert.NotAfter - (Get-Date)).TotalDays
        $sev = 'OK'
        if ($days -lt 0) { $sev = 'Critical' }
        elseif ($days -le [double]$t.CertExpiryDaysCritical) { $sev = 'Critical' }
        elseif ($days -le [double]$t.CertExpiryDaysWarning) { $sev = 'Warning' }

        if ($sev -ne 'OK') {
            $subject = [string]$cert.Subject
            Add-Finding -Category 'Certificate' -Server $Target.Name -Item ([string]$cert.Thumbprint) -Severity $sev `
                -Message ('Certificato "{0}" (servizi: {1}) scade il {2:yyyy-MM-dd} -> {3:N0} giorni.' -f $subject, $cert.Services, [datetime]$cert.NotAfter, $days) `
                -Value ('{0:N0} giorni' -f $days)
        }
    }

    # Non basta guardare le date: un certificato puo essere revocato, con catena
    # non attendibile o non ancora valido pur avendo una scadenza lontana.
    foreach ($cert in $relevant) {
        $status = [string]$cert.Status
        if ($status -and $status -ne 'Valid') {
            Add-Finding -Category 'Certificate' -Server $Target.Name -Item ('{0}-Status' -f $cert.Thumbprint) -Severity 'Warning' `
                -Message ('Certificato "{0}" (servizi: {1}) in stato {2}, pur essendo assegnato a servizi attivi.' -f $cert.Subject, $cert.Services, $status) `
                -Value $status
        }
    }

    # Due certificati con stesso Subject E stesso Issuer rendono ambiguo il
    # TlsCertificateName dei connector, che li identifica con "<I>Issuer<S>Subject"
    # e non per thumbprint: Exchange puo agganciare quello sbagliato, ed e una
    # causa tipica del monitor Transport.ServerCertMismatch.
    $duplicates = @($relevant |
        Group-Object { '{0}|{1}' -f $_.Subject, $_.Issuer } |
        Where-Object { $_.Count -gt 1 })

    foreach ($group in $duplicates) {
        $details = ($group.Group | ForEach-Object {
            '{0} (scade {1:yyyy-MM-dd}, stato {2})' -f $_.Thumbprint, [datetime]$_.NotAfter, $_.Status
        }) -join '; '
        $subject = [string]$group.Group[0].Subject
        Add-Finding -Category 'Certificate' -Server $Target.Name -Item ('Duplicato-{0}' -f $subject) -Severity 'Warning' `
            -Message ('{0} certificati con stesso Subject e Issuer ("{1}") assegnati a servizi: il TlsCertificateName dei connector diventa ambiguo. Thumbprint: {2}' -f $group.Count, $subject, $details) `
            -Value $group.Count
    }

    $expiring = @($relevant | Where-Object { ([datetime]$_.NotAfter - (Get-Date)).TotalDays -le [double]$t.CertExpiryDaysWarning })
    $invalid  = @($relevant | Where-Object { [string]$_.Status -and [string]$_.Status -ne 'Valid' })

    if ($expiring.Count -eq 0 -and $invalid.Count -eq 0 -and $duplicates.Count -eq 0) {
        Add-Finding -Category 'Certificate' -Server $Target.Name -Item 'AllCertificates' -Severity 'OK' `
            -Message ('Nessun certificato in scadenza entro {0} giorni, non valido o duplicato ({1} verificati).' -f $t.CertExpiryDaysWarning, $relevant.Count)
    }
}

function Invoke-HcMapiCheck {
    param([object]$Target)

    if (-not (Test-HcCommand 'Test-MapiConnectivity')) { return }
    if (-not $Target.IsMailbox) { return }

    $results = @(Test-MapiConnectivity -Server $Target.Name -ErrorAction Stop)
    foreach ($res in $results) {
        $result = [string]$res.Result
        if ($result -match 'Success') {
            Add-Finding -Category 'Mapi' -Server $Target.Name -Item ([string]$res.Database) -Severity 'OK' `
                -Message ('MAPI OK su {0} ({1} ms).' -f $res.Database, $res.Latency.TotalMilliseconds) -Value $result
        }
        else {
            Add-Finding -Category 'Mapi' -Server $Target.Name -Item ([string]$res.Database) -Severity 'Critical' `
                -Message ('Test MAPI fallito su {0}: {1}' -f $res.Database, $res.Error) -Value $result
        }
    }
}

#endregion

#region ------------------------------------------------------------ ALERT STATE

function Get-HcState {
    param([string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        try {
            $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
            if ($raw.Trim()) { return (ConvertFrom-Json $raw) }
        }
        catch {
            Write-HcLog ('State file illeggibile, riparto da zero: {0}' -f $_.Exception.Message) -Level WARN
        }
    }
    return [pscustomobject]@{ Alerts = [pscustomobject]@{}; LastHeartbeat = $null }
}

function Save-HcState {
    param([string]$Path, [object]$State)
    if (-not $Path) { return }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    try {
        ($State | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-HcLog ('Salvataggio state file fallito: {0}' -f $_.Exception.Message) -Level ERROR
    }
}

# Confronta i finding correnti con lo stato salvato e decide cosa notificare.
function Resolve-HcAlert {
    param([object]$State, [object[]]$Findings)

    $cooldown  = [double]$script:Config.Alerting.CooldownMinutes
    $minRank   = Get-SeverityRank ([string]$script:Config.Alerting.MinimumSeverityToMail)
    $notifyUnk = [bool]$script:Config.Alerting.NotifyOnUnknown
    $now       = Get-Date

    $active = @($Findings | Where-Object {
        $rank = Get-SeverityRank $_.Severity
        ($rank -ge $minRank) -and ($notifyUnk -or $_.Severity -ne 'Unknown')
    })

    $newAlerts      = New-Object System.Collections.Generic.List[object]
    $escalated      = New-Object System.Collections.Generic.List[object]
    $reminders      = New-Object System.Collections.Generic.List[object]
    $recovered      = New-Object System.Collections.Generic.List[object]
    $nextState      = [pscustomobject]@{}
    $activeKeys     = @{}

    foreach ($finding in $active) {
        $key = $finding.Key
        if ($activeKeys.ContainsKey($key)) { continue }
        $activeKeys[$key] = $true

        $previous = $null
        if ($State.Alerts.PSObject.Properties.Name -contains $key) { $previous = $State.Alerts.$key }

        if ($null -eq $previous) {
            $newAlerts.Add($finding) | Out-Null
            $entry = [pscustomobject]@{
                FirstSeen    = $now
                LastSeen     = $now
                LastNotified = $now
                Severity     = $finding.Severity
                Occurrences  = 1
                Message      = $finding.Message
            }
        }
        else {
            $prevRank = Get-SeverityRank ([string]$previous.Severity)
            $currRank = Get-SeverityRank $finding.Severity
            $lastNotified = $null
            if ($previous.LastNotified) { $lastNotified = [datetime]$previous.LastNotified }

            $shouldNotify = $false
            if ($currRank -gt $prevRank) { $escalated.Add($finding) | Out-Null; $shouldNotify = $true }
            elseif ($ForceMail) { $reminders.Add($finding) | Out-Null; $shouldNotify = $true }
            elseif ($null -eq $lastNotified -or ($now - $lastNotified).TotalMinutes -ge $cooldown) {
                $reminders.Add($finding) | Out-Null; $shouldNotify = $true
            }

            $entry = [pscustomobject]@{
                FirstSeen    = $previous.FirstSeen
                LastSeen     = $now
                LastNotified = if ($shouldNotify) { $now } else { $previous.LastNotified }
                Severity     = $finding.Severity
                Occurrences  = [int]$previous.Occurrences + 1
                Message      = $finding.Message
            }
        }

        Add-Member -InputObject $nextState -NotePropertyName $key -NotePropertyValue $entry -Force
    }

    # Chiavi presenti nello stato ma non piu attive = rientro
    foreach ($prop in $State.Alerts.PSObject.Properties) {
        if (-not $activeKeys.ContainsKey($prop.Name)) {
            $parts = $prop.Name -split '\|'
            $recovered.Add([pscustomobject]@{
                Key        = $prop.Name
                Category   = $parts[0]
                Server     = if ($parts.Count -gt 1) { $parts[1] } else { '-' }
                Item       = if ($parts.Count -gt 2) { $parts[2] } else { '-' }
                Severity   = [string]$prop.Value.Severity
                Message    = [string]$prop.Value.Message
                FirstSeen  = $prop.Value.FirstSeen
                Duration   = if ($prop.Value.FirstSeen) { ($now - [datetime]$prop.Value.FirstSeen) } else { $null }
            }) | Out-Null
        }
    }

    # ToArray(): espone array veri, non List[object]. Le liste generiche esposte
    # come proprieta di un PSCustomObject si comportano diversamente tra
    # Windows PowerShell 5.1 e PowerShell 7.x.
    return [pscustomobject]@{
        Active       = $active
        New          = $newAlerts.ToArray()
        Escalated    = $escalated.ToArray()
        Reminders    = $reminders.ToArray()
        Recovered    = $recovered.ToArray()
        NextAlerts   = $nextState
    }
}

#endregion

#region ------------------------------------------------------------------ MAIL

function ConvertTo-HcHtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Get-HcSeverityColor {
    param([string]$Severity)
    switch ($Severity) {
        'Critical' { '#c0392b' }
        'Unknown'  { '#8e44ad' }
        'Warning'  { '#e67e22' }
        'Info'     { '#2980b9' }
        default    { '#27ae60' }
    }
}

function New-HcFindingTable {
    param([string]$Title, [object[]]$Rows, [string]$Accent = '#34495e')

    if (-not $Rows -or $Rows.Count -eq 0) { return '' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<h3 style='font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:$Accent;margin:22px 0 6px 0;'>$(ConvertTo-HcHtmlText $Title) ($($Rows.Count))</h3>")
    [void]$sb.AppendLine("<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px;'>")
    [void]$sb.AppendLine("<tr style='background:#f4f6f7;color:#2c3e50;text-align:left;'><th style='border:1px solid #dfe4e6;'>Severita</th><th style='border:1px solid #dfe4e6;'>Server</th><th style='border:1px solid #dfe4e6;'>Categoria</th><th style='border:1px solid #dfe4e6;'>Oggetto</th><th style='border:1px solid #dfe4e6;'>Dettaglio</th></tr>")

    foreach ($row in ($Rows | Sort-Object @{Expression={Get-SeverityRank $_.Severity}; Descending=$true}, Server, Category)) {
        $color = Get-HcSeverityColor $row.Severity
        $template = "<tr>" +
            "<td style='border:1px solid #dfe4e6;'><span style='display:inline-block;padding:2px 8px;border-radius:3px;background:{0};color:#fff;font-weight:600;'>{1}</span></td>" +
            "<td style='border:1px solid #dfe4e6;font-weight:600;'>{2}</td>" +
            "<td style='border:1px solid #dfe4e6;'>{3}</td>" +
            "<td style='border:1px solid #dfe4e6;'>{4}</td>" +
            "<td style='border:1px solid #dfe4e6;'>{5}</td>" +
            "</tr>"
        # NB: l'espressione -f va risolta fuori dalla chiamata di metodo, altrimenti
        # le virgole vengono lette come separatori di argomenti di AppendLine().
        $html = $template -f $color, (ConvertTo-HcHtmlText $row.Severity), (ConvertTo-HcHtmlText $row.Server),
                              (ConvertTo-HcHtmlText $row.Category), (ConvertTo-HcHtmlText $row.Item),
                              (ConvertTo-HcHtmlText $row.Message)
        [void]$sb.AppendLine($html)
    }
    [void]$sb.AppendLine('</table>')
    return $sb.ToString()
}

# Stessa vista aggregata della sezione "Code di trasporto" della mail, ma a
# console: utile durante un giro a secco (-NoMail), dove la mail non parte e
# altrimenti delle code si vedrebbe solo il singolo finding quando supera soglia,
# non il quadro d'insieme del mail flow.
function Write-HcQueueSummaryConsole {
    param([object[]]$QueueSummary)

    if (-not $QueueSummary -or $QueueSummary.Count -eq 0) { return }

    $grandTotal = 0
    foreach ($row in $QueueSummary) { $grandTotal += [int64]$row.Total }

    Write-Host ("`nCode di trasporto - {0} messaggi su {1} server" -f $grandTotal, $QueueSummary.Count) -ForegroundColor White
    $header = '{0,-20} {1,9} {2,6} {3,10} {4,7} {5,7} {6,-26} {7,8}' -f `
        'Server', 'In coda', 'Code', 'Submiss.', 'Retry', 'Poison', 'NextHopDomain', 'Shadow'
    Write-Host $header -ForegroundColor Gray
    Write-Host ('-' * $header.Length) -ForegroundColor Gray

    foreach ($row in ($QueueSummary | Sort-Object { [int64]$_.Total } -Descending)) {
        $largest = '-'
        if ($row.LargestName) { $largest = '{0} ({1})' -f $row.LargestName, $row.LargestCount }
        if ($largest.Length -gt 26) { $largest = $largest.Substring(0, 23) + '...' }

        $line = '{0,-20} {1,9} {2,6} {3,10} {4,7} {5,7} {6,-26} {7,8}' -f `
            $row.Server, $row.Total, $row.QueueCount, $row.Submission, $row.RetryCount, $row.Poison, $largest, $row.ShadowTotal
        Write-Host $line -ForegroundColor (Get-HcConsoleColor $row.Severity)
    }
    Write-Host "(Shadow escluse dai totali: trattengono messaggi per progetto.)`n" -ForegroundColor DarkGray
}

function New-HcMailBody {
    param(
        [object]$AlertResult,
        [object[]]$AllFindings,
        [object[]]$Targets,
        [object[]]$QueueSummary,
        [switch]$Heartbeat
    )

    $orgName   = $script:Config.Organization.Name
    $duration  = ((Get-Date) - $script:StartTime).TotalSeconds
    $counts    = @{
        Critical = @($AllFindings | Where-Object { $_.Severity -eq 'Critical' }).Count
        Warning  = @($AllFindings | Where-Object { $_.Severity -eq 'Warning' }).Count
        Unknown  = @($AllFindings | Where-Object { $_.Severity -eq 'Unknown' }).Count
        Info     = @($AllFindings | Where-Object { $_.Severity -eq 'Info' }).Count
        Ok       = @($AllFindings | Where-Object { $_.Severity -eq 'OK' }).Count
    }

    $banner = if ($counts.Critical -gt 0) { '#c0392b' } elseif ($counts.Warning -gt 0 -or $counts.Unknown -gt 0) { '#e67e22' } else { '#27ae60' }
    $headline = if ($Heartbeat) { 'Riepilogo periodico - nessuna anomalia' }
                elseif ($counts.Critical -gt 0) { 'Anomalie CRITICHE rilevate' }
                elseif ($counts.Warning -gt 0 -or $counts.Unknown -gt 0) { 'Anomalie rilevate' }
                else { 'Ambiente in salute' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<html><body style='margin:0;padding:18px;background:#ffffff;'>")
    [void]$sb.AppendLine("<div style='font-family:Segoe UI,Arial,sans-serif;max-width:1100px;'>")
    [void]$sb.AppendLine("<div style='background:$banner;color:#fff;padding:14px 18px;border-radius:4px;'>")
    [void]$sb.AppendLine("<div style='font-size:18px;font-weight:600;'>$(ConvertTo-HcHtmlText $headline)</div>")
    [void]$sb.AppendLine("<div style='font-size:13px;opacity:.9;margin-top:3px;'>Organizzazione $(ConvertTo-HcHtmlText $orgName) - $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</div>")
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine("<table cellpadding='8' cellspacing='0' style='margin-top:14px;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>")
    [void]$sb.AppendLine("<tr>")
    foreach ($sev in @('Critical','Unknown','Warning','Info','Ok')) {
        $color = Get-HcSeverityColor $sev
        [void]$sb.AppendLine("<td style='border:1px solid #dfe4e6;text-align:center;min-width:92px;'><div style='color:$color;font-size:22px;font-weight:700;'>$($counts[$sev])</div><div style='color:#7f8c8d;font-size:11px;text-transform:uppercase;'>$sev</div></td>")
    }
    [void]$sb.AppendLine("</tr></table>")

    [void]$sb.AppendLine((New-HcFindingTable -Title 'Nuove anomalie' -Rows @($AlertResult.New) -Accent '#c0392b'))
    [void]$sb.AppendLine((New-HcFindingTable -Title 'Anomalie peggiorate' -Rows @($AlertResult.Escalated) -Accent '#c0392b'))
    [void]$sb.AppendLine((New-HcFindingTable -Title 'Anomalie ancora aperte' -Rows @($AlertResult.Reminders) -Accent '#e67e22'))

    $recovered = @($AlertResult.Recovered)
    if ($recovered.Count -gt 0) {
        [void]$sb.AppendLine("<h3 style='font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#27ae60;margin:22px 0 6px 0;'>Rientrate ($($recovered.Count))</h3>")
        [void]$sb.AppendLine("<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px;'>")
        [void]$sb.AppendLine("<tr style='background:#f4f6f7;text-align:left;'><th style='border:1px solid #dfe4e6;'>Server</th><th style='border:1px solid #dfe4e6;'>Categoria</th><th style='border:1px solid #dfe4e6;'>Oggetto</th><th style='border:1px solid #dfe4e6;'>Durata</th><th style='border:1px solid #dfe4e6;'>Ultimo messaggio</th></tr>")
        foreach ($rec in $recovered) {
            $dur = if ($rec.Duration) { '{0:N0}h {1:N0}m' -f [math]::Floor($rec.Duration.TotalHours), $rec.Duration.Minutes } else { 'n/d' }
            [void]$sb.AppendLine(("<tr><td style='border:1px solid #dfe4e6;font-weight:600;'>{0}</td><td style='border:1px solid #dfe4e6;'>{1}</td><td style='border:1px solid #dfe4e6;'>{2}</td><td style='border:1px solid #dfe4e6;'>{3}</td><td style='border:1px solid #dfe4e6;color:#7f8c8d;'>{4}</td></tr>" -f
                (ConvertTo-HcHtmlText $rec.Server), (ConvertTo-HcHtmlText $rec.Category), (ConvertTo-HcHtmlText $rec.Item), $dur, (ConvertTo-HcHtmlText $rec.Message)))
        }
        [void]$sb.AppendLine('</table>')
    }

    # Vista aggregata delle code: mostrata sempre, non solo in presenza di alert.
    # Il quadro del mail flow serve anche per confermare che sia tutto scorrevole.
    $queues = @($QueueSummary)
    if ($queues.Count -gt 0 -and $script:Config.Mail.IncludeQueueSummary) {
        $grandTotal = 0
        foreach ($row in $queues) { $grandTotal += [int64]$row.Total }

        [void]$sb.AppendLine("<h3 style='font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#34495e;margin:22px 0 6px 0;'>Code di trasporto - $grandTotal messaggi su $($queues.Count) server</h3>")
        [void]$sb.AppendLine("<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px;'>")
        [void]$sb.AppendLine("<tr style='background:#f4f6f7;text-align:left;'>" +
            "<th style='border:1px solid #dfe4e6;'>Server</th>" +
            "<th style='border:1px solid #dfe4e6;'>In coda</th>" +
            "<th style='border:1px solid #dfe4e6;'>Code attive</th>" +
            "<th style='border:1px solid #dfe4e6;'>Submission</th>" +
            "<th style='border:1px solid #dfe4e6;'>In retry</th>" +
            "<th style='border:1px solid #dfe4e6;'>Poison</th>" +
            "<th style='border:1px solid #dfe4e6;'>NextHopDomain</th>" +
            "<th style='border:1px solid #dfe4e6;'>Shadow</th></tr>")

        foreach ($row in ($queues | Sort-Object { [int64]$_.Total } -Descending)) {
            $color = Get-HcSeverityColor $row.Severity
            $largest = '-'
            if ($row.LargestName) { $largest = '{0} ({1})' -f $row.LargestName, $row.LargestCount }
            $poisonStyle = if ([int64]$row.Poison -gt 0) { "color:#e67e22;font-weight:600;" } else { '' }
            $retryStyle  = if ([int64]$row.RetryCount -gt 0) { "color:#e67e22;font-weight:600;" } else { '' }

            $template = "<tr>" +
                "<td style='border:1px solid #dfe4e6;font-weight:600;'>{0}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;color:{1};font-weight:600;'>{2}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;'>{3}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;'>{4}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;{5}'>{6}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;{7}'>{8}</td>" +
                "<td style='border:1px solid #dfe4e6;'>{9}</td>" +
                "<td style='border:1px solid #dfe4e6;text-align:right;color:#95a5a6;'>{10}</td></tr>"
            $html = $template -f (ConvertTo-HcHtmlText $row.Server), $color, $row.Total, $row.QueueCount,
                                 $row.Submission, $retryStyle, $row.RetryCount, $poisonStyle, $row.Poison,
                                 (ConvertTo-HcHtmlText $largest), $row.ShadowTotal
            [void]$sb.AppendLine($html)
        }
        [void]$sb.AppendLine('</table>')
        [void]$sb.AppendLine("<p style='font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#95a5a6;margin:4px 0 0 0;'>Le code Shadow Redundancy trattengono messaggi per progetto e non rientrano nei totali.</p>")
    }

    # Riepilogo per server: utile per capire a colpo d'occhio chi sta male
    [void]$sb.AppendLine("<h3 style='font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#34495e;margin:22px 0 6px 0;'>Riepilogo per server ($($Targets.Count))</h3>")
    [void]$sb.AppendLine("<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px;'>")
    [void]$sb.AppendLine("<tr style='background:#f4f6f7;text-align:left;'><th style='border:1px solid #dfe4e6;'>Server</th><th style='border:1px solid #dfe4e6;'>Ruolo</th><th style='border:1px solid #dfe4e6;'>DAG</th><th style='border:1px solid #dfe4e6;'>Critical</th><th style='border:1px solid #dfe4e6;'>Warning</th><th style='border:1px solid #dfe4e6;'>Stato</th></tr>")
    foreach ($tgt in ($Targets | Sort-Object Name)) {
        $srvFindings = @($AllFindings | Where-Object { $_.Server -eq $tgt.Name })
        $c = @($srvFindings | Where-Object { $_.Severity -eq 'Critical' }).Count
        $w = @($srvFindings | Where-Object { $_.Severity -in @('Warning','Unknown') }).Count
        $state = if ($c -gt 0) { 'CRITICO' } elseif ($w -gt 0) { 'ATTENZIONE' } else { 'OK' }
        $color = if ($c -gt 0) { '#c0392b' } elseif ($w -gt 0) { '#e67e22' } else { '#27ae60' }
        [void]$sb.AppendLine(("<tr><td style='border:1px solid #dfe4e6;font-weight:600;'>{0}</td><td style='border:1px solid #dfe4e6;'>{1}</td><td style='border:1px solid #dfe4e6;'>{2}</td><td style='border:1px solid #dfe4e6;text-align:center;'>{3}</td><td style='border:1px solid #dfe4e6;text-align:center;'>{4}</td><td style='border:1px solid #dfe4e6;color:{5};font-weight:600;'>{6}</td></tr>" -f
            (ConvertTo-HcHtmlText $tgt.Name), (ConvertTo-HcHtmlText $tgt.Role), (ConvertTo-HcHtmlText $tgt.Dag), $c, $w, $color, $state))
    }
    [void]$sb.AppendLine('</table>')

    [void]$sb.AppendLine("<p style='font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#95a5a6;margin-top:22px;border-top:1px solid #ecf0f1;padding-top:10px;'>")
    [void]$sb.AppendLine(("Eseguito da {0} come {1} - durata {2:N1}s - {3} controlli valutati.<br/>Script: {4}" -f
        $env:COMPUTERNAME, ($env:USERDOMAIN + '\' + $env:USERNAME), $duration, $AllFindings.Count,
        (ConvertTo-HcHtmlText (Join-Path $script:ScriptRoot 'Invoke-ExchangeHealthCheck.ps1'))))
    [void]$sb.AppendLine('</p></div></body></html>')

    return $sb.ToString()
}

function Send-HcMail {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [string[]]$Attachments,
        [ValidateSet('Low','Normal','High')][string]$Priority = 'Normal'
    )

    $mailCfg = $script:Config.Mail
    if (-not $mailCfg.Enabled) { Write-HcLog 'Invio mail disabilitato in configurazione.'; return $false }
    if ($NoMail) { Write-HcLog 'Invio mail saltato (-NoMail).'; return $false }

    $servers = @($mailCfg.SmtpServers)
    if ($servers.Count -eq 0) { Write-HcLog 'Nessun server SMTP configurato.' -Level ERROR; return $false }

    foreach ($smtpHost in $servers) {
        $message = $null
        $client  = $null
        try {
            $message = New-Object System.Net.Mail.MailMessage
            $message.From = New-Object System.Net.Mail.MailAddress($mailCfg.From, $mailCfg.FromDisplayName)
            foreach ($addr in @($mailCfg.To)) { if ($addr) { $message.To.Add($addr) } }
            foreach ($addr in @($mailCfg.Cc)) { if ($addr) { $message.CC.Add($addr) } }
            $message.Subject    = $Subject
            $message.Body       = $Body
            $message.IsBodyHtml = $true
            $message.Priority   = [System.Net.Mail.MailPriority]::$Priority

            foreach ($file in @($Attachments)) {
                if ($file -and (Test-Path -LiteralPath $file)) {
                    $message.Attachments.Add((New-Object System.Net.Mail.Attachment($file)))
                }
            }

            $client = New-Object System.Net.Mail.SmtpClient($smtpHost, [int]$mailCfg.Port)
            $client.EnableSsl = [bool]$mailCfg.UseSsl
            $client.Timeout   = 60000

            if ($mailCfg.CredentialFile) {
                $credPath = Resolve-HcPath $mailCfg.CredentialFile
                if (Test-Path -LiteralPath $credPath) {
                    $cred = Import-Clixml -LiteralPath $credPath
                    $client.Credentials = New-Object System.Net.NetworkCredential($cred.UserName, $cred.GetNetworkCredential().Password)
                }
                else {
                    Write-HcLog "CredentialFile non trovato: $credPath" -Level WARN
                }
            }
            elseif ($mailCfg.UseDefaultCredentials) {
                $client.UseDefaultCredentials = $true
            }

            $client.Send($message)
            Write-HcLog ('Mail inviata via {0} a: {1}' -f $smtpHost, (@($mailCfg.To) -join ', '))
            return $true
        }
        catch {
            # SmtpClient incapsula la causa vera: il messaggio esterno e quasi sempre
            # un generico "Failure sending mail", mentre "No such host is known",
            # "Unable to connect" o la risposta 5.7.x del server stanno nelle
            # eccezioni interne. Si risale tutta la catena.
            $reasons = @()
            $current = $_.Exception
            while ($current) {
                if ($current.Message -and ($reasons -notcontains $current.Message)) { $reasons += $current.Message }
                $current = $current.InnerException
            }
            Write-HcLog ('Invio via {0}:{1} fallito: {2}' -f $smtpHost, $mailCfg.Port, ($reasons -join ' -> ')) -Level WARN
        }
        finally {
            if ($message) { $message.Dispose() }
            if ($client -and $client -is [System.IDisposable]) { $client.Dispose() }
        }
    }

    Write-HcLog 'Invio mail fallito su tutti i server SMTP configurati.' -Level ERROR
    return $false
}

#endregion

#region ------------------------------------------------------------------ MAIN

$mutex = $null
$mutexAcquired = $false

try {
    # --- Configurazione
    if (-not $ConfigPath) { $ConfigPath = Join-Path $script:ScriptRoot 'ExchangeHealthCheck.config.json' }
    $defaultConfig = ConvertFrom-Json $DefaultConfigJson
    if (Test-Path -LiteralPath $ConfigPath) {
        $userConfig = ConvertFrom-Json (Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8)
        $script:Config = Merge-HcConfig -Default $defaultConfig -Override $userConfig
    }
    else {
        Write-Warning "File di configurazione non trovato ($ConfigPath): uso i valori di default."
        $script:Config = $defaultConfig
    }

    # --- Percorsi e logging
    $logDir    = Resolve-HcPath $script:Config.Paths.LogDirectory
    $reportDir = Resolve-HcPath $script:Config.Paths.ReportDirectory
    $stateFile = Resolve-HcPath $script:Config.Paths.StateFile
    foreach ($dir in @($logDir, $reportDir, (Split-Path -Parent $stateFile))) {
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }
    $script:LogFile = Join-Path $logDir ('healthcheck-{0}.log' -f (Get-Date -Format 'yyyyMMdd'))

    Write-HcLog '============================================================'
    Write-HcLog ('Avvio health check Exchange - organizzazione "{0}"' -f $script:Config.Organization.Name)

    # --- Mail di test
    if ($TestMail) {
        $mailCfg = $script:Config.Mail
        if (-not $mailCfg.Enabled) {
            Write-HcLog 'Test mail non eseguibile: Mail.Enabled e false nella configurazione. Impostalo a true e rilancia.' -Level ERROR
            return
        }

        Write-HcLog ('Test mail -> server {0} porta {1} | SSL {2} | autenticazione {3} | da {4} | a {5}' -f `
            (@($mailCfg.SmtpServers) -join ', '), $mailCfg.Port, [bool]$mailCfg.UseSsl,
            $(if ($mailCfg.CredentialFile) { 'CredentialFile' } elseif ($mailCfg.UseDefaultCredentials) { 'account corrente' } else { 'anonima' }),
            $mailCfg.From, (@($mailCfg.To) -join ', '))

        $body = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>" +
                "<p>Se ricevi questo messaggio, la configurazione SMTP dello health check Exchange funziona.</p>" +
                "<table cellpadding='4' style='border-collapse:collapse;font-size:12px;color:#555;'>" +
                "<tr><td>Inviato da</td><td><b>$env:COMPUTERNAME</b></td></tr>" +
                "<tr><td>Account</td><td>$env:USERDOMAIN\$env:USERNAME</td></tr>" +
                "<tr><td>Server SMTP configurati</td><td>$(ConvertTo-HcHtmlText (@($mailCfg.SmtpServers) -join ', '))</td></tr>" +
                "<tr><td>Porta / SSL</td><td>$($mailCfg.Port) / $([bool]$mailCfg.UseSsl)</td></tr>" +
                "<tr><td>Orario</td><td>$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</td></tr></table></div>"

        $ok = Send-HcMail -Subject ('{0} Test configurazione SMTP' -f $mailCfg.SubjectPrefix) -Body $body
        if ($ok) { Write-HcLog 'Test mail: messaggio accettato dal server SMTP. Verifica la ricezione nella casella di destinazione.' }
        else { Write-HcLog 'Test mail fallito: vedi il motivo nelle righe "Invio via ..." qui sopra.' -Level ERROR }
        return
    }

    # --- Lock: evita sovrapposizioni se un giro dura piu dell'intervallo
    $mutex = New-Object System.Threading.Mutex($false, 'Global\ExchangeHealthCheck')
    try { $mutexAcquired = $mutex.WaitOne(0) }
    catch [System.Threading.AbandonedMutexException] { $mutexAcquired = $true }
    if (-not $mutexAcquired) {
        Write-HcLog 'Un altro health check e gia in esecuzione: esco senza fare nulla.' -Level WARN
        return
    }

    # --- Connessione a Exchange
    Connect-HcExchange

    # --- Perimetro
    $targets = Get-HcTargetServer

    # --- Aggancio a un DC del dominio dei server (solo se non ne e stato imposto uno)
    Set-HcAutoDomainController -Targets $targets

    # --- Dati OS in parallelo
    $remoteData = @{}
    if ((Test-CheckEnabled 'Os') -or (Test-CheckEnabled 'Disk') -or (Test-CheckEnabled 'Services')) {
        $remoteData = Get-HcRemoteData -Targets $targets
    }

    foreach ($target in $targets) {
        $data = $null
        $upperName = $target.Name.ToUpperInvariant()
        if ($remoteData.ContainsKey($upperName)) { $data = $remoteData[$upperName] }
        $target.Online = ($null -ne $data)

        Write-HcLog ('--- Controllo server {0} ({1})' -f $target.Name, $target.Role)

        if (Test-CheckEnabled 'Os')       { Invoke-HcCheck -Category 'Os'       -ServerName $target.Name -Body { Invoke-HcOsCheck      -Target $target -Data $data } }
        if (Test-CheckEnabled 'Disk')     { Invoke-HcCheck -Category 'Disk'     -ServerName $target.Name -Body { Invoke-HcDiskCheck    -Target $target -Data $data } }
        if (Test-CheckEnabled 'Services') { Invoke-HcCheck -Category 'Service'  -ServerName $target.Name -Body { Invoke-HcServiceCheck -Target $target -Data $data } }

        # WinRM e i cmdlet Exchange sono due canali diversi: un server puo essere
        # vivo e sano ma avere WinRM chiuso. Di default si salta comunque, per non
        # restare appesi ai timeout RPC su un server davvero morto; chi sa che si
        # tratta solo di accesso bloccato puo proseguire.
        if (-not $target.Online) {
            $skipWhenOffline = $true
            if ($null -ne $script:Config.Servers.SkipExchangeChecksWhenOffline) {
                $skipWhenOffline = [bool]$script:Config.Servers.SkipExchangeChecksWhenOffline
            }
            if ($skipWhenOffline) {
                Write-HcLog ('{0} non raggiungibile via WinRM: salto i check Exchange-side.' -f $target.Name) -Level WARN
                continue
            }
            Write-HcLog ('{0} non raggiungibile via WinRM: proseguo comunque con i check Exchange-side.' -f $target.Name) -Level WARN
        }

        if (Test-CheckEnabled 'Components')   { Invoke-HcCheck -Category 'ComponentState'      -ServerName $target.Name -Body { Invoke-HcComponentCheck   -Target $target } }
        if (Test-CheckEnabled 'Health')       { Invoke-HcCheck -Category 'ManagedAvailability' -ServerName $target.Name -Body { Invoke-HcHealthCheck      -Target $target } }
        if (Test-CheckEnabled 'Certificates') { Invoke-HcCheck -Category 'Certificate'         -ServerName $target.Name -Body { Invoke-HcCertificateCheck -Target $target } }
        if (Test-CheckEnabled 'Queues')       { Invoke-HcCheck -Category 'Queue'               -ServerName $target.Name -Body { Invoke-HcQueueCheck       -Target $target } }
        if (Test-CheckEnabled 'BackPressure') { Invoke-HcCheck -Category 'BackPressure'        -ServerName $target.Name -Body { Invoke-HcBackPressureCheck -Target $target } }

        if ($target.IsMailbox) {
            if (Test-CheckEnabled 'Dag')         { Invoke-HcCheck -Category 'DatabaseCopy' -ServerName $target.Name -Body { Invoke-HcCopyStatusCheck   -Target $target } }
            if (Test-CheckEnabled 'Replication') { Invoke-HcCheck -Category 'Replication'  -ServerName $target.Name -Body { Invoke-HcReplicationCheck -Target $target } }
            if (Test-CheckEnabled 'Mapi')        { Invoke-HcCheck -Category 'Mapi'         -ServerName $target.Name -Body { Invoke-HcMapiCheck        -Target $target } }
        }
    }

    # --- Check a livello di organizzazione
    if (Test-CheckEnabled 'Dag')       { Invoke-HcDagCheck -Targets $targets }
    if (Test-CheckEnabled 'Databases') { Invoke-HcCheck -Category 'Database' -Body { Invoke-HcDatabaseCheck -Targets $targets } }

    $allFindings = $script:Findings.ToArray()
    $critical = @($allFindings | Where-Object { $_.Severity -eq 'Critical' })
    $warning  = @($allFindings | Where-Object { $_.Severity -in @('Warning','Unknown') })
    Write-HcLog ('Controlli completati: {0} finding ({1} critical, {2} warning/unknown).' -f $allFindings.Count, $critical.Count, $warning.Count)

    if ($script:Config.Console.ShowQueueSummary) {
        Write-HcQueueSummaryConsole -QueueSummary $script:QueueSummary.ToArray()
    }

    # --- Report su disco
    $stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
    $csvPath = Join-Path $reportDir ('healthcheck-{0}.csv' -f $stamp)
    try {
        $allFindings | Select-Object Timestamp, Severity, Server, Category, Item, Value, Message |
            Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-HcLog "Report CSV: $csvPath"
    }
    catch {
        Write-HcLog ('Scrittura CSV fallita: {0}' -f $_.Exception.Message) -Level WARN
        $csvPath = $null
    }

    # --- Stato e decisione di invio
    $state  = Get-HcState -Path $stateFile
    $alerts = Resolve-HcAlert -State $state -Findings $allFindings

    $toNotify = @($alerts.New) + @($alerts.Escalated) + @($alerts.Reminders)
    $recovered = @($alerts.Recovered)

    $heartbeatHours = [double]$script:Config.Alerting.HeartbeatHours
    $needHeartbeat = $false
    if ($heartbeatHours -gt 0) {
        if (-not $state.LastHeartbeat) { $needHeartbeat = $true }
        elseif (((Get-Date) - [datetime]$state.LastHeartbeat).TotalHours -ge $heartbeatHours) { $needHeartbeat = $true }
    }

    $sendRecovery = ([bool]$script:Config.Alerting.SendRecovery -and $recovered.Count -gt 0)
    $shouldSend   = ($toNotify.Count -gt 0) -or $sendRecovery -or $needHeartbeat -or $ForceMail

    $nextState = [pscustomobject]@{
        Alerts        = $alerts.NextAlerts
        LastHeartbeat = $state.LastHeartbeat
        LastRun       = Get-Date
    }

    $mailSent = $false
    if ($shouldSend) {
        $isHeartbeat = ($toNotify.Count -eq 0 -and -not $sendRecovery)
        $body = New-HcMailBody -AlertResult $alerts -AllFindings $allFindings -Targets $targets -QueueSummary $script:QueueSummary.ToArray() -Heartbeat:$isHeartbeat

        $criticalCount = @($toNotify | Where-Object { $_.Severity -eq 'Critical' }).Count
        $warningCount  = @($toNotify | Where-Object { $_.Severity -in @('Warning','Unknown') }).Count

        if ($criticalCount -gt 0) {
            $tag = 'CRITICO'; $priority = 'High'
        }
        elseif ($warningCount -gt 0) {
            $tag = 'WARNING'; $priority = 'Normal'
        }
        elseif ($sendRecovery) {
            $tag = 'RIENTRO'; $priority = 'Normal'
        }
        else {
            $tag = 'OK'; $priority = 'Low'
        }

        $subject = '{0} {1} - {2} - {3} critical / {4} warning' -f `
            $script:Config.Mail.SubjectPrefix, $tag, $script:Config.Organization.Name, $critical.Count, $warning.Count

        $attachments = @()
        if ($script:Config.Mail.AttachCsv -and $csvPath) { $attachments += $csvPath }

        $mailSent = [bool](Send-HcMail -Subject $subject -Body $body -Attachments $attachments -Priority $priority)

        # Se la mail NON e partita, lo stato non deve registrare la notifica:
        # altrimenti l'anomalia entrerebbe in cooldown senza che nessuno l'abbia
        # ricevuta, ed e proprio il caso in cui non ci si puo permettere silenzio.
        # Si riporta indietro LastNotified e si rimettono i rientri, cosi il giro
        # successivo riprova da capo.
        if (-not $mailSent) {
            Write-HcLog 'Mail non inviata: annullo le notifiche nello stato, il prossimo giro riprovera.' -Level WARN

            foreach ($finding in $toNotify) {
                if ($nextState.Alerts.PSObject.Properties.Name -contains $finding.Key) {
                    $previousNotified = $null
                    if ($state.Alerts.PSObject.Properties.Name -contains $finding.Key) {
                        $previousNotified = $state.Alerts.$($finding.Key).LastNotified
                    }
                    $nextState.Alerts.$($finding.Key).LastNotified = $previousNotified
                }
            }

            foreach ($rec in $recovered) {
                if ($state.Alerts.PSObject.Properties.Name -contains $rec.Key) {
                    Add-Member -InputObject $nextState.Alerts -NotePropertyName $rec.Key `
                        -NotePropertyValue $state.Alerts.$($rec.Key) -Force
                }
            }
        }
        elseif ($needHeartbeat) {
            $nextState.LastHeartbeat = Get-Date
        }
    }
    else {
        Write-HcLog 'Nessuna notifica da inviare (nessuna novita e cooldown non scaduto).'
    }

    Save-HcState -Path $stateFile -State $nextState

    Remove-HcOldFiles -Directory $logDir    -RetentionDays ([int]$script:Config.Paths.RetentionDays)
    Remove-HcOldFiles -Directory $reportDir -RetentionDays ([int]$script:Config.Paths.RetentionDays)

    Write-HcLog ('Health check terminato in {0:N1}s.' -f ((Get-Date) - $script:StartTime).TotalSeconds)

    if ($PassThru) { $allFindings }

    # Exit code utile allo scheduler: 2 = critical, 1 = warning, 0 = ok
    if ($critical.Count -gt 0) { $global:LASTEXITCODE = 2 }
    elseif ($warning.Count -gt 0) { $global:LASTEXITCODE = 1 }
    else { $global:LASTEXITCODE = 0 }
}
catch {
    $errorMessage = $_.Exception.Message
    Write-HcLog ('ERRORE FATALE: {0}' -f $errorMessage) -Level ERROR
    Write-HcLog ($_.ScriptStackTrace) -Level ERROR

    # Il monitor che muore in silenzio e peggio di un monitor assente
    try {
        if ($script:Config -and -not $NoMail) {
            $body = "<p style='font-family:Segoe UI,Arial'>L'health check Exchange si e interrotto con un errore:</p>" +
                    "<pre style='font-family:Consolas,monospace;background:#f8f9f9;padding:10px;border:1px solid #e5e8e8;'>" +
                    (ConvertTo-HcHtmlText $errorMessage) + "</pre>" +
                    "<p style='font-family:Segoe UI,Arial;font-size:12px;color:#7f8c8d'>Host: $env:COMPUTERNAME - $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</p>"
            Send-HcMail -Subject ('{0} ERRORE - health check non completato' -f $script:Config.Mail.SubjectPrefix) `
                -Body $body -Priority High | Out-Null
        }
    }
    catch { }

    $global:LASTEXITCODE = 3
}
finally {
    Disconnect-HcExchange
    if ($mutex) {
        if ($mutexAcquired) { try { $mutex.ReleaseMutex() } catch { } }
        $mutex.Dispose()
    }
}

exit $global:LASTEXITCODE

#endregion
