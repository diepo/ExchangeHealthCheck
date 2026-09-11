#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Registra (o aggiorna) lo scheduled task che esegue l'health check Exchange.

.DESCRIPTION
    Crea un'attivita pianificata che parte all'avvio del server e si ripete a
    intervalli regolari. Il task viene eseguito con l'account di servizio
    indicato, che deve avere:
      - ruolo RBAC "View-Only Organization Management" in Exchange
      - diritti di amministratore locale sui server Exchange (WinRM/CIM remoto)
      - "Log on as a batch job" sul server dove gira il task

.PARAMETER IntervalMinutes
    Intervallo di ripetizione. Default 15 minuti.

.PARAMETER UserName
    Account di servizio (DOMINIO\utente). Se omesso viene usato SYSTEM, che
    funziona solo se il task gira su un server Exchange e l'account macchina
    ha i permessi necessari: consigliato usare un account dedicato o un gMSA.

.PARAMETER TaskName
    Nome dell'attivita pianificata. Default: "Exchange Health Check".

.PARAMETER ScriptPath
    Percorso dello script di health check. Default: .\Invoke-ExchangeHealthCheck.ps1

.PARAMETER Unregister
    Rimuove l'attivita pianificata invece di crearla.

.EXAMPLE
    .\Install-ExchangeHealthCheckTask.ps1 -IntervalMinutes 15 -UserName 'CONTOSO\svc-exmonitor'

.EXAMPLE
    .\Install-ExchangeHealthCheckTask.ps1 -Unregister
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(1, 1440)]
    [int]    $IntervalMinutes = 15,
    [string] $UserName,
    [string] $TaskName = 'Exchange Health Check',
    [string] $ScriptPath,
    [switch] $Unregister
)

$ErrorActionPreference = 'Stop'
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ScriptPath) { $ScriptPath = Join-Path $root 'Invoke-ExchangeHealthCheck.ps1' }

if ($Unregister) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        if ($PSCmdlet.ShouldProcess($TaskName, 'Unregister-ScheduledTask')) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Host "Attivita '$TaskName' rimossa." -ForegroundColor Green
        }
    }
    else {
        Write-Host "Attivita '$TaskName' non presente." -ForegroundColor Yellow
    }
    return
}

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Script non trovato: $ScriptPath"
}

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $ScriptPath

$action = New-ScheduledTaskAction -Execute $psExe -Argument $arguments -WorkingDirectory $root

# Trigger: all'avvio + ripetizione infinita ogni N minuti, cosi il monitor
# riparte da solo dopo un reboot del server.
$trigger = New-ScheduledTaskTrigger -AtStartup
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)).Repetition

# Secondo trigger immediato, per non aspettare il prossimo reboot
$triggerNow = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes ([Math]::Max(30, $IntervalMinutes * 3))) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5)

if ($UserName) {
    $securePassword = Read-Host -Prompt "Password per $UserName" -AsSecureString
    $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))

    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $TaskName `
        -Action $action -Trigger @($trigger, $triggerNow) -Settings $settings `
        -User $UserName -Password $plain -RunLevel Highest `
        -Description 'Health check periodico dell''ambiente Exchange on-premises con alerting via e-mail.' | Out-Null
    $plain = $null
}
else {
    Write-Warning 'Nessun -UserName indicato: il task girera come SYSTEM. Verificare che abbia i permessi Exchange e WinRM necessari.'
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $TaskName `
        -Action $action -Trigger @($trigger, $triggerNow) -Settings $settings -Principal $principal `
        -Description 'Health check periodico dell''ambiente Exchange on-premises con alerting via e-mail.' | Out-Null
}

Write-Host ""
Write-Host "Attivita '$TaskName' registrata." -ForegroundColor Green
Write-Host "  Script    : $ScriptPath"
Write-Host "  Intervallo: ogni $IntervalMinutes minuti"
Write-Host "  Account   : $(if ($UserName) { $UserName } else { 'SYSTEM' })"
Write-Host ""
Write-Host "Esecuzione immediata di prova:" -ForegroundColor Cyan
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "  Get-ScheduledTaskInfo -TaskName '$TaskName'"
