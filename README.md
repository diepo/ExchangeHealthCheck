# Exchange Health Check

Monitoraggio periodico di un ambiente Exchange on-premises multi-server, con
alerting via e-mail, deduplica degli alert e notifiche di rientro.

```
ExchangeHealthCheck\
├── Invoke-ExchangeHealthCheck.ps1           # script principale
├── ExchangeHealthCheck.config.example.json  # template di configurazione
├── ExchangeHealthCheck.config.json          # la TUA configurazione (non versionata)
├── Install-ExchangeHealthCheckTask.ps1 # registra l'attivita pianificata
├── GUIDE.md                             # guida utente (sorgente del PDF)
├── ExchangeHealthCheck-Guide.pdf         # stessa guida, in formato documento
├── Build-Guide.py                       # rigenera il PDF da GUIDE.md
├── HANDOFF.md                           # note tecniche per chi riprende il progetto
├── Logs\                               # log giornalieri (creata al primo run)
├── Reports\                            # CSV di ogni esecuzione
└── State\alert-state.json              # stato degli alert tra un run e l'altro
```

`GUIDE.md`/`ExchangeHealthCheck-Guide.pdf` sono pensati per chi preferisce un
documento a un repository; `HANDOFF.md` è per chi riprende lo sviluppo dello
script (architettura, scelte di design, cronologia dei bug trovati con causa e
fix). Tutti e tre vanno aggiornati insieme al codice, non sono generati
automaticamente dai commit.

## Cosa controlla

| Categoria | Controllo | Come |
|---|---|---|
| Connectivity | server raggiungibile | WinRM / CIM |
| Uptime | reboot recenti, uptime | `Win32_OperatingSystem` |
| Memory / Cpu | memoria libera, carico | `Win32_OperatingSystem`, `Win32_Processor` |
| Os | riavvio in sospeso | registry (CBS / WU / PendingFileRename) |
| Disk | spazio libero **incluse le mount point** | `Win32_Volume` |
| Service | servizi `MSExchange*` in Auto ma fermi | `Win32_Service` |
| ServiceHealth | servizi richiesti dal ruolo | `Test-ServiceHealth` |
| ComponentState | server lasciato in maintenance mode | `Get-ServerComponentState` |
| ManagedAvailability | health set Unhealthy / Degraded | `Get-HealthReport` |
| Dag | membri fermi, witness, nodi cluster | `Get-DatabaseAvailabilityGroup`, `Get-ClusterNode` |
| DatabaseCopy | stato delle copie | `Get-MailboxDatabaseCopyStatus` |
| CopyQueue / ReplayQueue | log in arretrato | idem |
| ContentIndex | stato dell'indice di ricerca | idem |
| Replication | salute della replica | `Test-ReplicationHealth` |
| Database | database non montati | `Get-MailboxDatabase -Status` |
| Backup | eta dell'ultimo backup | `LastFullBackup` / `LastIncrementalBackup` |
| Activation | DB non sulla copia con preference 1 | `ActivationPreference` |
| Queue | code, submission, poison, retry | `Get-Queue` |
| BackPressure | back pressure del transport | `Get-ExchangeDiagnosticInfo` |
| Certificate | certificati in scadenza | `Get-ExchangeCertificate` |
| Mapi | connettivita MAPI (opzionale) | `Test-MapiConnectivity` |

Le code **ShadowRedundancy** sono escluse dai conteggi: trattengono messaggi per
progetto e non sono un'anomalia.

## Sola lettura

Lo script **non modifica nulla** in Exchange. Usa esclusivamente cmdlet `Get-*` e
`Test-*`; non monta/smonta database, non riavvia servizi, non sposta copie, non
tocca code o configurazione. Le uniche scritture sono **locali alla cartella dello
script**: log, CSV in `Reports\`, file di stato in `State\`, più la cancellazione
dei file più vecchi di `RetentionDays` dentro quelle due cartelle. A queste si
aggiunge l'invio delle mail di alert.

Due precisazioni:

* `Set-ADServerSettings -ViewEntireForest $true` ha effetto **solo sulla sessione
  PowerShell corrente**, non cambia impostazioni persistenti.
* `Test-MapiConnectivity` (check `Mapi`, **disabilitato di default**) non altera
  dati ma esegue un logon reale alla system mailbox di ogni database.

Di conseguenza l'account può restare con soli permessi **View-Only Organization
Management**: se qualcuno un domani aggiungesse un cmdlet di scrittura, fallirebbe
per mancanza di privilegi invece di agire.

## Comportamento con server irraggiungibili

Un server che non risponde **non interrompe l'esecuzione**:

1. la raccolta dati usa un unico `Invoke-Command` in fan-out con
   `-ErrorAction SilentlyContinue`: gli host che falliscono non sollevano
   eccezioni terminanti, finiscono in una error variable e vengono loggati;
2. per ognuno viene generato un finding `Connectivity` **Critical**
   ("Server non raggiungibile via WinRM/CIM");
3. i check Exchange-side su quel server vengono saltati, perché darebbero solo
   errori a cascata mascherando la causa vera;
4. gli altri server proseguono normalmente.

Inoltre **ogni singolo check è isolato** in un `try/catch` (`Invoke-HcCheck`): se
un cmdlet fallisce su un server, diventa un finding di severità `Unknown` e il
giro continua con il check successivo. Solo un errore nella fase di bootstrap
(configurazione illeggibile, impossibile connettersi a Exchange) ferma
l'esecuzione, e in quel caso parte una mail di errore dedicata.

`RemoteOpenTimeoutSeconds` (default 20) limita l'attesa di connessione WinRM: senza
di esso un server spento occuperebbe uno slot del fan-out per il timeout di
default (~45s).

## Requisiti

* Windows PowerShell 5.1 (testato anche su PowerShell 7.x)
* Cmdlet Exchange disponibili in uno di questi modi:
  * lo script gira **su un server Exchange** (carica lo snap-in da solo), oppure
  * `Organization.ConnectTo` valorizzato con un server Exchange: lo script apre
    una remote PowerShell session con autenticazione Kerberos
* Account di esecuzione con:
  * ruolo RBAC **View-Only Organization Management**
  * **amministratore locale** sui server Exchange (serve per WinRM/CIM remoto)
  * *Log on as a batch job* sul server che ospita l'attivita pianificata
* WinRM abilitato sui server target (`Test-WSMan <server>` per verificare)

## Installazione

1. Copia la cartella su un server di gestione (o su un Exchange non produttivo).
2. Crea la tua configurazione partendo dal template e personalizzala
   (`Organization`, `Mail`, `Servers`, soglie):

```powershell
Copy-Item .\ExchangeHealthCheck.config.example.json .\ExchangeHealthCheck.config.json
```

   `ExchangeHealthCheck.config.json` e escluso dal versionamento: contiene FQDN
   dei server, indirizzi e host SMTP dell'ambiente reale. Nel repo resta solo il
   template. Se il file manca, lo script parte sui default con un warning.
3. Verifica la configurazione SMTP:

```powershell
.\Invoke-ExchangeHealthCheck.ps1 -TestMail
```

4. Esegui un giro a vuoto e leggi il CSV prodotto in `Reports\`:

```powershell
.\Invoke-ExchangeHealthCheck.ps1 -NoMail -Verbose
```

5. Registra l'attivita pianificata (richiede una console elevata):

```powershell
.\Install-ExchangeHealthCheckTask.ps1 -IntervalMinutes 15 -UserName 'CONTOSO\svc-exmonitor'
```

Il task parte all'avvio del server e si ripete ogni N minuti; una seconda
schedulazione lo avvia entro due minuti senza aspettare il reboot. Un mutex
globale impedisce esecuzioni sovrapposte se un giro dura piu dell'intervallo.

Per rimuoverlo: `.\Install-ExchangeHealthCheckTask.ps1 -Unregister`

## Logica di alerting

Lo stato di ogni anomalia e persistito in `State\alert-state.json` con chiave
`Categoria|Server|Oggetto`. Ad ogni esecuzione:

* **nuova anomalia** → mail immediata
* **anomalia peggiorata** (es. Warning → Critical) → mail immediata
* **anomalia gia nota** → nessuna mail finche non scade `CooldownMinutes`
  (default 120 minuti), poi un promemoria
* **anomalia rientrata** → mail di rientro con la durata del disservizio
* **nessuna anomalia** → una mail di heartbeat ogni `HeartbeatHours` (default 24),
  per sapere che il monitor e vivo

Con un intervallo di 15 minuti e cooldown di 120, un disco pieno genera una mail
subito e poi una ogni 2 ore, non 8 all'ora.

Se lo script stesso va in errore invia una mail dedicata: un monitor che muore in
silenzio e peggio di un monitor assente.

Exit code: `0` tutto ok, `1` warning, `2` critical, `3` errore dello script.

## Soglie disco

Il controllo disco usa due soglie insieme, in modalita `And`: l'allarme scatta
solo se **sia** la percentuale **sia** i GB liberi sono sotto soglia. Serve a
evitare che un volume da 4 TB al 10% libero (400 GB) generi un critical inutile.

Per i volumi dove conta solo il valore assoluto si usa `VolumeOverrides` con
`"Mode": "Or"` e le percentuali a 0:

```json
{
  "ServerPattern": "*",
  "VolumePattern": "*ExchangeVolumes*",
  "Mode": "Or",
  "FreeGBWarning": 150,
  "FreeGBCritical": 75,
  "FreePercentWarning": 0,
  "FreePercentCritical": 0
}
```

**L'ordine conta**: vince il primo override che matcha. `VolumePattern` viene
confrontato con la stringa `<percorso> <lettera> <etichetta>`, quindi intercetta
anche le mount point senza lettera di unita (`C:\ExchangeVolumes\Vol1\`).

## Ridurre il rumore

* `Ignore.Services` — servizi in Auto ma legittimamente fermi (POP3, IMAP4...)
* `Ignore.HealthSets` — health set di Managed Availability notoriamente rumorosi
* `Ignore.Volumes` — volumi da non controllare (System Reserved, Recovery)
* `Ignore.Databases` — pattern di database da escludere (recovery DB, test)
* `Ignore.Keys` — silenzia un singolo controllo, con wildcard:
  `"Disk|EX-ARCHIVE-01|D:"`, `"Cpu|*|LoadPercent"`
* `Alerting.MinimumSeverityToMail` — `Warning` (default) o `Critical`

Nota sulla memoria: Exchange usa quasi tutta la RAM per la cache dello Store, per
progetto. Le soglie di default sono volutamente basse (6% / 3%); alzarle solo
dopo aver osservato i valori reali.

## Troubleshooting

### "object '*\SERVER' could not be found on <domain controller>"

Le ricerche AD della sessione stanno finendo su un domain controller che non vede
l'oggetto: tipicamente un DC di **un altro dominio della foresta o di un dominio
trusted**. Non e un problema del server Exchange, e di contesto AD.

Nell'ordine, in `Organization`:

1. **`"ViewEntireForest": false`** — la prima cosa da provare. Se
   l'organizzazione Exchange vive in un solo dominio, la visione dell'intera
   foresta non serve ed e proprio cio che allarga la ricerca ai DC sbagliati.
2. **`"PreferredDomainController": "dc01.contoso.local"`** — fissa il DC del
   dominio corretto per tutta la sessione (`Set-ADServerSettings -PreferredServer`).
3. **`"PreferredGlobalCatalog": "gc01.contoso.local"`** — se il problema e sul
   global catalog.

All'avvio il log riporta il contesto applicato: `Contesto AD: ViewEntireForest=True, ...`

Il check interessato fallisce da solo come finding `Unknown` senza fermare il
resto, ma finche il contesto AD e sbagliato quei controlli non producono dati.

### "Nessun DAG rilevato sui server selezionati"

Non e un errore: e il comportamento corretto su server standalone. I check DAG,
copie e `Test-ReplicationHealth` vengono semplicemente saltati; tutto il resto
(servizi, dischi, database, code, certificati) viene eseguito normalmente.

### Sembra bloccato dopo il messaggio sul DAG

Subito dopo viene eseguito il check dei database. Le interrogazioni sono limitate
ai server del perimetro e ai soli server che hanno risposto, proprio per evitare
attese: i cmdlet Exchange non hanno timeout, quindi un `Get-MailboxDatabase
-Status` su un server irraggiungibile resta appeso finche l'RPC non cede.

Se rallenta ancora, con `-Verbose` vedi a che punto e (`Recupero database di
<server>...`) e puoi isolare il check:

```powershell
.\Invoke-ExchangeHealthCheck.ps1 -Check Databases -NoMail -Verbose
```

Per misurare quanto costa davvero sul tuo ambiente:

```powershell
Measure-Command { Get-MailboxDatabase -Server EX-MBX-01 -Status } | Select-Object TotalSeconds
```

Se il tempo e concentrato li, disabilita `Databases` nei giri frequenti e tienilo
in una seconda attivita pianificata piu rada, come per `Replication`.

### "WinRM cannot find the computer ..." su un server che e acceso

E un problema di **risoluzione nome o di WinRM**, non di server giu. Lo script si
connette usando l'FQDN restituito da `Get-ExchangeServer`
(`Servers.UseFqdnForRemoting`, default `true`); se anche cosi fallisce, verifica
nell'ordine, dall'host che esegue lo script:

```powershell
Resolve-DnsName ex01.contoso.local          # il nome si risolve?
Test-NetConnection ex01.contoso.local -Port 5985   # la porta WinRM risponde?
Test-WSMan ex01.contoso.local               # WinRM e attivo e risponde?
Invoke-Command -ComputerName ex01.contoso.local -ScriptBlock { $env:COMPUTERNAME }
```

Cause tipiche, in ordine di frequenza:

* **il nome corto non si risolve** perche l'host di monitoraggio ha un suffisso
  DNS diverso: risolto usando l'FQDN (comportamento di default);
* **WinRM non abilitato** sul server target: `Enable-PSRemoting -Force` oppure
  `winrm quickconfig` sul server;
* **firewall**: TCP 5985 (HTTP) chiuso tra host di monitoraggio e server;
* **host non joinato al dominio**: Kerberos non funziona, servirebbero TrustedHosts
  o HTTPS. Esegui lo script da una macchina in dominio;
* **permessi**: l'account non e amministratore locale sul server target (l'errore
  in quel caso parla di accesso negato, non di computer non trovato).

Se in quell'ambiente funziona solo il nome corto, imposta
`"Servers": { "UseFqdnForRemoting": false }`.

Il log indica sempre **quale** server ha fallito: `Errore remoto su <server>: ...`

## Uso interattivo

```powershell
# solo dischi e code, senza inviare nulla
.\Invoke-ExchangeHealthCheck.ps1 -Check Disk,Queues -NoMail

# un solo server
.\Invoke-ExchangeHealthCheck.ps1 -Server EX-MBX-03 -NoMail -Verbose

# forza l'invio ignorando il cooldown
.\Invoke-ExchangeHealthCheck.ps1 -ForceMail

# usa i risultati in pipeline
.\Invoke-ExchangeHealthCheck.ps1 -NoMail -PassThru | Where-Object Severity -eq 'Critical'
```

## Cosa serve perche la mail parta

Lo script non ha un motore SMTP proprio: consegna il messaggio a un server che lo
accetta. Serve quindi, in ordine:

1. **Raggiungibilita di rete**: TCP 25 (o 587) aperto dall'host che esegue il task
   verso i server in `SmtpServers`, e risoluzione DNS dei nomi indicati.
   Verifica rapida: `Test-NetConnection ex01.contoso.local -Port 25`.
2. **Un receive connector che accetti il messaggio**. Qui il caso cambia a seconda
   dei destinatari:
   * **destinatari interni** (il caso normale: la mailbox del team IT): il
     connector *Default Frontend `<SERVER>`* presente di serie su ogni Mailbox
     server accetta gia submission anonime dirette a un accepted domain.
     **Non serve configurare nulla.**
   * **destinatari esterni** (un indirizzo di reperibilita, un SMS gateway): e
     relay, e il relay anonimo e negato per progetto. Due strade: autenticare con
     un account (vedi sotto), oppure creare un receive connector dedicato con
     `RemoteIPRanges` limitato all'IP dell'host di monitoraggio e il permesso
     `ms-Exch-SMTP-Accept-Any-Recipient` concesso a `NT AUTHORITY\ANONYMOUS LOGON`.
3. **Un indirizzo mittente**. Tecnicamente una submission anonima passa anche con
   un `From` inventato, ma conviene creare una mailbox o almeno un utente
   mail-enabled reale: altrimenti gli NDR non tornano a nessuno e le regole
   anti-spoofing possono scartare un mittente interno che arriva da un connector
   anonimo.
4. **Niente**, se `Mail.Enabled` e `false`: in quel caso restano log e CSV.

### Nota su TLS

`UseSsl: true` in `System.Net.Mail.SmtpClient` significa **STARTTLS**, non TLS
implicito: la **porta 465 non funziona**, usa 25 o 587. Il certificato del server
deve inoltre essere valido per il nome usato in `SmtpServers` e attendibile
dall'host di monitoraggio, altrimenti l'invio fallisce.

### Se l'invio fallisce

I server in `SmtpServers` vengono provati in ordine finche uno accetta. Se
falliscono tutti, la notifica **non viene registrata come inviata**: lo stato
riporta indietro `LastNotified` e il giro successivo riprova, invece di far
entrare l'anomalia in cooldown senza che nessuno l'abbia letta. Lo stesso vale per
le mail di rientro.

## SMTP autenticato

Se il relay richiede credenziali, salvale una volta **con lo stesso account che
esegue il task** (la cifratura DPAPI e legata all'utente):

```powershell
Get-Credential | Export-Clixml -Path .\State\smtp.cred.xml
```

poi in configurazione: `"CredentialFile": "State\\smtp.cred.xml"`.

`SmtpServers` accetta piu host: vengono provati in ordine finche uno accetta il
messaggio. Metterne almeno due evita di perdere l'alert proprio quando il primo
server e quello che ha il problema.

## Prestazioni

La raccolta dati OS e un unico fan-out `Invoke-Command` su tutti i server
(`Thresholds.ThrottleLimit`, default 24 in parallelo). I check Exchange-side sono
invece sequenziali per server: `Test-ReplicationHealth` e il piu lento (~5-15s per
server). Su molti server, se il giro si avvicina all'intervallo di schedulazione,
conviene disabilitare `Replication` nei giri frequenti e tenerlo in una seconda
attivita pianificata oraria con `-Check Replication`.
