# Handoff tecnico — Exchange Health Check

> Questo documento serve a un altro modello (o a un ingegnere umano) per riprendere
> il progetto senza dover rileggere l'intera cronologia di sviluppo. Descrive
> l'architettura, le scelte di design con la loro motivazione, l'elenco dei bug
> reali trovati durante il test su un ambiente Exchange vero con la loro causa e
> il fix, e i punti che restano da verificare. Non contiene alcun dato del
> cliente (nomi server, domini, indirizzi): il repository è **pubblico**.

**Repository:** https://github.com/diepo/ExchangeHealthCheck
**Linguaggio:** PowerShell 5.1+ (compatibile anche con PowerShell 7.x)
**Stato:** funzionante, in test attivo su un ambiente Exchange on-premises reale
(2 server di un DAG, verifica in corso su altri).

---

## 1. Cosa fa e perché esiste

Health check periodico per Exchange on-premises multi-server: servizi, spazio
disco (incluse le mount point), stato del DAG, copie di database, code di
trasporto, certificati, Managed Availability. Alerting via e-mail con
deduplica — non manda una mail per ogni giro su ogni anomalia già nota, solo su
novità, peggioramenti, rientri, e un heartbeat periodico se tutto tace.

Pensato per girare **su un server Exchange stesso** (o su un management server
con le remote PowerShell), lanciato a intervalli regolari da uno scheduled task.

Principio guida seguito in ogni scelta di design: **sola lettura**. Solo
cmdlet `Get-*` e `Test-*` verso Exchange. Le uniche scritture sono locali alla
cartella dello script (log, CSV, stato) più l'invio delle mail.

## 2. Struttura del repository

```
Invoke-ExchangeHealthCheck.ps1           # script principale, ~1900 righe, tutto in un file
ExchangeHealthCheck.config.example.json  # template di configurazione, versionato
ExchangeHealthCheck.config.json          # config reale dell'utente, NON versionato (.gitignore)
Install-ExchangeHealthCheckTask.ps1      # registra lo scheduled task
README.md                                # guida operativa (installazione, troubleshooting)
HANDOFF.md                               # questo file
GUIDE.md / ExchangeHealthCheck-Guide.pdf # guida utente in formato documento
.gitignore                               # esclude config.json, Logs/, Reports/, State/, *.cred.xml
```

Lo script è deliberatamente un unico file. È stata una scelta esplicita per
mantenere la distribuzione a un solo comando di copia, senza gestione di moduli
o percorsi relativi tra file — il costo è un file lungo, organizzato in
`#region` con intestazioni chiare (CONFIG, LOGGING, FINDINGS, EXCHANGE ACCESS,
i vari `Invoke-Hc*Check`, ALERT STATE, MAIL, MAIN).

## 3. Architettura

### 3.1 Configurazione: merge ricorsivo sui default incorporati

Lo script porta con sé una copia completa dei valori di default
(`$DefaultConfigJson`, in cima al file). Il file JSON dell'utente viene letto
e sovrapposto ricorsivamente (`Merge-HcConfig`), proprietà per proprietà — non
sostituito. Questo significa:

- il file utente può essere **parziale**: due righe bastano per cambiare una
  soglia, il resto prende il default;
- se il file manca, lo script parte comunque sui default con un warning;
- **eccezione importante**: gli **array** vengono sostituiti per intero, non
  concatenati. Se l'utente ridefinisce `Ignore.Services`, la sua lista
  rimpiazza quella di default — non ci si aggiunge.

Questa scelta ha causato un bug reale (vedi §5, "ViewEntireForest ignorato a
false"): un valore booleano `false` esplicito nel file utente, se il codice
applica l'impostazione solo quando è `true`, produce l'illusione che la
configurazione non abbia effetto. **Lezione**: quando un valore di
configurazione controlla uno stato binario di un sistema esterno (qui,
un'impostazione di sessione PowerShell/AD), il codice deve *impostare sempre*
il valore risolto, mai limitarsi ad attivarlo quando è true.

### 3.2 Selezione del perimetro (`Get-HcTargetServer`)

Parte sempre da `Get-ExchangeServer` (i server realmente registrati
nell'organizzazione). `Servers.Include`/`Exclude` **filtrano** quel risultato,
non lo estendono — un nome inventato in `Include` produce silenziosamente zero
server selezionati, non un errore. Supportano wildcard e sono confrontati sia
sul nome corto sia sull'FQDN.

### 3.3 Raccolta dati OS: fan-out remoto + eccezione per il nodo locale

`Get-HcRemoteData` fa un unico `Invoke-Command` in fan-out su tutti i server
del perimetro per raccogliere OS, volumi, servizi, memoria, reboot pending.
Punti di design rilevanti:

- **Timeout espliciti** su `New-PSSessionOption` (`RemoteOpenTimeoutSeconds`,
  default 20s; `RemoteOperationTimeoutSeconds`, default 120s). Senza, un
  server spento occupa uno slot del fan-out per il timeout WinRM di default
  (~45s), moltiplicato per ogni server morto.
- **Il server locale è un caso speciale** (§5, bug "loopback WinRM"): se lo
  script gira su un Exchange server, quel server viene esclus dal fan-out
  `Invoke-Command -ComputerName` e interrogato **eseguendo lo scriptblock
  direttamente in sessione** (`& $scriptBlock ...`). Una connessione WinRM
  verso il proprio nome richiede comunque autenticazione Kerberos, SPN
  corretti e supera il *loopback check* di sicurezza di Windows: fallisce per
  motivi che non hanno nulla a che vedere con la salute della macchina, e il
  fallimento colpiva proprio il server più sano di tutti (quello da cui si
  lanciava lo script).
- **FQDN per il remoting** (`Servers.UseFqdnForRemoting`, default true):
  `Get-ExchangeServer` restituisce il nome corto, che può non risolvere se
  l'host di monitoraggio ha un suffisso DNS diverso dai server Exchange. La
  mappa dei risultati è indicizzata sia sull'FQDN sia sul nome corto sia su
  `ComputerName`, così il resto del codice continua a cercare per nome corto
  senza saperlo.

### 3.4 Interrogazione AD/DC: rilevamento automatico

`Set-HcAutoDomainController` deduce il dominio dei server Exchange dal
suffisso FQDN più diffuso tra i target, e chiede ad Active Directory (via
`System.DirectoryServices.ActiveDirectory.DomainController]::FindOne()`, non
il modulo RSAT) un domain controller **che risponde in quel momento**. La
ricerca si ripete a ogni esecuzione: non fissa mai un DC che domani potrebbe
essere spento. `Organization.PreferredDomainController`, se valorizzato, ha
sempre la precedenza e disattiva l'automatismo per quella sessione.

Motivazione: in foreste multi-dominio, `ViewEntireForest` può far risolvere le
query AD su un domain controller di un dominio trusted che non replica il
Configuration NC di Exchange, con errori "object ... could not be found" che
sembrano un guasto ma sono un problema di contesto AD.

### 3.5 Ogni check è isolato (`Invoke-HcCheck`)

Ogni singolo controllo (per categoria, per server) gira dentro un
`try/catch`: un cmdlet che fallisce su un server diventa un finding di
severità `Unknown`, non interrompe il giro. Solo un errore nella fase di
bootstrap (config illeggibile, impossibile connettersi a Exchange) ferma
l'esecuzione — e in quel caso parte comunque una mail di errore dedicata,
perché un monitor che muore in silenzio è peggio di un monitor assente.

### 3.6 Macchina a stati degli alert (`Resolve-HcAlert`)

Il cuore dell'anti-mail-storm. Stato persistito in `State\alert-state.json`
con chiave `Categoria|Server|Oggetto`. Ad ogni giro, per ogni finding con
severità sopra soglia:

- **non presente nello stato precedente** → nuovo alert, notifica immediata;
- **presente, severità peggiorata** (es. Warning→Critical) → notifica
  immediata, a prescindere dal cooldown;
- **presente, stessa severità o migliorata ma ancora sopra soglia** →
  notifica solo se il cooldown (`Alerting.CooldownMinutes`, default 120) è
  scaduto dall'ultima notifica;
- **presente nello stato precedente ma non più nei finding correnti** →
  "rientrato", con la durata calcolata da `FirstSeen`.

**Punto critico corretto durante lo sviluppo**: il conteggio dell'invio non
deve mai essere ottimistico. Se `Send-HcMail` ritorna `false` (SMTP giù, host
irraggiungibile, credenziali sbagliate...), lo stato **non viene aggiornato
come notificato** — altrimenti l'anomalia entrerebbe in cooldown per due ore
senza che nessuno l'abbia letta, il worst case possibile per un sistema di
allerta. Il codice riporta `LastNotified` al valore precedente e rimette i
rientri nello stato se l'invio fallisce, così il giro successivo riprova.

### 3.7 Mail: HTML + fallback su più host SMTP

`Send-HcMail` usa `System.Net.Mail.SmtpClient` (non `Send-MailMessage`,
deprecato). `Mail.SmtpServers` è una lista provata in ordine: il primo host che
accetta il messaggio vince. Motivazione: la mail di alert esce **attraverso
l'Exchange che si sta monitorando** — un solo host configurato significa
perdere l'alert proprio quando quell'host ha un problema.

Il corpo HTML (`New-HcMailBody`) ha tre livelli di vista, deliberatamente
sovrapposti e non alternativi:
1. contatori aggregati (Critical/Warning/Info/OK) su tutto l'ambiente;
2. tabelle per categoria di evento (Nuove/Peggiorate/Ancora aperte/Rientrate),
   una riga per finding con colonna Server, ordinate per severità decrescente;
3. riepilogo per server in fondo (chi sta male, a colpo d'occhio) e una
   sezione aggregata per le code di trasporto (vedi §3.9).

La stessa vista aggregata delle code è disponibile anche a console
(`Write-HcQueueSummaryConsole`), per non doverla vedere solo in mail durante un
giro a secco (`-NoMail`).

### 3.8 Logging a console: colore per severità, non per livello di log

I finding sono l'unità di informazione principale, non le righe di log
generiche. `Add-Finding` sceglie colore console (rosso Critical, magenta
Unknown, giallo Warning, ciano Info, verde OK — fattorizzato in
`Get-HcConsoleColor`, riusato anche dal riepilogo code) e decide se la riga
compare senza `-Verbose` (solo severità ≥ Warning) o solo con `-Verbose`
(tutto, incluso OK). Prima di questa scelta, tutto passava dal generico
Write-Host/Write-Verbose di PowerShell, che colora in giallo anche le righe
`-Verbose`: Critical, Warning e diagnostica erano visivamente indistinguibili.

### 3.9 Vista aggregata delle code

`Invoke-HcQueueCheck` non produce solo i finding per soglia superata: popola
sempre (`$script:QueueSummary`, una lista a livello di script) una fotografia
per server — totale in coda, code attive, submission, poison, code in retry,
il totale delle code shadow (tenuto separato, mai sommato al totale: le shadow
trattengono messaggi per progetto e non sono un'anomalia), e la "coda
maggiore" (`NextHopDomain` col conteggio più alto, escludendo submission e
poison che non hanno un destinatario esterno significativo).

**`NextHopDomain` può essere un IP** (smart host configurato per indirizzo
anziché FQDN, o consegna diretta). `Get-HcNextHopLabel` prova una risoluzione
PTR (`System.Net.Dns.BeginGetHostEntry` con timeout esplicito via
`AsyncWaitHandle.WaitOne`, mai la forma sincrona bloccante) e, se risolve,
affianca l'hostname tra parentesi quadre senza nascondere l'IP. Se il PTR non
risponde, prova un **fallback NetBIOS** (`Resolve-HcNetBiosName`, `nbtstat -A`
con lo stesso pattern di timeout esplicito e uccisione del processo): su reti
dove le zone di reverse lookup DNS non sono tenute aggiornate, `ping -a` risolve
comunque il nome perché il resolver di Windows prova anche NetBIOS — `System.Net.Dns`
da solo no. Cache per IP a livello di sessione per entrambi i meccanismi (la
stessa destinazione ricorre su più code/server). **Importante**: la chiave di
deduplica degli alert (`Item`) resta sempre l'IP grezzo — solo il messaggio
testuale mostra l'hostname risolto, perché una risoluzione DNS/NetBIOS può
cambiare da un giro all'altro e non deve far perdere lo stato di un alert.

`Get-Queue` espone già il GUID del send connector in uso per ogni coda
(`NextHopConnector`): `Get-HcSendConnectorName` lo risolve con
`Get-SendConnector -Identity <guid>` invece di indovinarlo dal dominio o dagli
`AddressSpaces`. Per la consegna interna (via DAG/database, es. `NextHopDomain`
= nome del DAG) quel GUID non corrisponde a un send connector reale: la
ricerca fallisce con grazia e non si mostra nulla in più, invece di inventare
un nome. L'etichetta finale è composta da `Get-HcQueueDestinationLabel`, che
unisce risoluzione IP e nome del connector.

### 3.10 Stato aggregato per categoria e mail dedicata alle novità

Oltre al riepilogo "per server" (chi sta male), esiste un riepilogo **"per
categoria"** (`Get-HcCategorySummary`, con rendering `New-HcCategorySummaryTable`
per la mail e `Write-HcCategorySummaryConsole` per la console): quante cose
sono Critical/Warning/Unknown/Info in ciascuna categoria (Disk, Service, Queue,
Certificate, ...), ordinate per gravità. Risponde a "cosa non va
nell'infrastruttura" a colpo d'occhio, senza dover scorrere ogni server. In
mail è posizionato subito dopo i contatori aggregati, prima delle tabelle di
dettaglio.

**Conta occorrenze distinte, non ogni server** (bug reale, corretto dopo il
primo test su ambiente vero: un giro mostrava "25" per `ManagedAvailability`
quando i problemi distinti erano solo 3, ripetuti su più server dello stesso
DAG). La deduplica raggruppa per **`Item` + testo del `Message`**, non sul solo
`Item`: un health set Unhealthy, un witness irraggiungibile o un certificato in
scadenza vengono spesso rilevati **identici** su più server (stessa causa,
stesso messaggio) e vanno contati una volta sola. Ma `Item` da solo non basta
come chiave — due dischi `C:` pieni su server diversi condividono l'etichetta
pur essendo due problemi realmente distinti (GB liberi reali diversi nel
messaggio), e la stessa cosa vale per `Queue` (`TotalMessages` è lo stesso Item
su ogni server, ma il messaggio riporta il conteggio reale di quel server).
Deduplicare sul solo `Item` li avrebbe fatti sparire per errore.

**Due canali di notifica distinti, con scopi diversi**:

1. **Mail di riepilogo** (`New-HcMailBody`, invariata nella sua logica di invio):
   copre sempre tutto il quadro — nuove, peggiorate, promemoria, rientrate,
   riepilogo per categoria e per server, code. È l'unica il cui esito
   determina lo stato degli alert (§3.6).
2. **Mail dedicata alle sole novità** (`Get-HcUrgentFindings` +
   `New-HcUrgentAlertBody`), aggiunta su richiesta esplicita dell'utente:
   parte in più, con oggetto distinto (default `WARNING FOUND`,
   `Mail.SeparateAlertSubjectTag`) e corpo ridotto alle sole righe rilevanti.
   Non scatta mai su un promemoria (anomalia già nota, in cooldown) — solo su
   `$alerts.New` o `$alerts.Escalated` di `Resolve-HcAlert`. Regola di innesco,
   decisa esplicitamente dall'utente per evitare rumore:
   - qualunque categoria diversa da `Queue`: solo un nuovo `Critical` (o un
     `Warning` che diventa `Critical`) la fa scattare;
   - categoria `Queue`: si guarda il **valore numerico** del finding (messaggi
     in coda), non la sua severità — una coda che passa da 15 a 18 non deve
     generare nulla; serve superare `Thresholds.QueueSubjectThreshold` (default
     200), a prescindere che sia già `Warning` o `Critical`.

   Il suo esito (inviata o no) **non tocca lo stato degli alert**: la mail di
   riepilogo resta l'unica fonte di verità per il cooldown, così un fallimento
   di questo canale aggiuntivo non altera la logica anti-mail-storm già
   verificata in §3.6.

**Oggetto della mail di riepilogo**: guadagna un tag aggiuntivo, indipendente
da CRITICO/WARNING, quando il totale in coda di un qualunque server raggiunge
`Thresholds.QueueSubjectThreshold` (la stessa soglia usata sopra): default
`ATTENZIONE CODE` (`Mail.QueueAlertSubjectTag`). Calcolato sulla fotografia
corrente delle code (`$script:QueueSummary`), non sul solo insieme delle
notifiche di questo giro: resta visibile anche se quella coda è già nota e in
cooldown — a differenza della mail dedicata alle novità, che invece su una
coda già nota non scatterebbe.

## 4. Schema di configurazione (riferimento completo)

Vedi `ExchangeHealthCheck.config.example.json` per i valori concreti. Sezioni:

| Sezione | Contenuto |
|---|---|
| `Organization` | Nome, server a cui connettersi (`ConnectTo`, usato solo se non si è già in Exchange Management Shell), `ViewEntireForest`, rilevamento/override del domain controller |
| `Servers` | Perimetro (`Include`/`Exclude`/`SiteFilter`/`IncludeEdge`), `UseFqdnForRemoting`, `SkipExchangeChecksWhenOffline` |
| `Checks` | Un booleano per famiglia di controllo (Os, Disk, Services, Components, Health, Dag, Replication, Databases, Queues, BackPressure, Certificates, Mapi) |
| `Thresholds` | Tutte le soglie numeriche: disco (con `DiskMode` And/Or), memoria, CPU, code, copy/replay queue del DAG, età backup, scadenza certificati, timeout di rete, `QueueSubjectThreshold` (soglia condivisa tra il tag "ATTENZIONE CODE" in oggetto e l'innesco della mail dedicata alle novità) |
| `VolumeOverrides` | Soglie disco per pattern di server/volume, con precedenza sul primo match |
| `HealthReport` | `IncludeFailingMonitors` (arricchisce l'alert con i monitor Managed Availability in errore), `MaxMonitorsPerHealthSet` |
| `Queues` | `ResolveNextHopHostnames`, `ReverseDnsTimeoutMs`, `TryNetBiosFallback`, `NetBiosTimeoutMs`, `ResolveSendConnectorName` |
| `Console` | `ShowCategorySummary`, `ShowQueueSummary` |
| `Ignore` | Liste di esclusione: Services, ServerComponents, HealthSets, Volumes, Databases, Keys (pattern esatto `Categoria\|Server\|Oggetto`, con wildcard) |
| `ExtraServices` | Servizi non-Exchange da includere nel check Services (es. W3SVC, WinRM) |
| `Alerting` | Cooldown, heartbeat, notifica di rientro, severità minima da notificare |
| `Mail` | SMTP, autenticazione, mittente/destinatari, allegato CSV, vista code, `QueueAlertSubjectTag`, `SeparateAlertSubjectTag` (vedi §3.10) |
| `Paths` | Cartelle di log/report/stato, retention |

## 5. Bug reali trovati durante il test su ambiente vero — con causa e fix

Questo è probabilmente l'elenco più utile del documento: sono tutti bug emersi
**testando lo script contro un Exchange reale**, non ipotesi. Ogni riga
riporta sintomo osservato → causa reale → correzione. Commit nel repository
pubblico, cronologici.

| # | Sintomo osservato | Causa reale | Fix |
|---|---|---|---|
| 1 | `WinRM cannot find the computer` su server perfettamente attivi | `Get-ExchangeServer` restituisce il nome corto; l'host di monitoraggio aveva un suffisso DNS diverso e Kerberos non risolveva | Remoting via FQDN, mappa risultati indicizzata su più chiavi |
| 2 | Lo script sembrava bloccato dopo il check del DAG | `Get-MailboxDatabase -Status` **senza** `-Server`, con `ViewEntireForest` attivo, enumera l'intera foresta e contatta ogni server proprietario per lo stato di mount; i cmdlet Exchange non hanno timeout | Interrogazione per singolo server, solo su quelli online, con deduplica per GUID (un DB con più copie DAG comparirebbe più volte) |
| 3 | Health set con nome vuoto (`Health set "" Unhealthy`) | Uso delle proprietà di `Get-ServerHealth` (`HealthSetName`/`FirstAlertObservedTime`) su un oggetto di `Get-HealthReport`, che le espone con nomi diversi (e variabili tra versioni Exchange) | `Get-HcFirstValue` prova più nomi candidati; se nessuno risponde, logga lo schema reale dell'oggetto invece di mostrare un placeholder |
| 4 | `object '*\SERVER' could not be found on <DC>` nonostante il DC fosse della foresta giusta | **Due cause distinte, sovrapposte**: (a) `ViewEntireForest: false` in config non veniva applicato perché il codice lo impostava solo quando era `true`; (b) `Get-MailboxDatabaseCopyStatus -Server X` risponde "could not be found" anche quando il server semplicemente non ospita copie di database, o quando l'oggetto è fuori dallo scope RBAC — non è necessariamente un guasto | `ViewEntireForest` impostato sempre in entrambe le direzioni; l'errore "could not be found" su quel cmdlet diventa un finding Info, non Unknown |
| 5 | Back pressure segnalata come attiva su ogni server, sempre | Campi del componente `ResourceThrottling` letti invertiti: `CurrentResourceUse` è lo stato (Low/Medium/High), `Pressure` è la misura numerica. Il codice confrontava `Pressure` con la stringa `'Normal'`, sempre vera | Legge `CurrentResourceUse`; `Low` (stato sano) non genera alert, solo Medium/High |
| 6 | `ForwardSyncDaemon`/`ProvisioningRps` sempre Inactive su ogni server | Sono componenti usati solo dal datacenter Microsoft (Exchange Online); su on-premises sono Inactive per progetto | Aggiunti alle esclusioni di default (`Ignore.ServerComponents`) |
| 7 | `error opening cluster DAG1` nonostante il DAG fosse sano (8/8 membri attivi via Exchange) | Un DAG creato senza Administrative Access Point (default sui DAG moderni) non ha IP né nome di cluster risolvibile | `Get-ClusterNode` tentato prima sul nome del DAG poi su un nodo membro; se nessuno risponde il finding è Info (lo stato dei membri è già verificato dai cmdlet Exchange), non un allarme |
| 8 | Il server locale (quello da cui gira lo script) risultava `Connectivity Critical` | `Invoke-Command` verso il proprio nome apre comunque una connessione WinRM di loopback, soggetta ad autenticazione Kerberos, SPN e al loopback security check di Windows | Il target che coincide col computer locale viene eseguito in sessione diretta, non via WinRM |
| 9 | Due certificati pubblici con lo stesso nome, uno scaduto, causa probabile di `Transport.ServerCertMismatch` | I connector referenziano il certificato per `TlsCertificateName` (`<I>Issuer<S>Subject`), non per thumbprint: due certificati con stesso Subject e Issuer sono indistinguibili per Exchange, che può agganciare quello sbagliato | Il check certificati segnala i duplicati con stesso Subject+Issuer (elencando i thumbprint) e lo Status diverso da `Valid` sui certificati assegnati a servizi |
| 10 | Test mail fallito con solo "Failure sending mail", nessuna causa utile | `SmtpClient` incapsula la causa reale nelle `InnerException` | Il log risale tutta la catena di eccezioni; `-TestMail` stampa prima dell'invio i parametri effettivi in uso |
| 11 | Riepilogo per categoria mostrava "25" per `ManagedAvailability` con solo 3 problemi distinti | La stessa anomalia (health set Unhealthy, witness irraggiungibile, certificato in scadenza) viene spesso rilevata identica su più server dello stesso DAG; sommare ogni finding conta una volta per server invece che una volta per problema | Deduplica per `Item`+`Message` insieme (non sul solo `Item`, che avrebbe fatto sparire per errore dischi/code realmente diversi con la stessa etichetta) |

## 6. Verificato vs. non verificato contro Exchange reale

Per trasparenza verso chi riprende il progetto: **la maggior parte della
logica pura è stata testata con unit test basati su mock** (funzioni Exchange
sostituite con stub PowerShell che restituiscono oggetti con lo schema atteso),
non contro un ambiente reale, per la semplice ragione che l'agente di sviluppo
non ha accesso a un Exchange server. Sono stati verificati con mock: soglie
disco, macchina a stati degli alert, generazione HTML della mail, ogni singolo
check Exchange-side con dati costruiti a mano.

**Sono stati verificati contro un Exchange reale**, dall'utente, in sessioni
successive: raggiungibilità WinRM/FQDN, enumerazione DAG e database, Managed
Availability (nomi degli health set, che sono infatti quelli con più bug: §5
righe 3-4), back pressure, stato dei componenti, cluster del DAG, certificati.

**Conseguenza pratica per chi continua lo sviluppo**: quando un cmdlet
Exchange nuovo o poco usato viene collegato per la prima volta, è quasi
garantito che qualche nome di proprietà sia diverso da quello atteso (è successo
due volte: Managed Availability e ResourceThrottling). Il pattern difensivo già
in uso (`Get-HcFirstValue` con più candidati, log dello schema reale se nessuno
risponde) va replicato per ogni nuovo cmdlet, non assunto per scontato dal
primo tentativo.

## 7. Aree ancora aperte / da tenere d'occhio

- **Il check `Mapi`** (`Test-MapiConnectivity`) è disabilitato di default e non
  è mai stato testato contro un ambiente reale in questo sviluppo — è l'unico
  check che esegue un logon reale (system mailbox), va verificato con
  attenzione prima di attivarlo su produzione.
- **Il costo di `Test-ReplicationHealth`** (5-15s per server, sequenziale) non
  è stato misurato su più di 2 server reali. Su ambienti con decine di server
  può avvicinarsi all'intervallo di schedulazione; il README suggerisce di
  separarlo in un secondo scheduled task più rado.
- **La risoluzione PTR** (`Get-HcNextHopLabel`) è stata verificata solo in un
  ambiente sandbox senza uscita di rete (fallback confermato, ma la
  risoluzione vera non è mai stata osservata con successo). Da confermare che
  l'hostname compaia davvero su un ambiente con DNS funzionante.
- **Versioni Exchange**: lo sviluppo e i fix di schema sono stati fatti contro
  un ambiente reale di cui non è nota la versione esatta (probabilmente
  Exchange 2019). Nomi di proprietà possono differire su 2016/2013.
- **Il pattern feature-comment del CSV** (colonne `Timestamp, Severity, Server,
  Category, Item, Value, Message`) non ha subito modifiche da quando è stato
  scritto: nessun bug noto lì, ma nemmeno un test dedicato oltre alla verifica
  di sintassi `Export-Csv`.

## 8. Convenzioni di sviluppo osservate finora

- Ogni fix è stato **verificato con un test mirato prima del commit** (mock
  della funzione Exchange coinvolta, o esecuzione diretta dello script con
  `-NoMail`), non solo controllato a vista.
- Ogni commit ha un messaggio esteso che spiega *perché* (causa del bug), non
  solo *cosa* è cambiato — pensato per essere letto senza il contesto della
  conversazione originale.
- Nessun nome di server, dominio o indirizzo reale del cliente è mai stato
  scritto in un commit, in un file versionato, o in questo documento. Gli
  esempi usano sempre nomi generici (`EX-MBX-01`, `contoso.local`, `azienda.com`).
- La configurazione reale (`ExchangeHealthCheck.config.json`) non è mai
  versionata; solo il template lo è.

## 9. Se riprendi questo progetto: da dove iniziare

1. Leggi `README.md` per la prospettiva utente (installazione, troubleshooting
   già scritto per i problemi più comuni).
2. Leggi questo file per il perché delle scelte di design.
3. Se devi aggiungere un nuovo check Exchange-side: usa `Invoke-HcCheck` come
   wrapper per l'isolamento degli errori, `Add-Finding` per registrare i
   risultati, e **non fidarti dei nomi di proprietà della documentazione
   Microsoft** senza un modo per verificarli o degradare con grazia (vedi
   `Get-HcFirstValue` e il logging dello schema in caso di proprietà non
   trovata, §5 riga 3).
4. Se tocchi la macchina a stati degli alert (`Resolve-HcAlert`) o l'invio
   mail (`Send-HcMail`), ricontrolla sempre il caso "invio fallito": lo stato
   non deve mai registrare una notifica che non è realmente partita.
5. **Mantieni aggiornati `HANDOFF.md` e la guida utente (`GUIDE.md` /
   `ExchangeHealthCheck-Guide.pdf`) ad ogni modifica funzionale**: aggiungi la
   riga alla tabella §5 se hai corretto un bug, aggiorna §4 se hai aggiunto una
   chiave di configurazione, rigenera il PDF dalla guida.
