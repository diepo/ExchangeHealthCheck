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

### 3.11 Timeout duro sui cmdlet RPC senza timeout proprio

`Get-ExchangeCertificate`, `Get-HealthReport` e `Get-ServerHealth` non
espongono alcun parametro di timeout: se il servizio a cui fanno RPC su un
server è lento o inceppato, la chiamata resta appesa finché non risponde —
osservato in produzione, `Get-ExchangeCertificate` fermo per 300,3s (il
timeout RPC di default di Windows) su un server con un problema di backend
IIS, bloccando l'intero giro sugli altri server ancora in coda. Diverso dal
caso WinRM/DNS (§3.3, §3.9), dove un `-SessionOption`/timeout esplicito basta:
questi cmdlet Exchange non hanno un parametro equivalente da passare.

`Invoke-HcExchangeWithTimeout` risolve il problema isolando la chiamata in un
**job separato** (un processo `powershell.exe` indipendente): un job, a
differenza di una chiamata diretta nella sessione corrente, può essere
**terminato con forza** (`Stop-Job`) se supera il limite. Il job deve
riottenere l'accesso a Exchange da solo — parte in un processo nuovo che non
eredita né lo snap-in né la sessione remota già aperti nel processo
principale — rifacendo lì dentro lo stesso identico percorso a due vie di
`Connect-HcExchange` (snap-in locale se registrato, altrimenti
`New-PSSession -ConfigurationName Microsoft.Exchange` verso
`Organization.ConnectTo`).

Applicata a:
- `Get-ExchangeCertificate` → `Thresholds.CertificateCheckTimeoutSeconds`
  (default 30s)
- `Get-HealthReport` → `Thresholds.ManagedAvailabilityTimeoutSeconds`
  (default 30s)
- `Get-ServerHealth` (drill-down dei monitor, §3.10) →
  `HealthReport.MonitorDetailTimeoutSeconds` (default 30s, tenuto separato
  perché questa chiamata può ripetersi una volta per ogni health set non sano
  sullo stesso server, e non deve sommare più timeout lunghi in sequenza)

Il default di 30s non è arbitrario: una risposta sana di `Get-HealthReport`
osservata in produzione impiegava ~5s, quindi 30s lascia margine ampio prima
di dichiarare il servizio non responsivo, restando comunque molto sotto i
300s del timeout RPC di sistema che altrimenti si subirebbe per intero.

**Costo**: ogni chiamata protetta spende qualche secondo in più per l'avvio
del processo del job e il ricaricamento dell'accesso Exchange, anche nel caso
sano — accettabile per due chiamate isolate per server, non pensato per
essere applicato a ogni singolo cmdlet Exchange dello script senza
valutazione caso per caso.

**Fallback quando il job non può autenticarsi da solo (bug 21)**: il job
gira in un processo nuovo che non eredita nulla del processo principale —
se non trova né lo snap-in locale né `Organization.ConnectTo` valorizzato,
non ha alcun modo di accedere a Exchange, a differenza del processo
principale che può avere i cmdlet disponibili tramite una sessione esterna
già aperta (Exchange Management Shell, o `Connect-ExchangeServer`/
`RemoteExchange.ps1` lanciati a mano prima di avviare lo script — proprio il
motivo per cui `Organization.ConnectTo` non serve in quel caso). In quella
configurazione il job falliva **sempre**, per ogni chiamata a
`Get-ExchangeCertificate`/`Get-HealthReport`/`Get-ServerHealth`, con
`Cannot bind parameter 'ConnectionUri'` — non un sintomo di sessione stale
ma un fallimento strutturale e deterministico, che però innescava
inutilmente (e distruttivamente, vedi §3.16) il recupero pensato per
sessioni davvero cadute: `Remove-HcStaleExchangeProxy` rimuoveva il modulo
Exchange condiviso, valido, della sessione esterna, lasciando il resto del
giro a lavorare su una sessione a metà smontata (`Access is denied` sparso
su check che non c'entravano nulla, es. `Get-Queue`). Fix: quando né
snap-in né `ConnectTo` sono disponibili, `Invoke-HcExchangeWithTimeout`
salta del tutto il job e chiama il cmdlet **direttamente nel processo
corrente**, dove i cmdlet sono già disponibili — perdendo la protezione da
timeout duro per queste chiamate specifiche in questa configurazione, ma
evitando un fallimento garantito e la distruzione collaterale della sessione
di lavoro.

### 3.12 Riepiloghi aggregati: cluster per DAG e copie database per server

Stesso principio della vista aggregata delle code (§3.9): una fotografia
raccolta **sempre**, non solo quando c'è un problema, per rispondere a colpo
d'occhio a "come sta l'infrastruttura" senza dover aprire ogni singola
tabella di dettaglio.

- **Stato cluster per DAG** (`$script:ClusterSummary`, popolato dentro
  `Invoke-HcDagCheck`): una riga per DAG con totale nodi, quanti `Up`, quanti
  non `Up` (con nome e stato di ciascuno), stato del witness. Usa la stessa
  fonte dati (`Get-HcClusterNode`) già usata per i singoli finding `Cluster`.
- **Copie database per server** (`$script:DatabaseCopySummary`, popolato
  dentro `Invoke-HcCopyStatusCheck`): una riga per server con totale copie
  ospitate, quante sane, quante non sane (con nome del DB e stato di
  ciascuna). Stessa fonte dati già usata per i finding `DatabaseCopy`.

Renderizzati sia a console (`Write-HcClusterSummaryConsole`,
`Write-HcDatabaseSummaryConsole`, sullo stesso modello di
`Write-HcQueueSummaryConsole`) sia in mail, posizionati subito prima della
vista aggregata delle code — infrastruttura e dati vengono prima del traffico
nell'ordine di lettura. Disattivabili singolarmente
(`Console.ShowClusterSummary`/`ShowDatabaseSummary`,
`Mail.IncludeClusterSummary`/`IncludeDatabaseSummary`).

**Bug di visibilità corretto nello stesso commit**: se `Get-ClusterNode` non è
disponibile del tutto sulla macchina da cui gira lo script (manca il modulo
FailoverClusters, cioè lo strumento RSAT "Failover Clustering Tools"), l'intero
controllo cluster veniva **saltato in silenzio** — nessun errore, nessun
finding, nessuna riga di log visibile. Un nodo cluster `Down` sarebbe passato
inosservato senza che nulla lo segnalasse. Ora l'assenza della capacità
produce essa stessa un finding `Warning`, con l'indicazione esplicita di quale
componente installare (`RSAT-Clustering-PowerShell`). Lezione generale,
coerente con §8: **un `if (Test-HcCommand ...) { ... }` senza un `else` che
segnali l'assenza è un modo comune per far sparire silenziosamente un'intera
famiglia di controlli** — va sempre verificato che manchi un ramo visibile per
il caso "capacità non disponibile", non solo per il caso "controllo eseguito
ma fallito".

### 3.13 Sessione Exchange stale su giri lunghi

Segnalato dall'utente: un errore che sembrava di rete —
`Cannot bind parameter 'ConnectionUri' ... Invalid URI: The hostname could
not be parsed` — su `ComponentState`, un check (`Get-ServerComponentState`)
che non costruisce alcun URI proprio. La causa reale sta un livello più in
basso: su un giro lungo (molti server), la sessione PowerShell remota di
Exchange aperta all'inizio dello script (`$script:ExSession`, percorso
`New-PSSession -ConfigurationName Microsoft.Exchange` quando non c'è lo
snap-in locale) può diventare stale/disconnessa — timeout di idle di WinRM o
dell'endpoint Exchange — prima che lo script finisca di girare su tutti i
server.

**Perché è fuorviante**: `Import-PSSession` lascia le funzioni proxy nella
sessione corrente anche quando la sessione remota sottostante non è più
viva — `Test-HcCommand 'Get-ServerComponentState'` continua a trovarle e a
restituire vero. Alla prima chiamata reale, il proxy tenta una riconnessione
interna automatica, e **è quel tentativo interno a fallire** con un errore di
URI/rete, mascherando completamente l'errore originale (che sarebbe stato
semplicemente "sessione non più valida").

**Perché richiamare `Connect-HcExchange` da solo non basta**: il suo unico
controllo è `if (Test-HcCommand 'Get-ExchangeServer')` — verifica che la
*funzione* esista, non che la *sessione* dietro sia viva. Con le funzioni
proxy ancora presenti, quel controllo risulta vero e la funzione non fa
nulla.

**Fix**: `Test-HcExchangeSessionHealthy` legge lo stato reale dell'oggetto
sessione (`$script:ExSession.State -eq 'Opened'`; sempre vero quando si usa
lo snap-in locale, dove non esiste un oggetto sessione che possa invecchiare
allo stesso modo) e `Repair-HcExchangeSession` chiude esplicitamente la
sessione stale e ne apre una nuova. Verificato prima dei check Exchange-side
di **ogni singolo server** nel loop principale, non solo all'avvio dello
script — un giro su decine di server non si porta più dietro una sessione
morta senza accorgersene fino al primo errore confuso.

### 3.14 Soglia del totale-code separata dalla soglia per singola coda

Segnalato dall'utente: 150 messaggi in coda su un server, spalmati su 12 code
(media ~12,5 a coda, nessuna singolarmente anomala), generavano lo stesso
finding `Warning` di una singola coda realmente bloccata a 150 messaggi.
Causa: `Invoke-HcQueueCheck` sommava `MessageCount` di tutte le code reali del
server (`$total`) e confrontava quella somma con `Thresholds.QueueWarning`/
`QueueCritical` — le stesse soglie già usate, poche righe sotto, per giudicare
*una singola coda*. Le due grandezze misurano cose diverse (salute
complessiva del mail flow sul server vs. una coda specifica in stallo verso
una destinazione) e non hanno motivo di condividere la soglia.

**Fix**: nuove chiavi `Thresholds.QueueTotalWarning`/`QueueTotalCritical`
(default 300/1000, deliberatamente più alte delle soglie per singola coda)
usate solo per il finding aggregato `TotalMessages`; `QueueWarning`/
`QueueCritical` restano invariate per il giudizio per-coda. Non tocca la
soglia `QueueSubjectThreshold` usata per l'oggetto mail "ATTENZIONE CODE" e
per l'innesco della mail urgente, che è un meccanismo indipendente.

### 3.15 Riepilogo per database (chi è attivo, chi è passivo, chi è indietro)

Richiesto dall'utente dopo aver visto il riepilogo copie-per-server (§3.12):
utile per capire quante copie non sono sane su un server, ma non risponde
alla domanda più naturale — "questo database, su chi è montato in questo
momento, e le altre copie sono allineate?" — senza incrociare a mano le righe
di più server.

**Perché non basta iterare `Invoke-HcCopyStatusCheck`**: quella funzione gira
una volta per server con `Get-MailboxDatabaseCopyStatus -Server X`, che
restituisce solo le copie *ospitate su quel server*. Nessuna singola chiamata
vede l'intero quadro di un database (tutte le sue copie, su tutti i server).

**Fix**: ogni copia esaminata da `Invoke-HcCopyStatusCheck` viene anche
registrata, grezza, in `$script:DatabaseCopyDetail` (Server/Status/Severity/
ContentIndexState/CopyQueueLength/ReplayQueueLength). Solo a fine giro, dopo
che tutti i server sono stati processati, `Get-HcDatabaseSummary` raggruppa
quei dati per nome database e produce, per ciascuno: il server con la copia
`Mounted` (etichettata `Active`), l'elenco di tutte le copie con il loro
`Status` grezzo di Exchange (che è già la risposta a "è in sync?" — `Healthy`
= sì, `Seeding`/`Suspended`/`Failed` = no), e un suffisso `CI:<stato>` solo
quando il content index di quella copia non è nello stato normale (Healthy/
NotApplicable/Disabled), per non appesantire la riga nel caso comune.

Due casi limite gestiti esplicitamente: **nessuna copia montata** (il
database non è servibile da nessun server) → `Severity` forzata a `Critical`
anche se ogni singola copia risultasse `Healthy` presa da sola; **più di una
copia montata** (anomalia che Exchange stesso non dovrebbe permettere, ma i
dati raccolti potrebbero rifletterla in una finestra di transizione) → almeno
`Warning`, con entrambi i server elencati e un'etichetta esplicita.

Vista aggiunta sia in console (`Write-HcDatabasePerDbSummaryConsole`,
gate `Console.ShowDatabasePerDbSummary`) sia in mail (sezione "Stato per
database", gate `Mail.IncludeDatabasePerDbSummary`), subito dopo la vista
per-server esistente. Non sostituisce quella vista: le due rispondono a
domande diverse (salute di un server vs. salute di un database).

**Ordinamento (bug 18)**: la prima versione ordinava per severita poi per
nome, come le viste code/cluster - li ha senso perche sono viste "cosa non
va", ma qui e un inventario fisso (tutti i database, sempre) che si vuole
scorrere sempre nello stesso ordine, non uno che salta di posizione a ogni
giro in base a chi ha un problema in quel momento. In piu il confronto per
nome era lessicografico puro sulla stringa, quindi `"DAG5-DB12"` ordinava
prima di `"DAG5-DB2"` (non tutti gli ambienti hanno i numeri con zero-padding
a lunghezza fissa). `Get-HcNaturalSortKey` sostituisce ogni sequenza di
cifre nel nome con la sua versione zero-paddata a 20 caratteri prima del
confronto, ottenendo un ordine numerico anche senza padding esplicito nei
nomi reali. La colonna dettaglio copie, inoltre, troncava a 60 caratteri:
insufficiente con nomi server realistici e 4+ copie per database; il
troncamento e stato rimosso, la riga si espande a quanto serve.

**Severita che ignorava le code (bug 19)**: la severita per riga (sia
per-server sia per-database) veniva calcolata solo dallo `Status` testuale
della copia. Su un DB reale con una copia `Status: Healthy` ma
`ReplayQueueLength: 8576` (soglia critica di default 100), il finding
`ReplayQueue` era gia Critical tra i Findings, ma i due riepiloghi aggregati
mostravano quella riga come sana - `Status: Healthy` di Exchange descrive
solo che il meccanismo di copia funziona, non che sia allineata. Ora la
severita di riga e il massimo tra severita di stato, severita copy queue e
severita replay queue, calcolate sempre (non solo quando il copy e in stato
diverso da `Mounted`, come accadeva prima per l'emissione dei finding
dedicati, che restano invariati). Il riepilogo per database mostra inoltre
sempre `RQ:<valore>` per ogni copia non attiva - non solo quando supera la
soglia - insieme a `CQ:<valore>` quando diverso da zero e `Suspend:"..."`
quando la copia riporta un `SuspendComment`: dati diagnostici che l'utente ha
chiesto esplicitamente dopo aver dovuto controllare `Get-MailboxDatabaseCopyStatus`
a mano per capire perche una riga "Healthy" nascondesse un problema reale.

### 3.16 Recupero reattivo, origin-agnostic, di una sessione Exchange caduta

Il fix §3.13 (`Test-HcExchangeSessionHealthy`/`Repair-HcExchangeSession`) presuppone che
sia questo script ad aver aperto la sessione remota, tracciata in
`$script:ExSession`. Un utente senza `Organization.ConnectTo` configurato
(caso comune: lo script gira dentro una Exchange Management Shell, o dentro
una sessione aperta a mano con `Connect-ExchangeServer`/`RemoteExchange.ps1`
**prima** di lanciare lo script) ha riportato lo stesso identico sintomo
**dopo** il fix §3.13: `Connect-HcExchange` trova i cmdlet Exchange gia
disponibili (`Test-HcCommand 'Get-ExchangeServer'` vero) e non fa nulla,
quindi `$script:ExSession` resta `$null` per l'intero giro. Lo script non ha
alcuna visibilita sulla sessione esterna: non puo sapere quando cade, non
puo ripararla per riferimento perche non la possiede.

**Fix, in `Invoke-HcCheck`** (il wrapper centrale che gia isola le eccezioni
di ogni check): il corpo del check viene ora eseguito dentro un try/catch
interno che riconosce il sintomo per **impronta del messaggio di errore**
(`Test-HcStaleExchangeSessionSymptom`: contiene `'ConnectionUri'` e `Uri`),
non per provenienza della sessione. Se riconosciuto, `Remove-HcStaleExchangeProxy`
ripulisce gli artefatti **a prescindere da chi li abbia creati**: rimuove
ogni `PSSession` non `Opened` trovata con `Get-PSSession`, chiude
`$script:ExSession` se tracciata, e soprattutto rimuove per nome (via
`Get-Module | Where-Object { $_.ExportedCommands.ContainsKey('Get-ExchangeServer') }`)
qualunque modulo di implicit remoting esponga quel cmdlet — che sia stato
generato da `Connect-HcExchange` o da un `Connect-ExchangeServer` lanciato a
mano dall'utente, il modulo proxy che ne risulta espone gli stessi nomi di
cmdlet ed e indistinguibile. Dopo la pulizia, `Test-HcCommand 'Get-ExchangeServer'`
torna correttamente falso (il modulo non c'e piu), quindi un secondo
`Connect-HcExchange` tenta per davvero una riconnessione (snap-in se
registrato, `Organization.ConnectTo` se valorizzato) invece di limitarsi a
constatare "gia disponibile". Il check viene poi rieseguito **una sola
volta**: se va a buon fine il giro prosegue senza alcun finding; se fallisce
di nuovo (nessuna via di riconnessione disponibile: ne snap-in ne ConnectTo,
la sessione era genuinamente esterna e non recuperabile da questo script),
il finding `Unknown` risultante lo dice esplicitamente, invece di ripetere
l'errore di URI fuorviante per ogni check successivo dell'intero giro.

**Nota (rumore in console)**: `Remove-Module -Force` su un modulo di
implicit remoting Exchange stampa l'elenco dei centinaia di cmdlet che
disattiva - rumore atteso, non un errore, ma confuso da vedere apparire
subito dopo il log "sessione Exchange caduta". La chiamata usa `*> $null`
(tutti gli stream, non solo gli errori) per silenziarlo: `-ErrorAction
SilentlyContinue` da solo non basta, perché quell'output non passa dallo
stream di errore.

### 3.17 Check dischi fisici (`Get-PhysicalDisk`) e riepilogo dedicato

Nato da un incidente reale: un database (DAG5-DB19) aveva il replay queue
bloccato a ~9600 log mentre altri 9 database sullo stesso server restavano
sani. La causa, trovata solo guardando Server Manager, era un singolo disco
fisico in stato degradato — l'ambiente usa storage **JBOD** (un disco SAS
dedicato per database, niente RAID sui volumi dati: la ridondanza la
fornisce già il DAG) su server fisici (Lenovo SR650 V3), non storage
condiviso/di rete. In un layout così, un disco che degrada colpisce
esattamente il database che ci vive sopra — un sintomo che a lungo sembra
"un problema di quel database" e non "il server ha un disco da sostituire",
finché qualcuno non apre Server Manager o `Get-PhysicalDisk` per caso.

**Raccolta**: `Get-PhysicalDisk` è stato aggiunto allo stesso fan-out remoto
già esistente per OS/volumi/servizi (`Get-HcRemoteData`, §"RACCOLTA DATI
OS") — non apre una connessione remota aggiuntiva, riusa quella già in corso
per server. Se il modulo `Storage` non è disponibile (raro, ma possibile su
configurazioni minimali) degrada in silenzio a un array vuoto, senza far
fallire l'intero fan-out.

**Severità**: `HealthStatus` mappato esplicitamente — `Healthy` → OK,
`Warning` → Warning, `Unhealthy` → Critical, qualunque valore non
riconosciuto → Unknown (stessa filosofia di `ContentIndexState` in §5 riga
14: mai un default che dichiara Critical su un valore mai visto). Il finding
viene emesso per **ogni** disco, incluso quelli sani — coerente con il
principio già usato per code/cluster/database: la fotografia serve anche
quando conferma che va tutto bene, non solo quando c'è un problema.

**Riepilogo dedicato** (`$script:PhysicalDiskSummary`, popolato dentro
`Invoke-HcPhysicalDiskCheck`): una riga per disco con server, ID, modello,
tipo (SSD/HDD), dimensione, stato di salute e stato operativo — in console
ordinato per severità (i pochi dischi malati in cima, a differenza del
riepilogo per database che resta in ordine fisso: qui su un server con
decine di dischi l'obiettivo è vedere subito i pochi problemi, non scorrere
un inventario). Vista aggiunta anche in mail, sezione "Dischi fisici", subito
dopo quella per-database.

**Nuovo check**: `PhysicalDisk`, aggiunto al `ValidateSet` del parametro
`-Check` e a `Checks.PhysicalDisk` (default `true`) — indipendente da `Disk`
(che resta il check sui volumi/spazio libero, §3 "CHECK DISK"): sono due
livelli diversi, un volume pieno e un disco fisico degradato sono due
problemi distinti anche se a volte collegati.

### 3.18 Rimosso: check e alert sul backup

Su richiesta esplicita dell'utente (2026-09-23), rimosso l'intero blocco
"Backup" da `Invoke-HcDatabaseCheck`: non genera più finding di categoria
`Backup` (né "nessun backup registrato" né "ultimo backup troppo vecchio").
Rimosse anche le soglie `Thresholds.BackupAgeHoursWarning`/
`BackupAgeHoursCritical` da `$DefaultConfigJson` e da entrambi i file di
config. Il resto del check database (stato di mount, copia attiva fuori
preferenza 1) resta invariato — il backup era solo una delle tre cose
verificate da quella funzione, non l'intero check.

### 3.19 Esclusi dai check: dischi virtuali VMware e servizio RemoteRegistry

Su richiesta esplicita dell'utente (2026-09-23), due sorgenti di falsi
positivi rimosse dai check attivi:

- **Dischi "VMware Virtual disk"** ignorati nel check `PhysicalDisk` (§3.17):
  su un server virtualizzato, `Get-PhysicalDisk` espone il disco virtuale
  presentato dall'hypervisor, non lo storage fisico reale sottostante (gestito
  da vSphere/dalla SAN, fuori dalla visibilità di Windows). Il suo
  `HealthStatus` non riflette lo stato reale dell'array e genera solo rumore
  su VM — filtrato per `FriendlyName` esatto in `Get-HcRemoteData`, prima
  ancora che il disco entri nel finding o nel riepilogo. Sui server fisici
  (dove il problema §3.17 è nato) il comportamento non cambia: nessun disco
  reale ha quel `FriendlyName`.
- **`RemoteRegistry` rimosso da `ExtraServices`**: non più incluso tra i
  servizi extra monitorati dal check `Service` (§"CHECK SERVIZI" — servizio
  Automatico ma non in esecuzione ⇒ Critical). Rimosso dal default in
  `$DefaultConfigJson` e da entrambi i file di config (example + locale): il
  servizio smette semplicemente di essere raccolto e valutato, nessun cambio
  di logica nel check stesso, che resta generico per qualunque nome in lista.

### 3.20 Soppressione dinamica di `Search` sui server senza copie di database

Su richiesta esplicita dell'utente (2026-09-23): in un DAG con più membri non
tutti i server ospitano copie di ogni database, e in ambienti con diversi
membri "capacità futura" può capitare che un server non ospiti **nessuna**
copia (né attiva né passiva). In quel caso l'health set `Search` in
`Unhealthy` non riflette alcun impatto reale — non c'è alcun indice locale da
mantenere.

**Perché solo `Search` e non gli altri health set citati insieme** (ActiveSync,
Imap, OWA.Calendar.Proxy, Outlook.Proxy/MapiHttp.Proxy, RemoteMonitoring):
`Search` è l'unico legato 1:1 alla presenza di copie locali (ogni copia ha il
suo catalogo). Gli health set `*.Proxy`/protocollo citati sono invece funzioni
di **front-end**: un server le esercita per qualunque utente dell'organizzazione
instradato su di lui dal bilanciatore, indipendentemente da quali cassette
ospita — sopprimerli sulla stessa base rischierebbe di nascondere un incidente
reale su utenti reali. `RemoteMonitoring` non è legato ai database in alcun
modo. Nessuno di questi è stato quindi incluso nella soppressione.

**Implementazione (pigra, per-server — vedi bug #23)**: `Test-HcServerHasDatabaseCopy`
interroga `Get-MailboxDatabase -Server <nome>` (senza `-Status`: filtro lato AD,
nessun contatto con il server) solo per il singolo server richiesto, e solo
quando serve davvero — cioè solo se in `Invoke-HcHealthCheck` esiste almeno un
health set il cui nome compare in `Ignore.HealthSetsWithoutDatabaseCopy`
(default `["Search"]`) **e** il cui `AlertValue` è Unhealthy/Degraded. Nella
maggior parte dei giri (Search Healthy ovunque) questa funzione non viene mai
chiamata: zero query aggiuntive. Il risultato è cached per server
(`$script:ServerHasDatabaseCopyCache`) per non ripetere la query se lo stesso
server viene rivalutato più volte nello stesso giro. Fail-open: se la query
fallisce, non si sopprime nulla per quel server (meglio un falso allarme in
più che un problema nascosto per un dubbio sulla rilevazione).

Il filtro in `Invoke-HcHealthCheck` non fa sparire nulla in silenzio: gli
health set soppressi generano comunque un finding `Info` esplicito ("...
ma ignorato: X non ospita alcuna copia di database"), visibile per chi
controlla i finding grezzi, ma senza contribuire a `$bad`/`$degraded` e
quindi senza alert.

**Bug preesistente scoperto durante l'implementazione** (non introdotto oggi,
presente da quando la funzione è stata scritta): `ActivationPreference` è un
`IDictionary` (sia `Hashtable` che `Dictionary` generico) e **PowerShell non lo
srotola** in coppie chiave/valore quando viene passato a `@()` o a un `foreach`
diretto — viene trattato come un singolo oggetto scalare (comportamento
documentato di PowerShell per qualunque tipo che implementa `IDictionary`, non
solo per gli `Hashtable` letterali). Il check "copia attiva fuori preferenza 1"
in `Invoke-HcDatabaseCheck` (§3.18) usava esattamente questo pattern
(`@($db.ActivationPreference)` poi `.Key`/`.Value` in un foreach), quindi il
suo `$prefs.Count` era sempre 1 e il controllo **non è mai scattato una sola
volta** da quando è stato scritto — nessun errore visibile, il blocco è
avvolto in un try/catch che logga solo a livello DEBUG in caso di eccezione,
e qui non ne genera nessuna. Corretto enumerando `.Keys` (un `ICollection`
normale, non soggetto allo stesso comportamento) e leggendo il valore tramite
l'indicizzatore (`$db.ActivationPreference[$server]`) invece di `.Value` su un
foreach diretto. Stesso fix applicato anche alla nuova
`Initialize-HcServersWithDatabaseCopy`, scritta da zero in questa stessa
sessione: il bug è stato notato mentre si scriveva codice nuovo con lo stesso
pattern e non ha mai raggiunto un commit.

### 3.20bis Soglia Critical disco più stringente

Su richiesta esplicita dell'utente (2026-09-23): `DiskFreePercentCritical`
10→**6**, `DiskFreeGBCritical` 25→**15** (in `$DefaultConfigJson`, in
`ExchangeHealthCheck.config.example.json` e nel config locale). `DiskMode`
resta `And` (default preesistente, invariato): il Critical scatta solo
quando **entrambe** le condizioni sono vere insieme (percentuale libera sotto
il 6% **e** meno di 15 GB liberi), non con l'una o l'altra da sola — coerente
con la logica già in uso per evitare falsi positivi su volumi molto grandi
(10% di un volume da 4 TB sono comunque 400 GB liberi, non un'emergenza).
`DiskFreePercentWarning`/`DiskFreeGBWarning` (20% / 60 GB) restano invariati.

### 3.21 Alert "sostenuto": non notificare code Replay/Copy che si risolvono da sole

Su richiesta esplicita dell'utente (2026-09-23): `ReplayQueue`/`CopyQueue`
oscillano normalmente in Warning durante una replica pesante (es. dopo un
riavvio del servizio replica, un failover, un picco di scrittura) e spesso
rientrano da soli entro pochi minuti — non un incidente, solo il DAG che sta
smaltendo il backlog. Notificare ogni volta genera fatica da allarme senza
informazione utile: quello che conta e' sapere se resta indietro **a lungo**,
non se e' stato Warning per tre minuti.

**Implementazione**: `Resolve-HcAlert` (macchina a stati esistente, §5 riga 12
e dintorni — gia' persisteva `FirstSeen` per ogni finding tra un giro e
l'altro in `State\alert-state.json`) ora, per le categorie elencate in
`Alerting.SustainedCategories` (default `["ReplayQueue", "CopyQueue"]`), non
notifica piu' un finding alla sua prima comparsa: si limita a iniziare a
tracciarne la durata. Diventa un alert vero (finisce in "Nuove anomalie") solo
se resta ininterrottamente elevato (Warning o Critical) per almeno
`Alerting.SustainedDurationMinutes` (default 120 = 2 ore, 0 disattiva la
funzione). Se rientra prima, la voce sparisce dallo stato senza aver mai
generato una notifica — e senza comparire nemmeno tra i "rientri", perché non
si può recuperare da un allarme mai dato (altrimenti l'utente riceverebbe un
"risolto" per qualcosa che non gli era mai stato segnalato).

Una volta superata la soglia e notificato per la prima volta, il finding torna
al comportamento normale: un'ulteriore escalation di severità o il cooldown
per i promemoria funzionano esattamente come per qualunque altro alert.
`-ForceMail`/`-TestMail` bypassano sempre l'attesa (comodo per verificare la
configurazione senza dover aspettare 2 ore). La soglia è **per categoria**,
non globale: gli altri health set/categorie (ActiveSync, certificati,
componenti, ecc.) restano notificati subito come sempre, non solo Replay/Copy.

## 4. Schema di configurazione (riferimento completo)

Vedi `ExchangeHealthCheck.config.example.json` per i valori concreti. Sezioni:

| Sezione | Contenuto |
|---|---|
| `Organization` | Nome, server a cui connettersi (`ConnectTo`, usato solo se non si è già in Exchange Management Shell), `ViewEntireForest`, rilevamento/override del domain controller |
| `Servers` | Perimetro (`Include`/`Exclude`/`SiteFilter`/`IncludeEdge`), `UseFqdnForRemoting`, `SkipExchangeChecksWhenOffline` |
| `Checks` | Un booleano per famiglia di controllo (Os, Disk, PhysicalDisk (§3.17), Services, Components, Health, Dag, Replication, Databases, Queues, BackPressure, Certificates, Mapi) |
| `Thresholds` | Tutte le soglie numeriche: disco (con `DiskMode` And/Or), memoria, CPU, code (`QueueWarning`/`QueueCritical` per singola coda, `QueueTotalWarning`/`QueueTotalCritical` per il totale-server, §3.14), copy/replay queue del DAG, scadenza certificati, timeout di rete, `QueueSubjectThreshold` (soglia condivisa tra il tag "ATTENZIONE CODE" in oggetto e l'innesco della mail dedicata alle novità), `CertificateCheckTimeoutSeconds`/`ManagedAvailabilityTimeoutSeconds` (timeout duro via job separato, §3.11) |
| `VolumeOverrides` | Soglie disco per pattern di server/volume, con precedenza sul primo match |
| `HealthReport` | `IncludeFailingMonitors` (arricchisce l'alert con i monitor Managed Availability in errore), `MaxMonitorsPerHealthSet`, `MonitorDetailTimeoutSeconds` (§3.11) |
| `Queues` | `ResolveNextHopHostnames`, `ReverseDnsTimeoutMs`, `TryNetBiosFallback`, `NetBiosTimeoutMs`, `ResolveSendConnectorName` |
| `Console` | `ShowCategorySummary`, `ShowClusterSummary`, `ShowDatabaseSummary` (per server), `ShowDatabasePerDbSummary` (per database, §3.15), `ShowPhysicalDiskSummary` (§3.17), `ShowQueueSummary` |
| `Ignore` | Liste di esclusione: Services, ServerComponents, HealthSets (statico, per nome), HealthSetsWithoutDatabaseCopy (dinamico, solo su server senza copie — §3.20), Volumes, Databases, Keys (pattern esatto `Categoria\|Server\|Oggetto`, con wildcard) |
| `ExtraServices` | Servizi non-Exchange da includere nel check Services (default W3SVC, WinRM — RemoteRegistry rimosso, §3.19) |
| `Alerting` | Cooldown, heartbeat, notifica di rientro, severità minima da notificare, `SustainedDurationMinutes`/`SustainedCategories` (alert "sostenuto" per Replay/Copy queue, §3.21) |
| `Mail` | SMTP, autenticazione, mittente/destinatari, allegato CSV, vista code, `QueueAlertSubjectTag`, `SeparateAlertSubjectTag` (vedi §3.10), `IncludeClusterSummary`, `IncludeDatabaseSummary` (§3.12), `IncludeDatabasePerDbSummary` (§3.15), `IncludePhysicalDiskSummary` (§3.17) |
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
| 12 | `Cannot convert value ... DisplayHint ... to type System.DateTime` alla lettura dello stato | Sotto **Windows PowerShell 5.1** (non riproducibile in pwsh 7), `Get-Date` chiamato **direttamente** dentro un literal `@{ Chiave = Get-Date }` o assegnato direttamente a una proprietà esistente (`$obj.Prop = Get-Date`) produce un oggetto che `ConvertTo-Json` serializza come `{"value":..., "DisplayHint":2, "DateTime":...}` invece di una data semplice; il cast `[datetime]` al giro successivo fallisce | Passare prima da una variabile (`$now = Get-Date`, poi usare `$now`) o da un cast esplicito `[datetime](Get-Date)`: entrambi verificati sicuri con un test dedicato. Vedi §8 per la regola generale |
| 13 | Il giro sembrava bloccato indefinitamente; il log mostrava solo l'ultimo check completato (`ManagedAvailability`) senza altro per minuti | In realtà il check *successivo* (`Certificate`) era fermo, non quello loggato per ultimo: la durata si stampa solo a fine check, quindi l'ultimo nome visibile in log durante un blocco non è quello bloccato, è quello appena prima. `Get-ExchangeCertificate` era appeso 300,3s (il timeout RPC di default di Windows) su un server con un problema di backend IIS pre-esistente | `Invoke-HcExchangeWithTimeout` (§3.11): la chiamata gira in un job separato terminabile con forza entro un limite configurabile (default 30s) |
| 14 | `ContentIndex` segnalato Critical su `NotApplicable` | Il codice trattava "qualunque cosa diversa da Crawling/Seeding/Suspended" come Critical per default; `NotApplicable` è invece uno stato normale — tipicamente una copia ritardata (lagged copy, `ReplayLagTime` > 0) che Exchange non indicizza di proposito perché non è pensata per servire ricerche live | Mappatura esplicita per stato (`Failed`/`FailedAndSuspended` → Critical, `Crawling`/`Seeding`/`Suspended`/`Unknown` → Warning, `NotApplicable`/`Disabled` → Info); un valore mai visto prima diventa `Unknown`, non più Critical per default |
| 15 | Un nodo cluster `Down` reale (incidente su un membro DAG) non risultava da nessuna parte nel report | `Get-ClusterNode` non era disponibile sulla macchina da cui girava lo script (modulo FailoverClusters/RSAT mancante); il controllo era scritto come `if (Test-HcCommand ...) { ... }` senza alcun ramo per il caso "comando non disponibile", quindi l'intero controllo cluster spariva senza lasciare traccia | Aggiunto un `else` che produce un finding `Warning` esplicito quando la capacità manca del tutto, con l'indicazione di quale componente RSAT installare |
| 16 | `Cannot bind parameter 'ConnectionUri' ... Invalid URI: The hostname could not be parsed` su un check (`ComponentState`) che non costruisce alcun URI | Su un giro lungo (molti server) la sessione remota Exchange (`$script:ExSession`) diventa stale/disconnessa prima che lo script finisca; le funzioni proxy di `Import-PSSession` restano richiamabili anche a sessione morta, e il loro tentativo di riconnessione interna fallisce con un errore di URI che maschera la vera causa | `Test-HcExchangeSessionHealthy`/`Repair-HcExchangeSession` (§3.13): stato reale della sessione verificato e riparato prima dei check Exchange-side di ogni server, non solo all'avvio |
| 17 | 150 messaggi totali su 12 code (~12,5 a coda, nessuna anomala) segnalati come `Warning` code, l'utente si aspettava l'allarme solo per una singola coda realmente alta | Il finding `TotalMessages` sommava tutte le code del server e lo confrontava con le **stesse** soglie (`QueueWarning`/`QueueCritical`) usate per giudicare una singola coda; tante code piccole sommate superavano la soglia pensata per un'unica coda bloccata | Soglie separate `QueueTotalWarning`/`QueueTotalCritical` (default 300/1000) per il totale-server, distinte da `QueueWarning`/`QueueCritical` che restano invariate per la singola coda |
| 18 | Riepilogo per database (§3.15) mostrato in un ordine confuso (es. DB12 prima di DB01) e con la colonna "Copie (Server:Stato)" troncata con `...` su ambienti con nomi server lunghi e 4+ copie per database | L'ordinamento era severita-poi-nome (utile per le viste "solo problemi" come code/cluster, fuorviante per un inventario fisso che si scorre sempre allo stesso modo); il confronto per nome era lessicografico puro (`"DB12" < "DB2"` come stringhe); la colonna dettaglio era troncata a 60 caratteri, insufficiente con nomi server realistici (es. `GRPI-EXC-PCxx`) su 4 copie | Ordinamento per solo nome database con `Get-HcNaturalSortKey` (zero-padding delle sequenze numeriche, cosi l'ordine e numerico e non lessicografico); troncamento rimosso, la colonna si espande al contenuto |
| 19 | Un database reale (DAG5-DB19) aveva una copia con `Status Healthy` ma `ReplayQueueLength 8576` (soglia critica di default 100): il finding `ReplayQueue` era gia Critical tra i Findings, ma sia il riepilogo per server sia quello per database (§3.15) mostravano quella riga come sana, senza alcun segnale | La severita usata nei due riepiloghi aggregati veniva calcolata **solo** dallo Status testuale della copia (`Healthy`/`Seeding`/`Suspended`/...); `Status: Healthy` descrive solo che il meccanismo di copia funziona, non che la copia sia allineata - CopyQueueLength/ReplayQueueLength non entravano mai nel calcolo di quella severita, restavano confinati al loro finding dedicato | La severita per riga (sia per-server sia per-database) e ora il massimo tra severita di stato, severita copy queue e severita replay queue; il `ProblemDetail` per-server mostra esplicitamente "Healthy ma code indietro (copy=X, replay=Y)" invece di limitarsi a ripetere lo Status; il riepilogo per database mostra sempre `RQ:<valore>` per le copie non attive (non solo quando fuori soglia) e aggiunge `Suspend:"..."` quando presente |
| 20 | Il fix §3.13 (stale-session repair) non risolveva il sintomo per un utente con `Organization.ConnectTo` **non configurato**: `Cannot bind parameter 'ConnectionUri' ... hostname could not be parsed` continuava a presentarsi identico dopo il fix | §3.13 traccia e ripara solo `$script:ExSession`, popolata unicamente quando e `Connect-HcExchange` (di questo script) ad aprire la sessione remota. Ma se `Organization.ConnectTo` non serve perche lo script gira dentro una Exchange Management Shell o una sessione aperta a mano con `Connect-ExchangeServer`/`RemoteExchange.ps1` **prima** di lanciare lo script, `Test-HcCommand 'Get-ExchangeServer'` trova i cmdlet gia disponibili, `Connect-HcExchange` non fa nulla, e `$script:ExSession` resta `$null` per l'intero giro: lo script non ha alcuna visibilita sulla sessione esterna e non puo mai sapere che e caduta | §3.16: recupero reattivo e origin-agnostic in `Invoke-HcCheck`, che non dipende dal tracciare una sessione specifica |
| 21 | Dopo il fix del bug 20, `Access is denied` comparso su `Get-Queue` su **tutti** i server (prima non succedeva), e `ManagedAvailability` diventato lento, nello stesso ambiente senza `Organization.ConnectTo` | Il fix del bug 20 riconosce il sintomo per impronta del messaggio ovunque compaia, ma `Invoke-HcExchangeWithTimeout` (usata da Certificate/ManagedAvailability/Get-ServerHealth) lo produce **sempre**, deterministicamente, in un ambiente senza ConnectTo ne snap-in: il job che apre gira in un processo nuovo che non eredita la sessione esterna del processo principale, quindi non ha modo di autenticarsi da solo. Il recupero del bug 20, innescato da questo fallimento strutturale (non da una sessione davvero stale), rimuoveva il modulo Exchange condiviso e valido della sessione esterna - da cui `Access is denied` su `Get-Queue` e altri check che non c'entravano nulla con Certificate/ManagedAvailability, e la lentezza (ogni job impiega secondi per fallire nel modo sbagliato) | Quando ne snap-in ne ConnectTo sono disponibili, `Invoke-HcExchangeWithTimeout` salta il job e chiama il cmdlet diretto nel processo corrente (§3.11), dove i cmdlet della sessione esterna sono gia disponibili: il fallimento deterministico non si presenta piu, quindi il recupero del bug 20 non viene piu innescato da questa causa |
| 22 | Il check "copia attiva fuori preferenza 1" (§3.18) non generava mai un finding `Activation`, nemmeno con database volutamente attivi su una copia diversa dalla preferenza 1 | `ActivationPreference` e' un `IDictionary`; `@($db.ActivationPreference)` non lo srotola in coppie chiave/valore (trattato come un singolo oggetto scalare, comportamento PowerShell per qualunque `IDictionary`), quindi `$prefs.Count` restava sempre 1 e la condizione `-gt 1` non era mai vera | Enumerazione tramite `.Keys` (un `ICollection` normale) e lettura del valore con l'indicizzatore (`$db.ActivationPreference[$server]`), sia nel check esistente sia nella nuova `Initialize-HcServersWithDatabaseCopy` (§3.20) scritta con lo stesso pattern |
| 23 | Il giro rallentava percettibilmente subito prima di iniziare i check ManagedAvailability per-server, su ogni esecuzione | La prima versione della soppressione dinamica di Search (§3.20) faceva una query ``Get-MailboxDatabase`` **senza** filtro server, una volta per ogni giro, indipendentemente dal fatto che servisse davvero (nella maggior parte dei giri Search e' gia' Healthy ovunque e la soppressione non serve a nulla) | Query resa pigra e per-singolo-server: ``Test-HcServerHasDatabaseCopy`` interroga ``Get-MailboxDatabase -Server X`` solo quando esiste davvero un health set Unhealthy/Degraded che la richiede, con risultato cached per server per lo stesso giro |

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
- **Sviluppato e testato prevalentemente in PowerShell 7 (pwsh), ma eseguito in
  produzione su Windows PowerShell 5.1** (quella con cui gira Exchange
  Management Shell): un bug reale (§5 riga 12) esisteva SOLO sotto 5.1 e non si
  riproduceva affatto in pwsh 7. Regola pratica per non ripeterlo: **`Get-Date`
  non va mai usato direttamente come valore di una proprietà** — né dentro un
  literal `@{ Chiave = Get-Date }` né come assegnazione diretta
  (`$obj.Prop = Get-Date`) — se quella proprietà può finire in
  `ConvertTo-Json`/`Save-HcState`. Assegnarlo prima a una variabile
  (`$now = Get-Date`, poi usare `$now`) o castarlo esplicitamente
  (`[datetime](Get-Date)`). Più in generale: quando un difetto tocca la
  serializzazione, il parsing di date/numeri o il comportamento dei cmdlet di
  base, **verificare sotto Windows PowerShell 5.1 reale** (`powershell.exe`,
  non solo `pwsh`) prima di dare per buono un fix — i due engine non si
  comportano sempre allo stesso modo, come già successo qui.
- **Mai un default che dichiara Critical su "qualunque valore non riconosciuto"**
  per un campo enum-like di Exchange (`ContentIndexState`, `Status` di una
  copia database, e simili). §5 riga 14: `NotApplicable` — uno stato normale
  per le copie ritardate — veniva segnalato Critical solo perché non era uno
  dei pochi valori esplicitamente gestiti come non-Critical. La lista completa
  dei valori possibili di questi enum non è sempre nota in anticipo (vedi §6):
  un valore mai visto va verso `Unknown`, mai verso `Critical` per default —
  un falso allarme costa fiducia nello strumento, un falso silenzio su
  `Unknown` resta comunque visibile nel report.
- **Versione dichiarata nello script, stampata a ogni esecuzione**: dopo un
  episodio in cui un fix era stato pushato ma l'utente stava ancora testando
  una copia non aggiornata sulla macchina di esecuzione (dubbio impossibile
  da sciogliere a distanza senza un riferimento diretto), `$script:ScriptVersion`
  (vicino agli altri `$script:` di inizializzazione) va incrementato a ogni
  commit che tocca il comportamento dello script, e il file `VERSION` alla
  radice del repository tenuto allineato allo stesso valore. La versione è
  anche la prima riga di log di ogni esecuzione (`Invoke-ExchangeHealthCheck.ps1
  - versione X.Y.Z`): basta guardare l'inizio di un log per sapere con
  certezza quale versione ha davvero girato, senza dover confrontare il
  contenuto del file.

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
