# hMailServer Event Handlers (JScript)

*[English](#english) | [Français](#français)*

A set of hMailServer event handlers written in JScript. It filters inbound
connections and auto-bans abusive hosts (geographic filtering, AbuseIPDB
reputation, failed-login and recipient-harvesting protection), cleans up a few
outbound headers, and writes a single, column-aligned log file per day.

---

## English

### Features

* Country blocking on every port (SMTP, IMAP, POP).
* Geographic restriction on non-SMTP ports: mailbox access (IMAP/POP) is allowed
  only from a configurable list of countries.
* AbuseIPDB reputation check on the submission ports (587/465).
* Auto-ban of unauthenticated / failed logins.
* Recipient-probe (directory harvesting) protection: bans an IP after a number
  of unknown recipients within a sliding window. Uses a small counter file, no
  database.
* Outbound `Received` header anonymisation (removes the client IP).
* Fills a missing `Message-Id` and reformats `X-Spam-Report` so it is readable.
* Optional SMTP rejection of messages whose spam score is above a threshold. It
  reads the score your anti-spam layer already added (for example SpamAssassin's
  `X-Spam-Score`, or the `score=` field of `X-Spam-Status`). Off by default.
* Daily log with fixed-width, aligned columns, UTF-8 without BOM. One line per
  event.

Each filtering feature can be enabled or disabled, and set to **ban** (a temporary
IP range in hMailServer, so a banned host is dropped before it reaches the script
again) or **reject** (refuse only the current session with a professional SMTP
message, without a persistent ban).

### Requirements

* hMailServer with scripting set to **JScript**. Tested with RvdH's community
  build **5.7**: https://d-fault.nl/files/hMailServer-Builds/Installers
* **Disconnect.exe** to drop the current session of a banned host. Copy RvdH's
  `Disconnect.exe` to the hMailServer `Events` folder: https://d-fault.nl/files/
* For the country lookup and AbuseIPDB checks, two COM components must be
  registered on the server:
  * `DNSLibrary.DNSResolver` (country lookup)
  * `AbuseIPDBComponent.AbuseIPDBRestClient` (AbuseIPDB)

  These are part of the RvdH community tools (see the d-fault.nl files above). If
  a component is not present, the matching check is skipped, the connection is
  allowed, and the error is written to the log. The rest keeps working.
* An AbuseIPDB API key, if you use the AbuseIPDB check.

Country lookup uses the public reverse-DNS service
`country.junkemailfilter.com`.

### Installation

1. In hMailServer Administrator, go to **Settings → Advanced → Scripts**, set the
   language to **JScript** and enable scripting.
2. Copy `EventHandlers.js` into the `Events` folder (default
   `C:\Program Files\hMailServer\Events`). The file must keep that exact name.
3. Copy `Disconnect.exe` into the same `Events` folder.
4. Open `EventHandlers.js` and edit the settings at the top of the file (see
   below).
5. Back in the Administrator, click **Save**, then **Reload script**.

### Configuration

All settings sit at the top of the file:

* Admin credentials used to create the bans (`ADMIN`, `PASSWORD`).
* `LOGDIR`, `LOGPREFIX`, `DISCONNECT_EXE`: log folder, log file prefix and path
  to `Disconnect.exe`.
* One object per filtering feature — `GEOBLOCK`, `GEORESTRICT`, `ABUSEIPDB`,
  `UNKNOWNUSER`, `RCPTPROBE`, `SPAMREJECT` — each with:
  * `enabled`: turn the feature on or off.
  * `action`: `"ban"` (temporary IP-range ban + disconnect) or `"reject"` (refuse
    only the current session with a professional SMTP message, no ban).
  * `ban`: `{ qty, unit }` duration (unit `d`/`h`/`n`/`s`), used in `"ban"` mode;
    `qty` `0` disconnects without banning.
  * `msg`: the SMTP message returned in `"reject"` mode.
  * plus feature-specific fields: `ABUSEIPDB` has `apikey`, `maxConfidence`,
    `maxAgeDays`; `RCPTPROBE` has `threshold`, `windowMin`; `SPAMREJECT` has
    `score`, `header`. `SPAMREJECT` is off by default.
* `RECEIVEDANON_ENABLED`, `MESSAGEID_ENABLED`, `SPAMREPORT_ENABLED`: on/off flags
  for the header-transform features.
* `BLOCKED_COUNTRIES`, `ALLOWED_GEO`, `SMTP_PORTS`, `SUBMISSION_PORTS`:
  pipe-delimited lists such as `"|cn|cz|ru|"`.
* `BAN_PRIORITY`, `LOG_SOURCES`, local network exemption (`LOCAL_IP_PREFIX`,
  `LOCALHOST_IP`), and the recipient-probe counter files (`RCPT_DATA_FILE`,
  `RCPT_LOCK_FILE`, `LOCK_TRIES`, `LOCK_STALE_SEC`).

### Order of checks on connect

1. Local network and localhost are skipped.
2. Blocked country: rejected on any port.
3. Non-SMTP port from a country outside the allowed list: rejected.
4. Submission port with an IP listed on AbuseIPDB: rejected.

### Logging

The log is written to `LOGDIR\<prefix>_YYYYMMDD.log`, one file per day, in fixed
columns:

```
2026-07-04 00:39:59  AutoBan     111.18.196.17    cn   +7d    Bad Country 143
2026-07-04 05:28:56  AutoBan     85.121.183.244   ro   +1d    Unknown USER 587
2026-07-04 17:39:24  AutoBan     20.12.240.184    us   +1d    AbuseIPDB 587
```

Columns: date and time, source, IP, country, ban duration, detail. Fields are
separated by two or more spaces, so a log line can be split on `/ {2,}/`. The
`source` column is one of `AutoBan`, `Disconnect`, `Reject`, `Spam`, `AbuseIPDB`,
`GeoLookup`, `Debug`, `System`.

### Notes

* Geographic restriction on non-SMTP ports can also lock out a legitimate user
  who reads mail over IMAP/POP from abroad. Keep the ban short, set
  `GEORESTRICT.ban.qty` to `0` to only disconnect, or `GEORESTRICT.action` to
  `"reject"` to avoid a persistent ban.
* Bans are stored as IP ranges and are removed automatically when they expire.
  Long ban durations on high-volume sources will accumulate more ranges.
* Two simultaneous connections from the same host can both try to create the same
  ban; the duplicate is handled safely and logged once at `Debug` level, not as an
  error.
* Spam-score rejection is off by default and needs an anti-spam layer that adds
  a score header before the message is accepted (SpamAssassin does this). If you
  also use hMailServer's built-in spam delete threshold, keep only one of the
  two: the built-in threshold deletes silently, while the script returns a 550
  to the sender.
* The script is written for the classic JScript engine used by hMailServer.
  `oMessage.HeaderValue("X") = value` is the correct way to set a header there
  and matches hMailServer's own tests; it is not standard browser JavaScript.

### License

MIT. See `LICENSE`.

---

## Français

Ensemble de gestionnaires d'évènements hMailServer écrits en JScript. Il filtre
les connexions entrantes et bannit automatiquement les hôtes abusifs (filtrage
géographique, réputation AbuseIPDB, protection contre les échecs
d'authentification et la collecte d'adresses), nettoie quelques en-têtes
sortants et écrit un seul fichier de log par jour, en colonnes alignées.

### Fonctionnalités

* Blocage par pays sur tous les ports (SMTP, IMAP, POP).
* Restriction géographique sur les ports non-SMTP : l'accès aux boîtes
  (IMAP/POP) n'est autorisé que depuis une liste de pays configurable.
* Contrôle de réputation AbuseIPDB sur les ports de soumission (587/465).
* Bannissement automatique des connexions non authentifiées.
* Protection contre les sondes de destinataires (collecte d'adresses) : bannit
  une IP après un certain nombre de destinataires inconnus dans une fenêtre
  glissante. Utilise un petit fichier compteur, sans base de données.
* Anonymisation de l'en-tête `Received` sortant (retire l'IP du client).
* Complète un `Message-Id` manquant et remet en forme `X-Spam-Report` pour le
  rendre lisible.
* Rejet SMTP optionnel des messages dont le score de spam dépasse un seuil. Il
  lit le score déjà ajouté par ta couche anti-spam (par exemple le `X-Spam-Score`
  de SpamAssassin, ou le champ `score=` de `X-Spam-Status`). Désactivé par défaut.
* Log quotidien en colonnes de largeur fixe, UTF-8 sans BOM. Une ligne par
  évènement.

Chaque fonction de filtrage peut être activée ou désactivée, et réglée sur **ban**
(plage d'IP temporaire dans hMailServer : l'hôte banni est coupé avant même
d'atteindre le script à sa tentative suivante) ou **reject** (refus de la seule
session courante avec un message SMTP professionnel, sans ban persistant).

### Prérequis

* hMailServer avec le langage de script réglé sur **JScript**. Testé avec la
  build communautaire **5.7** de RvdH :
  https://d-fault.nl/files/hMailServer-Builds/Installers
* **Disconnect.exe** pour couper la session en cours d'un hôte banni. Copiez le
  `Disconnect.exe` de RvdH dans le dossier `Events` de hMailServer :
  https://d-fault.nl/files/
* Pour la résolution de pays et le contrôle AbuseIPDB, deux composants COM
  doivent être enregistrés sur le serveur :
  * `DNSLibrary.DNSResolver` (résolution de pays)
  * `AbuseIPDBComponent.AbuseIPDBRestClient` (AbuseIPDB)

  Ils font partie des outils communautaires RvdH (voir les fichiers d-fault.nl
  ci-dessus). Si un composant est absent, le contrôle correspondant est ignoré,
  la connexion est acceptée et l'erreur est écrite dans le log. Le reste
  continue de fonctionner.
* Une clé d'API AbuseIPDB, si vous utilisez le contrôle AbuseIPDB.

La résolution de pays utilise le service DNS inversé public
`country.junkemailfilter.com`.

### Installation

1. Dans hMailServer Administrator, allez dans **Settings → Advanced → Scripts**,
   réglez le langage sur **JScript** et activez les scripts.
2. Copiez `EventHandlers.js` dans le dossier `Events` (par défaut
   `C:\Program Files\hMailServer\Events`). Le fichier doit garder ce nom exact.
3. Copiez `Disconnect.exe` dans ce même dossier `Events`.
4. Ouvrez `EventHandlers.js` et modifiez les paramètres en tête de fichier (voir
   ci-dessous).
5. De retour dans l'Administrator, cliquez sur **Save**, puis sur
   **Reload script**.

### Configuration

Tous les paramètres se trouvent en tête de fichier :

* Identifiants d'administration servant à créer les bans (`ADMIN`, `PASSWORD`).
* `LOGDIR`, `LOGPREFIX`, `DISCONNECT_EXE` : dossier des logs, préfixe du fichier
  et chemin vers `Disconnect.exe`.
* Un objet par fonction de filtrage — `GEOBLOCK`, `GEORESTRICT`, `ABUSEIPDB`,
  `UNKNOWNUSER`, `RCPTPROBE`, `SPAMREJECT` — chacun avec :
  * `enabled` : active ou désactive la fonction.
  * `action` : `"ban"` (ban temporaire par plage d'IP + déconnexion) ou `"reject"`
    (refus de la seule session courante avec un message SMTP professionnel, sans ban).
  * `ban` : durée `{ qty, unit }` (unité `d`/`h`/`n`/`s`), utilisée en mode `"ban"` ;
    `qty` à `0` déconnecte sans bannir.
  * `msg` : le message SMTP renvoyé en mode `"reject"`.
  * plus des champs propres à la fonction : `ABUSEIPDB` a `apikey`, `maxConfidence`,
    `maxAgeDays` ; `RCPTPROBE` a `threshold`, `windowMin` ; `SPAMREJECT` a `score`,
    `header`. `SPAMREJECT` est désactivé par défaut.
* `RECEIVEDANON_ENABLED`, `MESSAGEID_ENABLED`, `SPAMREPORT_ENABLED` : drapeaux
  d'activation des fonctions de transformation d'en-têtes.
* `BLOCKED_COUNTRIES`, `ALLOWED_GEO`, `SMTP_PORTS`, `SUBMISSION_PORTS` : listes
  délimitées par des barres, par exemple `"|cn|cz|ru|"`.
* `BAN_PRIORITY`, `LOG_SOURCES`, exemption du réseau local (`LOCAL_IP_PREFIX`,
  `LOCALHOST_IP`) et les fichiers du compteur de sondes (`RCPT_DATA_FILE`,
  `RCPT_LOCK_FILE`, `LOCK_TRIES`, `LOCK_STALE_SEC`).

### Ordre des contrôles à la connexion

1. Réseau local et localhost sont ignorés.
2. Pays bloqué : rejet sur tous les ports.
3. Port non-SMTP depuis un pays hors liste autorisée : rejet.
4. Port de soumission avec une IP listée sur AbuseIPDB : rejet.

### Journalisation

Le log est écrit dans `LOGDIR\<préfixe>_AAAAMMJJ.log`, un fichier par jour, en
colonnes fixes :

```
2026-07-04 00:39:59  AutoBan     111.18.196.17    cn   +7d    Bad Country 143
2026-07-04 05:28:56  AutoBan     85.121.183.244   ro   +1d    Unknown USER 587
2026-07-04 17:39:24  AutoBan     20.12.240.184    us   +1d    AbuseIPDB 587
```

Colonnes : date et heure, source, IP, pays, durée de ban, détail. Les champs
sont séparés par au moins deux espaces, on peut donc découper une ligne sur
`/ {2,}/`. La colonne `source` vaut `AutoBan`, `Disconnect`, `Reject`, `Spam`,
`AbuseIPDB`, `GeoLookup`, `Debug` ou `System`.

### Remarques

* La restriction géographique sur les ports non-SMTP peut aussi bloquer un
  utilisateur légitime qui relève son courrier en IMAP/POP depuis l'étranger.
  Gardez une durée de ban courte, mettez `GEORESTRICT.ban.qty` à `0` pour seulement
  déconnecter, ou `GEORESTRICT.action` à `"reject"` pour éviter un ban persistant.
* Les bans sont stockés en plages d'IP et supprimés automatiquement à
  expiration. Des durées longues sur des sources à fort volume accumulent
  davantage de plages.
* Deux connexions simultanées d'un même hôte peuvent tenter de créer le même ban ;
  le doublon est géré proprement et journalisé une fois en `Debug`, pas en erreur.
* Le rejet sur score de spam est désactivé par défaut et suppose une couche
  anti-spam qui ajoute un en-tête de score avant l'acceptation du message
  (SpamAssassin le fait). Si tu utilises aussi le seuil de suppression natif de
  hMailServer, n'en garde qu'un seul : le seuil natif supprime silencieusement,
  le script renvoie un 550 à l'expéditeur.
* Le script vise le moteur JScript classique utilisé par hMailServer.
  `oMessage.HeaderValue("X") = valeur` y est la bonne façon de définir un
  en-tête et correspond aux propres tests de hMailServer ; ce n'est pas du
  JavaScript navigateur standard.

### Licence

MIT. Voir `LICENSE`.
