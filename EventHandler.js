/****************************************************************************
 *  EventHandlers.js  (Microsoft JScript for hMailServer)
 *  Converted from EventHandlers.vbs
 *
 *  Activation: hMailServer Admin -> Settings -> Advanced -> Scripts
 *              Language = JScript, file named EventHandlers.js.
 ****************************************************************************/

/* ==========================  Settings  =============================== */
var ADMIN    = "YOUR_ADMIN_USERNAME";
var PASSWORD = "YOUR_ADMIN_PASSWORD";
var ABUSEIPDB_APIKEY = "YOUR_ABUSEIPDB_API_KEY";

/**
 * AbuseIPDB settings.
 * @const {number} ABUSEIPDB_MAX_CONFIDENCE  max tolerated confidence score (above -> blocked)
 * @const {number} ABUSEIPDB_MAX_AGE_DAYS    max age (days) of reports taken into account
 * @const {string} ABUSEIPDB_REJECT_MSG      SMTP message returned on rejection
 */
var ABUSEIPDB_MAX_CONFIDENCE = 40;
var ABUSEIPDB_MAX_AGE_DAYS   = 90;
var ABUSEIPDB_REJECT_MSG = "4.7.1 Your access to this mail system has been rejected due to the sending MTA's poor reputation (AbuseIPDB). If you believe that this failure is in error, please contact the intended recipient via alternate means.";
var LOGDIR    = "C:\\path\\to\\your\\logs";
var LOGPREFIX = "events";
var DISCONNECT_EXE = "C:\\path\\to\\Disconnect.exe";

/**
 * Local network exempt from all filtering.
 * @const {string} LOCAL_IP_PREFIX  LAN IP prefix
 * @const {string} LOCALHOST_IP     loopback address
 */
var LOCAL_IP_PREFIX = "172.16.";
var LOCALHOST_IP    = "127.0.0.1";

/**
 * Enabled features: set to false to disable a feature as needed. Each feature's
 * parameters remain in their own blocks below.
 * @const {boolean} GEOBLOCK_ENABLED      country blocking (BLOCKED_COUNTRIES), all ports
 * @const {boolean} GEORESTRICT_ENABLED   non-SMTP ports restricted to ALLOWED_GEO countries
 * @const {boolean} ABUSEIPDB_ENABLED     AbuseIPDB reputation check (submission ports)
 * @const {boolean} UNKNOWNUSER_ENABLED   ban of unauthenticated connections (OnClientLogon)
 * @const {boolean} RCPTPROBE_ENABLED     recipient-probe counter (OnRecipientUnknown)
 * @const {boolean} RECEIVEDANON_ENABLED  outbound Received header anonymisation
 * @const {boolean} MESSAGEID_ENABLED     fills a missing Message-Id
 * @const {boolean} SPAMREPORT_ENABLED    makes the X-Spam-Report header readable
 * @const {boolean} SPAM_REJECT_ENABLED   SMTP rejection on high spam score (SPAM_REJECT_* config below)
 */
var GEOBLOCK_ENABLED     = true;
var GEORESTRICT_ENABLED  = true;
var ABUSEIPDB_ENABLED    = true;
var UNKNOWNUSER_ENABLED  = true;
var RCPTPROBE_ENABLED    = true;
var RECEIVEDANON_ENABLED = true;
var MESSAGEID_ENABLED    = true;
var SPAMREPORT_ENABLED   = true;
var SPAM_REJECT_ENABLED  = false;

/**
 * Auto-ban durations per scenario. DateAdd unit: "d"=day, "h"=hour, "n"=minute, "s"=second.
 * Set qty to 0 to NOT ban (disconnect only).
 * @const {{qty:number,unit:string}} BAN_BADCOUNTRY   blocked countries (BLOCKED_COUNTRIES)
 * @const {{qty:number,unit:string}} BAN_GEORESTRICT  non-SMTP port outside ALLOWED_GEO
 * @const {{qty:number,unit:string}} BAN_ABUSEIPDB    IP listed on AbuseIPDB
 * @const {{qty:number,unit:string}} BAN_UNKNOWNUSER  authentication failure (OnClientLogon)
 * @const {number} BAN_PRIORITY  priority of created ranges (must exceed the "Internet" ranges)
 */
var BAN_BADCOUNTRY  = { qty: 7, unit: "d" };
var BAN_GEORESTRICT = { qty: 1, unit: "d" };
var BAN_ABUSEIPDB   = { qty: 1, unit: "d" };
var BAN_UNKNOWNUSER = { qty: 1, unit: "d" };
var BAN_PRIORITY    = 20;

/**
 * Recipient probes (OnRecipientUnknown): bans an IP once a number of unknown
 * recipients is reached within a sliding window, using a counter file
 * (no database required).
 * @const {number} RCPT_UNKNOWN_THRESHOLD   unknown recipients per IP before ban
 * @const {number} RCPT_UNKNOWN_WINDOW_MIN  sliding window, in minutes, to reach the threshold
 * @const {{qty:number,unit:string}} RCPT_UNKNOWN_BAN  ban duration once the threshold is reached
 * @const {string} RCPT_DATA_FILE  counter file name (inside LOGDIR)
 * @const {string} RCPT_LOCK_FILE  lock file name (inside LOGDIR)
 * @const {number} LOCK_TRIES      number of lock acquisition attempts
 * @const {number} LOCK_STALE_SEC  age (s) beyond which a lock is considered orphaned
 */
var RCPT_UNKNOWN_THRESHOLD  = 3;
var RCPT_UNKNOWN_WINDOW_MIN = 10;
var RCPT_UNKNOWN_BAN        = { qty: 1, unit: "d" };
var RCPT_DATA_FILE = "rcpt_unknown.dat";
var RCPT_LOCK_FILE = "rcpt_unknown.lock";
var LOCK_TRIES     = 20;
var LOCK_STALE_SEC = 30;

/**
 * SMTP rejection of messages with a spam score that is too high, in OnAcceptMessage.
 * hMailServer's anti-spam tests (SpamAssassin included) run BEFORE this event, so the
 * score header is already present. Enabled via SPAM_REJECT_ENABLED (features panel);
 * check the header name produced by your configuration.
 * @const {number}  SPAM_REJECT_SCORE    score at or above which the message is rejected
 * @const {string}  SPAM_REJECT_HEADER   header holding the score (X-Spam-Score or X-Spam-Status)
 * @const {string}  SPAM_REJECT_MSG      SMTP refusal text (enhanced code, no "550"; the score is appended)
 */
var SPAM_REJECT_SCORE   = 15.0;
var SPAM_REJECT_HEADER  = "X-Spam-Score";
var SPAM_REJECT_MSG     = "5.7.1 Message rejected: spam score too high";

/**
 * Geographic filtering (2-letter ISO country codes in "|xx|yy|" format; "zz" = unknown country).
 * @const {string} BLOCKED_COUNTRIES  countries rejected immediately, all ports
 * @const {string} ALLOWED_GEO        only countries allowed on non-SMTP ports
 * @const {string} SMTP_PORTS         SMTP ports (transport), "|xx|yy|" format
 * @const {string} SUBMISSION_PORTS   submission/SSL ports, "|xx|yy|" format
 * @const {string} GEO_DNS_SUFFIX     domain of the DNS geolocation provider
 */
var BLOCKED_COUNTRIES = "|cn|cz|ru|";
var ALLOWED_GEO       = "|fr|zz|";
var SMTP_PORTS        = "|25|587|465|";
var SUBMISSION_PORTS  = "|587|465|";
var GEO_DNS_SUFFIX    = "country.junkemailfilter.com";

/**
 * Logging per source: set to false to silence a log category.
 * @const {object} LOG_SOURCES  source -> boolean table (GeoLookup, AutoBan, AbuseIPDB, Disconnect, Debug, System)
 */
var LOG_SOURCES = {
    GeoLookup:  true,
    AutoBan:    true,
    AbuseIPDB:  true,
    Disconnect: true,
    Debug:      true,
    Spam:       true,
    System:     true
};

/**
 * Log column widths (alignment guaranteed by space padding).
 * @const {number} COL_SOURCE  width of the source column
 * @const {number} COL_IP      width of the IP column
 * @const {number} COL_GEO     width of the country column
 * @const {number} COL_DUR     width of the duration column
 */
var COL_SOURCE = 10, COL_IP = 15, COL_GEO = 3, COL_DUR = 5;

/* ==========================  Utilities  ============================== */

/**
 * Trims leading/trailing whitespace.
 * @param {*} s  value to clean
 * @returns {string}
 */
function Trim(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

/**
 * JScript equivalent of VBScript DateAdd: adds a duration to a date.
 * @param {string} interval  "yyyy"=year,"m"=month,"d"=day,"h"=hour,"n"=minute,"s"=second
 * @param {number} n         quantity to add
 * @param {Date}   d         reference date
 * @returns {Date}
 */
function DateAdd(interval, n, d) {
    var r = new Date(d.getTime());
    switch (interval) {
        case "yyyy": r.setFullYear(r.getFullYear() + n); break;
        case "m":    r.setMonth(r.getMonth() + n);       break;
        case "d":    r.setDate(r.getDate() + n);         break;
        case "h":    r.setHours(r.getHours() + n);       break;
        case "n":    r.setMinutes(r.getMinutes() + n);   break;
        case "s":    r.setSeconds(r.getSeconds() + n);   break;
    }
    return r;
}

/**
 * Converts a JScript Date to an OLE Automation serial (VT_DATE; days since 1899-12-30).
 * @param {Date} d
 * @returns {number}
 */
function toOADate(d) {
    var localMs = d.getTime() - d.getTimezoneOffset() * 60000;
    return localMs / 86400000 + 25569;
}

/**
 * Sanitizes a string for an ASCII/UTF-8 log.
 * Security: escapes any non-printable or non-ASCII character (neutralizes
 * log injection via HELO/username); tab -> space (preserves columns).
 * @param {*} s
 * @returns {string} pure ASCII string
 */
function logSanitize(s) {
    s = String(s);
    var out = "", code, hex;
    for (var i = 0; i < s.length; i++) {
        code = s.charCodeAt(i);
        if (code === 9) {
            out += " ";
        } else if (code < 32 || code > 126) {
            hex = code.toString(16).toUpperCase();
            while (hex.length < 4) hex = "0" + hex;
            out += "\\u" + hex;
        } else {
            out += s.charAt(i);
        }
    }
    return out;
}

/**
 * Pads a string with spaces on the right up to the desired width.
 * Never truncates: an over-long value only shifts its own line.
 * @param {*} s
 * @param {number} w
 * @returns {string}
 */
function padRight(s, w) {
    s = String(s);
    while (s.length < w) s += " ";
    return s;
}

/**
 * Extracts a numeric score from a spam header. Handles a raw score ("15.2") as well
 * as the X-Spam-Status format ("Yes, score=15.2 required=5.0 ...").
 * @param {*} raw  header value
 * @returns {number} the score, or NaN if missing/unreadable
 */
function parseSpamScore(raw) {
    var s = Trim(raw);
    if (s === "") return NaN;
    var m = s.match(/score=(-?\d+(\.\d+)?)/i);
    if (m) return parseFloat(m[1]);
    return parseFloat(s);
}

/**
 * Writes a log line in fixed-width columns (alignment guaranteed by space
 * padding, no tab):
 *   YYYY-MM-DD HH:MM:SS  SOURCE  IP  GEO  DURATION  DETAIL
 * Daily UTF-8 file without BOM (FSO in ASCII mode, code 0). Best-effort:
 * retries on concurrent lock then falls back to EventLog.Write; never lets an
 * exception reach the mail flow.
 * @param {string} source  category (see LOG_SOURCES)
 * @param {string} ip      IP involved ("" -> "-")
 * @param {string} geo     country code ("" -> "-")
 * @param {string} dur     ban duration e.g. "+1d" ("" -> "-")
 * @param {*}      detail  free text (reason, error message...)
 */
function Log(source, ip, geo, dur, detail) {
    if (LOG_SOURCES[source] === false) return;

    ip     = (ip     === undefined || ip     === "") ? "-" : ip;
    geo    = (geo    === undefined || geo    === "") ? "-" : geo;
    dur    = (dur    === undefined || dur    === "") ? "-" : dur;
    detail = (detail === undefined) ? "" : detail;

    var d = new Date();
    function p2(n) { return (n < 10 ? "0" : "") + n; }
    var day  = d.getFullYear() + p2(d.getMonth() + 1) + p2(d.getDate());
    var ts   = d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate()) +
               " " + p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds());
    var path = LOGDIR + "\\" + LOGPREFIX + "_" + day + ".log";

    var line = ts + "  " +
               padRight(source, COL_SOURCE) + "  " +
               padRight(logSanitize(ip),  COL_IP)  + "  " +
               padRight(logSanitize(geo), COL_GEO) + "  " +
               padRight(logSanitize(dur), COL_DUR) + "  " +
               logSanitize(detail) + "\r\n";

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(LOGDIR)) { try { fso.CreateFolder(LOGDIR); } catch (eF) {} }

        var ForAppending = 8, done = false;
        for (var attempt = 0; attempt < 5 && !done; attempt++) {
            try {
                var f = fso.OpenTextFile(path, ForAppending, true, 0);
                f.Write(line);
                f.Close();
                done = true;
            } catch (eW) {}
        }
        if (!done) EventLog.Write("[Log fallback] " + line);
    } catch (e) {
        try { EventLog.Write("[Log error] " + e.description + " | " + line); } catch (e2) {}
    }
}

/* ====================  Security and inbound filtering  =============== */

/**
 * Returns whether an IP belongs to the exempt local network (LAN or localhost).
 * @param {string} ip
 * @returns {boolean}
 */
function isLocalIP(ip) {
    return ip.substring(0, LOCAL_IP_PREFIX.length) === LOCAL_IP_PREFIX || ip === LOCALHOST_IP;
}

/**
 * Validates that a string is a well-formed IP (v4/v6) before any shell call.
 * Security: blocks any command injection (spaces, ;, &, quotes...).
 * @param {string} ip
 * @returns {boolean}
 */
function isValidIP(ip) {
    var s = String(ip);
    var v4 = /^(25[0-5]|2[0-4]\d|1?\d?\d)(\.(25[0-5]|2[0-4]\d|1?\d?\d)){3}$/;
    if (v4.test(s)) return true;
    if (s.indexOf(":") === -1) return false;
    if (/[^0-9A-Fa-f:.]/.test(s)) return false;
    if ((s.match(/::/g) || []).length > 1) return false;
    if (s.replace(/[^:]/g, "").length > 7) return false;
    return true;
}

/**
 * Checks an IP's reputation via the AbuseIPDB COM component.
 * Robust: any error (missing component, network) is logged and treated as
 * "not listed" (fail-open) so legitimate mail is never blocked.
 * @param {string} strIP
 * @returns {boolean}  true if the IP must be blocked
 */
function ListedInAbuseIPDB(strIP) {
    try {
        var client = new ActiveXObject("AbuseIPDBComponent.AbuseIPDBRestClient");
        client.SetApiKey(ABUSEIPDB_APIKEY);
        client.SetMaxConfidenceScore(ABUSEIPDB_MAX_CONFIDENCE);
        client.SetMaxAgeInDays(ABUSEIPDB_MAX_AGE_DAYS);
        return client.BlockEndpoint(strIP) ? true : false;
    } catch (e) {
        Log("AbuseIPDB", strIP, "", "", "error - " + e.description);
        return false;
    }
}

/**
 * Resolves an IPv4's country code via junkemailfilter.com (reverse DNS TXT).
 * Near-pure function: logs only on error (System source); the normal GeoLookup
 * log is emitted at the block point in OnClientConnect.
 * Robust: any error returns "zz" (fail-open, allowed by ALLOWED_GEO).
 * @param {string} strIP
 * @returns {string}  country code (2 letters) or "zz" if unknown
 */
function GeoLookup(strIP) {
    var geo = "zz";
    try {
        var a = String(strIP).split(".");
        var name = a[3] + "." + a[2] + "." + a[1] + "." + a[0] + "." + GEO_DNS_SUFFIX;
        var resolver = new ActiveXObject("DNSLibrary.DNSResolver");
        var strLookup = resolver.TXT(name);

        if (Trim(strLookup) === "") return geo;

        var group = String(strLookup).split("\r\n");
        if (group.length > 1) {
            for (var i = 0; i < group.length; i++) {
                if (Trim(group[i]) !== "") {
                    if (geo === "zz") geo = Trim(group[i]);
                }
            }
        } else {
            geo = Trim(group[0]);
        }
    } catch (e) {
        Log("System", strIP, "", "", "GeoLookup error - " + e.description);
    }
    return geo;
}

/* ========================  Actions (ban / disconnect)  ============== */

/**
 * Adds a time-limited security range (ban) if it does not already exist.
 * Robust: any COM error (auth, add, save) is logged and not propagated.
 * @param {string} sIPAddress
 * @param {string} sReason    reason (used as unique name)
 * @param {number} iDuration  duration
 * @param {string} sType      DateAdd unit ("d","h","n"...)
 * @param {string} [geo]      country code (logging only)
 * @returns {boolean}  true if the ban was created
 */
function AutoBan(sIPAddress, sReason, iDuration, sType, geo) {
    var added = false;
    try {
        var oApp = new ActiveXObject("hMailServer.Application");
        oApp.Authenticate(ADMIN, PASSWORD);

        var ranges = oApp.Settings.SecurityRanges;
        var name = "(" + sReason + ") " + sIPAddress;

        var exists = false;
        try {
            var existing = ranges.ItemByName(name);
            if (existing != null) exists = true;
        } catch (eLookup) {
            exists = false;
        }

        if (!exists) {
            var r = ranges.Add();
            r.Name        = name;
            r.LowerIP     = sIPAddress;
            r.UpperIP     = sIPAddress;
            r.Priority    = BAN_PRIORITY;
            r.AllowSMTPConnections            = false;
            r.AllowPOP3Connections            = false;
            r.AllowIMAPConnections            = false;
            r.AllowDeliveryFromLocalToLocal   = false;
            r.AllowDeliveryFromLocalToRemote  = false;
            r.AllowDeliveryFromRemoteToLocal  = false;
            r.AllowDeliveryFromRemoteToRemote = false;
            r.Expires     = true;
            r.ExpiresTime = toOADate(DateAdd(sType, iDuration, new Date()));
            r.Save();
            added = true;
            Log("AutoBan", sIPAddress, geo, "+" + iDuration + sType, sReason);
        }
        ranges.Refresh();
    } catch (e) {
        Log("AutoBan", sIPAddress, geo, "", "error - " + e.description);
    }
    return added;
}

/**
 * Bans the IP according to a {qty, unit} config; qty<=0 => no ban (disconnect only).
 * @param {string} ip
 * @param {string} reason
 * @param {{qty:number,unit:string}} cfg
 * @param {string} [geo]  country code (logging only)
 */
function BanBy(ip, reason, cfg, geo) {
    if (cfg && cfg.qty > 0) AutoBan(ip, reason, cfg.qty, cfg.unit, geo);
}

/**
 * Bans (per cfg) then disconnects the current session. The disconnect is logged
 * only when no ban was applied (qty=0): a ban already emits an AutoBan line and
 * the disconnect is merely its logical follow-up.
 * @param {string} ip
 * @param {string} reason
 * @param {{qty:number,unit:string}} cfg
 * @param {string} [geo]
 */
function BanAndDisconnect(ip, reason, cfg, geo) {
    var willBan = (cfg && cfg.qty > 0);
    BanBy(ip, reason, cfg, geo);
    Disconnect(ip, reason, geo, !willBan);
}

/**
 * Disconnects an IP via Disconnect.exe and logs the action.
 * Robust: IP validated before the shell call (anti-injection) and shell errors logged.
 * @param {string} sIPAddress
 * @param {string} reason
 * @param {string} [geo]   already-resolved country code (avoids a redundant GeoLookup)
 * @param {boolean} [doLog] log the action (default true; false when a ban was already logged)
 */
function Disconnect(sIPAddress, reason, geo, doLog) {
    if (doLog === undefined) doLog = true;
    if (typeof geo === "undefined") geo = GeoLookup(sIPAddress);

    if (!isValidIP(sIPAddress)) {
        Log("Disconnect", sIPAddress, "", "", "skipped invalid IP");
        return;
    }
    try {
        var shell = new ActiveXObject("WScript.Shell");
        shell.Run("\"" + DISCONNECT_EXE + "\" " + sIPAddress, 0, true);
        if (doLog) Log("Disconnect", sIPAddress, geo, "", reason);
    } catch (e) {
        Log("System", sIPAddress, "", "", "Disconnect error - " + e.description);
    }
}

/* ==============  Recipient-probe counter (file + lock)  ============== */

/**
 * Acquires an exclusive lock via atomic creation of a .lock file.
 * Breaks a stale lock (> LOCK_STALE_SEC: dead thread) to avoid a permanent deadlock.
 * @param {object} fso       FileSystemObject instance
 * @param {string} lockPath  lock file path
 * @param {number} tries     number of attempts
 * @returns {boolean} true if the lock is acquired
 */
function acquireLock(fso, lockPath, tries) {
    for (var i = 0; i < tries; i++) {
        try {
            var lf = fso.CreateTextFile(lockPath, false);
            lf.Close();
            return true;
        } catch (e) {
            try {
                var f = fso.GetFile(lockPath);
                var ageSec = (new Date().getTime() - f.DateLastModified.getTime()) / 1000;
                if (ageSec > LOCK_STALE_SEC) fso.DeleteFile(lockPath);
            } catch (e2) {}
        }
    }
    return false;
}

/**
 * Releases the lock (deletes the .lock file).
 * @param {object} fso
 * @param {string} lockPath
 */
function releaseLock(fso, lockPath) {
    try { if (fso.FileExists(lockPath)) fso.DeleteFile(lockPath); } catch (e) {}
}

/**
 * Tolerant read of the counter file -> map ip -> {count, first}.
 * Ignores corrupt lines and purges out-of-window entries.
 * Line format: "ip<TAB>count<TAB>firstEpochSec".
 * @param {object} fso
 * @param {string} path
 * @param {number} now        current epoch seconds
 * @param {number} windowSec  window in seconds
 * @returns {object} map ip -> {count, first}
 */
function readCounter(fso, path, now, windowSec) {
    var map = {};
    try {
        if (!fso.FileExists(path)) return map;
        var f = fso.OpenTextFile(path, 1);
        var content = f.AtEndOfStream ? "" : f.ReadAll();
        f.Close();
        var lines = content.split("\r\n");
        for (var i = 0; i < lines.length; i++) {
            var ln = Trim(lines[i]);
            if (ln === "") continue;
            var p = ln.split("\t");
            if (p.length < 3) continue;
            var cnt = parseInt(p[1], 10), first = parseInt(p[2], 10);
            if (isNaN(cnt) || isNaN(first)) continue;
            if ((now - first) > windowSec) continue;
            map[p[0]] = { count: cnt, first: first };
        }
    } catch (e) {}
    return map;
}

/**
 * Rewrites the counter file (ASCII, no BOM) from the map.
 * @param {object} fso
 * @param {string} path
 * @param {object} map  ip -> {count, first}
 */
function writeCounter(fso, path, map) {
    try {
        var f = fso.CreateTextFile(path, true);
        for (var ip in map) {
            if (map.hasOwnProperty(ip)) {
                f.Write(ip + "\t" + map[ip].count + "\t" + map[ip].first + "\r\n");
            }
        }
        f.Close();
    } catch (e) {}
}

/**
 * Increments the unknown-recipient counter for the IP; bans at the threshold.
 * Everything is best-effort and under lock: never lets an exception reach the mail flow.
 * @param {string} ip
 */
function RcptUnknownCount(ip) {
    if (!isValidIP(ip)) return;

    var fso, gotLock = false;
    var dataPath = LOGDIR + "\\" + RCPT_DATA_FILE;
    var lockPath = LOGDIR + "\\" + RCPT_LOCK_FILE;
    try {
        fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(LOGDIR)) { try { fso.CreateFolder(LOGDIR); } catch (eF) {} }

        gotLock = acquireLock(fso, lockPath, LOCK_TRIES);
        if (!gotLock) { Log("Debug", ip, "", "", "rcpt-counter lock busy, skipped"); return; }

        var now = Math.round(new Date().getTime() / 1000);
        var windowSec = RCPT_UNKNOWN_WINDOW_MIN * 60;
        var map = readCounter(fso, dataPath, now, windowSec);

        var e = map[ip];
        if (e && (now - e.first) <= windowSec) { e.count = e.count + 1; }
        else                                   { e = { count: 1, first: now }; }
        map[ip] = e;

        if (e.count >= RCPT_UNKNOWN_THRESHOLD) {
            delete map[ip];
            writeCounter(fso, dataPath, map);
            releaseLock(fso, lockPath); gotLock = false;
            var geo = GeoLookup(ip);
            BanAndDisconnect(ip, "Rcpt probe", RCPT_UNKNOWN_BAN, geo);
        } else {
            writeCounter(fso, dataPath, map);
        }
    } catch (eMain) {
        try { Log("System", ip, "", "", "rcpt-counter error - " + eMain.description); } catch (e2) {}
    } finally {
        if (gotLock) releaseLock(fso, lockPath);
    }
}

/* ======================  hMailServer event triggers  ================= */

/**
 * Bans + disconnects unauthenticated connections. Excludes LAN and localhost.
 * @param {object} oClient  hMailServer client info
 */
function OnClientLogon(oClient) {
    try {
        if (!UNKNOWNUSER_ENABLED) return;
        var ip = String(oClient.IPAddress);
        if (isLocalIP(ip)) return;

        if (!oClient.Authenticated) {
            var geo = GeoLookup(ip);
            BanAndDisconnect(ip, "Unknown USER " + oClient.Port + " - " + oClient.Username, BAN_UNKNOWNUSER, geo);
        }
    } catch (e) {
        try { Log("System", "", "", "", "OnClientLogon error - " + e.description); } catch (e2) {}
    }
}

/**
 * Connection-time filtering (first excludes LAN 172.16.x and localhost):
 * 1) rejects countries in BLOCKED_COUNTRIES;
 * 2) on non-SMTP ports, only allows countries in ALLOWED_GEO;
 * 3) on submission ports (587/465), rejects IPs listed on AbuseIPDB.
 * Each rejection bans (per config) then disconnects.
 * @param {object} oClient  hMailServer client info
 */
function OnClientConnect(oClient) {
    try {
        var ip = String(oClient.IPAddress);
        if (isLocalIP(ip)) return;

        var port = String(oClient.Port);
        var geo  = (GEOBLOCK_ENABLED || GEORESTRICT_ENABLED || ABUSEIPDB_ENABLED) ? GeoLookup(ip) : "-";

        if (GEOBLOCK_ENABLED && BLOCKED_COUNTRIES.indexOf("|" + geo + "|") > -1) {
            Result.Value = 1;
            BanAndDisconnect(ip, "Bad Country " + port, BAN_BADCOUNTRY, geo);
            return;
        }

        if (GEORESTRICT_ENABLED && SMTP_PORTS.indexOf("|" + port + "|") === -1) {
            if (ALLOWED_GEO.indexOf("|" + geo + "|") === -1) {
                Result.Value = 1;
                BanAndDisconnect(ip, "Geo restrict " + port, BAN_GEORESTRICT, geo);
                return;
            }
        }

        if (ABUSEIPDB_ENABLED && SUBMISSION_PORTS.indexOf("|" + port + "|") > -1) {
            if (ListedInAbuseIPDB(ip)) {
                Result.Message = ABUSEIPDB_REJECT_MSG;
                Result.Value = 2;
                BanAndDisconnect(ip, "AbuseIPDB " + port, BAN_ABUSEIPDB, geo);
                return;
            }
        }
    } catch (e) {
        try { Log("System", "", "", "", "OnClientConnect error - " + e.description); } catch (e2) {}
    }
}

/**
 * Rejects, at SMTP level, messages whose spam score exceeds the threshold.
 * Inactive while SPAM_REJECT_ENABLED is false. Anti-spam tests run before this
 * event, so the score is already available in the headers.
 * @param {object} oClient
 * @param {object} oMessage
 */
function OnAcceptMessage(oClient, oMessage) {
    try {
        if (!SPAM_REJECT_ENABLED) return;

        var score = parseSpamScore(oMessage.HeaderValue(SPAM_REJECT_HEADER));
        if (!isNaN(score) && score >= SPAM_REJECT_SCORE) {
            Result.Value = 2;
            Result.Message = SPAM_REJECT_MSG + " (" + score + ")";
            var ip = oClient ? String(oClient.IPAddress) : "-";
            Log("Spam", ip, "", "", "rejected score=" + score);
        }
    } catch (e) {
        try { Log("System", "", "", "", "OnAcceptMessage error - " + e.description); } catch (e2) {}
    }
}

/**
 * Anonymizes the Received header of outbound mail (keeps only from "by " onward).
 * @param {object} oMessage
 */
function OnDeliveryStart(oMessage) {
    try {
        if (!RECEIVEDANON_ENABLED) return;
        var strReceived = oMessage.HeaderValue("Received");
        if (strReceived &&
            (strReceived.indexOf("ESMTPSA") !== -1 || strReceived.indexOf("ESMTPA") !== -1)) {
            var pos = strReceived.indexOf("by ");
            if (pos !== -1) {
                oMessage.HeaderValue("Received") = strReceived.substring(pos);
                oMessage.Save();
            }
        }
    } catch (e) {
        try { Log("System", "", "", "", "OnDeliveryStart error - " + e.description); } catch (e2) {}
    }
}

/**
 * Fills a missing Message-Id, then makes the X-Spam-Report header readable
 * (a line break before each asterisk).
 * @param {object} oMessage
 */
function OnDeliverMessage(oMessage) {
    try {
        if (MESSAGEID_ENABLED && oMessage.HeaderValue("Message-Id") === "") {
            var parts = String(oMessage.FromAddress).split("@");
            if (parts.length > 1 && parts[1] !== "") {
                var utils = new ActiveXObject("hMailServer.Utilities");
                var guid = utils.GenerateGUID();
                oMessage.HeaderValue("Message-Id") = "<" + guid.substr(1, 36) + "@" + parts[1] + ">";
                oMessage.Save();
            }
        }

        if (SPAMREPORT_ENABLED) {
            var spam = oMessage.HeaderValue("X-Spam-Report");
            if (spam !== "") {
                oMessage.HeaderValue("X-Spam-Report") = String(spam).replace(/\*/g, "\r\n *");
                oMessage.Save();
            }
        }
    } catch (e) {
        try { Log("System", "", "", "", "OnDeliverMessage error - " + e.description); } catch (e2) {}
    }
}

/**
 * Recipient probes: counts unknown RCPT TO per IP and bans at the threshold.
 * Excludes LAN, localhost and authenticated senders (a legitimate user's typo
 * must never count).
 * @param {object} oClient
 * @param {object} oMessage
 */
function OnRecipientUnknown(oClient, oMessage) {
    try {
        if (!RCPTPROBE_ENABLED) return;
        var ip = String(oClient.IPAddress);
        if (isLocalIP(ip)) return;
        if (oClient.Authenticated) return;

        RcptUnknownCount(ip);
    } catch (e) {
        try { Log("System", "", "", "", "OnRecipientUnknown error - " + e.description); } catch (e2) {}
    }
}
