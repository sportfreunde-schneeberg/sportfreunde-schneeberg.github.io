/* Vereinsausschuss-Portal der Sportfreunde Schneeberg 1928 e. V.
 *
 * Anmeldung mit dem Vereinskonto (MSAL), danach Lesen und Schreiben der Liegenschaften-Excel
 * über Microsoft Graph. Die Excel-Logik ist aus der iPhone-App übernommen
 * (SpfrLiegenschaften/Data/DataStore.swift und Model/*.swift):
 *   - alle Blätter: Kopfzeile in Zeile 4, Daten ab Zeile 5, Datumsangaben als Excel-Seriennummer
 *   - Ausleihliste: Spalte F (Listenpreis) ist eine VLOOKUP-Formel und wird nie geschrieben
 *   - Kabinenreinigung: Stunden × Satz, auf 5 € aufgerundet, plus Sonderzahlung; Zahlungsspalten I–K bleiben unberührt
 * Vor jedem Schreiben wird das Blatt frisch gelesen und geprüft, ob die Zielzeile noch passt.
 */
(function(){
"use strict";

var CFG = {
  clientId: "5c899010-742f-41cd-94a4-5135c670a30d",
  tenantId: "e5e88ea6-f980-4f32-ab46-41a7b04829be",
  host: "spfr1928.sharepoint.com",
  site: "/sites/TEAMLiegenschaften",
  /* Ordner der Verwaltungs-Excel. Genommen wird die höchste Fassung „Spfr_Liegenschaften_Verwaltung_vX.Y.xlsx“
     (Stand 23.11.2026: v1.4, laut Uebergabe/Export_Verwaltung_Projekt_Setup.md). */
  ordner: "General/Verwaltung",
  muster: /^Spfr_Liegenschaften_Verwaltung_v(\d+)\.(\d+)\.xlsx$/i
};
/* Hauptschalter fürs Schreiben (false = Portal nur lesend). Freigegeben nach Prüfung der Excel v1.4. */
var SCHREIBEN_FREI = true; /* freigegeben 23.11.2026 nach Test (Kabinenreinigung Zeile 35 geschrieben, geprüft, geleert) */
var SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All", "Calendars.ReadWrite", "Calendars.ReadWrite.Shared", "Tasks.ReadWrite", "Group.Read.All", "User.ReadBasic.All"];
var GR = "https://graph.microsoft.com/v1.0";
var ERSTE = 5; /* erste Datenzeile */
var MONATE = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];
var BLATT = { inventar:"Übersicht", reinigung:"Kabinenreinigung", typen:"Schlüsseltypen", schluessel:"Schlüsselliste", schloesser:"Schlösser" };
function ausleihBlatt(j){ return "Ausleihliste " + j; }

/* ---------- Hilfen ---------- */
var $ = function(id){ return document.getElementById(id); };
function el(tag, cls, text){ var e = document.createElement(tag); if(cls) e.className = cls; if(text != null) e.textContent = text; return e; }
function add(p){ for(var i = 1; i < arguments.length; i++){ var c = arguments[i]; if(c) p.appendChild(typeof c === "string" ? document.createTextNode(c) : c); } return p; }
function btn(cls, text, fn){ var b = el("button", "btn " + (cls || ""), text); b.type = "button"; if(fn) b.addEventListener("click", fn); return b; }
function txt(v){ return v == null ? "" : String(v).trim(); }
function zahl(v){ if(v === "" || v == null) return null; if(typeof v === "number") return v; var n = parseFloat(String(v).replace(/\./g, "").replace(",", ".")); return isNaN(n) ? null : n; }
function eingabeZahl(s){ s = String(s || "").trim(); if(!s) return null; var n = parseFloat(s.replace(",", ".")); return isNaN(n) ? null : n; }
function euro(v){ return v == null ? "–" : v.toLocaleString("de-DE", {style:"currency", currency:"EUR"}); }
function zahlText(v){ return v == null ? "" : String(v).replace(".", ","); }
function rund(v){ return Math.round(v * 100) / 100; }
function serieZuDatum(n){ return new Date(Math.round((n - 25569) * 86400000)); }
function datumZuSerie(d){ return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000 + 25569); }
function datumText(v){ if(typeof v === "number" && v > 20000) { var d = serieZuDatum(v); return pad(d.getUTCDate()) + "." + pad(d.getUTCMonth() + 1) + "." + d.getUTCFullYear(); } return txt(v); }
function pad(n){ return (n < 10 ? "0" : "") + n; }
function heuteISO(){ var d = new Date(); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
function isoZuDatum(s){ var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ""); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }
function isoZuText(s){ var d = isoZuDatum(s); return d ? pad(d.getDate()) + "." + pad(d.getMonth() + 1) + "." + d.getFullYear() : ""; }
function spalte(i){ var s = ""; i++; while(i > 0){ var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function zelle(werte, zeile, sp){ var r = werte[zeile - 1]; return r && sp < r.length ? r[sp] : ""; }
function leer(v){ return v === "" || v == null; }
function naechsteFreieZeile(werte, bisSpalte){
  for(var z = werte.length; z >= ERSTE; z--){ for(var s = 0; s <= bisSpalte; s++){ if(!leer(zelle(werte, z, s))) return z + 1; } }
  return ERSTE;
}
function zeigeToast(t){ var e = $("toast"); e.textContent = t; e.classList.add("on"); clearTimeout(zeigeToast.t); zeigeToast.t = setTimeout(function(){ e.classList.remove("on"); }, 3200); }

/* ---------- Anmeldung ---------- */
var pca = null, konto = null;
var TEST = (location.hostname === "localhost" || location.hostname === "127.0.0.1") && /[?&]test\b/.test(location.search);
function msalBereit(){
  if(pca) return true;
  if(!window.msal || location.protocol === "file:") return false;
  pca = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: location.origin + "/" },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: true }
  });
  return true;
}
function token(){
  return pca.acquireTokenSilent({scopes: SCOPES, account: konto})
    .catch(function(e){ /* Weiterleitung statt Popup: Popups scheitern auf dem Handy (In-App-Browser, Safari). */
      if(e && /interaction_required|consent_required|login_required|no_tokens_found|monitor_window_timeout/i.test(e.errorCode || "")){ pca.acquireTokenRedirect({scopes: SCOPES, account: konto, state: location.hash}); return new Promise(function(){}); }
      throw e; })
    .then(function(r){ return r.accessToken; });
}
function anmelden(){
  if(!msalBereit()){ fehlerZeigen(location.protocol === "file:" ? "Die Anmeldung funktioniert nur über die Web-Adresse, nicht als lokale Datei." : "Die Microsoft-Anmeldung konnte nicht geladen werden."); return; }
  /* Anmeldung per Weiterleitung (funktioniert auch auf dem Handy); die Seite kommt danach mit derselben Ansicht zurück. */
  pca.loginRedirect({scopes: SCOPES, prompt: "select_account", state: location.hash || "#/"})
    .catch(function(e){ if(e && e.errorCode === "user_cancelled") return; fehlerZeigen("Die Anmeldung hat nicht geklappt." + (e && e.errorMessage ? " (" + e.errorMessage.split("\n")[0] + ")" : "")); });
}
function abmelden(){ var k = konto; konto = null; kontoHash = ""; ids = null; sess = null; DATEN = null; kalInfo = null; prDrive = null; if(pca && k){ pca.logoutRedirect({account: k, postLogoutRedirectUri: location.origin + "/"}).catch(function(){}); return; } kopfzeile(); location.hash = "#/"; route(); }
function angemeldet(){ return TEST || !!konto; }
function kopfzeile(){ $("wer").textContent = TEST ? "Testmodus" : (konto ? (konto.name || konto.username || "") : ""); $("abmelden").hidden = !konto; }

/* ---------- Graph ---------- */
var ids = null, sess = null;
function warte(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
function gf(pfad, o, versuch){
  o = o || {}; versuch = versuch || 0;
  return token().then(function(t){
    var hd = {Authorization: "Bearer " + t}; if(o.headers) Object.keys(o.headers).forEach(function(k){ hd[k] = o.headers[k]; });
    if(o.body) hd["Content-Type"] = "application/json";
    if(o.sitzung) hd["workbook-session-id"] = o.sitzung; else if(sess && o.wb) hd["workbook-session-id"] = sess;
    return fetch(pfad.indexOf("https://") === 0 ? pfad : GR + pfad, {method: o.method || "GET", headers: hd, body: o.body ? JSON.stringify(o.body) : undefined});
  }).then(function(r){
    if((r.status === 429 || r.status === 503) && versuch < 5){ var w = (parseInt(r.headers.get("Retry-After"), 10) || 2) * 1000; return warte(Math.min(w, 30000)).then(function(){ return gf(pfad, o, versuch + 1); }); }
    if(!r.ok) return r.text().then(function(tx){ var e = new Error("Graph " + r.status); e.status = r.status; try{ e.detail = JSON.parse(tx).error.message; }catch(_){} throw e; });
    return r.status === 204 ? null : r.text().then(function(tx){ return tx ? JSON.parse(tx) : null; });
  });
}
function pfadK(p){ return p.split("/").map(encodeURIComponent).join("/"); }
function blattPfad(n){ return "/worksheets('" + encodeURIComponent(n.replace(/'/g, "''")) + "')"; }
function finden(){
  if(ids) return Promise.resolve(ids);
  return gf("/sites/" + CFG.host + ":" + CFG.site + "?$select=id").then(function(s){
    return gf("/sites/" + s.id + "/drive?$select=id"); }).then(function(d){
    return gf("/drives/" + d.id + "/root:/" + pfadK(CFG.ordner) + ":/children?$select=id,name&$top=200").then(function(l){
      var beste = null, rang = -1;
      (l.value || []).forEach(function(it){ var m = CFG.muster.exec(it.name); if(m){ var r = +m[1] * 1000 + +m[2]; if(r > rang){ rang = r; beste = it; } } });
      if(!beste){ var e = new Error("Datei"); e.status = 404; e.detail = "Im Ordner „" + CFG.ordner + "“ wurde keine Datei „Spfr_Liegenschaften_Verwaltung_vX.Y.xlsx“ gefunden"; throw e; }
      DATEINAME = beste.name;
      return ids = {drive: d.id, wb: "/drives/" + d.id + "/items/" + beste.id + "/workbook"}; }); });
}
var DATEINAME = "";
/* Schreibrechte je Bereich. Alle anderen Angemeldeten dürfen nur lesen.
   Damit im öffentlichen Code keine Adressen stehen, sind die Vereinskonten nur als SHA-256 von
   "spfr1928:" + Adresse (klein geschrieben) hinterlegt. Liste und Zuordnung: Uebergabe/README.md (privates Repo).
   Hinweis: Das gilt im Portal. Wer in Teams Bearbeitungsrechte auf die Excel hat, kann dort weiterhin direkt ändern. */
var RECHTE = {
  verwaltung: ["d7b9f3a32c445ded96bf3bcb146aeb0827ee68db2ca7cc12e34dd180f6e8ac75"],
  teamwear: ["b9653028cf8b498aa3747d9fa3f8e64e50073f7de9fa4b5b8d36f642ef9256bc", "d7b9f3a32c445ded96bf3bcb146aeb0827ee68db2ca7cc12e34dd180f6e8ac75"]
};
var kontoHash = "";
var fremdKonto = false, fehlerAnmeldung = "";
/* Nur Vereinskonten (…@1928schneeberg.de). Gäste aus Teams (z. B. #EXT#) werden abgewiesen und abgemeldet. */
function vereinskonto(acc){ return /@1928schneeberg\.de$/i.test(String(acc && acc.username || "")); }
function kontoSetzen(acc){
  if(acc && !vereinskonto(acc)){ fremdKonto = true; acc = null; }
  else if(acc) fremdKonto = false;
  konto = acc || null; kontoHash = "";
  if(!konto || !window.crypto || !crypto.subtle) return Promise.resolve();
  var daten = new TextEncoder().encode("spfr1928:" + String(konto.username || "").toLowerCase());
  return crypto.subtle.digest("SHA-256", daten).then(function(h){
    kontoHash = Array.prototype.map.call(new Uint8Array(h), function(b){ return ("0" + b.toString(16)).slice(-2); }).join(""); }).catch(function(){});
}
function darf(bereich){
  if(TEST){ var rolle = (/[?&]rolle=([a-z]+)/.exec(location.search) || [])[1] || "verwaltung"; return rolle === bereich; }
  return !!konto && !!kontoHash && (RECHTE[bereich] || []).indexOf(kontoHash) > -1;
}
function nurLesenHinweis(bereich){ return el("div", "hinweis", bereich === "teamwear" ? "Du kannst hier nur lesen. Abhaken kann nur der Teamwear-Verantwortliche." : "Du kannst hier nur lesen. Eintragen kann nur die Verwaltung – bitte Änderungen dort melden."); }
/* Gibt true zurück und zeigt einen Hinweis, wenn nicht geschrieben werden darf (Sperre oder fehlendes Recht). */
function gesperrt(f, bereich){
  bereich = bereich || "verwaltung";
  if(!darf(bereich)){ add(f, nurLesenHinweis(bereich)); return true; }
  if(TEST || SCHREIBEN_FREI) return false;
  add(f, el("div", "hinweis", "Eintragen ist im Portal noch gesperrt, bis es an die neue Excel-Fassung angepasst ist. Bitte vorerst wie bisher eintragen.")); return true; }
function ausrichten(adr, values){
  var m = /!?\$?([A-Z]+)\$?(\d+)/.exec((adr || "").split("!").pop() || ""); if(!m) return values;
  var c0 = 0; for(var i = 0; i < m[1].length; i++) c0 = c0 * 26 + (m[1].charCodeAt(i) - 64); c0--;
  var r0 = +m[2] - 1; if(!c0 && !r0) return values;
  var out = []; for(var r = 0; r < r0; r++) out.push([]);
  values.forEach(function(a){ var b = []; for(var c = 0; c < c0; c++) b.push(""); out.push(b.concat(a)); });
  return out;
}
function sitzung(I){ return sess ? Promise.resolve() : gf(I.wb + "/createSession", {method: "POST", body: {persistChanges: true}}).then(function(a){ sess = a.id; }); }

var Live = {
  namen: function(){ return finden().then(function(I){ return gf(I.wb + "/worksheets?$select=name", {wb: 1}); }).then(function(l){ return l.value.map(function(x){ return x.name; }); }); },
  blatt: function(n){ return finden().then(function(I){ return gf(I.wb + blattPfad(n) + "/usedRange(valuesOnly=true)?$select=address,values", {wb: 1}); }).then(function(b){ return ausrichten(b.address, b.values || []); }); },
  /* werte: 2D-Liste; null lässt eine Zelle unverändert, "" leert sie. */
  schreiben: function(n, adresse, werte, formate){
    var einmal = function(){ return finden().then(function(I){ return sitzung(I).then(function(){
      var body = {values: werte}; if(formate) body.numberFormat = formate;
      return gf(I.wb + blattPfad(n) + "/range(address='" + adresse + "')", {method: "PATCH", wb: 1, body: body}); }); }); };
    return einmal().catch(function(e){ if(sess && e.status && e.status !== 401 && e.status !== 403){ sess = null; return einmal(); } throw e; });
  }
};

/* Testquelle: nur auf localhost mit ?test. Lädt die Fixture der iPhone-App, schreibt in den Speicher
 * und protokolliert jeden Schreibaufruf in window.__schreib. Die Fixture wird nie veröffentlicht. */
var Test = {
  db: null,
  laden: function(){ if(Test.db) return Promise.resolve(Test.db);
    return fetch("/SpfrLiegenschaften/Resources/fixture.json").then(function(r){ return r.json(); }).then(function(j){ Test.db = {}; Object.keys(j.sheets).forEach(function(k){ Test.db[k] = ausrichten(j.sheets[k].address, j.sheets[k].values); }); return Test.db; }); },
  namen: function(){ return Test.laden().then(function(db){ return Object.keys(db); }); },
  blatt: function(n){ return Test.laden().then(function(db){ return JSON.parse(JSON.stringify(db[n] || [])); }); },
  schreiben: function(n, adresse, werte){ Test.db[n] = Test.db[n] || []; return Test.schreibenIn(Test.db[n], n, adresse, werte); },
  schreibenIn: function(b, n, adresse, werte){
    window.__schreib = window.__schreib || []; window.__schreib.push({blatt: n, adresse: adresse, werte: werte});
    var m = /^([A-Z]+)(\d+)/.exec(adresse), c0 = 0; for(var i = 0; i < m[1].length; i++) c0 = c0 * 26 + (m[1].charCodeAt(i) - 64); c0--; var r0 = +m[2];
    werte.forEach(function(zeile, dz){ var z = r0 + dz; while(b.length < z) b.push([]); var row = b[z - 1]; zeile.forEach(function(v, ds){ if(v === null) return; while(row.length <= c0 + ds) row.push(""); row[c0 + ds] = v; }); });
    return Promise.resolve();
  }
};
function Q(){ return TEST ? Test : Live; }

/* ---------- Daten der Verwaltung ---------- */
var DATEN = null; /* {namen:[...], blaetter:{name:werte}} */
var WICHTIG = [BLATT.inventar, BLATT.reinigung, BLATT.typen, BLATT.schluessel, BLATT.schloesser];
function laden(neu){
  if(DATEN && !neu) return Promise.resolve(DATEN);
  return Q().namen().then(function(namen){
    var holen = namen.filter(function(n){ return WICHTIG.indexOf(n) > -1 || /^Ausleihliste \d{4}$/.test(n); });
    var out = {namen: namen, blaetter: {}};
    return Promise.all(holen.map(function(n){ return Q().blatt(n).then(function(w){ out.blaetter[n] = w; }); })).then(function(){ DATEN = out; return out; });
  });
}
function frisch(n){ return Q().blatt(n).then(function(w){ if(DATEN) DATEN.blaetter[n] = w; return w; }); }

function inventar(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var b = txt(zelle(w, z, 1)); if(!b) continue;
  out.push({zeile: z, nr: zahl(zelle(w, z, 0)), bez: b, anzahl: zahl(zelle(w, z, 2)), anzahlRoh: zelle(w, z, 2), ort1: txt(zelle(w, z, 3)), ort2: txt(zelle(w, z, 4)), gebuehr: zahl(zelle(w, z, 5)), wer: txt(zelle(w, z, 6)), bem: txt(zelle(w, z, 7))}); }
  return out; }
function reinigung(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var m = txt(zelle(w, z, 1)); if(!m) continue;
  out.push({zeile: z, jahr: zahl(zelle(w, z, 0)), monat: m, stunden: zahl(zelle(w, z, 2)), satz: zahl(zelle(w, z, 3)), summe: zahl(zelle(w, z, 4)), gerundet: zahl(zelle(w, z, 5)), sonder: zahl(zelle(w, z, 6)), gesamt: zahl(zelle(w, z, 7)),
    z1: txt(zelle(w, z, 8)), z2: txt(zelle(w, z, 9)), z3: txt(zelle(w, z, 10)), bem: txt(zelle(w, z, 11))}); }
  return out; }
function positionen(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var wer = txt(zelle(w, z, 3)), g = txt(zelle(w, z, 4)); if(!wer && !g) continue;
  /* v1.3+: Spalte L = Anzahl (leer = 1). F ist der Einzel-Listenpreis, G der Gesamtpreis der Zeile. */
  var anz = zahl(zelle(w, z, 11)) || 1, liste = zahl(zelle(w, z, 5));
  out.push({zeile: z, nr: zahl(zelle(w, z, 0)), sub: zahl(zelle(w, z, 1)), datum: zelle(w, z, 2), wer: wer, gegenstand: g, anzahl: anz, liste: liste != null ? liste * anz : null, preis: zahl(zelle(w, z, 6)),
    bezahlt: txt(zelle(w, z, 7)), am: zelle(w, z, 8), bei: txt(zelle(w, z, 9)), bem: txt(zelle(w, z, 10))}); }
  return out; }
/* Höchstmenge je Gegenstand: nur wenn im Inventar eine reine Zahl steht (wie die GUI v1.4). */
function hoechstmenge(i){ return i && typeof i.anzahlRoh === "number" && i.anzahlRoh > 0 ? i.anzahlRoh : null; }
function buchungen(w){ var g = {}; positionen(w).forEach(function(p){ if(p.nr == null) return; (g[p.nr] = g[p.nr] || []).push(p); });
  return Object.keys(g).map(function(k){ var ps = g[k].sort(function(a, b){ return (a.sub || 0) - (b.sub || 0); });
    return {nr: +k, datum: ps[0].datum, wer: ps[0].wer, pos: ps,
      summe: ps.reduce(function(s, p){ return s + (p.preis || 0); }, 0),
      liste: ps.reduce(function(s, p){ return s + (p.liste != null ? p.liste : (p.preis || 0)); }, 0),
      bezahlt: ps.every(function(p){ return /^ja/i.test(p.bezahlt); }),
      abgesagt: ps.every(function(p){ return /abgesagt/i.test(p.bem); })}; })
    .sort(function(a, b){ return b.nr - a.nr; }); }
function schluessel(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var nn = txt(zelle(w, z, 1)), vn = txt(zelle(w, z, 2)); if(!nn && !vn) continue;
  out.push({zeile: z, nr: zahl(zelle(w, z, 0)), nn: nn, vn: vn, name: [nn, vn].filter(Boolean).join(", "), funktion: txt(zelle(w, z, 3)), status: txt(zelle(w, z, 4)), gen: zahl(zelle(w, z, 5)) || 0, t1: zahl(zelle(w, z, 6)) || 0, t2: zahl(zelle(w, z, 7)) || 0,
    sonst: txt(zelle(w, z, 8)), best: datumText(zelle(w, z, 9)), aus: datumText(zelle(w, z, 10)), bem: txt(zelle(w, z, 11))}); }
  return out; }
function schloesser(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var t = txt(zelle(w, z, 2)); if(!t) continue;
  out.push({standort: txt(zelle(w, z, 1)), tuer: t, typ: txt(zelle(w, z, 3)), ort: txt(zelle(w, z, 4)), bem: txt(zelle(w, z, 8))}); }
  return out; }
function typen(w){ var out = {}; for(var z = ERSTE; z <= w.length; z++){ var n = txt(zelle(w, z, 1)); if(n) out[n] = txt(zelle(w, z, 2)); } return out; }
function kennzahlen(w, sB, sW, von, bis){ var out = []; for(var z = von; z <= Math.min(bis, w.length); z++){ var b = txt(zelle(w, z, sB)), v = zelle(w, z, sW); if(b && !leer(v)) out.push([b, v]); } return out; }

/* Reinigung rechnen wie DataStore.reinigungRechnen */
function reinigungRechnen(stunden, satz, sonder){ var summe = stunden * satz, gerundet = Math.ceil(summe / 5 - 1e-9) * 5; return {summe: summe, gerundet: gerundet, gesamt: gerundet + (sonder || 0)}; }
/* Rabatt anteilig am Listenpreis verteilen, Rest auf Position 1 (wie BuchungBearbeitenView.verteileRabatt) */
function rabattVerteilen(pos, rabatt){
  var basis = pos.reduce(function(s, p){ return s + (p.liste != null ? p.liste : (p.preis || 0)); }, 0);
  if(basis <= 0) return pos.map(function(p){ return p.preis; });
  var faktor = Math.max(0, (basis - rabatt) / basis), neu = pos.map(function(p){ return rund((p.liste != null ? p.liste : (p.preis || 0)) * faktor); });
  var rest = (basis - rabatt) - neu.reduce(function(s, v){ return s + v; }, 0);
  if(Math.abs(rest) >= 0.005 && neu.length) neu[0] = rund(neu[0] + rest);
  return neu;
}

/* ---------- Seiten ---------- */
var seite = function(){ return $("seite"); };
function fehlerZeigen(t){ var s = seite(); s.textContent = ""; add(s, el("div", "fehler", t)); }
function fehlerText(e){
  if(e && (e.status === 403 || e.status === 404)) return "Kein Zugriff auf die Liegenschaften-Excel. Du musst Mitglied im Team „TEAM - Liegenschaften“ sein – bitte bei der Verwaltung melden.";
  if(e && e.status === 401) return "Die Anmeldung ist abgelaufen. Bitte melde dich neu an.";
  if(e && e.status === 423) return "Die Excel ist gerade gesperrt. Bitte in einer Minute noch einmal versuchen.";
  return "Das hat nicht geklappt" + (e && e.detail ? ": " + e.detail : (e && e.message ? ": " + e.message : ".")) ;
}
function pfad(teile){ var p = $("pfad"); p.textContent = ""; var vorher = teile.length > 1 ? teile[teile.length - 2] : null;
  if(vorher && vorher[1]){ var z = el("a", "zurueck", "‹ Zurück"); z.href = vorher[1]; add(p, z); }
  teile.forEach(function(t, i){ if(i) add(p, el("span", null, "›")); if(t[1]){ var a = el("a", null, t[0]); a.href = t[1]; add(p, a); } else add(p, el("span", null, t[0])); }); }
function kachel(sym, titel, text, href, gross, extern){ var a = el("a", "kachel" + (gross ? " gross" : "")); a.href = href; if(extern) a.rel = "noopener"; add(a, el("span", "sym", sym), el("b", null, titel), el("span", null, text)); return a; }
function ladeHinweis(s){ s.textContent = ""; add(s, el("div", "laden", "Daten werden aus der Excel geladen …")); }

function seiteWillkommen(s){
  pfad([]);
  var w = el("section", "willkommen");
  var img = el("img"); img.src = "img/wappen.png"; img.alt = ""; img.width = 110; img.height = 110;
  add(w, img, el("h1", null, "Vereinsausschuss-Portal"),
    el("p", "lead", "Interner Bereich des Vereinsausschusses. Melde dich mit deinem Vereinskonto an (…@1928schneeberg.de)."),
    fremdKonto ? el("div", "fehler", "Dieses Konto ist kein Vereinskonto. Bitte mit einer Adresse …@1928schneeberg.de anmelden.") : null,
    fehlerAnmeldung ? el("div", "fehler", fehlerAnmeldung) : null,
    btn("primary", "Mit Vereinskonto anmelden", anmelden));
  add(s, w);
}
function seiteStart(s){
  pfad([["Start"]]);
  add(s, el("h1", null, "Hallo" + (konto && konto.name ? " " + konto.name.split(" ")[0] : "") + "!"), el("p", "lead", "Wähle einen Bereich."));
  var k = el("div", "kacheln");
  add(k, kachel("🏠", "Liegenschaften", "Kabinenreinigung, Ausleihliste, Schlüssel, Inventar, Aufgaben und Bewässerung", "#/verwaltung", true),
    kachel("🎉", "Festausschuss", "Festplaner 100 Jahre Sportfreunde (2.–5. Juni 2028)", "/festplaner/", true),
    kachel("👕", "Teamwear-Status", "Bestellungen und ihr Stand: bestellt, ausgehändigt, bezahlt", "#/teamwear", true),
    kachel("📄", "Protokolle", "Ausschusssitzungen, Klausurtagungen und Generalversammlung nach Jahr und Monat", "#/protokolle", true),
    kachel("📅", "Vereinskalender", "Termine wie Ausschusssitzungen ansehen, eintragen, ändern und löschen", "#/kalender", true),
    kachel("🎪", "Veranstaltungen", "Schlachtfest, Ortsmeisterschaft, Kerb, Weihnachtsfeier: Unterlagen und Dienstpläne je Jahr", "#/veranstaltungen", true),
    kachel("🧒", "Jugend", "Alle Jugendmannschaften mit Trainingszeiten und Links", "#/jugend", true),
    kachel("🥅", "Aktive", "1. und 2. Mannschaft mit Trainingszeiten und Links", "#/aktive", true));
  add(s, k);
}
function seiteVerwaltung(s){
  pfad([["Start", "#/"], ["Liegenschaften"]]);
  add(s, el("h1", null, "Liegenschaften"), el("p", "lead", "Die Daten stehen in der Liegenschaften-Excel im Team „TEAM - Liegenschaften“. Änderungen hier landen direkt dort."));
  var k = el("div", "kacheln");
  add(k, kachel("🧽", "Kabinenreinigung", "Monat eintragen, Stunden und Beträge", "#/verwaltung/reinigung"),
    kachel("📦", "Ausleihliste", "Buchungen anlegen, Preise, Rabatt und Zahlung", "#/verwaltung/ausleihe"),
    kachel("🔑", "Schlüssel", "Wer hat welchen Schlüssel, Schlösser", "#/verwaltung/schluessel"),
    kachel("🗂️", "Inventar", "Gegenstände, Lagerorte und Leihgebühren", "#/verwaltung/inventar"),
    kachel("✅", "Aufgaben", "Offene Aufgaben aus Planner und der iCloud-Liste für die monatliche Ausschusssitzung", "#/verwaltung/aufgaben"),
    kachel("💧", "Bewässerungssteuerung", "Alter und neuer Sportplatz: Kamera und Bewässerung", "#/verwaltung/bewaesserung"));
  add(s, k);
  laden().then(function(D){
    if(DATEINAME) add(s, el("p", "lead", "Gelesen wird: " + DATEINAME + (SCHREIBEN_FREI ? "" : " · Eintragen ist im Portal noch gesperrt.")));
    var w = D.blaetter[BLATT.inventar]; if(!w) return;
    var kz = kennzahlen(w, 9, 10, 1, 12); if(!kz.length) return;
    var c = el("div", "karte"); add(c, el("b", null, "Kasse laut Übersicht"));
    var t = el("table", "tabelle"); kz.forEach(function(r){ var tr = el("tr"); add(tr, el("td", null, r[0]), el("td", "z", typeof r[1] === "number" ? euro(r[1]) : txt(r[1]))); add(t, tr); });
    add(c, t); add(s, c);
  }).catch(function(e){ add(s, el("div", "fehler", fehlerText(e))); });
}

/* --- Kabinenreinigung --- */
function seiteReinigung(s){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Kabinenreinigung"]]);
  ladeHinweis(s);
  laden().then(function(D){
    s.textContent = "";
    var w = D.blaetter[BLATT.reinigung]; if(!w){ add(s, el("div", "fehler", "Das Blatt „Kabinenreinigung“ fehlt in der Excel.")); return; }
    var liste = reinigung(w);
    add(s, el("h1", null, "Kabinenreinigung"), el("p", "lead", "Stunden × Satz, auf volle 5 € aufgerundet, plus Sonderzahlung. Bezahlt wird seit Mai 2026 in einem Schritt (Spalte K in der Excel)."));

    var heute = new Date(), letzter = liste[liste.length - 1];
    var f = el("form", "karte"); f.noValidate = true;
    var fJahr = el("input"); fJahr.type = "number"; fJahr.inputMode = "numeric"; fJahr.value = heute.getFullYear();
    var fMonat = el("select"); MONATE.forEach(function(m){ var o = el("option", null, m); o.value = m; add(fMonat, o); }); fMonat.value = MONATE[heute.getMonth()];
    var fStd = el("input"); fStd.inputMode = "decimal"; fStd.placeholder = "z. B. 7,5";
    var fSatz = el("input"); fSatz.inputMode = "decimal"; fSatz.value = zahlText(letzter && letzter.satz != null ? letzter.satz : 15);
    var fSonder = el("input"); fSonder.inputMode = "decimal"; fSonder.placeholder = "0";
    var fBem = el("input"); fBem.placeholder = "optional";
    function fld(t, i, cls){ var l = el("label", "fld" + (cls ? " " + cls : ""), t); add(l, i); return l; }
    var gr = el("div", "formular"); add(gr, fld("Jahr", fJahr), fld("Monat", fMonat), fld("Stunden", fStd), fld("Satz (€/h)", fSatz), fld("Sonderzahlung (€)", fSonder), fld("Bemerkung", fBem, "w2"));
    var vorschau = el("div", "summe"), info = el("div", "hinweis"); info.hidden = true;
    var speichern = el("button", "btn primary", "Monat speichern"); speichern.type = "submit";
    add(f, el("b", null, "Monat eintragen oder ändern"), gr, vorschau, info, add(el("div", "reihe"), speichern));
    if(darf("verwaltung")) add(s, f); else add(s, nurLesenHinweis("verwaltung"));

    function vorhanden(){ var j = parseInt(fJahr.value, 10); return liste.filter(function(r){ return r.jahr === j && r.monat === fMonat.value; })[0]; }
    var letzteWahl = "";
    function vorbelegen(){ var key = fJahr.value + "|" + fMonat.value; if(key === letzteWahl) return; letzteWahl = key; var v = vorhanden();
      if(v){ fStd.value = v.stunden != null ? zahlText(v.stunden) : ""; fSatz.value = v.satz != null ? zahlText(v.satz) : "15"; fSonder.value = v.sonder ? zahlText(v.sonder) : ""; fBem.value = v.bem || ""; }
      else { fStd.value = ""; fSonder.value = ""; fBem.value = ""; } }
    function neu(){ vorbelegen(); var st = eingabeZahl(fStd.value), sa = eingabeZahl(fSatz.value), so = eingabeZahl(fSonder.value) || 0;
      vorschau.textContent = "";
      if(st != null && sa != null){ var r = reinigungRechnen(st, sa, so); add(vorschau, add(el("span"), "Summe ", el("b", null, euro(r.summe))), add(el("span"), "Gerundet ", el("b", null, euro(r.gerundet))), add(el("span"), "Gesamt ", el("b", null, euro(r.gesamt)))); }
      var v = vorhanden(); info.hidden = !v; if(v){ info.textContent = v.monat + " " + v.jahr + " ist schon eingetragen (" + (v.stunden != null ? zahlText(v.stunden) + " h, " : "") + euro(v.gesamt) + "). Speichern überschreibt Stunden, Satz und Beträge; die Zahlungsspalten bleiben."; }
      speichern.textContent = v ? "Monat aktualisieren" : "Monat speichern"; }
    [fJahr, fMonat, fStd, fSatz, fSonder].forEach(function(i){ i.addEventListener("input", neu); i.addEventListener("change", neu); });
    neu();
    f.addEventListener("submit", function(ev){ ev.preventDefault(); if(gesperrt(f)) return;
      var j = parseInt(fJahr.value, 10), st = eingabeZahl(fStd.value), sa = eingabeZahl(fSatz.value), so = eingabeZahl(fSonder.value) || 0;
      if(!j || j < 2000 || j > 2100){ zeigeToast("Bitte ein gültiges Jahr eingeben."); return; }
      if(st == null || st < 0){ zeigeToast("Bitte die Stunden eintragen."); fStd.focus(); return; }
      if(sa == null || sa < 0){ zeigeToast("Bitte den Satz eintragen."); fSatz.focus(); return; }
      var r = reinigungRechnen(st, sa, so), monat = fMonat.value, bem = fBem.value.trim();
      speichern.disabled = true;
      frisch(BLATT.reinigung).then(function(w2){
        var aktuell = reinigung(w2).filter(function(x){ return x.jahr === j && x.monat === monat; })[0];
        var z = aktuell ? aktuell.zeile : naechsteFreieZeile(w2, 11);
        /* v1.4: E (Summe), F (Gerundet) und H (Gesamt) sind Formeln. Geschrieben werden A–D, G und L;
           E, F, H bekommen die Standardformel der Excel (repariert auch Zeilen mit festen Werten).
           I–K (Zahlungen) bleiben unverändert (null). */
        var zeile = [j, monat, st, sa,
          '=IF(C' + z + '="","",C' + z + '*IF(D' + z + '="",15,D' + z + '))',
          '=IF(C' + z + '="","",CEILING(E' + z + ',5))',
          so > 0 ? so : "",
          '=IF(AND(C' + z + '="",G' + z + '=""),"",N(F' + z + ')+N(G' + z + '))',
          null, null, null, bem ? bem : (aktuell && aktuell.bem ? "" : null)];
        return Q().schreiben(BLATT.reinigung, "A" + z + ":L" + z, [zeile]).then(function(){ return frisch(BLATT.reinigung); });
      }).then(function(){ zeigeToast(monat + " " + j + " gespeichert ✓"); route(); })
        .catch(function(e){ speichern.disabled = false; add(f, el("div", "fehler", fehlerText(e))); });
    });

    add(s, el("h2", null, "Bisherige Monate"));
    var t = el("table", "tabelle"), kopf = el("tr");
    ["Monat", "Stunden", "Satz", "Gesamt", "Zahlungen", "Bemerkung"].forEach(function(h, i){ add(kopf, el("th", i > 0 && i < 4 ? "z" : null, h)); });
    add(t, add(el("thead"), kopf)); var tb = el("tbody");
    liste.slice().reverse().forEach(function(r){
      /* Zahlweg laut Excel (Zeile 2): bis April 2026 zwei Schritte (I und J), ab Mai 2026 ein Schritt (K).
         Bezahlt ist ein Monat, wenn K = Ja oder (I und J) = Ja. */
      var ja = function(v){ return /^(ja|n\/a)/i.test(v); };
      var ok = ja(r.z3) || (ja(r.z1) && ja(r.z2));
      var tr = el("tr"); add(tr, el("td", null, r.monat + " " + (r.jahr || "")), el("td", "z", r.stunden != null ? zahlText(r.stunden) : ""), el("td", "z", r.satz != null ? euro(r.satz) : ""),
        el("td", "z", euro(r.gesamt)), add(el("td"), el("span", "chip " + (ok ? "ok" : "offen"), ok ? "bezahlt" : "offen")), el("td", null, r.bem));
      add(tb, tr); });
    add(t, tb); add(s, add(el("div", "scroll"), t));
  }).catch(function(e){ fehlerZeigen(fehlerText(e)); });
}

/* --- Ausleihliste --- */
var AUSWAHL_JAHR = null;
function seiteAusleihe(s, unter){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Ausleihliste"]]);
  ladeHinweis(s);
  laden().then(function(D){
    s.textContent = "";
    var jahre = D.namen.filter(function(n){ return /^Ausleihliste \d{4}$/.test(n); }).map(function(n){ return +n.slice(13); }).sort().reverse();
    if(!jahre.length){ add(s, el("div", "fehler", "In der Excel gibt es kein Blatt „Ausleihliste JJJJ“.")); return; }
    var jahr = AUSWAHL_JAHR && jahre.indexOf(AUSWAHL_JAHR) > -1 ? AUSWAHL_JAHR : jahre[0];
    var w = D.blaetter[ausleihBlatt(jahr)] || [], liste = buchungen(w);
    if(unter === "neu" && darf("verwaltung")){ formularNeueBuchung(s, D, jahre); return; }
    var bm = /^(\d+)$/.exec(unter || ""); if(bm && darf("verwaltung")){ var b = liste.filter(function(x){ return x.nr === +bm[1]; })[0]; if(b){ formularBuchung(s, b, jahr); return; } }

    add(s, el("h1", null, "Ausleihliste"));
    var kopf = el("div", "reihe"), sel = el("select"); sel.style.maxWidth = "160px";
    jahre.forEach(function(j){ var o = el("option", null, String(j)); o.value = j; add(sel, o); }); sel.value = jahr;
    sel.addEventListener("change", function(){ AUSWAHL_JAHR = +sel.value; route(); });
    var neuBtn = el("a", "btn primary", "+ Neue Buchung"); neuBtn.href = "#/verwaltung/ausleihe/neu"; if(!darf("verwaltung")) neuBtn = null;
    add(kopf, add(el("label", "fld", "Jahr"), sel), el("span", "sp"), neuBtn); add(s, kopf);
    var kz = kennzahlen(w, 12, 13, 1, 14);
    if(kz.length){ var c = el("div", "summe"); kz.forEach(function(r){ add(c, add(el("span"), r[0] + " ", el("b", null, typeof r[1] === "number" ? euro(r[1]) : txt(r[1])))); }); add(s, c); }
    if(!liste.length) add(s, el("p", "lead", "Noch keine Buchungen in " + jahr + "."));
    /* Tabelle wie in der GUI v1.4: eine Zeile je Position, Buchungen durch eine Linie getrennt. */
    if(liste.length){
      var schreibt = darf("verwaltung");
      var t = el("table", "tabelle"), kopf2 = el("tr");
      ["Nr.", "Sub", "Datum", "Wer", "Gegenstand", "Anz.", "Listenpreis", "Preis", "Bezahlt", "Bezahlt am", "bei", "Bemerkung", ""].forEach(function(h, i){ add(kopf2, el("th", i === 5 || i === 6 || i === 7 ? "z" : null, h)); });
      add(t, add(el("thead"), kopf2)); var tb = el("tbody");
      liste.forEach(function(b){
        b.pos.forEach(function(p, i){
          var tr = el("tr", i === 0 ? "gruppe" : null);
          var bez = /^ja/i.test(p.bezahlt) ? el("span", "chip ok", "Ja") : (b.abgesagt ? el("span", "chip grau", "abgesagt") : (p.bezahlt ? el("span", "chip offen", p.bezahlt) : null));
          var aktion = el("td");
          if(i === 0 && schreibt){ var ed = el("a", "btn klein", "Bearbeiten"); ed.href = "#/verwaltung/ausleihe/" + b.nr; add(aktion, ed); }
          add(tr, el("td", null, i === 0 ? String(b.nr) : ""), el("td", null, p.sub != null ? String(p.sub) : ""), el("td", null, i === 0 ? datumText(b.datum) : ""),
            el("td", null, i === 0 ? b.wer : ""), el("td", null, p.gegenstand || "–"), el("td", "z", String(p.anzahl || 1)),
            el("td", "z", p.liste != null ? euro(p.liste) : ""), el("td", "z", p.preis != null ? euro(p.preis) : ""), add(el("td"), bez),
            el("td", null, datumText(p.am)), el("td", null, p.bei), el("td", null, p.bem), aktion);
          add(tb, tr); }); });
      add(t, tb); add(s, add(el("div", "scroll"), t));
    }
  }).catch(function(e){ fehlerZeigen(fehlerText(e)); });
}
function formularNeueBuchung(s, D, jahre){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Ausleihliste", "#/verwaltung/ausleihe"], ["Neue Buchung"]]);
  var inv = inventar(D.blaetter[BLATT.inventar] || []).filter(function(i){ return !/nicht/i.test(i.wer); });
  add(s, el("h1", null, "Neue Buchung"), el("p", "lead", "Der Listenpreis kommt automatisch aus dem Inventar. Weicht der Preis ab, ist die Differenz der Rabatt."));
  var f = el("form", "karte"); f.noValidate = true;
  var fDatum = el("input"); fDatum.type = "date"; fDatum.value = heuteISO();
  var fWer = el("input"); fWer.placeholder = "Name oder Verein"; fWer.autocomplete = "off";
  var kopf = el("div", "formular"); add(kopf, add(el("label", "fld", "Datum"), fDatum), add(el("label", "fld w2", "Wer leiht aus?"), fWer));
  var posBox = el("div");
  function posZeile(){
    var z = el("div", "pos"), g = el("select"), o0 = el("option", null, "Gegenstand wählen"); o0.value = ""; add(g, o0);
    inv.forEach(function(i){ var o = el("option", null, i.bez + (i.gebuehr != null ? " (" + euro(i.gebuehr) + ")" : "")); o.value = i.bez; add(g, o); });
    var a = el("input"); a.type = "number"; a.inputMode = "numeric"; a.min = 1; a.value = 1;
    var p = el("input"); p.inputMode = "decimal"; p.placeholder = "Gesamtpreis €";
    var b = el("input"); b.placeholder = "Bemerkung (optional)";
    var hinweis = el("small"); hinweis.style.color = "var(--muted)";
    function artikel(){ return inv.filter(function(x){ return x.bez === g.value; })[0]; }
    function neuRechnen(){ var i = artikel(), max = hoechstmenge(i), n = parseInt(a.value, 10) || 1;
      a.max = max || ""; hinweis.textContent = max ? "höchstens " + max : "";
      if(max && n > max){ n = max; a.value = max; }
      p.value = i && i.gebuehr != null ? zahlText(rund(i.gebuehr * n)) : ""; summe(); }
    g.addEventListener("change", neuRechnen); a.addEventListener("input", neuRechnen);
    p.addEventListener("input", summe);
    var x = btn("klein", "Entfernen", function(){ if(posBox.children.length > 1){ z.remove(); summe(); } });
    var la = add(el("label", "fld", "Anzahl"), a, hinweis);
    add(z, add(el("label", "fld breit", "Gegenstand"), g), la, add(el("label", "fld", "Preis gesamt (€)"), p), add(el("label", "fld breit", "Bemerkung"), b), x);
    z._g = g; z._a = a; z._p = p; z._b = b; z._max = function(){ return hoechstmenge(artikel()); }; add(posBox, z); }
  var sumEl = el("div", "summe");
  function summe(){ var t = 0; [].forEach.call(posBox.children, function(z){ t += eingabeZahl(z._p.value) || 0; }); sumEl.textContent = ""; add(sumEl, add(el("span"), "Summe ", el("b", null, euro(t)))); }
  posZeile(); summe();
  var speichern = el("button", "btn primary", "Buchung speichern"); speichern.type = "submit";
  var abbrechen = el("a", "btn", "Abbrechen"); abbrechen.href = "#/verwaltung/ausleihe";
  add(f, kopf, el("h2", null, "Positionen"), posBox, btn("klein", "+ Weitere Position", posZeile), sumEl, add(el("div", "reihe"), speichern, abbrechen));
  add(s, f);
  f.addEventListener("submit", function(ev){ ev.preventDefault(); if(gesperrt(f)) return;
    var d = isoZuDatum(fDatum.value), wer = fWer.value.trim();
    var pos = [].map.call(posBox.children, function(z){ return {g: z._g.value, n: parseInt(z._a.value, 10) || 1, max: z._max(), p: eingabeZahl(z._p.value), b: z._b.value.trim()}; }).filter(function(x){ return x.g; });
    var zuViel = pos.filter(function(x){ return x.max && x.n > x.max; })[0];
    if(zuViel){ zeigeToast("Von „" + zuViel.g + "“ gibt es höchstens " + zuViel.max + "."); return; }
    if(!d){ zeigeToast("Bitte ein Datum wählen."); return; }
    if(!wer){ zeigeToast("Bitte eintragen, wer ausleiht."); fWer.focus(); return; }
    if(!pos.length){ zeigeToast("Bitte mindestens einen Gegenstand wählen."); return; }
    var blatt = ausleihBlatt(d.getFullYear());
    if(jahre.indexOf(d.getFullYear()) < 0){ add(f, el("div", "fehler", "In der Excel gibt es das Blatt „" + blatt + "“ noch nicht. Bitte zuerst in der Excel anlegen.")); return; }
    speichern.disabled = true;
    frisch(blatt).then(function(w){
      var nr = positionen(w).reduce(function(m, p){ return Math.max(m, p.nr || 0); }, 0) + 1, serie = datumZuSerie(d), kette = Promise.resolve(), werte = w;
      pos.forEach(function(p, i){ kette = kette.then(function(){
        var z = naechsteFreieZeile(werte, 10);
        /* A–E und G–L getrennt schreiben, F (Listenpreis, Formel) nie anfassen. L = Anzahl (v1.3+). */
        return Q().schreiben(blatt, "A" + z + ":E" + z, [[nr, i + 1, serie, wer, p.g]], [[null, null, "dd.mm.yyyy", null, null]])
          .then(function(){ return Q().schreiben(blatt, "G" + z + ":L" + z, [[p.p != null ? p.p : "", "Nein", "", "", p.b, p.n]]); })
          .then(function(){ return frisch(blatt); }).then(function(w2){ werte = w2; }); }); });
      return kette.then(function(){ return nr; });
    }).then(function(nr){ AUSWAHL_JAHR = d.getFullYear(); zeigeToast("Buchung Nr. " + nr + " gespeichert ✓"); location.hash = "#/verwaltung/ausleihe"; })
      .catch(function(e){ speichern.disabled = false; add(f, el("div", "fehler", fehlerText(e))); });
  });
}
function formularBuchung(s, b, jahr){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Ausleihliste", "#/verwaltung/ausleihe"], ["Nr. " + b.nr]]);
  add(s, el("h1", null, "Buchung Nr. " + b.nr), el("p", "lead", b.wer + " · " + datumText(b.datum)));
  var f = el("form", "karte"); f.noValidate = true;
  var felder = b.pos.map(function(p){ var i = el("input"); i.inputMode = "decimal"; i.value = p.preis != null ? zahlText(p.preis) : ""; return i; });
  var t = el("table", "tabelle"), kopf = el("tr"); ["Gegenstand", "Listenpreis", "Preis"].forEach(function(h, i){ add(kopf, el("th", i ? "z" : null, h)); }); add(t, add(el("thead"), kopf));
  var tb = el("tbody"); b.pos.forEach(function(p, i){ var tr = el("tr"); add(tr, el("td", null, (p.anzahl > 1 ? p.anzahl + " × " : "") + (p.gegenstand || "–")), el("td", "z", euro(p.liste)), add(el("td", "z"), felder[i])); add(tb, tr); }); add(t, tb);
  var fRabatt = el("input"); fRabatt.inputMode = "decimal"; fRabatt.placeholder = "0";
  var aktRabatt = b.liste - b.summe; if(aktRabatt > 0.005) fRabatt.value = zahlText(rund(aktRabatt));
  fRabatt.addEventListener("input", function(){ var r = eingabeZahl(fRabatt.value) || 0; var neu = rabattVerteilen(b.pos, r); neu.forEach(function(v, i){ felder[i].value = v != null ? zahlText(v) : ""; }); summe(); });
  felder.forEach(function(i){ i.addEventListener("input", summe); });
  var sumEl = el("div", "summe");
  function summe(){ var tt = felder.reduce(function(a, i){ return a + (eingabeZahl(i.value) || 0); }, 0); sumEl.textContent = ""; add(sumEl, add(el("span"), "Listenwert ", el("b", null, euro(b.liste))), add(el("span"), "Endbetrag ", el("b", null, euro(tt)))); }
  summe();
  var fBez = el("input"); fBez.type = "checkbox"; fBez.checked = b.bezahlt;
  var fAm = el("input"); fAm.type = "date"; fAm.value = heuteISO();
  var fBei = el("input"); fBei.placeholder = "z. B. Vorname"; fBei.value = (b.pos.filter(function(p){ return p.bei; })[0] || {}).bei || "";
  var zahlung = el("div", "formular"), lAm = add(el("label", "fld", "Bezahlt am"), fAm), lBei = add(el("label", "fld", "Bezahlt bei wem?"), fBei);
  var lBez = el("label", "fld"); lBez.style.flexDirection = "row"; lBez.style.alignItems = "center"; add(lBez, fBez, " Bezahlt");
  add(zahlung, lBez, lAm, lBei);
  function sichtbar(){ lAm.hidden = lBei.hidden = !fBez.checked; } fBez.addEventListener("change", sichtbar); sichtbar();
  if(b.bezahlt){ var am = (b.pos.filter(function(p){ return p.am; })[0] || {}).am; var m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(datumText(am)); if(m) fAm.value = m[3] + "-" + m[2] + "-" + m[1]; }
  var speichern = el("button", "btn primary", "Änderungen speichern"); speichern.type = "submit";
  var zurueck = el("a", "btn", "Abbrechen"); zurueck.href = "#/verwaltung/ausleihe";
  add(f, add(el("div", "scroll"), t), add(el("div", "formular"), add(el("label", "fld", "Rabatt gesamt (€)"), fRabatt)), sumEl,
    el("p", "lead", "Der Rabatt wird anteilig auf die Positionen verteilt. Beispiel Feuerwehr: Listenwert 50 €, Rabatt 20 € ergibt 30 €."),
    el("h2", null, "Zahlung"), zahlung, add(el("div", "reihe"), speichern, zurueck));
  add(s, f);
  f.addEventListener("submit", function(ev){ ev.preventDefault(); if(gesperrt(f)) return;
    /* v1.4: „Bezahlt am“ (I) ist ein echtes Datum (Seriennummer mit Datumsformat). */
    var amDatum = isoZuDatum(fAm.value);
    var blatt = ausleihBlatt(jahr), bez = fBez.checked, stempel = bez && amDatum ? datumZuSerie(amDatum) : "", bei = bez ? fBei.value.trim() : "";
    speichern.disabled = true;
    frisch(blatt).then(function(w){
      var aktuell = positionen(w), kette = Promise.resolve();
      b.pos.forEach(function(p, i){
        var jetzt = aktuell.filter(function(x){ return x.zeile === p.zeile; })[0];
        if(!jetzt || jetzt.nr !== p.nr || jetzt.sub !== p.sub){ throw new Error("Die Buchung wurde inzwischen in der Excel verändert. Bitte die Seite neu laden."); }
        var preis = eingabeZahl(felder[i].value);
        /* nur G–K schreiben; Bemerkung (K) bleibt wie sie ist. Ohne Haken und ohne bisherigen Eintrag
           bleiben H–J unverändert (sonst würde aus „leer“ ein „Nein“ und die offenen Posten ändern sich). */
        var warBezahlt = /^ja/i.test(p.bezahlt), unberuehrt = !bez && !p.bezahlt;
        var h = bez ? "Ja" : (unberuehrt ? null : "Nein"), iW = bez ? stempel : (warBezahlt ? "" : null), jW = bez ? bei : (warBezahlt ? "" : null);
        kette = kette.then(function(){ return Q().schreiben(blatt, "G" + p.zeile + ":K" + p.zeile, [[preis != null ? preis : "", h, iW, jW, null]], [[null, null, "dd.mm.yyyy", null, null]]); });
      });
      return kette.then(function(){ return frisch(blatt); });
    }).then(function(){ zeigeToast("Buchung Nr. " + b.nr + " gespeichert ✓"); location.hash = "#/verwaltung/ausleihe"; })
      .catch(function(e){ speichern.disabled = false; add(f, el("div", "fehler", e && e.status ? fehlerText(e) : (e.message || "Fehler beim Speichern."))); });
  });
}

/* --- Schlüssel --- */
function seiteSchluessel(s){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Schlüssel"]]);
  ladeHinweis(s);
  laden().then(function(D){
    s.textContent = "";
    var w = D.blaetter[BLATT.schluessel]; if(!w){ add(s, el("div", "fehler", "Das Blatt „Schlüsselliste“ fehlt in der Excel.")); return; }
    var leute = schluessel(w), farben = typen(D.blaetter[BLATT.typen] || []);
    var schreibt = darf("verwaltung");
    add(s, el("h1", null, "Schlüssel"), schreibt ? el("p", "lead", "Schlüsselträger ändern oder neu anlegen. Die Bestandsübersicht rechnet die Excel selbst.") : nurLesenHinweis("verwaltung"));
    if(schreibt) add(s, add(el("div", "reihe"), btn("primary", "+ Neuer Schlüsselträger", function(){ schluesselBearbeiten(null, leute); })));
    /* Bestandsübersicht aus N–Q. Nur Zeilen mit mindestens einer Zahl und nur Spalten, in denen etwas steht
       (in der Excel ist der Bestand noch nicht gepflegt, dann bleibt nur „Im Umlauf“). */
    var statZeilen = [], spaltenAn = [false, false, false];
    for(var z = 1; z <= Math.min(12, w.length); z++){ var b = txt(zelle(w, z, 13)), v = [zelle(w, z, 14), zelle(w, z, 15), zelle(w, z, 16)];
      if(!b || /bestand/i.test(txt(v[0]))) continue; if(!v.some(function(x){ return typeof x === "number"; })) continue;
      v.forEach(function(x, i){ if(!leer(x)) spaltenAn[i] = true; }); statZeilen.push([b, v]); }
    if(statZeilen.length){ var c = el("div", "karte"), t0 = el("table", "tabelle"), k0 = el("tr"), namen = ["Bestand", "Im Umlauf", "Frei"];
      add(k0, el("th", null, "Schlüssel")); namen.forEach(function(h, i){ if(spaltenAn[i]) add(k0, el("th", "z", h)); }); add(t0, add(el("thead"), k0));
      var tb0 = el("tbody"); statZeilen.forEach(function(r){ var tr = el("tr"); add(tr, el("td", null, r[0])); r[1].forEach(function(x, i){ if(spaltenAn[i]) add(tr, el("td", "z", txt(x))); }); add(tb0, tr); }); add(t0, tb0);
      add(c, add(el("div", "scroll"), t0)); add(s, c); }
    var such = el("input", "suche"); such.type = "search"; such.placeholder = "Name suchen …";
    var nurAktiv = el("input"); nurAktiv.type = "checkbox"; nurAktiv.checked = true;
    var la = el("label", "fld"); la.style.flexDirection = "row"; la.style.alignItems = "center"; add(la, nurAktiv, " nur aktive");
    add(s, add(el("div", "reihe"), such, la));
    var box = el("div", "scroll"); add(s, box);
    function zeigen(){ box.textContent = ""; var q = such.value.trim().toLowerCase();
      var t = el("table", "tabelle"), k = el("tr"); ["Name", "Funktion", "Status", "General", "Tür 1", "Tür 2", "Sonstige", "Bestätigt", ""].forEach(function(h, i){ add(k, el("th", i > 2 && i < 6 ? "z" : null, h)); }); add(t, add(el("thead"), k));
      var tb = el("tbody"), n = 0;
      leute.forEach(function(p){ if(nurAktiv.checked && !/^aktiv/i.test(p.status)) return; if(q && (p.name + " " + p.funktion).toLowerCase().indexOf(q) < 0) return; n++;
        var tr = el("tr"); add(tr, el("td", null, p.name), el("td", null, p.funktion), add(el("td"), el("span", "chip " + (/^aktiv/i.test(p.status) ? "ok" : "grau"), p.status || "–")),
          el("td", "z", p.gen || ""), el("td", "z", p.t1 || ""), el("td", "z", p.t2 || ""), el("td", null, p.sonst), el("td", null, p.best),
          add(el("td"), schreibt ? btn("klein", "Ändern", function(){ schluesselBearbeiten(p, leute); }) : null)); add(tb, tr); });
      add(t, tb); add(box, t); if(!n) add(box, el("p", "lead", "Keine Treffer.")); }
    such.addEventListener("input", zeigen); nurAktiv.addEventListener("change", zeigen); zeigen();
    var sl = schloesser(D.blaetter[BLATT.schloesser] || []);
    if(sl.length){ add(s, el("h2", null, "Schlösser – welche Tür braucht welchen Schlüssel"));
      var t2 = el("table", "tabelle"), k2 = el("tr"); ["Standort", "Tür", "Schlüssel", "Lagerort"].forEach(function(h){ add(k2, el("th", null, h)); }); add(t2, add(el("thead"), k2));
      var tb2 = el("tbody"); sl.forEach(function(x){ var tr = el("tr"), td = el("td"), pk = el("span", "punkt"); pk.style.background = /^#?[0-9a-f]{6}$/i.test(farben[x.typ] || "") ? (farben[x.typ][0] === "#" ? "" : "#") + farben[x.typ] : "transparent";
        add(td, pk, x.typ); add(tr, el("td", null, x.standort), el("td", null, x.tuer), td, el("td", null, x.ort)); add(tb2, tr); }); add(t2, tb2); add(s, add(el("div", "scroll"), t2)); }
  }).catch(function(e){ fehlerZeigen(fehlerText(e)); });
}

/* Schlüsselträger ändern oder neu anlegen. Geschrieben wird A–L der Schlüsselliste; N–Q (Bestand, Formeln) bleiben.
   Unveränderte Felder werden als null geschickt, damit z. B. Datumswerte ihren Zelltyp behalten. */
function schluesselBearbeiten(p, leute){
  var d = el("dialog"), f = el("form", "karte"); f.noValidate = true;
  var stati = ["Aktiv", "Zurückgegeben", "Warteliste"]; leute.forEach(function(x){ if(x.status && stati.indexOf(x.status) < 0) stati.push(x.status); });
  function feld(t, v, cls, typ){ var i = el("input"); if(typ) i.inputMode = typ; i.value = v == null ? "" : v; i._alt = i.value; var l = el("label", "fld" + (cls ? " " + cls : ""), t); add(l, i); l._i = i; return l; }
  var sel = el("select"); stati.forEach(function(x){ var o = el("option", null, x); o.value = x; add(sel, o); }); sel.value = p ? (p.status || "Aktiv") : "Aktiv"; sel._alt = sel.value;
  var fN = feld("Nachname", p && p.nn), fV = feld("Vorname", p && p.vn), fF = feld("Funktion", p && p.funktion, "w2"),
    fG = feld("Generalschlüssel", p && p.gen ? p.gen : "", null, "numeric"), fT1 = feld("Hütte Tür 1", p && p.t1 ? p.t1 : "", null, "numeric"), fT2 = feld("Hütte Tür 2", p && p.t2 ? p.t2 : "", null, "numeric"),
    fS = feld("Sonstige", p && p.sonst, "w2"), fB = feld("Bestätigt am/von", p ? p.best : "", null), fA = feld("Austritt am/von", p ? p.aus : "", null), fBe = feld("Bemerkung", p && p.bem, "w2");
  fB._i.placeholder = "z. B. 23.11.2026 Vorname"; fA._i.placeholder = "z. B. 23.11.2026 Vorname";
  var lS = add(el("label", "fld", "Status"), sel);
  var ok = el("button", "btn primary", "Speichern"); ok.type = "submit";
  add(f, el("h2", null, p ? "Schlüsselträger ändern" : "Neuer Schlüsselträger"), add(el("div", "formular"), fN, fV, fF, lS, fG, fT1, fT2, fS, fB, fA, fBe),
    add(el("div", "reihe"), ok, btn("", "Abbrechen", function(){ d.close(); d.remove(); })));
  add(d, f); document.body.appendChild(d); d.showModal(); fN._i.focus();
  f.addEventListener("submit", function(ev){ ev.preventDefault(); if(gesperrt(f)) return;
    var nn = fN._i.value.trim(), vn = fV._i.value.trim(); if(!nn && !vn){ zeigeToast("Bitte einen Namen eintragen."); return; }
    function wert(l, zahlFeld){ var v = l._i.value.trim(); if(p && v === String(l._i._alt).trim()) return null; if(!v) return ""; if(zahlFeld){ var n = eingabeZahl(v); return n == null ? "" : n; } return v; }
    ok.disabled = true;
    frisch(BLATT.schluessel).then(function(w2){
      var akt = schluessel(w2), z, nr;
      if(p){ var j = akt.filter(function(x){ return x.zeile === p.zeile; })[0]; if(!j || j.nn !== p.nn || j.vn !== p.vn) throw new Error("Die Zeile wurde inzwischen in der Excel verändert. Bitte die Seite neu laden."); z = p.zeile; nr = null; }
      else { z = naechsteFreieZeile(w2, 11); nr = akt.reduce(function(m, x){ return Math.max(m, x.nr || 0); }, 0) + 1; }
      var zeile = [nr, p ? wert(fN) : nn, p ? wert(fV) : vn, wert(fF), p && sel.value === sel._alt ? null : sel.value,
        wert(fG, true), wert(fT1, true), wert(fT2, true), wert(fS), wert(fB), wert(fA), wert(fBe)];
      return Q().schreiben(BLATT.schluessel, "A" + z + ":L" + z, [zeile]);
    }).then(function(){ return frisch(BLATT.schluessel); }).then(function(){ d.close(); d.remove(); zeigeToast((vn || nn) + " gespeichert ✓"); route(); })
      .catch(function(e){ ok.disabled = false; add(f, el("div", "fehler", e && e.status ? fehlerText(e) : (e.message || "Fehler beim Speichern."))); });
  });
}

/* --- Inventar --- */
function seiteInventar(s){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Inventar"]]);
  ladeHinweis(s);
  laden().then(function(D){
    s.textContent = "";
    var w = D.blaetter[BLATT.inventar]; if(!w){ add(s, el("div", "fehler", "Das Blatt „Übersicht“ fehlt in der Excel.")); return; }
    var liste = inventar(w);
    add(s, el("h1", null, "Inventar"), el("p", "lead", "Die Leihgebühr hier ist der Listenpreis in der Ausleihliste."));
    if(darf("verwaltung")) add(s, add(el("div", "reihe"), btn("primary", "+ Neuer Gegenstand", function(){ bearbeiten(null); }))); else add(s, nurLesenHinweis("verwaltung"));
    var t = el("table", "tabelle"), k = el("tr"); ["Nr.", "Bezeichnung", "Anzahl", "Lagerort", "Leihgebühr", "Ausleihe", ""].forEach(function(h, i){ add(k, el("th", i === 2 || i === 4 ? "z" : null, h)); }); add(t, add(el("thead"), k));
    var tb = el("tbody"); liste.forEach(function(i){ var tr = el("tr"); add(tr, el("td", null, i.nr != null ? String(i.nr) : ""), el("td", null, i.bez), el("td", "z", i.anzahl != null ? String(i.anzahl) : ""),
      el("td", null, [i.ort1, i.ort2].filter(Boolean).join(" / ")), el("td", "z", i.gebuehr != null ? euro(i.gebuehr) : ""), el("td", null, i.wer), add(el("td"), darf("verwaltung") ? btn("klein", "Ändern", function(){ bearbeiten(i); }) : null)); add(tb, tr); });
    add(t, tb); add(s, add(el("div", "scroll"), t));

    function bearbeiten(pos){
      var d = el("dialog"), f = el("form", "karte"); f.noValidate = true;
      function feld(t, v, cls, typ){ var i = el("input"); if(typ) i.inputMode = typ; i.value = v == null ? "" : v; i._alt = i.value; var l = el("label", "fld" + (cls ? " " + cls : ""), t); add(l, i); l._i = i; return l; }
      var fB = feld("Bezeichnung", pos && pos.bez, "w2"), fA = feld("Anzahl", pos ? (typeof pos.anzahlRoh === "number" ? pos.anzahlRoh : txt(pos.anzahlRoh)) : "", null, "numeric"), fO1 = feld("Lagerort 1", pos && pos.ort1), fO2 = feld("Lagerort 2", pos && pos.ort2),
        fG = feld("Leihgebühr (€)", pos && pos.gebuehr != null ? zahlText(pos.gebuehr) : "", null, "decimal"), fW = feld("Wer darf ausleihen?", pos ? pos.wer : "Alle Mitglieder", "w2"), fBe = feld("Bemerkung", pos && pos.bem, "w2");
      var ok = el("button", "btn primary", "Speichern"); ok.type = "submit";
      add(f, el("h2", null, pos ? "Gegenstand ändern" : "Neuer Gegenstand"), add(el("div", "formular"), fB, fA, fO1, fO2, fG, fW, fBe), add(el("div", "reihe"), ok, btn("", "Abbrechen", function(){ d.close(); d.remove(); })));
      add(d, f); document.body.appendChild(d); d.showModal(); fB._i.focus();
      f.addEventListener("submit", function(ev){ ev.preventDefault(); if(gesperrt(f)) return;
        var bez = fB._i.value.trim(); if(!bez){ zeigeToast("Bitte eine Bezeichnung eintragen."); return; }
        ok.disabled = true;
        frisch(BLATT.inventar).then(function(w2){
          var akt = inventar(w2), z, nr;
          if(pos){ var j = akt.filter(function(x){ return x.zeile === pos.zeile; })[0]; if(!j || j.bez !== pos.bez) throw new Error("Der Gegenstand wurde inzwischen in der Excel verändert. Bitte die Seite neu laden."); z = pos.zeile; nr = pos.nr; }
          else { z = naechsteFreieZeile(w2, 7); nr = akt.reduce(function(m, x){ return Math.max(m, x.nr || 0); }, 0) + 1; }
          /* Unveränderte Felder bleiben unangetastet (null). Geänderte Zahlenfelder: Zahl, wenn lesbar, sonst der Text. */
          function wert(l, zahlFeld){ var v = l._i.value.trim(); if(pos && v === String(l._i._alt).trim()) return null; if(!v) return ""; if(zahlFeld){ var n = eingabeZahl(v); return n == null ? v : n; } return v; }
          return Q().schreiben(BLATT.inventar, "A" + z + ":H" + z, [[pos ? null : nr, pos ? wert(fB) : bez, wert(fA, true), wert(fO1), wert(fO2), wert(fG, true), wert(fW), wert(fBe)]]);
        }).then(function(){ return frisch(BLATT.inventar); }).then(function(){ d.close(); d.remove(); zeigeToast("„" + bez + "“ gespeichert ✓"); route(); })
          .catch(function(e){ ok.disabled = false; add(f, el("div", "fehler", e && e.status ? fehlerText(e) : (e.message || "Fehler beim Speichern."))); });
      });
    }
  }).catch(function(e){ fehlerZeigen(fehlerText(e)); });
}

/* ---------- Teamwear-Status ----------
 * Excel „Teamwear_Bestellungen.xlsx“ im Team Vorstandschaft, Tabelle „Bestellungen“ (Kopf in Zeile 4).
 * Spalten: A Eingang · B Bestell-Nr · C Nachname · D Vorname · E E-Mail · F Telefon · G Mannschaft · H Artikel · I Größe
 * J Anzahl · K Initialen · L–N Preise (Formeln) · O Bestellt · P Geliefert · Q Ausgehändigt · R Bezahlt · S Storniert
 * T Status (Formel) · U Bemerkung. Geschrieben werden nur O–S und U, und nur vom Teamwear-Verantwortlichen.
 * Die Datei nie ersetzen (Power Automate hängt an der Datei-ID). */
var TW = { gruppe: "Vorstandschaft", datei: "General/Teamwear/Teamwear_Bestellungen.xlsx" };
var STUFEN = [["Bestellt", 14], ["Geliefert", 15], ["Ausgehändigt", 16], ["Bezahlt", 17], ["Storniert", 18]];
/* „Geliefert“ wird im Portal nicht mehr gezeigt (Entscheidung 23.11.2026, läuft intern). Damit die Status-Formel in Spalte T
   stimmt, setzt „Ausgehändigt“ die Spalte Geliefert (P) automatisch mit. */
var TW_SICHTBAR = [0, 2, 3, 4];
var twIds = null, twSess = null;
function twFinden(){
  if(twIds) return Promise.resolve(twIds);
  return gf("/groups?$filter=" + encodeURIComponent("displayName eq '" + TW.gruppe + "'") + "&$select=id").then(function(r){
    var g = (r.value || [])[0]; if(!g){ var e = new Error("Team"); e.status = 404; throw e; }
    return gf("/groups/" + g.id + "/drive?$select=id"); }).then(function(d){
    return gf("/drives/" + d.id + "/root:/" + pfadK(TW.datei) + "?$select=id").then(function(it){ return twIds = {wb: "/drives/" + d.id + "/items/" + it.id + "/workbook"}; }); });
}
var TwLive = {
  /* Lesen in derselben Excel-Sitzung wie das Schreiben, sonst kommt kurz nach dem Abhaken noch der alte Stand zurück. */
  lesen: function(){ return twFinden().then(function(I){ var pf = I.wb + "/tables('Bestellungen')/range?$select=address,values";
    return (twSess ? gf(pf, {sitzung: twSess}).catch(function(e){ if(e && (e.status === 401 || e.status === 403)) throw e; twSess = null; return gf(pf); }) : gf(pf)); }).then(function(b){ return ausrichten(b.address, b.values || []); }); },
  preise: function(){ return twFinden().then(function(I){ return gf(I.wb + "/tables('Artikel')/range?$select=values"); }).then(function(b){ return twPreisliste(b.values || []); }); },
  schreiben: function(adresse, werte){
    var einmal = function(){ return twFinden().then(function(I){
      return (twSess ? Promise.resolve() : gf(I.wb + "/createSession", {method: "POST", body: {persistChanges: true}}).then(function(a){ twSess = a.id; })).then(function(){
        return gf(I.wb + "/worksheets('Bestellungen')/range(address='" + adresse + "')", {method: "PATCH", sitzung: twSess, body: {values: werte}}); }); }); };
    return einmal().catch(function(e){ if(twSess && e.status && e.status !== 401 && e.status !== 403){ twSess = null; return einmal(); } throw e; });
  }
};
var TwTest = {
  db: null,
  lesen: function(){ if(TwTest.db) return Promise.resolve(JSON.parse(JSON.stringify(TwTest.db)));
    return fetch("/Teamwear/test.json").then(function(r){ return r.json(); }).then(function(j){ TwTest.db = j; return JSON.parse(JSON.stringify(j)); }); },
  schreiben: function(adresse, werte){ return Test.schreibenIn(TwTest.db, "Teamwear", adresse, werte); },
  preise: function(){ return Promise.resolve(twPreisliste([["Artikel", "Preis Kind", "Preis Erwachsene"], ["Jersey", 20, 22], ["Rucksack", 18, 18], ["„1928“ Shirt", 20, 20]])); }
};
function TWQ(){ return TEST ? TwTest : TwLive; }
function twPreisliste(v){ var out = {}; v.slice(1).forEach(function(r){ var n = txt(r[0]); if(n) out[n] = {kind: zahl(r[1]), erw: zahl(r[2])}; }); return out; }
function darfTeamwearSchreiben(){ return darf("teamwear"); }
function twZeilen(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var art = txt(zelle(w, z, 7)); if(!art) continue;
  out.push({zeile: z, eingang: zelle(w, z, 0), nr: txt(zelle(w, z, 1)), name: [txt(zelle(w, z, 3)), txt(zelle(w, z, 2))].filter(Boolean).join(" "), mannschaft: txt(zelle(w, z, 6)),
    artikel: art, groesse: txt(zelle(w, z, 8)), anzahl: zahl(zelle(w, z, 9)), initialen: txt(zelle(w, z, 10)), einzel: zahl(zelle(w, z, 11)), aufpreis: zahl(zelle(w, z, 12)), summe: zahl(zelle(w, z, 13)),
    stufen: STUFEN.map(function(s){ return /^ja/i.test(txt(zelle(w, z, s[1]))); }), status: txt(zelle(w, z, 19)).replace(/^geliefert$/, "bestellt"), bem: txt(zelle(w, z, 20))}); }
  return out; }
/* Status wie die Formel in Spalte T, damit die Anzeige nach dem Klicken sofort stimmt. */
function twStatus(st){ if(st[4]) return "storniert"; if(st[0] && st[1] && st[2] && st[3]) return "abgeschlossen";
  if(!st[0]) return "zu bestellen"; if(!st[2]) return "bestellt"; if(!st[3]) return "Zahlung offen"; return "abgeschlossen"; }
var TW_FILTER = "offen";
/* Lampenfarbe je Status: rot = noch nichts passiert, gelb = läuft, grün = fertig, grau = storniert */
function twLampe(st){ return st === "abgeschlossen" ? "gruen" : (st === "storniert" ? "grau" : (st === "zu bestellen" ? "rot" : "gelb")); }
/* Preisprüfung gegen das Blatt „Artikel“: Kindergrößen (reine Zahl bis 176) zum Kinderpreis, sonst Erwachsene; Initialen +4 €.
 * Gibt einen Hinweistext zurück, wenn Einzelpreis, Aufpreis oder Summe nicht zur Preisliste passen. */
function twPreisPruefen(r, preise){
  var p = preise && preise[r.artikel]; if(!p) return preise ? "Artikel nicht in der Preisliste" : "";
  var kind = /^\d+$/.test(r.groesse) && +r.groesse <= 176, soll = p.kind === p.erw ? p.kind : (kind ? p.kind : p.erw);
  if(!/^\d+$/.test(r.groesse) && /\d+\s*-\s*\d+/.test(r.groesse) && p.kind !== p.erw) soll = null; /* z. B. Sockengrößen: nicht eindeutig */
  var aufSoll = r.initialen && r.initialen !== "-" ? 4 : 0, n = r.anzahl || 1, f = [];
  if(soll != null && r.einzel != null && Math.abs(r.einzel - soll) > 0.005) f.push("Einzelpreis " + euro(r.einzel) + " statt " + euro(soll));
  if(r.aufpreis != null && Math.abs(r.aufpreis - aufSoll) > 0.005) f.push("Aufpreis Initialen " + euro(r.aufpreis) + " statt " + euro(aufSoll));
  var basis = r.einzel != null ? r.einzel : soll;
  if(basis != null && r.summe != null && Math.abs(r.summe - n * (basis + (r.aufpreis != null ? r.aufpreis : aufSoll))) > 0.005) f.push("Summe passt nicht zu Anzahl × Preis");
  return f.join(" · ");
}
function seiteTeamwear(s){
  pfad([["Start", "#/"], ["Teamwear-Status"]]);
  ladeHinweis(s);
  var darf = darfTeamwearSchreiben();
  Promise.all([TWQ().lesen(), TWQ().preise().catch(function(){ return null; })]).then(function(x){
    var w = x[0], preise = x[1];
    s.textContent = "";
    var zeilen = twZeilen(w);
    add(s, el("h1", null, "Teamwear-Status"),
      el("p", "lead", darf ? "Du kannst die Stufen hier abhaken. Jeder Klick wird sofort in der Excel gespeichert." : "Stand aller Bestellungen. Abhaken kann nur der Teamwear-Verantwortliche."));
    var c = el("div", "summe"); add(s, c);
    function zaehlen(){ var zaehl = {}; zeilen.forEach(function(r){ zaehl[r.status] = (zaehl[r.status] || 0) + 1; }); c.textContent = "";
      ["zu bestellen", "bestellt", "Zahlung offen", "abgeschlossen", "storniert"].forEach(function(k){ add(c, add(el("span"), el("i", "lampe " + twLampe(k)), k + " ", el("b", null, String(zaehl[k] || 0)))); }); }
    var filter = el("select"); filter.style.maxWidth = "220px";
    [["offen", "Alle offenen"], ["alle", "Alle (ohne stornierte)"], ["zu bestellen", "zu bestellen"], ["bestellt", "bestellt"], ["Zahlung offen", "Zahlung offen"], ["abgeschlossen", "abgeschlossen"], ["pruefen", "Preis prüfen"]]
      .forEach(function(o){ var op = el("option", null, o[1]); op.value = o[0]; add(filter, op); });
    filter.value = TW_FILTER;
    var such = el("input", "suche"); such.type = "search"; such.placeholder = "Name, Artikel oder Mannschaft …";
    add(s, add(el("div", "reihe"), add(el("label", "fld", "Anzeigen"), filter), such));
    var box = el("div", "twliste"), archiv = el("div", "twarchiv"); add(s, box, archiv);
    zeilen.forEach(function(r){ r.pruefen = twPreisPruefen(r, preise); });
    function passt(r, q){
      if(r.status === "storniert") return false; /* stornierte stehen nur unten im Archiv */
      if(TW_FILTER === "offen" && r.status === "abgeschlossen") return false;
      if(TW_FILTER === "pruefen" && !r.pruefen) return false;
      if(TW_FILTER !== "offen" && TW_FILTER !== "alle" && TW_FILTER !== "pruefen" && r.status !== TW_FILTER) return false;
      return !q || (r.name + " " + r.artikel + " " + r.mannschaft).toLowerCase().indexOf(q) > -1; }
    /* Bestellungen (gleiche Bestell-Nr) zusammenfassen, neueste zuerst */
    function gruppieren(liste){ var gruppen = [], idx = {};
      liste.slice().reverse().forEach(function(r){ var k = r.nr || ("z" + r.zeile);
        if(!(k in idx)){ idx[k] = gruppen.length; gruppen.push({r: r, pos: []}); } gruppen[idx[k]].pos.unshift(r); });
      return gruppen; }
    var ARCHIV_OFFEN = false;
    function zeigen(){ zaehlen(); box.textContent = ""; archiv.textContent = ""; TW_FILTER = filter.value; var q = such.value.trim().toLowerCase();
      var gruppen = gruppieren(zeilen.filter(function(r){ return passt(r, q); }));
      gruppen.forEach(function(g){ add(box, twBestellung(g, false)); });
      if(!gruppen.length) add(box, el("p", "lead", "Keine Bestellungen in dieser Ansicht."));
      /* Archiv: stornierte Positionen, zugeklappt ganz unten */
      var storno = gruppieren(zeilen.filter(function(r){ return r.status === "storniert" && (!q || (r.name + " " + r.artikel + " " + r.mannschaft).toLowerCase().indexOf(q) > -1); }));
      var det = el("details", "karte"); det.open = ARCHIV_OFFEN; det.addEventListener("toggle", function(){ ARCHIV_OFFEN = det.open; });
      add(det, add(el("summary"), el("b", null, "🗄️ Archiv: storniert (" + storno.length + (storno.length === 1 ? " Bestellung" : " Bestellungen") + ")")),
        el("p", "grau", "Stornierte Bestellungen bleiben in der Excel stehen, werden hier aber nicht mehr bei den offenen gezeigt." + (darf ? " „Storno zurücknehmen“ holt sie zurück." : "")));
      storno.forEach(function(g){ add(det, twBestellung(g, true)); });
      if(!storno.length) add(det, el("p", "lead", "Keine stornierten Bestellungen."));
      add(archiv, det); }
    function twBestellung(g, imArchiv){
      var b = el("div", "twbest" + (imArchiv ? " storniert" : "")), r = g.r, h = el("div", "twkopf"), summe = g.pos.reduce(function(m, p){ return m + (p.summe || 0); }, 0);
      var knopf = null;
      if(darf){ knopf = imArchiv ? btn("klein", "Storno zurücknehmen", function(){ bestellungStorno(g, false, knopf); })
        : btn("klein gefahr", "Bestellung stornieren", function(){ bestellungStorno(g, true, knopf); }); }
      add(h, el("b", null, r.name || "–"), el("span", null, r.mannschaft), el("span", null, datumText(r.eingang)), r.nr ? el("span", null, "Nr. " + r.nr) : null, el("span", "sp"), add(el("span"), "Summe ", el("b", null, euro(summe))), knopf);
      add(b, h); g.pos.forEach(function(p){ add(b, imArchiv ? twArchivPosition(p) : twPosition(p)); }); return b; }
    function twArchivPosition(r){ var z = el("div", "twpos"), links = el("div", "twtext");
      add(links, add(el("div", "twart"), el("i", "lampe grau", ""), el("b", null, (r.anzahl > 1 ? r.anzahl + " × " : "") + r.artikel), el("span", null, [r.groesse, r.initialen && r.initialen !== "-" ? "Initialen " + r.initialen : ""].filter(Boolean).join(" · "))),
        add(el("div", "twmeta"), el("span", null, "storniert"), el("span", null, euro(r.summe))));
      if(r.bem) add(links, el("div", "twmeta", r.bem));
      add(z, links); return z; }
    /* Ganze Bestellung stornieren bzw. Storno zurücknehmen: Spalte S (Storniert) jeder Position. Vorher frisch lesen und prüfen. */
    function bestellungStorno(g, an, knopf){
      if(an && !confirm("Bestellung " + (g.r.nr ? "Nr. " + g.r.nr + " " : "") + "von " + (g.r.name || "–") + " mit " + g.pos.length + (g.pos.length === 1 ? " Position" : " Positionen") + " stornieren?\n\nSie bleibt in der Excel stehen und wandert unten ins Archiv.")) return;
      if(!darf){ return; }
      if(!TEST && !SCHREIBEN_FREI){ zeigeToast("Eintragen ist im Portal noch gesperrt."); return; }
      knopf.disabled = true; var sp = spalte(STUFEN[4][1]);
      TWQ().lesen().then(function(w2){ var jetzt = twZeilen(w2);
        g.pos.forEach(function(p){ var x = jetzt.filter(function(y){ return y.zeile === p.zeile; })[0];
          if(!x || x.nr !== p.nr || x.artikel !== p.artikel) throw new Error("Die Bestellung hat sich in der Excel verändert. Bitte die Seite neu laden."); });
        return g.pos.reduce(function(k, p){ return k.then(function(){ return TWQ().schreiben(sp + p.zeile + ":" + sp + p.zeile, [[an ? "Ja" : ""]]); }); }, Promise.resolve());
      }).then(function(){ g.pos.forEach(function(p){ p.stufen[4] = an; p.status = twStatus(p.stufen); });
        zeigeToast(an ? "Bestellung storniert – steht jetzt im Archiv" : "Storno zurückgenommen"); zeigen(); })
        .catch(function(e){ knopf.disabled = false; zeigeToast(e && e.status ? twFehler(e) : (e.message || "Fehler beim Speichern.")); });
    }
    function twPosition(r){
      var z = el("div", "twpos"), links = el("div", "twtext");
      add(links, add(el("div", "twart"), el("i", "lampe " + twLampe(r.status), ""), el("b", null, (r.anzahl > 1 ? r.anzahl + " × " : "") + r.artikel), el("span", null, [r.groesse, r.initialen && r.initialen !== "-" ? "Initialen " + r.initialen : ""].filter(Boolean).join(" · "))),
        add(el("div", "twmeta"), el("span", null, r.status || "–"), el("span", null, euro(r.summe))));
      if(r.pruefen) add(links, el("div", "twwarn", "⚠ Preis prüfen: " + r.pruefen));
      if(r.bem) add(links, el("div", "twmeta", r.bem));
      var st = el("div", "twstufen"), naechste = [0, 2, 3].filter(function(k){ return !r.stufen[k]; })[0];
      TW_SICHTBAR.forEach(function(i){ var s2 = STUFEN[i];
        var an = r.stufen[i], farbe = i === 4 ? (an ? "rot" : "aus") : (an ? "gruen" : (r.stufen[4] ? "aus" : (i === naechste ? "gelb" : "aus")));
        var bt = el("button", "twstufe"); bt.type = "button"; bt.disabled = !darf;
        add(bt, el("i", "lampe gross " + farbe, an ? (i === 4 ? "✕" : "✓") : ""), el("span", null, s2[0]));
        bt.setAttribute("aria-pressed", an ? "true" : "false"); bt.setAttribute("aria-label", s2[0] + (an ? " – erledigt" : " – offen"));
        if(darf) bt.addEventListener("click", function(){ umschalten(r, i, z); });
        add(st, bt); });
      add(z, links, st); return z; }
    function umschalten(r, i, pos){
      if(gesperrt(pos, "teamwear")) return;
      var neu = r.stufen.slice(); neu[i] = !neu[i];
      pos.querySelectorAll(".twstufe").forEach(function(b){ b.disabled = true; });
      TWQ().lesen().then(function(w2){
        var jetzt = twZeilen(w2).filter(function(x){ return x.zeile === r.zeile; })[0];
        if(!jetzt || jetzt.nr !== r.nr || jetzt.artikel !== r.artikel) throw new Error("Die Zeile hat sich in der Excel verändert. Bitte die Seite neu laden.");
        var sp = spalte(STUFEN[i][1]);
        if(i === 2){ neu[1] = neu[2]; return TWQ().schreiben(spalte(STUFEN[1][1]) + r.zeile + ":" + sp + r.zeile, [[neu[2] ? "Ja" : "", neu[2] ? "Ja" : ""]]); }
        return TWQ().schreiben(sp + r.zeile + ":" + sp + r.zeile, [[neu[i] ? "Ja" : ""]]);
      }).then(function(){ r.stufen = neu; r.status = twStatus(neu); zeigeToast(r.artikel + ": " + STUFEN[i][0] + (neu[i] ? " ✓" : " entfernt") + (i === 4 && neu[i] ? " – steht jetzt im Archiv" : "")); if(i === 4) zeigen(); else pos.replaceWith(twPosition(r)); })
        .catch(function(e){ pos.querySelectorAll(".twstufe").forEach(function(b){ b.disabled = false; }); add(pos, el("div", "fehler", e && e.status ? twFehler(e) : (e.message || "Fehler beim Speichern."))); });
    }
    filter.addEventListener("change", zeigen); such.addEventListener("input", zeigen); zeigen();
  }).catch(function(e){ fehlerZeigen(twFehler(e)); });
}
function twFehler(e){
  if(e && (e.status === 403 || e.status === 404)) return "Kein Zugriff auf die Teamwear-Excel. Du musst Mitglied im Team „Vorstandschaft“ sein – bitte bei der Verwaltung melden.";
  return fehlerText(e);
}

/* ---------- Protokolle ----------
 * Nur lesen. Quelle: Team „Sportfreunde Schneeberg 1928 e.V.“, Dokumente › Schriftführer:
 *   05_Protokolle/Ausschusssitzungen/<Jahr>/, 05_Protokolle/Klausurtagungen/ (Jahr im Dateinamen),
 *   04_Generalversammlungen/Generalversammlung <Jahr>/.
 * Gezeigt werden nur Dateien mit „Protokoll“ im Namen (keine Einladungen, Vorlagen, Kopien). PDF vor Word. */
var PR = { site: "/sites/SportfreundeSchneeberg1928e.V", basis: "Schriftführer" };
var PR_MON = ["januar","februar","märz","april","mai","juni","juli","august","september","oktober","november","dezember"];
var prDrive = null;
function prFinden(){ if(prDrive) return Promise.resolve(prDrive);
  return gf("/sites/" + CFG.host + ":" + PR.site + "?$select=id").then(function(s){ return gf("/sites/" + s.id + "/drive?$select=id"); }).then(function(d){ return prDrive = d.id; }); }
function prKinder(pfadTeil){
  return prFinden().then(function(d){ return gf("/drives/" + d + "/root:/" + pfadK(PR.basis + "/" + pfadTeil) + ":/children?$select=name,webUrl,folder,file&$top=999"); })
    .then(function(r){ return (r && r.value) || []; }).catch(function(e){ if(e && e.status === 404) return []; throw e; }); }
var PrLive = {
  jahre: function(){ return Promise.all([prKinder("05_Protokolle/Ausschusssitzungen"), prKinder("04_Generalversammlungen"), prKinder("05_Protokolle/Klausurtagungen")]).then(function(r){
    var j = {}; r[0].forEach(function(x){ if(x.folder && /^\d{4}$/.test(x.name)) j[x.name] = 1; });
    r[1].forEach(function(x){ var m = /(\d{4})/.exec(x.name); if(x.folder && m) j[m[1]] = 1; });
    r[2].forEach(function(x){ var y = prJahrAusName(x.name); if(x.file && y) j[y] = 1; });
    return Object.keys(j).map(Number).sort(function(a, b){ return b - a; }); }); },
  dateien: function(jahr){ return Promise.all([prKinder("05_Protokolle/Ausschusssitzungen/" + jahr), prKinder("04_Generalversammlungen/Generalversammlung " + jahr), prKinder("05_Protokolle/Klausurtagungen")]).then(function(r){
    var out = [];
    r[0].forEach(function(x){ if(x.file) out.push({name: x.name, url: x.webUrl, art: /klausur|tagung/i.test(x.name) ? "Klausurtagung" : "Ausschusssitzung"}); });
    r[1].forEach(function(x){ if(x.file) out.push({name: x.name, url: x.webUrl, art: "Generalversammlung"}); });
    r[2].forEach(function(x){ if(x.file && prJahrAusName(x.name) === jahr) out.push({name: x.name, url: x.webUrl, art: "Klausurtagung"}); });
    return out; }); }
};
var PrTest = {
  jahre: function(){ return Promise.resolve([2026, 2025, 2017]); },
  dateien: function(j){ var t = j === 2017 ? ["Protokolle 2017.pdf"] : ["01_Ausschusssitzung_Protokoll_Januar_" + j + ".pdf", "01_Ausschusssitzung_Protokoll_Januar_" + j + ".docx", "02_Ausschusssitzung_Protokoll_März_" + j + ".docx",
      "05_Einladung_Juli Sitzung.docx", "Vorlage Einladung.docx", "03_Protokoll_April – Kopie.docx", "06_Ausschusssitzung_Protokoll_September_" + j + ".pdf", "Klausurtagung_Protokoll_November_" + j + ".pdf"];
    var out = t.map(function(n){ return {name: n, url: "#", art: /klausur/i.test(n) ? "Klausurtagung" : "Ausschusssitzung"}; });
    if(j === 2026){ out.push({name: "Protokoll_Generalversammlung_09.10.2026.pdf", url: "#", art: "Generalversammlung"}); out.push({name: j + "_Klausurtagung_Protokoll_November.pdf", url: "#", art: "Klausurtagung"}); }
    return Promise.resolve(out); }
};
function PRQ(){ return TEST ? PrTest : PrLive; }
function prJahrAusName(n){ var m = /^(\d{4})[_ -]/.exec(n) || /\d{1,2}[._]\d{1,2}[._](\d{4})/.exec(n) || /_(\d{4})\.\w+$/.exec(n); return m ? +m[1] : null; }
function prMonat(n){ var l = n.toLowerCase().replace("july", "juli");
  for(var i = 0; i < 12; i++) if(l.indexOf(PR_MON[i]) > -1) return i;
  var m = /(\d{1,2})[._](\d{1,2})[._]\d{4}/.exec(l); if(m && +m[2] >= 1 && +m[2] <= 12) return +m[2] - 1;
  return null; }
/* Dateien eines Jahres zu Einträgen bündeln: je Monat und Art ein Eintrag, PDF bevorzugt. */
function prEintraege(dateien){
  var g = {};
  dateien.forEach(function(x){
    if(!/protokoll/i.test(x.name) || /kopie|vorlage|einladung/i.test(x.name)) return;
    if(!/\.(pdf|docx?)$/i.test(x.name)) return;
    var mon = prMonat(x.name), key = (mon == null ? "J" : mon) + "|" + x.art;
    var e = g[key] = g[key] || {monat: mon, art: x.art, pdf: null, word: null};
    if(/\.pdf$/i.test(x.name)){ if(!e.pdf) e.pdf = x; } else if(!e.word) e.word = x; });
  return Object.keys(g).map(function(k){ return g[k]; });
}
var PR_JAHR = null;
function seiteProtokolle(s, jahrTeil){
  pfad([["Start", "#/"], ["Protokolle"]]);
  ladeHinweis(s);
  PRQ().jahre().then(function(jahre){
    s.textContent = "";
    if(!jahre.length){ add(s, el("h1", null, "Protokolle"), el("p", "lead", "Keine Protokolle gefunden.")); return; }
    var jahr = +jahrTeil || PR_JAHR || jahre[0]; if(jahre.indexOf(jahr) < 0) jahr = jahre[0]; PR_JAHR = jahr;
    pfad([["Start", "#/"], ["Protokolle", "#/protokolle"], [String(jahr)]]);
    add(s, el("h1", null, "Protokolle " + jahr), el("p", "lead", "Ausschusssitzungen, Klausurtagungen und Generalversammlung aus dem Ordner des Schriftführers. Ein Klick öffnet das Protokoll."));
    var nav = el("div", "jahre"); jahre.forEach(function(j){ var a = el("a", "chip" + (j === jahr ? " ok" : ""), String(j)); a.href = "#/protokolle/" + j; if(j === jahr) a.setAttribute("aria-current", "page"); add(nav, a); }); add(s, nav);
    var box = el("div"); add(s, box); add(box, el("div", "laden", "Protokolle werden geladen …"));
    return PRQ().dateien(jahr).then(function(dateien){
      box.textContent = "";
      var ein = prEintraege(dateien), rang = {Ausschusssitzung: 0, Klausurtagung: 1, Generalversammlung: 2};
      ein.sort(function(a, b){ return rang[a.art] - rang[b.art]; });
      var t = el("table", "tabelle"), tb = el("tbody");
      add(t, add(el("thead"), add(el("tr"), el("th", null, "Monat"), el("th", null, "Protokoll"))), tb);
      ein.filter(function(e){ return e.monat == null; }).forEach(function(e){ add(tb, prZeile("Ganzes Jahr", [e])); });
      for(var m = 0; m < 12; m++){ var dieser = ein.filter(function(e){ return e.monat === m; }); add(tb, prZeile(MONATE[m], dieser)); }
      add(box, add(el("div", "scroll"), t));
    });
  }).catch(function(e){ fehlerZeigen(prFehler(e)); });
}
function prZeile(titel, eintraege){
  var tr = el("tr" + ""), td = el("td");
  if(!eintraege.length){ tr.className = "leer"; add(td, el("span", "grau", "keine Sitzung")); }
  eintraege.forEach(function(e){ var d = e.pdf || e.word, a = el("a", "doc", e.monat == null ? "Alle Protokolle des Jahres" : e.art); a.href = d.url; a.target = "_blank"; a.rel = "noopener";
    add(td, add(el("div", "docreihe"), a, el("span", "chip" + (e.pdf ? "" : " grau"), e.pdf ? "PDF" : "Word"), e.art !== "Ausschusssitzung" ? el("span", "chip offen", "Sondertermin") : null)); });
  add(tr, el("td", null, titel), td); return tr;
}
function prFehler(e){
  if(e && (e.status === 403 || e.status === 404)) return "Kein Zugriff auf die Protokolle. Du musst Mitglied im Team „Sportfreunde Schneeberg 1928 e.V.“ sein – bitte bei der Verwaltung melden.";
  return fehlerText(e);
}

/* ---------- Vereinskalender ----------
 * Outlook-Kalender „Vereinskalender_Spfr1928“. Er ist in Outlook für die Ausschussmitglieder freigegeben;
 * das Portal findet ihn unter den eigenen Kalendern des angemeldeten Kontos (nach Annahme der Freigabe).
 * Anlegen, Ändern und Löschen nur, wenn Outlook dem Konto „Kann bearbeiten“ gibt (canEdit). Keine Einladungen. */
var KAL = { name: /vereinskalender/i, zone: "W. Europe Standard Time" };
var kalInfo = null, KAL_MONAT = null;
var KalLive = {
  finden: function(){ if(kalInfo) return Promise.resolve(kalInfo);
    return gf("/me/calendars?$select=id,name,canEdit&$top=200").then(function(r){ var c = ((r && r.value) || []).filter(function(x){ return KAL.name.test(x.name); })[0];
      if(!c){ var e = new Error("Kalender"); e.kal = true; throw e; } return kalInfo = {id: c.id, name: c.name, darf: !!c.canEdit}; }); },
  termine: function(von, bis){ return KalLive.finden().then(function(k){
    return gf("/me/calendars/" + encodeURIComponent(k.id) + "/calendarView?startDateTime=" + von + "&endDateTime=" + bis + "&$select=id,subject,start,end,isAllDay,location,bodyPreview&$orderby=start/dateTime&$top=500", {headers: {Prefer: 'outlook.timezone="' + KAL.zone + '"'}}); })
    .then(function(r){ return (r && r.value) || []; }); },
  speichern: function(id, daten){ return KalLive.finden().then(function(k){ var p = "/me/calendars/" + encodeURIComponent(k.id) + "/events";
    return id ? gf(p + "/" + encodeURIComponent(id), {method: "PATCH", body: daten}) : gf(p, {method: "POST", body: daten}); }); },
  loeschen: function(id){ return KalLive.finden().then(function(k){ return gf("/me/calendars/" + encodeURIComponent(k.id) + "/events/" + encodeURIComponent(id), {method: "DELETE"}); }); }
};
var KalTest = {
  liste: null,
  finden: function(){ return Promise.resolve({id: "test", name: "Vereinskalender_Spfr1928", darf: (/[?&]rolle=([a-z]+)/.exec(location.search) || [])[1] !== "gast"}); },
  termine: function(von, bis){ if(!KalTest.liste){ var j = new Date().getFullYear(), m = pad(new Date().getMonth() + 1);
      KalTest.liste = [{id: "1", subject: "Ausschusssitzung", start: {dateTime: j + "-" + m + "-03T19:30:00"}, end: {dateTime: j + "-" + m + "-03T21:30:00"}, isAllDay: false, location: {displayName: "Sportheim"}, bodyPreview: "Beispieldaten"},
        {id: "2", subject: "Arbeitseinsatz Sportplatz", start: {dateTime: j + "-" + m + "-14T00:00:00"}, end: {dateTime: j + "-" + m + "-15T00:00:00"}, isAllDay: true, location: {displayName: ""}, bodyPreview: ""}]; }
    window.__kal = window.__kal || [];
    return Promise.resolve(KalTest.liste.filter(function(t){ return t.start.dateTime.slice(0, 10) >= von.slice(0, 10) && t.start.dateTime.slice(0, 10) < bis.slice(0, 10); }).sort(function(a, b){ return a.start.dateTime < b.start.dateTime ? -1 : 1; })); },
  speichern: function(id, d){ window.__kal = window.__kal || []; window.__kal.push({aktion: id ? "ändern" : "anlegen", id: id, daten: d});
    var t = {id: id || String(Date.now()), subject: d.subject, start: d.start, end: d.end, isAllDay: d.isAllDay, location: d.location, bodyPreview: d.body ? d.body.content : ""};
    KalTest.liste = KalTest.liste.filter(function(x){ return x.id !== id; }).concat([t]); return Promise.resolve(t); },
  loeschen: function(id){ window.__kal.push({aktion: "löschen", id: id}); KalTest.liste = KalTest.liste.filter(function(x){ return x.id !== id; }); return Promise.resolve(); }
};
function KALQ(){ return TEST ? KalTest : KalLive; }
var WT = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
function kalDatum(dt){ var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(dt || ""); return m ? {j: +m[1], mo: +m[2], t: +m[3], h: m[4], mi: m[5], iso: m[1] + "-" + m[2] + "-" + m[3], zeit: m[4] + ":" + m[5]} : null; }
function isoPlus(iso, tage){ var d = isoZuDatum(iso); d.setDate(d.getDate() + tage); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
function seiteKalender(s, monatTeil){
  pfad([["Start", "#/"], ["Vereinskalender"]]);
  ladeHinweis(s);
  var heute = new Date(), mm = /^(\d{4})-(\d{2})$/.exec(monatTeil || "") || /^(\d{4})-(\d{2})$/.exec(KAL_MONAT || "");
  var j = mm ? +mm[1] : heute.getFullYear(), mo = mm ? +mm[2] : heute.getMonth() + 1; KAL_MONAT = j + "-" + pad(mo);
  var von = j + "-" + pad(mo) + "-01T00:00:00", nj = mo === 12 ? j + 1 : j, nm = mo === 12 ? 1 : mo + 1, bis = nj + "-" + pad(nm) + "-01T00:00:00";
  var vj = mo === 1 ? j - 1 : j, vm = mo === 1 ? 12 : mo - 1;
  KALQ().finden().then(function(k){ return KALQ().termine(von, bis).then(function(liste){
    s.textContent = "";
    add(s, el("h1", null, "Vereinskalender"), el("p", "lead", k.darf ? "Termine des Vereins, z. B. Ausschusssitzungen. Du kannst Termine anlegen, ändern und löschen – das landet direkt im Outlook-Kalender." : "Termine des Vereins, z. B. Ausschusssitzungen. Du kannst hier nur lesen."));
    var kopf = el("div", "reihe monat"), zur = el("a", "btn klein", "‹ " + MONATE[vm - 1]), vor = el("a", "btn klein", MONATE[nm - 1] + " ›");
    zur.href = "#/kalender/" + vj + "-" + pad(vm); vor.href = "#/kalender/" + nj + "-" + pad(nm);
    var heuteL = el("a", "btn klein", "Heute"); heuteL.href = "#/kalender/" + heute.getFullYear() + "-" + pad(heute.getMonth() + 1);
    add(kopf, zur, el("h2", null, MONATE[mo - 1] + " " + j), vor, heuteL, el("span", "sp"), k.darf ? btn("primary", "+ Neuer Termin", function(){ kalBearbeiten(null, j + "-" + pad(mo) + "-" + pad(Math.min(28, heute.getDate()))); }) : null);
    add(s, kopf);
    if(!liste.length){ add(s, el("p", "lead", "Keine Termine in diesem Monat.")); return; }
    var t = el("table", "tabelle kal"), tb = el("tbody");
    add(t, add(el("thead"), add(el("tr"), el("th", null, "Datum"), el("th", null, "Uhrzeit"), el("th", null, "Termin"), el("th", "ort", "Ort"), k.darf ? el("th", null, "") : null)), tb);
    liste.forEach(function(ev){
      var a = kalDatum(ev.start.dateTime), e = kalDatum(ev.end.dateTime); if(!a) return;
      var d = isoZuDatum(a.iso), zeit = ev.isAllDay ? "ganztägig" : a.zeit + (e ? "–" + e.zeit : "");
      var tag = WT[d.getDay()] + ", " + pad(a.t) + "." + pad(a.mo) + ".";
      if(ev.isAllDay && e && isoPlus(a.iso, 1) !== e.iso){ var ende = isoZuDatum(isoPlus(e.iso, -1)); tag += " – " + pad(ende.getDate()) + "." + pad(ende.getMonth() + 1) + "."; }
      var tit = el("td"); add(tit, el("b", null, ev.subject || "(ohne Titel)")); var fe = festVonTitel(ev.subject); if(fe){ var fl = el("a", "chip festlink", fe.sym + " " + fe.name + " " + a.j + " ›"); fl.href = vaLink(fe, a.j); add(tit, " ", fl); } var ortT = (ev.location && ev.location.displayName) || ""; if(ortT) add(tit, el("div", "notiz ortm", "📍 " + ortT)); if(ev.bodyPreview) add(tit, el("div", "notiz", ev.bodyPreview));
      var tr = el("tr"); add(tr, el("td", null, tag), el("td", null, zeit), tit, el("td", "ort", ortT),
        k.darf ? add(el("td", "z"), btn("klein", "Ändern", function(){ kalBearbeiten(ev); })) : null);
      add(tb, tr); });
    add(s, add(el("div", "scroll"), t));
  }); }).catch(function(e){ fehlerZeigen(kalFehler(e)); });
}
function kalBearbeiten(ev, isoVorschlag){
  var d = el("dialog"), f = el("form", "karte"); f.noValidate = true;
  var a = ev && kalDatum(ev.start.dateTime), e = ev && kalDatum(ev.end.dateTime);
  function feld(t, typ, v, cls){ var i = el("input"); i.type = typ; i.value = v || ""; var l = el("label", "fld" + (cls ? " " + cls : ""), t); add(l, i); l._i = i; return l; }
  var fT = feld("Titel", "text", ev ? ev.subject : "", "w2"), fD = feld("Datum", "date", a ? a.iso : isoVorschlag),
    fGanz = el("input"); fGanz.type = "checkbox"; fGanz.checked = !!(ev && ev.isAllDay);
  var endeIso = ev && ev.isAllDay && e ? isoPlus(e.iso, -1) : (a ? a.iso : "");
  var fBisD = feld("Bis Datum (mehrtägig)", "date", ev && ev.isAllDay && endeIso !== a.iso ? endeIso : ""),
    fVon = feld("Beginn", "time", a && !ev.isAllDay ? a.zeit : "19:30"), fBis = feld("Ende", "time", e && !ev.isAllDay ? e.zeit : "21:00"),
    fO = feld("Ort", "text", ev && ev.location ? ev.location.displayName : "", "w2"),
    fN = el("textarea"); fN.rows = 3; fN.value = ev && ev.bodyPreview ? ev.bodyPreview : "";
  var lGanz = add(el("label", "fld haken"), fGanz, " Ganztägig"), lN = add(el("label", "fld w2", "Notiz"), fN);
  function sicht(){ fVon.hidden = fBis.hidden = fGanz.checked; fBisD.hidden = !fGanz.checked; }
  fGanz.addEventListener("change", sicht); sicht();
  var ok = el("button", "btn primary", "Speichern"); ok.type = "submit";
  var loe = ev ? btn("gefahr", "Löschen", function(){
    if(!confirm("Termin „" + (ev.subject || "") + "“ wirklich löschen? Er verschwindet für alle aus dem Vereinskalender.")) return;
    loe.disabled = true; KALQ().loeschen(ev.id).then(function(){ d.close(); d.remove(); zeigeToast("Termin gelöscht"); route(); })
      .catch(function(x){ loe.disabled = false; add(f, el("div", "fehler", kalFehler(x))); }); }) : null;
  add(f, el("h2", null, ev ? "Termin ändern" : "Neuer Termin"), add(el("div", "formular"), fT, fD, lGanz, fVon, fBis, fBisD, fO, lN),
    add(el("div", "reihe"), ok, btn("", "Abbrechen", function(){ d.close(); d.remove(); }), el("span", "sp"), loe));
  add(d, f); document.body.appendChild(d); d.showModal(); fT._i.focus();
  f.addEventListener("submit", function(x){ x.preventDefault();
    var titel = fT._i.value.trim(), tag = fD._i.value; if(!titel || !tag){ zeigeToast("Bitte Titel und Datum eintragen."); return; }
    var daten = {subject: titel, isAllDay: fGanz.checked, location: {displayName: fO._i.value.trim()}, body: {contentType: "text", content: fN.value.trim()}};
    if(fGanz.checked){ var bisTag = fBisD._i.value && fBisD._i.value >= tag ? fBisD._i.value : tag;
      daten.start = {dateTime: tag + "T00:00:00", timeZone: KAL.zone}; daten.end = {dateTime: isoPlus(bisTag, 1) + "T00:00:00", timeZone: KAL.zone}; }
    else { var v = fVon._i.value || "19:30", b = fBis._i.value || v; if(b <= v){ zeigeToast("Das Ende muss nach dem Beginn liegen."); return; }
      daten.start = {dateTime: tag + "T" + v + ":00", timeZone: KAL.zone}; daten.end = {dateTime: tag + "T" + b + ":00", timeZone: KAL.zone}; }
    if(ev && fN.value.trim() === (ev.bodyPreview || "").trim()) delete daten.body; /* Vorschau ist gekürzt: unverändert nie zurückschreiben */
    ok.disabled = true;
    KALQ().speichern(ev ? ev.id : null, daten).then(function(){ d.close(); d.remove(); zeigeToast(titel + " gespeichert ✓"); KAL_MONAT = tag.slice(0, 7); var ziel = "#/kalender/" + KAL_MONAT; if(location.hash !== ziel) location.hash = ziel; else route(); })
      .catch(function(y){ ok.disabled = false; add(f, el("div", "fehler", kalFehler(y))); });
  });
}
function kalFehler(e){
  if(e && e.kal) return "Der Vereinskalender ist für dein Konto noch nicht sichtbar. Er muss dir in Outlook freigegeben sein und du musst die Freigabe-Einladung in Outlook angenommen haben – bitte bei der Verwaltung melden.";
  if(e && (e.status === 403 || e.status === 404)) return "Kein Zugriff auf den Vereinskalender (oder keine Bearbeitungsrechte) – bitte bei der Verwaltung melden.";
  return fehlerText(e);
}

/* ---------- Aufgaben Liegenschaften (Planner, vorerst nur anzeigen) ----------
 * Plan im Team „TEAM Liegenschaften“ (Gruppe mailNickname TEAMLiegenschaften). Bevorzugt der Plan mit „Aufgaben“ im Titel,
 * sonst der erste Plan. Gepflegt wird in Planner (Teams); das Portal zeigt den Stand für die monatliche Ausschusssitzung. */
var AU = { gruppe: "TEAMLiegenschaften", titel: /aufgaben/i };
var AU_FILTER = "offen";
function alle(pfadA){ var out = []; function weiter(u){ return gf(u).then(function(r){ out = out.concat((r && r.value) || []); return r && r["@odata.nextLink"] ? weiter(r["@odata.nextLink"]) : out; }); } return weiter(pfadA); }
var AuLive = {
  laden: function(){
    return gf("/groups?$filter=" + encodeURIComponent("mailNickname eq '" + AU.gruppe + "'") + "&$select=id").then(function(r){
      var g = r && r.value && r.value[0]; if(!g){ var e = new Error("Gruppe"); e.status = 404; throw e; }
      return gf("/groups/" + g.id + "/planner/plans?$select=id,title").then(function(p){
        var plaene = (p && p.value) || [], plan = plaene.filter(function(x){ return AU.titel.test(x.title); })[0] || plaene[0];
        if(!plan) return {plan: null};
        return Promise.all([alle("/planner/plans/" + plan.id + "/tasks"), alle("/planner/plans/" + plan.id + "/buckets"), alle("/groups/" + g.id + "/members?$select=id,displayName")])
          .then(function(x){ var leute = {}; x[2].forEach(function(m){ leute[m.id] = m.displayName; });
            return {plan: plan, aufgaben: x[0], buckets: x[1], leute: leute}; }); }); }); }
};
var AuTest = {
  laden: function(){ var j = new Date().getFullYear();
    return Promise.resolve({plan: {id: "p", title: "Liegenschaften – Aufgaben"}, leute: {u1: "Max Muster", u2: "Erika Beispiel"},
      buckets: [{id: "b1", name: "Sportheim", orderHint: "1"}, {id: "b2", name: "Sportplatz", orderHint: "2"}, {id: "b3", name: "Sonstiges", orderHint: "3"}],
      aufgaben: [
        {id: "t1", bucketId: "b1", title: "Beispiel: Feuchtigkeit Treppenhaus prüfen", percentComplete: 50, priority: 1, dueDateTime: (j - 1) + "-12-01T00:00:00Z", assignments: {u1: {}}},
        {id: "t2", bucketId: "b2", title: "Beispiel: Torbefestigungen erneuern", percentComplete: 0, priority: 5, dueDateTime: null, assignments: {u2: {}}},
        {id: "t3", bucketId: "b2", title: "Beispiel: Motorsensen warten", percentComplete: 0, priority: 5, dueDateTime: j + 1 + "-03-15T00:00:00Z", assignments: {}},
        {id: "t4", bucketId: "b3", title: "Beispiel: Bandenwerbung abrechnen", percentComplete: 100, priority: 5, dueDateTime: null, completedDateTime: j + "-10-02T10:00:00Z", assignments: {u1: {}}}]}); }
};
function AUQ(){ return TEST ? AuTest : AuLive; }
function auLink(t){ return "https://tasks.office.com/" + CFG.tenantId + "/Home/Task/" + t.id + "?Type=TaskLink"; }
function auPlanLink(p){ return "https://tasks.office.com/" + CFG.tenantId + "/Home/PlanViews/" + p.id + "?Type=PlanLink"; }
function auStatus(t){ return t.percentComplete >= 100 ? "erledigt" : (t.percentComplete > 0 ? "in Arbeit" : "offen"); }
function auDatum(dt){ var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(dt || ""); return m ? m[3] + "." + m[2] + "." + m[1] : ""; }
function seiteAufgaben(s){
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Aufgaben"]]);
  ladeHinweis(s);
  AUQ().laden().then(function(D){
    s.textContent = "";
    add(s, el("h1", null, "Aufgaben Liegenschaften"));
    if(!D.plan){ add(s, el("p", "lead", "Es gibt noch keinen Planner-Plan im Team „TEAM Liegenschaften“."),
      el("div", "hinweis", "Anlegen in Teams: Team „TEAM Liegenschaften“ › Kanal „General“ › „+“ › Planner › neuer Plan „Liegenschaften – Aufgaben“. Danach erscheinen die Aufgaben hier automatisch.")); erinnerungenKarte(s); return; }
    var heute = new Date().toISOString().slice(0, 10);
    var lead = el("p", "lead", "Stand aus Planner („" + D.plan.title + "“) für die monatliche Ausschusssitzung. Bearbeitet wird in Planner. ");
    var pl = el("a", null, "Plan in Planner öffnen"); pl.href = auPlanLink(D.plan); pl.target = "_blank"; pl.rel = "noopener"; add(lead, pl); add(s, lead);
    var zaehl = {offen: 0, "in Arbeit": 0, erledigt: 0, ueber: 0};
    D.aufgaben.forEach(function(t){ var st = auStatus(t); zaehl[st]++; if(st !== "erledigt" && t.dueDateTime && t.dueDateTime.slice(0, 10) < heute) zaehl.ueber++; });
    var c = el("div", "summe"); [["offen", zaehl.offen], ["in Arbeit", zaehl["in Arbeit"]], ["überfällig", zaehl.ueber], ["erledigt", zaehl.erledigt]].forEach(function(k){ add(c, add(el("span"), k[0] + " ", el("b", null, String(k[1])))); }); add(s, c);
    var filter = el("select"); filter.style.maxWidth = "220px";
    [["offen", "Alle offenen"], ["ueber", "Überfällig"], ["alle", "Alle"], ["erledigt", "Erledigt"]].forEach(function(o){ var op = el("option", null, o[1]); op.value = o[0]; add(filter, op); });
    filter.value = AU_FILTER;
    add(s, add(el("div", "reihe"), add(el("label", "fld", "Anzeigen"), filter), btn("klein", "Drucken für die Sitzung", function(){ window.print(); })));
    var box = el("div"); add(s, box);
    var buckets = D.buckets.slice().sort(function(a, b){ return (a.orderHint || "") < (b.orderHint || "") ? -1 : 1; }), namen = {};
    buckets.forEach(function(b){ namen[b.id] = b.name; });
    function zeigen(){ box.textContent = ""; AU_FILTER = filter.value; var n = 0;
      var gruppen = buckets.map(function(b){ return b.id; }).concat(["?"]);
      gruppen.forEach(function(bid){
        var liste = D.aufgaben.filter(function(t){ var b = namen[t.bucketId] ? t.bucketId : "?"; if(b !== bid) return false; var st = auStatus(t);
          if(AU_FILTER === "offen") return st !== "erledigt"; if(AU_FILTER === "erledigt") return st === "erledigt";
          if(AU_FILTER === "ueber") return st !== "erledigt" && t.dueDateTime && t.dueDateTime.slice(0, 10) < heute; return true; });
        if(!liste.length) return;
        liste.sort(function(a, b){ return (a.dueDateTime || "9999") < (b.dueDateTime || "9999") ? -1 : 1; });
        var k = el("div", "karte"); add(k, el("h2", null, namen[bid] || "Ohne Bucket"));
        var t = el("table", "tabelle aufg"), tb = el("tbody");
        add(t, add(el("thead"), add(el("tr"), el("th", null, "Aufgabe"), el("th", null, "Zuständig"), el("th", null, "Fällig"), el("th", null, "Stand"))), tb);
        liste.forEach(function(a){ n++;
          var st = auStatus(a), ueber = st !== "erledigt" && a.dueDateTime && a.dueDateTime.slice(0, 10) < heute;
          var wer = Object.keys(a.assignments || {}).map(function(id){ return D.leute[id] || ""; }).filter(Boolean).join(", ");
          var titel = el("a", null, a.title); titel.href = auLink(a); titel.target = "_blank"; titel.rel = "noopener";
          var tr = el("tr"); add(tr, add(el("td"), a.priority <= 1 ? el("span", "chip offen", "dringend") : null, a.priority <= 1 ? " " : null, titel), el("td", null, wer || "–"),
            el("td", ueber ? "faellig" : null, auDatum(a.dueDateTime) || "–"), add(el("td"), el("span", "chip " + (st === "erledigt" ? "ok" : (st === "in Arbeit" ? "" : "offen")), st)));
          add(tb, tr); });
        add(k, add(el("div", "scroll"), t)); add(box, k); });
      if(!n) add(box, el("p", "lead", "Keine Aufgaben in dieser Ansicht.")); }
    filter.addEventListener("change", zeigen); zeigen();
    erinnerungenKarte(s);
  }).catch(function(e){ fehlerZeigen(e && (e.status === 403 || e.status === 404) ? "Kein Zugriff auf die Aufgaben. Du musst Mitglied im Team „TEAM Liegenschaften“ sein – bitte bei der Verwaltung melden." : fehlerText(e)); });
}

/* ---------- Erinnerungen (iCloud-Liste, vom Mac abgeglichen) ----------
 * Der Mac schreibt alle 30 Minuten die offenen Erinnerungen als JSON in den Teams-Ordner (Tools/erinnerungen_sync.sh).
 * Das Portal liest die Datei nur. Bearbeitet wird in der Erinnerungen-App. */
var ER_PFAD = "General/Aufgabenliste/Erinnerungen_offen.json";
function erinnerungenLaden(){
  if(TEST){ var j = new Date().getFullYear(); return Promise.resolve({liste: "Beispielliste", stand: j + "-11-24T09:30+01:00", aufgaben: [
    {titel: "Beispiel: Mülltonne raus", faellig: j + "-12-14", prioritaet: "hoch", notiz: ""}, {titel: "Beispiel: Torbefestigungen", faellig: null, prioritaet: "", notiz: "Netzhaken fehlen"}]}); }
  return finden().then(function(i){ return gf("/drives/" + i.drive + "/root:/" + pfadK(ER_PFAD) + "?$select=id,@microsoft.graph.downloadUrl"); })
    .then(function(it){ return fetch(it["@microsoft.graph.downloadUrl"]); })
    .then(function(r){ if(!r.ok){ var e = new Error("Datei " + r.status); e.status = r.status; throw e; } return r.json(); });
}
function erinnerungenKarte(s){
  var k = el("div", "karte"); add(k, el("h2", null, "Aus der iCloud-Erinnerungsliste"), el("p", "lead", "Wird geladen …")); add(s, k);
  erinnerungenLaden().then(function(D){
    k.textContent = ""; add(k, el("h2", null, "Aus der iCloud-Erinnerungsliste"));
    var st = D.stand ? new Date(D.stand) : null;
    add(k, el("p", "lead", (D.aufgaben || []).length + " offen · zuletzt geändert " + (st && !isNaN(st) ? st.toLocaleString("de-DE", {dateStyle: "medium", timeStyle: "short"}) : "–") + ". Bearbeitet wird in der Erinnerungen-App; der Abgleich läuft alle 30 Minuten, solange der Mac der Verwaltung an ist."));
    var heute = heuteISO(), t = el("table", "tabelle aufg"), tb = el("tbody");
    add(t, add(el("thead"), add(el("tr"), el("th", null, "Aufgabe"), el("th", null, "Notiz"), el("th", null, "Fällig"))), tb);
    (D.aufgaben || []).forEach(function(a){ var ueber = a.faellig && a.faellig < heute;
      add(tb, add(el("tr"), add(el("td"), a.prioritaet === "hoch" ? el("span", "chip offen", "dringend") : null, a.prioritaet === "hoch" ? " " : null, a.titel), el("td", null, a.notiz || "–"), el("td", ueber ? "faellig" : null, a.faellig ? isoZuText(a.faellig) : "–"))); });
    add(k, add(el("div", "scroll"), t));
  }).catch(function(e){ k.textContent = ""; add(k, el("h2", null, "Aus der iCloud-Erinnerungsliste"),
    el("p", "lead", e && e.status === 404 ? "Noch kein Abgleich vorhanden (Datei „Aufgabenliste/Erinnerungen_offen.json“ fehlt)." : "Die Erinnerungen konnten nicht geladen werden: " + fehlerText(e))); });
}

/* ---------- Veranstaltungen (Team „Veranstaltungen“, nur lesen) ----------
 * Je Fest ein Kanal-Ordner im Team, darin Ordner je Jahr (z. B. „Weihnachtsfeier 2025“ oder „2025“).
 * Das Portal zeigt je Fest die Jahre, je Jahr runde Knöpfe für die Bereiche. Die Bereiche sind die Unterordner
 * des Jahresordners und die Dateien, nach Namen sortiert (Dienstplan, Getränke, Essen, Tombola …).
 * Kalendertermine mit dem Festnamen im Titel verlinken auf das Fest im passenden Jahr.
 * Termine, deren Tag noch nicht feststeht, stehen nur hier als „Vorschlag“ (nicht im Kalender). */
var VA = { gruppe: "Veranstaltungen" };
var FESTE = [
  {key: "schlachtfest", name: "Schlachtfest", sym: "🐖", ordner: /schlachtfest/i, titel: /schlachtfest/i, monat: "Februar oder März", mon: 3, alle: 1},
  {key: "ortsmeisterschaft", name: "Ortsmeisterschaft", sym: "⚽", ordner: /ortsmeister/i, titel: /ortsmeister|(^|\W)OM(\W|$)/i, monat: "Juni", mon: 6, alle: 1},
  {key: "biergartenfest", name: "Biergartenfest", sym: "🍻", ordner: /biergarten/i, titel: /biergarten/i, monat: "", alle: 2, ab: 2029},
  {key: "kerb", name: "Kerb", sym: "🎡", ordner: /^kerb/i, titel: /(^|\W)kerb/i, monat: "Oktober", mon: 10, alle: 1},
  {key: "weihnachtsfeier", name: "Weihnachtsfeier", sym: "🎄", ordner: /weihnacht/i, titel: /weihnacht/i, monat: "Dezember", mon: 12, alle: 1},
  {key: "jubilaeum", name: "100 Jahre Sportfreunde", sym: "🎉", ordner: /jubil|julil|100\s*jahre/i, titel: /jubil|100 jahre/i, monat: "2.–5. Juni 2028", festplaner: true, ohneJahre: true}
];
/* Reihenfolge zählt: der erste passende Bereich nimmt die Datei. */
var BEREICHE = [
  {key: "dienstplan", name: "Dienstplan", sym: "🗓️", muster: /dienstplan|^dienst/i},
  {key: "tombola", name: "Tombola", sym: "🎁", muster: /tombola|tompola/i},
  {key: "anmeldungen", name: "Anmeldungen", sym: "📝", muster: /anmeld|mitglieder f/i},
  {key: "genehmigungen", name: "Genehmigungen", sym: "📋", muster: /antrag|bescheid|gestattung|verkehrsanordnung|sicherheitskonzept|nutzungsmeld|jugendschutz/i},
  {key: "rechnungen", name: "Rechnungen", sym: "🧾", muster: /rechnung|gutschrift|^sr\d|^ar\s?\d|^re-\d/i},
  {key: "getraenke", name: "Getränke", sym: "🍺", muster: /getränk|getraenk|bierdeckel|moschd|\bbar\b/i},
  {key: "essen", name: "Essen", sym: "🍽️", muster: /speise|essen|hähnchen|haehnchen|schlachtplatte|kuchen|kaffee|flammkuchen|kerbbacken|bedienung/i},
  {key: "aufbau", name: "Aufbau & Zelt", sym: "⛺", muster: /zelt|garnitur|pavil|toilette|elektro|wasser|auf-?abbau/i},
  {key: "werbung", name: "Werbung", sym: "📣", muster: /plakat|flyer|einladung|werbung/i},
  {key: "organisation", name: "Organisation", sym: "🧭", muster: /orga|vorbereitung|ansprechpartner|turnier|bestellung|angebot|reservier|public viewing|preisliste/i},
  {key: "rueckblick", name: "Rückblick", sym: "🔁", muster: /verbesser|optimier|feedback|verbrauch|vergleich/i},
  {key: "bilder", name: "Bilder", sym: "📷", muster: /bild|foto|\.(jpe?g|png|heic|gif)$|^img|^whatsapp|^\d{8}_\d+/i},
  {key: "sonstiges", name: "Sonstiges", sym: "📁", muster: /./}
];
function festVonTitel(t){ t = t || ""; for(var i = 0; i < FESTE.length; i++) if(FESTE[i].titel.test(t)) return FESTE[i]; return null; }
function festVonKey(k){ for(var i = 0; i < FESTE.length; i++) if(FESTE[i].key === k) return FESTE[i]; return null; }
function jahrAus(n){ var m = /(19|20)\d\d/.exec(n || ""); return m ? +m[0] : null; }
function bereichVon(name){ for(var i = 0; i < BEREICHE.length; i++) if(BEREICHE[i].muster.test(name)) return BEREICHE[i]; return BEREICHE[BEREICHE.length - 1]; }
/* „Dezember 2027 – Tag offen (Vorschlag)“, ohne bekannten Monat „2029 – Monat und Tag offen (Vorschlag)“ */
function offenText(f, j){ return f.monat ? f.monat + " " + j + " – Tag offen (Vorschlag)" : j + " – Monat und Tag offen (Vorschlag)"; }
function vaLink(f, jahr, ber){ return "#/veranstaltungen/" + f.key + (jahr ? "/" + jahr : "") + (ber ? "/" + encodeURIComponent(ber) : ""); }
/* Nächste geplante Jahre ohne festen Termin: jährlich ab nächstem Jahr, sonst ab „ab“ im Abstand „alle“. */
function geplanteJahre(f, bis){ var out = [], j = new Date().getFullYear(); if(!f.alle) return out;
  var start = f.ab || (f.mon && new Date().getMonth() + 1 > f.mon ? j + 1 : j); for(var y = start; y <= bis; y += f.alle) if(y >= j) out.push(y); return out; }

var vaCache = null;
var VaLive = {
  drive: function(){ if(vaCache) return Promise.resolve(vaCache);
    return gf("/groups?$filter=" + encodeURIComponent("mailNickname eq '" + VA.gruppe + "'") + "&$select=id").then(function(r){
      var g = r && r.value && r.value[0]; if(!g){ var e = new Error("Gruppe"); e.status = 404; throw e; }
      return gf("/groups/" + g.id + "/drive?$select=id,webUrl"); }).then(function(d){ return vaCache = {drive: d.id, web: d.webUrl}; }); },
  kinder: function(itemId){ return VaLive.drive().then(function(D){
    return alle("/drives/" + D.drive + "/" + (itemId ? "items/" + itemId : "root") + "/children?$select=id,name,folder,file,webUrl,size,lastModifiedDateTime&$top=200"); }); },
  blatt: function(itemId){ return VaLive.drive().then(function(D){ var wb = "/drives/" + D.drive + "/items/" + itemId + "/workbook";
    return gf(wb + "/worksheets?$select=name").then(function(l){ var n = l.value[0].name;
      return gf(wb + blattPfad(n) + "/usedRange(valuesOnly=true)?$select=address,values").then(function(b){ return {name: n, werte: ausrichten(b.address, b.values || [])}; }); }); }); },
  schreiben: function(itemId, blatt, adresse, werte){ return VaLive.drive().then(function(D){
    return gf("/drives/" + D.drive + "/items/" + itemId + "/workbook" + blattPfad(blatt) + "/range(address='" + adresse + "')", {method: "PATCH", body: {values: werte}}); }); },
  termine: function(){ var j = new Date().getFullYear(), jahre = [j - 1, j, j + 1, j + 2];
    return KalLive.finden().then(function(){ return Promise.all(jahre.map(function(y){ return KalLive.termine(y + "-01-01T00:00:00", (y + 1) + "-01-01T00:00:00"); })); })
      .then(function(l){ return [].concat.apply([], l); }).catch(function(){ return []; }); }
};
var VaTest = {
  baum: null,
  bauen: function(){ if(VaTest.baum) return VaTest.baum; var j = new Date().getFullYear(), n = 0;
    function d(name, kinder){ return {id: "i" + (++n), name: name, folder: {childCount: (kinder || []).length}, webUrl: "#", kinder: kinder || []}; }
    function f(name){ return {id: "i" + (++n), name: name, file: {}, webUrl: "#", size: 20000, lastModifiedDateTime: j + "-10-01T10:00:00Z"}; }
    VaTest.baum = d("root", [d("General", [d("Aufgaben Veranstaltungen", [f("Aufgabe Getränkestand.xlsx")])]),
      d("Weihnachtsfeier", [d("Weihnachtsfeier " + (j - 1), [f("Anmeldungen.xlsx"), f("Tombolapreise.docx"), f("Getränke.docx"), f("Bestellung Essen.docx"), f("Rückblick Verbesserungen.docx")]), d("Weihnachtsfeier " + j, [f("Tombola " + j + ".docx"), f("Dienstplan Weihnachtsfeier.xlsx")]),
        d("Weihnachtsfeier " + (j + 1), [d("Dienstplan"), d("Getränke"), d("Essen"), d("Tombola", [f("Tombola " + (j + 1) + ".xlsx")])])]),
      d("Schlachtfest", [d("Schlachtfest " + j, [f("Dienstplan Schlachtfest " + j + ".xlsx"), f("Plakat.pdf"), f("Rechnung Beispiel.pdf")])]),
      d("Kerb", [d("Kerb " + j, [f("Dienstplan Kerb.xlsx"), f("Speisekarte " + j + ".docx")])])]);
    return VaTest.baum; },
  suchen: function(id, k){ k = k || VaTest.bauen(); if(k.id === id) return k; for(var i = 0; i < (k.kinder || []).length; i++){ var r = VaTest.suchen(id, k.kinder[i]); if(r) return r; } return null; },
  drive: function(){ return Promise.resolve({drive: "test", web: "#"}); },
  kinder: function(id){ var k = id ? VaTest.suchen(id) : VaTest.bauen(); return Promise.resolve((k && k.kinder) || []); },
  tombola: null,
  blatt: function(id){ var it = VaTest.suchen(id);
    if(it && /tombola/i.test(it.name)){ if(!VaTest.tombola){ VaTest.tombola = [["Tombola (Beispieldaten)"], ["Hinweis"], [], ["Nr", "Preis", "Firma / Spender", "Ansprechpartner", "Telefon", "Angefragt am", "Angefragt von", "Zugesagt", "Abgeholt", "Wert (€)", "Bemerkung"]];
        for(var i = 1; i <= 6; i++) VaTest.tombola.push([i, i === 1 ? "Gutschein" : "", i < 4 ? "Firma " + String.fromCharCode(64 + i) : "", "", "", "", i === 1 ? "AB" : "", i === 1 ? "Ja" : "offen", "", i === 1 ? 25 : "", ""]); }
      return Promise.resolve({name: "Tombola", werte: JSON.parse(JSON.stringify(VaTest.tombola))}); }
    return Promise.resolve({name: "Dienstplan", werte: [["Dienstplan (Beispieldaten)", "", ""], ["Zeit", "Theke", "Kasse"], ["14–16 Uhr", "Person A", "Person B"], ["16–18 Uhr", "Person C", "Person D"]]}); },
  schreiben: function(id, blatt, adresse, werte){ window.__tombola = (window.__tombola || []).concat([{adresse: adresse, werte: werte}]); return Test.schreibenIn(VaTest.tombola, blatt, adresse, werte); },
  termine: function(){ var j = new Date().getFullYear(); return Promise.resolve([{subject: "Weihnachtsfeier", start: {dateTime: j + "-12-19T18:00:00"}, end: {dateTime: j + "-12-19T23:00:00"}, isAllDay: false}]); }
};
function VAQ(){ return TEST ? VaTest : VaLive; }
function vaFehler(e){ return e && (e.status === 403 || e.status === 404) ? "Kein Zugriff auf das Team „Veranstaltungen“. Du musst Mitglied in diesem Team sein – bitte bei der Verwaltung melden." : fehlerText(e); }
function terminText(ev){ var a = kalDatum(ev.start.dateTime); if(!a) return ""; var d = isoZuDatum(a.iso);
  return WT[d.getDay()] + ", " + pad(a.t) + "." + pad(a.mo) + "." + a.j + (ev.isAllDay ? "" : ", " + a.zeit + " Uhr"); }
function terminLink(ev){ var a = kalDatum(ev.start.dateTime), l = el("a", null, "📅 " + terminText(ev)); l.href = "#/kalender/" + a.j + "-" + pad(a.mo); return l; }

/* Ordnerstruktur eines Fests: Jahresordner und Übriges („Allgemein“). */
function vaFestLaden(f){
  return VAQ().kinder(null).then(function(wurzel){
    var o = wurzel.filter(function(x){ return x.folder && f.ordner.test(x.name); })[0];
    if(!o) return {ordner: null, jahre: {}, allgemein: []};
    return VAQ().kinder(o.id).then(function(k){ var jahre = {}, allgemein = [];
      k.forEach(function(x){ var j = x.folder && !f.ohneJahre ? jahrAus(x.name) : null; if(j) jahre[j] = x; else allgemein.push(x); });
      return {ordner: o, jahre: jahre, allgemein: allgemein}; }); });
}
function seiteVeranstaltungen(s, key, jahr, ber){
  if(key) return seiteFest(s, key, jahr, ber);
  pfad([["Start", "#/"], ["Veranstaltungen"]]);
  s.textContent = "";
  add(s, el("h1", null, "Veranstaltungen und Dienstpläne"), el("p", "lead", "Alle Feste des Vereins. Die Unterlagen liegen im Team „Veranstaltungen“ in Teams – je Fest und Jahr ein Ordner. Termine ohne festen Tag sind Vorschläge."));
  var k = el("div", "kacheln"); add(s, k);
  var hinweis = el("div"); add(s, hinweis);
  var heute = heuteISO();
  VAQ().termine().then(function(termine){
    FESTE.forEach(function(f){
      var naechster = termine.filter(function(ev){ return festVonTitel(ev.subject) === f && kalDatum(ev.start.dateTime).iso >= heute; })
        .sort(function(a, b){ return a.start.dateTime < b.start.dateTime ? -1 : 1; })[0];
      var text = naechster ? "Nächster Termin: " + terminText(naechster) : (f.festplaner ? f.monat : (function(){ var g = geplanteJahre(f, new Date().getFullYear() + 4)[0];
        var dieses = new Date().getFullYear(); if(g === dieses && termine.some(function(ev){ return festVonTitel(ev.subject) === f && kalDatum(ev.start.dateTime).j === dieses; })) g = geplanteJahre(f, dieses + 4)[1];
        return g ? "Nächstes Mal: " + offenText(f, g) : ""; })());
      add(k, kachel(f.sym, f.name, text, vaLink(f), true));
    });
    return VAQ().kinder(null).then(function(wurzel){
      var fremd = wurzel.filter(function(x){ return x.folder && !/^general$/i.test(x.name) && !FESTE.some(function(f){ return f.ordner.test(x.name); }); });
      var gen = wurzel.filter(function(x){ return x.folder && /^general$/i.test(x.name); })[0];
      if(gen) add(k, kachel("🗂️", "Allgemein", "Aufgabenlisten der Stände und allgemeine Unterlagen", "#/veranstaltungen/allgemein", true));
      if(fremd.length) add(hinweis, el("div", "hinweis", "Weitere Ordner im Team, die noch keinem Fest zugeordnet sind: " + fremd.map(function(x){ return x.name; }).join(", ") + "."));
    });
  }).catch(function(e){ add(hinweis, el("div", "fehler", vaFehler(e))); });
}
function seiteFest(s, key, jahr, ber){
  if(key === "allgemein") return seiteVaAllgemein(s, jahr);
  var f = festVonKey(key);
  if(!f){ pfad([["Start", "#/"], ["Veranstaltungen", "#/veranstaltungen"], ["?"]]); fehlerZeigen("Dieses Fest gibt es nicht."); return; }
  if(jahr) return seiteFestJahr(s, f, jahr, ber);
  pfad([["Start", "#/"], ["Veranstaltungen", "#/veranstaltungen"], [f.name]]);
  ladeHinweis(s); seite().firstChild.textContent = "Ordner werden aus Teams geladen …";
  Promise.all([vaFestLaden(f), VAQ().termine()]).then(function(x){
    var D = x[0], termine = x[1].filter(function(ev){ return festVonTitel(ev.subject) === f; }), dieses = new Date().getFullYear();
    s.textContent = "";
    add(s, el("h1", null, f.sym + " " + f.name));
    var lead = el("p", "lead", (f.alle === 2 ? "Alle zwei Jahre" + (f.ab ? ", nächstes Mal " + f.ab : "") : (f.alle ? "Jedes Jahr" : "Einmalig")) + (f.monat ? ", " + f.monat : "") + (f.alle ? " (Vorschlag, solange kein Termin im Kalender steht)" : "") + ". ");
    if(D.ordner && D.ordner.webUrl){ var o = el("a", null, "Ordner in Teams öffnen"); o.href = D.ordner.webUrl; o.target = "_blank"; o.rel = "noopener"; add(lead, o); }
    add(s, lead);
    if(f.festplaner){ var fp = el("a", "btn primary", "Zum Festplaner 2028"); fp.href = "/festplaner/"; add(s, add(el("div", "reihe"), fp)); }
    if(!D.ordner) add(s, el("div", "hinweis", "Im Team „Veranstaltungen“ gibt es für dieses Fest noch keinen Kanal bzw. Ordner. Anlegen in Teams: Team „Veranstaltungen“ › „+“ Kanal „" + f.name + "“."));
    var jahre = Object.keys(D.jahre).map(Number);
    termine.forEach(function(ev){ var j = kalDatum(ev.start.dateTime).j; if(jahre.indexOf(j) < 0) jahre.push(j); });
    geplanteJahre(f, dieses + 2).forEach(function(j){ if(jahre.indexOf(j) < 0) jahre.push(j); });
    jahre.sort(function(a, b){ return b - a; });
    if(!jahre.length){ if(D.allgemein.length) add(s, el("h2", null, "Unterlagen"), dateiListe(D.allgemein)); return; }
    var t = el("table", "tabelle"), tb = el("tbody");
    add(t, add(el("thead"), add(el("tr"), el("th", null, "Jahr"), el("th", null, "Termin"), el("th", null, "Unterlagen"))), tb);
    jahre.forEach(function(j){
      var evs = termine.filter(function(ev){ return kalDatum(ev.start.dateTime).j === j; }), ord = D.jahre[j];
      var tdT = el("td"); if(evs.length) evs.forEach(function(ev, i){ if(i) add(tdT, el("br")); add(tdT, terminLink(ev)); });
      else add(tdT, el("span", "grau", geplanteJahre(f, j).indexOf(j) > -1 ? offenText(f, j) : "nicht im Kalender"));
      var a = el("a", null, String(j)); a.href = vaLink(f, j);
      add(tb, add(el("tr"), add(el("td"), add(el("b"), a)), tdT, el("td", ord ? null : "grau", ord ? ((ord.folder && ord.folder.childCount) || 0) + " Einträge" : "noch kein Ordner"))); });
    add(s, add(el("div", "scroll"), t));
    if(D.allgemein.length){ add(s, el("h2", null, "Allgemein (für alle Jahre)")); add(s, dateiListe(D.allgemein)); }
  }).catch(function(e){ fehlerZeigen(vaFehler(e)); });
}
function dateiListe(liste){
  var t = el("table", "tabelle"), tb = el("tbody");
  add(t, add(el("thead"), add(el("tr"), el("th", null, "Datei"), el("th", null, "Geändert"))), tb);
  liste.slice().sort(function(a, b){ return (b.folder ? 1 : 0) - (a.folder ? 1 : 0) || a.name.localeCompare(b.name, "de"); }).forEach(function(x){
    var a = el("a", "doc", (x.folder ? "📁 " : dateiSym(x.name)) + x.name); a.href = x.webUrl || "#"; a.target = "_blank"; a.rel = "noopener";
    add(tb, add(el("tr"), add(el("td"), a), el("td", "grau", x.lastModifiedDateTime ? auDatum(x.lastModifiedDateTime) : ""))); });
  if(!liste.length) add(tb, add(el("tr", "leer"), el("td", null, "Keine Dateien."), el("td")));
  return add(el("div", "scroll"), t);
}
function dateiSym(n){ return /\.xlsx?$|\.ods$/i.test(n) ? "📊 " : (/\.docx?$|\.odt$/i.test(n) ? "📄 " : (/\.pdf$/i.test(n) ? "📕 " : (/\.(jpe?g|png|heic|gif)$/i.test(n) ? "🖼️ " : "📎 "))); }
function seiteFestJahr(s, f, jahr, ber){
  jahr = +jahr; ber = ber ? decodeURIComponent(ber) : "";
  pfad([["Start", "#/"], ["Veranstaltungen", "#/veranstaltungen"], [f.name, vaLink(f)], [String(jahr)]]);
  ladeHinweis(s); seite().firstChild.textContent = "Ordner werden aus Teams geladen …";
  Promise.all([vaFestLaden(f), VAQ().termine()]).then(function(x){
    var D = x[0], evs = x[1].filter(function(ev){ return festVonTitel(ev.subject) === f && kalDatum(ev.start.dateTime).j === jahr; }), ord = D.jahre[jahr];
    return (ord ? VAQ().kinder(ord.id) : Promise.resolve([])).then(function(inhalt){
      s.textContent = "";
      add(s, el("h1", null, f.sym + " " + f.name + " " + jahr));
      var lead = el("p", "lead");
      if(evs.length) evs.forEach(function(ev){ add(lead, terminLink(ev), " "); });
      else add(lead, geplanteJahre(f, jahr).indexOf(jahr) > -1 ? offenText(f, jahr) + ". " : "Kein Termin im Vereinskalender. ");
      if(ord && ord.webUrl){ var o = el("a", null, "Ordner in Teams öffnen"); o.href = ord.webUrl; o.target = "_blank"; o.rel = "noopener"; add(lead, " · ", o); }
      add(s, lead);
      if(!ord){ add(s, el("div", "hinweis", "Für " + jahr + " gibt es im Team „Veranstaltungen“ noch keinen Ordner. Anlegen in Teams im Kanal „" + (D.ordner ? D.ordner.name : f.name) + "“: Ordner „" + f.name + " " + jahr + "“.")); return; }
      /* Bereiche: Dateien nach Namen sortiert. Ein Unterordner, dessen Name zu einem Bereich passt (z. B. „Dienstplan“, „Tombola“),
         gehört zu diesem Knopf; andere Unterordner bekommen einen eigenen Knopf. */
      var gruppen = {}, reihenfolge = [];
      function rein(k, name, sym, it){ if(!gruppen[k]){ gruppen[k] = {key: k, name: name, sym: sym, dateien: [], ordner: []}; reihenfolge.push(k); } if(it) gruppen[k].dateien.push(it); return gruppen[k]; }
      BEREICHE.forEach(function(b){ if(b.key === "dienstplan" || b.key === "getraenke" || b.key === "essen" || b.key === "tombola") rein(b.key, b.name, b.sym); });
      inhalt.forEach(function(it){
        var b = bereichVon(it.name);
        if(it.folder){ if(b.key !== "sonstiges" && b.key !== "bilder" && b.name.toLowerCase().indexOf(it.name.toLowerCase().split(" ")[0]) === 0) rein(b.key, b.name, b.sym).ordner.push(it);
          else rein("ordner:" + it.name, it.name, b.key === "sonstiges" ? "📁" : b.sym).ordner.push(it); return; }
        rein(b.key, b.name, b.sym, it); });
      var knoepfe = el("div", "bereiche");
      reihenfolge.forEach(function(k){ var g = gruppen[k], n = g.dateien.length + g.ordner.reduce(function(m, o){ return m + ((o.folder && o.folder.childCount) || 0); }, 0);
        var a = el("a", "bereich" + (ber === k ? " an" : "") + (n ? "" : " leer")); a.href = vaLink(f, jahr, k);
        add(a, add(el("span", "rund"), g.sym, n ? el("span", "zahl", String(n)) : null), el("span", "name", g.name)); add(knoepfe, a); });
      add(s, knoepfe);
      var box = el("div"); add(s, box);
      if(!ber){ add(box, el("p", "lead", "Wähle einen Bereich.")); return; }
      var g = gruppen[ber]; if(!g){ add(box, el("p", "lead", "Diesen Bereich gibt es in " + jahr + " nicht.")); return; }
      var karte = el("div", "karte"); add(box, karte); add(karte, el("h2", null, g.sym + " " + g.name));
      if(g.ordner.length){ var lp = el("p", "lead"); g.ordner.forEach(function(o, i){ var lo = el("a", null, "Ordner „" + o.name + "“ in Teams öffnen"); lo.href = o.webUrl; lo.target = "_blank"; lo.rel = "noopener"; if(i) add(lp, " · "); add(lp, lo); }); add(karte, lp); }
      var platz = el("div", "laden", "Wird geladen …"); add(karte, platz);
      Promise.all(g.ordner.map(function(o){ return VAQ().kinder(o.id); })).then(function(l){
        var alleD = g.dateien.concat([].concat.apply([], l));
        if(!alleD.length){ platz.replaceWith(el("p", "lead", "Noch keine Unterlagen für " + g.name + " " + jahr + ". Dateien dafür " + (g.ordner.length ? "in den Ordner „" + g.ordner[0].name + "“" : "in den Jahresordner") + " in Teams legen; im Jahresordner entscheidet der Name über den Bereich (z. B. „Dienstplan …“, „Getränke …“, „Tombola …“).")); return; }
        platz.replaceWith(dateiListe(alleD));
        /* Dienstplan und Tombola: Excel direkt als Tabelle zeigen */
        if(ber === "dienstplan" || ber === "tombola") alleD.filter(function(d){ return !d.folder && /\.xlsx$/i.test(d.name); }).forEach(function(d){
          var platz2 = el("div", "laden", "„" + d.name + "“ wird geladen …"); add(karte, platz2);
          VAQ().blatt(d.id).then(function(b){ var c = el("div");
              if(ber === "tombola" && tombolaListe(b.werte)) add(c, el("h2", null, d.name), tombolaEditor(d, b));
              else add(c, el("h2", null, d.name + " · Blatt „" + b.name + "“"), tabelleAus(b.werte));
              platz2.replaceWith(c); })
            .catch(function(){ platz2.replaceWith(el("p", "grau", "„" + d.name + "“ ließ sich nicht als Tabelle anzeigen – bitte in Excel öffnen.")); }); });
      }).catch(function(e){ platz.replaceWith(el("div", "fehler", vaFehler(e))); });
    });
  }).catch(function(e){ fehlerZeigen(vaFehler(e)); });
}
/* Tombola-Liste (Vorlage Tools/make_tombola.py): Kopf in Zeile 4, Einträge ab Zeile 5 mit Nr in Spalte A.
   Spalten B–K: Preis, Firma, Ansprechpartner, Telefon, Angefragt am, Angefragt von, Zugesagt, Abgeholt, Wert, Bemerkung.
   Geschrieben wird je Zeile B–K; Nr (A) und die Zähler unter der Liste (Formeln) bleiben unberührt. */
var TB_SP = ["preis", "firma", "ansprech", "tel", "am", "von", "zugesagt", "abgeholt", "wert", "bem"];
function tombolaListe(w){ var k = (w[3] || []).map(function(x){ return txt(x).toLowerCase(); }); return k[0] === "nr" && k[1] === "preis" && /firma/.test(k[2] || ""); }
function tombolaZeilen(w){ var out = []; for(var z = 5; z <= w.length; z++){ var nr = zelle(w, z, 0); if(typeof nr !== "number") continue;
  var o = {zeile: z, nr: nr}; TB_SP.forEach(function(k, i){ o[k] = zelle(w, z, i + 1); }); out.push(o); } return out; }
function tombolaEditor(d, b){
  var box = el("div"), sum = el("div", "summe"), t = el("table", "tabelle tombola"), tb = el("tbody"), zeilen = tombolaZeilen(b.werte);
  var ex = el("a", null, "In Excel öffnen"); ex.href = d.webUrl; ex.target = "_blank"; ex.rel = "noopener";
  add(box, add(el("p", "lead"), "Direkt hier bearbeiten: Feld ändern, dann „Speichern“ in der Zeile. Gespeichert wird in der Excel in Teams. ", ex), sum);
  function zaehlen(){ var ja = 0, offen = 0, wert = 0; zeilen.forEach(function(z){ if(!txt(z.firma) && !txt(z.preis)) return; if(/^ja/i.test(txt(z.zugesagt))) ja++; else if(!/^nein/i.test(txt(z.zugesagt))) offen++; wert += zahl(z.wert) || 0; });
    sum.textContent = ""; add(sum, add(el("span"), "zugesagt ", el("b", null, String(ja))), add(el("span"), "offen ", el("b", null, String(offen))), add(el("span"), "Wert ", el("b", null, euro(wert)))); }
  add(t, add(el("thead"), add(el("tr"), el("th", null, "Nr"), el("th", null, "Firma / Spender"), el("th", null, "Preis"), el("th", null, "Ansprechpartner · Telefon"), el("th", null, "Angefragt"), el("th", null, "Zugesagt"), el("th", null, "Abgeholt"), el("th", null, "Wert €"), el("th", null, "Bemerkung"), el("th"))), tb);
  function eingabe(v, typ, ph){ var i = el("input"); i.type = typ || "text"; i.value = v == null ? "" : v; if(ph) i.placeholder = ph; return i; }
  function auswahl(v, werte){ var s2 = el("select"); werte.forEach(function(w){ var o = el("option", null, w || "–"); o.value = w; add(s2, o); }); s2.value = werte.indexOf(txt(v)) > -1 ? txt(v) : werte[0]; return s2; }
  function zeile(z){
    var f = {firma: eingabe(txt(z.firma)), preis: eingabe(txt(z.preis)), ansprech: eingabe(txt(z.ansprech), "text", "Name"), tel: eingabe(txt(z.tel), "tel", "Telefon"),
      am: eingabe(typeof z.am === "number" ? (function(dd){ return dd.getUTCFullYear() + "-" + pad(dd.getUTCMonth() + 1) + "-" + pad(dd.getUTCDate()); })(serieZuDatum(z.am)) : "", "date"),
      von: eingabe(txt(z.von), "text", "von"), zugesagt: auswahl(z.zugesagt, ["offen", "Ja", "Nein"]), abgeholt: auswahl(z.abgeholt, ["", "Ja", "Nein"]),
      wert: eingabe(z.wert === "" || z.wert == null ? "" : zahlText(z.wert), "text", "0"), bem: eingabe(txt(z.bem))};
    f.wert.inputMode = "decimal";
    var sp = btn("primary klein", "Speichern", function(){ speichern(z, f, sp, tr); }); sp.disabled = true;
    Object.keys(f).forEach(function(k){ f[k].addEventListener("input", function(){ sp.disabled = false; tr.classList.add("geaendert"); }); f[k].addEventListener("change", function(){ sp.disabled = false; tr.classList.add("geaendert"); }); });
    var tr = add(el("tr", /^ja/i.test(txt(z.zugesagt)) ? "zugesagt" : null), el("td", "z", String(z.nr)), add(el("td"), f.firma), add(el("td"), f.preis), add(el("td"), f.ansprech, f.tel),
      add(el("td"), f.am, f.von), add(el("td"), f.zugesagt), add(el("td"), f.abgeholt), add(el("td"), f.wert), add(el("td"), f.bem), add(el("td"), sp));
    return tr; }
  function speichern(z, f, sp, tr){
    var wert = eingabeZahl(f.wert.value), d0 = isoZuDatum(f.am.value);
    if(f.wert.value.trim() && wert == null){ zeigeToast("Wert bitte als Zahl eingeben."); return; }
    var werte = [[f.preis.value.trim(), f.firma.value.trim(), f.ansprech.value.trim(), f.tel.value.trim(), d0 ? datumZuSerie(d0) : "", f.von.value.trim(), f.zugesagt.value, f.abgeholt.value, wert == null ? "" : wert, f.bem.value.trim()]];
    sp.disabled = true;
    /* vorher frisch lesen: steht in der Zeile noch dieselbe Nr? */
    VAQ().blatt(d.id).then(function(b2){ if(zelle(b2.werte, z.zeile, 0) !== z.nr) throw new Error("Die Liste wurde inzwischen in Excel geändert. Bitte die Seite neu laden.");
      return VAQ().schreiben(d.id, b2.name, "B" + z.zeile + ":K" + z.zeile, werte); })
      .then(function(){ TB_SP.forEach(function(k, i){ z[k] = werte[0][i]; }); tr.classList.remove("geaendert"); tr.classList.toggle("zugesagt", z.zugesagt === "Ja"); zaehlen(); zeigeToast((z.firma || "Zeile " + z.nr) + " gespeichert ✓"); })
      .catch(function(e){ sp.disabled = false; zeigeToast(e && e.status ? vaFehler(e) : (e.message || "Speichern hat nicht geklappt.")); });
  }
  zeilen.forEach(function(z){ add(tb, zeile(z)); });
  add(box, add(el("div", "scroll"), t));
  var frei = zeilen.filter(function(z){ return !txt(z.firma) && !txt(z.preis); }).length;
  add(box, el("p", "grau", frei ? frei + " freie Zeilen am Ende der Liste. Mehr Zeilen: in Excel über der Zusammenfassung Zeilen einfügen und Nr fortführen." : "Keine freie Zeile mehr – in Excel über der Zusammenfassung Zeilen einfügen und Nr fortführen."));
  zaehlen(); return box;
}
function tabelleAus(werte){
  var zeilen = werte.filter(function(r){ return r.some(function(v){ return !leer(v); }); }).slice(0, 300);
  var breite = zeilen.reduce(function(m, r){ var l = r.length; while(l && leer(r[l - 1])) l--; return Math.max(m, l); }, 0);
  var t = el("table", "tabelle"), tb = el("tbody"); add(t, tb);
  zeilen.forEach(function(r){ var tr = el("tr"); for(var i = 0; i < breite; i++){ var v = r[i]; add(tr, el("td", null, typeof v === "number" && v > 40000 && v < 60000 ? datumText(v) : (typeof v === "number" && v > 0 && v < 1 ? zeitText(v) : txt(v)))); } add(tb, tr); });
  return add(el("div", "scroll"), t);
}
function zeitText(v){ var m = Math.round(v * 1440); return pad(Math.floor(m / 60)) + ":" + pad(m % 60); }
function seiteVaAllgemein(s){
  pfad([["Start", "#/"], ["Veranstaltungen", "#/veranstaltungen"], ["Allgemein"]]);
  ladeHinweis(s);
  VAQ().kinder(null).then(function(w){ var gen = w.filter(function(x){ return x.folder && /^general$/i.test(x.name); })[0];
    return gen ? VAQ().kinder(gen.id).then(function(l){ return {gen: gen, l: l}; }) : {gen: null, l: []}; }).then(function(D){
    s.textContent = ""; add(s, el("h1", null, "🗂️ Allgemein"), el("p", "lead", "Kanal „General“ im Team „Veranstaltungen“: Aufgabenlisten der Stände und allgemeine Unterlagen."));
    add(s, dateiListe(D.l));
  }).catch(function(e){ fehlerZeigen(vaFehler(e)); });
}

/* ---------- Jugend und Aktive ----------
 * Stand: Homepage 1928schneeberg.de (gelesen am 24.11.2026). Hier stehen bewusst keine Namen –
 * Trainer, Betreuer und Kontakte stehen auf der Homepage (Link je Mannschaft). */
var MANNSCHAFTEN = {
  aktive: {titel: "Aktive", sym: "🥅", lead: "Herrenmannschaften, Saison 2026/2027.", gruppen: [
    {name: "1. & 2. Mannschaft", link: "https://1928schneeberg.de/1-2-mannschaft/", teams: [
      {name: "1. Mannschaft", info: "A-Klasse"}, {name: "2. Mannschaft", info: "B-Klasse"}],
      training: ["Dienstag 19:00–20:45 Uhr", "Freitag 18:30–20:15 Uhr (anschließend Spielersitzung)"], ort: "Ort auf der Homepage nicht angegeben"}]},
  jugend: {titel: "Jugend", sym: "🧒", lead: "Alle Jugendmannschaften mit Trainingszeiten.", gruppen: [
    {name: "Großfeld", link: "https://1928schneeberg.de/jugend-grossfeld/", teams: [
      {name: "A-Jugend", info: "U19", training: ["Montag 19:00–20:30 Uhr (Amorbach)", "Mittwoch 19:00–20:30 Uhr (Amorbach)"]},
      {name: "B-Jugend", info: "U17", training: ["Montag 18:00–19:30 Uhr (Kirchzell)", "Mittwoch 18:00–19:30 Uhr (Kirchzell)"]},
      {name: "C-Jugend", info: "U15", training: ["Dienstag 18:00–19:30 Uhr (Amorbach)", "Donnerstag 18:00–19:30 Uhr (Schneeberg)"]},
      {name: "D-Jugend", info: "U13", training: ["Dienstag 17:30–19:00 Uhr (Schneeberg)", "Donnerstag 17:30–19:00 Uhr (Amorbach)"]}]},
    {name: "Kleinfeld", link: "https://1928schneeberg.de/jugend-kleinfeld/", teams: [
      {name: "E-Jugend", info: "U11", training: ["Montag 17:00–18:30 Uhr (Schneeberg)", "Mittwoch 17:00–18:30 Uhr (Schneeberg)"]},
      {name: "F-Jugend", info: "U9", training: ["Dienstag 17:00–18:30 Uhr (Schneeberg)", "Donnerstag 17:00–18:30 Uhr (Schneeberg)"]},
      {name: "Bambini", info: "", training: ["Donnerstag 16:45–18:00 Uhr (Schneeberg)"]}]},
    {name: "Mädchen", link: "https://1928schneeberg.de/maedchenmannschaft/", teams: [
      {name: "Mädchenmannschaft", info: "12 bis 18 Jahre", training: ["Mittwoch 19:00–20:15 Uhr (Sportplatz Schneeberg)"]}]}]}
};
/* Daten aus Teams: Team „TEAM Jugend“ (Gruppe TEAMJugend), Ordner General.
 *  - „Mannschaften.xlsx“, Blatt „Mannschaften“ (Kopf Zeile 4): A Bereich · B Gruppe · C Mannschaft · D Klasse/Alter · E Trainer
 *    F Betreuer · G Ansprechpartner · H Telefon · I E-Mail · J Homepage · K Bemerkung
 *  - „Übersicht Training_Spiel.xlsx“: Wochenplan je Saison (Blattname z. B. „2026_27“); angezeigt wird die neueste Saison.
 * Ohne Zugriff auf das Team zeigt das Portal die Angaben von der Homepage (MANNSCHAFTEN, ohne Namen). */
var MS = { gruppe: "TEAMJugend", liste: "General/Mannschaften.xlsx", plan: "General/Übersicht Training_Spiel.xlsx" };
var msCache = null;
function msArt(name){ name = txt(name); if(/^u\s?\d|bambini|jugend|mädchen|maedchen/i.test(name)) return "jugend"; if(/mannschaft|damen|herren|^ah\b|alte herren/i.test(name)) return "aktive"; return ""; }
var MsLive = {
  laden: function(){ if(msCache) return Promise.resolve(msCache);
    return gf("/groups?$filter=" + encodeURIComponent("mailNickname eq '" + MS.gruppe + "'") + "&$select=id").then(function(r){
      var g = r && r.value && r.value[0]; if(!g){ var e = new Error("Gruppe"); e.status = 404; throw e; }
      var dr = "/groups/" + g.id + "/drive";
      var liste = gf(dr + "/root:/" + pfadK(MS.liste) + "?$select=id,webUrl").then(function(it){
        return gf(dr + "/items/" + it.id + "/workbook/worksheets('Mannschaften')/usedRange(valuesOnly=true)?$select=address,values").then(function(b){ return {web: it.webUrl, werte: ausrichten(b.address, b.values || [])}; }); });
      var plan = gf(dr + "/root:/" + pfadK(MS.plan) + "?$select=id,webUrl").then(function(it){ var wb = dr + "/items/" + it.id + "/workbook";
        return gf(wb + "/worksheets?$select=name").then(function(l){
          var saisons = l.value.map(function(x){ return x.name; }).filter(function(n){ return /^\d{4}_\d{2}$/.test(n); }).sort(), n = saisons[saisons.length - 1] || l.value[0].name;
          return gf(wb + blattPfad(n) + "/usedRange(valuesOnly=true)?$select=values").then(function(b){ return {web: it.webUrl, blatt: n, werte: b.values || []}; }); }); }).catch(function(){ return null; });
      return Promise.all([liste, plan]).then(function(x){ return msCache = {liste: x[0], plan: x[1]}; }); }); }
};
var MsTest = { laden: function(){ return Promise.resolve({
  liste: {web: "#", werte: [[], [], [], ["Bereich"], ["Aktive", "Herren", "1. Mannschaft", "A-Klasse", "Trainer A (Beispiel)", "", "Trainer A", "0170/000000", "", "https://1928schneeberg.de/1-2-mannschaft/", ""],
    ["Jugend", "Großfeld", "C-Jugend", "U15", "Trainer B, Trainer C", "Betreuerin D", "Jugendleiter E", "0170/111111", "e@example.org", "https://1928schneeberg.de/jugend-grossfeld/", "Beispieldaten"],
    ["Jugend", "Kleinfeld", "Bambini", "", "Trainer F", "", "", "", "", "", ""]]},
  plan: {web: "#", blatt: "2026_27", werte: [["Übersicht Training & Spiel - Saison 2026/27"], ["", "", "Montag", "Dienstag", "Mittwoch"], ["Bambini", "Training", "", "", "17.00 - 18.30"], ["", "Spiel", "Turniere"], ["U 15 - C", "Training", "A 18.30", "", "S 18.30"], ["1. Mannschaft", "Training", "", "x", ""], ["Schneeberg", "Juli/August", "November"]]}}); } };
function MSQ(){ return TEST ? MsTest : MsLive; }
function seiteMannschaften(s, art){
  var M = MANNSCHAFTEN[art];
  pfad([["Start", "#/"], [M.titel]]);
  ladeHinweis(s); seite().firstChild.textContent = "Mannschaften werden aus Teams geladen …";
  MSQ().laden().then(function(D){
    s.textContent = "";
    add(s, el("h1", null, M.sym + " " + M.titel));
    var lead = el("p", "lead", "Aus dem Team „TEAM Jugend“: Liste „Mannschaften“ (Trainer, Betreuer, Ansprechpartner) und „Übersicht Training_Spiel“ (Trainingszeiten). Geändert wird dort. ");
    var l1 = el("a", null, "Liste öffnen"); l1.href = D.liste.web; l1.target = "_blank"; l1.rel = "noopener"; add(lead, l1);
    if(D.plan){ var l2 = el("a", null, "Übersicht öffnen"); l2.href = D.plan.web; l2.target = "_blank"; l2.rel = "noopener"; add(lead, " · ", l2); }
    add(s, lead);
    var w = D.liste.werte, gruppen = {}, reihe = [];
    for(var z = ERSTE; z <= w.length; z++){ var name = txt(zelle(w, z, 2)); if(!name) continue;
      var bereich = txt(zelle(w, z, 0)).toLowerCase(); if((art === "jugend") !== (bereich === "jugend")) continue;
      var gr = txt(zelle(w, z, 1)) || "Mannschaften"; if(!gruppen[gr]){ gruppen[gr] = []; reihe.push(gr); }
      gruppen[gr].push({name: name, klasse: txt(zelle(w, z, 3)), trainer: txt(zelle(w, z, 4)), betreuer: txt(zelle(w, z, 5)), ansprech: txt(zelle(w, z, 6)), tel: txt(zelle(w, z, 7)), mail: txt(zelle(w, z, 8)), link: txt(zelle(w, z, 9)), bem: txt(zelle(w, z, 10))}); }
    reihe.forEach(function(gr){ var k = el("div", "karte"); add(s, k); add(k, el("h2", null, gr));
      var t = el("table", "tabelle"), tb = el("tbody");
      add(t, add(el("thead"), add(el("tr"), el("th", null, "Mannschaft"), el("th", null, "Trainer"), el("th", null, "Betreuer"), el("th", null, "Ansprechpartner"))), tb);
      gruppen[gr].forEach(function(m){
        var td1 = add(el("td"), el("b", null, m.name)); if(m.klasse) add(td1, el("div", "grau", m.klasse));
        if(m.link){ var hl = el("a", "grau", "Homepage"); hl.href = m.link; hl.target = "_blank"; hl.rel = "noopener"; add(td1, el("div"), hl); }
        var td4 = add(el("td"), m.ansprech || "–");
        if(m.tel){ var tl = el("a", null, "📞 " + m.tel); tl.href = "tel:" + m.tel.replace(/[^\d+]/g, ""); add(td4, el("br"), tl); }
        if(m.mail){ var ml = el("a", null, "✉️ " + m.mail); ml.href = "mailto:" + m.mail; add(td4, el("br"), ml); }
        var tr = add(el("tr"), td1, el("td", null, m.trainer || "–"), el("td", null, m.betreuer || "–"), td4); add(tb, tr);
        if(m.bem) add(tb, add(el("tr"), add(el("td"), ""), add(el("td"), el("span", "grau", m.bem)), el("td"), el("td"))); });
      add(k, add(el("div", "scroll"), t)); });
    if(!reihe.length) add(s, el("p", "lead", "In der Liste „Mannschaften“ steht noch keine Mannschaft mit Bereich „" + (art === "jugend" ? "Jugend" : "Aktive") + "“."));
    /* Wochenübersicht Training & Spiel: Zeilen der passenden Mannschaften, Kopfzeilen und übrige Zeilen bleiben */
    var k2 = el("div", "karte"); add(s, k2); add(k2, el("h2", null, "Training und Spiele" + (D.plan ? " · Saison " + D.plan.blatt.replace("_", "/") : "")));
    if(!D.plan){ add(k2, el("p", "lead", "„Übersicht Training_Spiel.xlsx“ wurde im Team „TEAM Jugend“ nicht gefunden.")); return; }
    var akt = "", zeilen = [];
    D.plan.werte.forEach(function(r, i){ var a = txt(r[0]); if(a) akt = msArt(a) || (i < 3 || /montag|dienstag/i.test(r.join(" ")) ? "kopf" : "sonst");
      else if(!r.some(function(v){ return !leer(v); })) return;
      if(akt === "kopf" || akt === art) zeilen.push(r); });
    add(k2, tabelleAus(zeilen));
  }).catch(function(e){
    s.textContent = "";
    add(s, el("h1", null, M.sym + " " + M.titel), el("div", "hinweis", e && (e.status === 403 || e.status === 404) ? "Kein Zugriff auf das Team „TEAM Jugend“ – angezeigt werden die Angaben von der Homepage ohne Namen. Für Trainer und Kontakte bitte bei der Verwaltung melden." : "Die Liste aus Teams konnte nicht geladen werden (" + fehlerText(e) + "). Angezeigt werden die Angaben von der Homepage."));
    mannschaftenHomepage(s, M); });
}
function mannschaftenHomepage(s, M){
  add(s, el("p", "lead", M.lead + " Stand laut Homepage (gelesen am 24.11.2026)."));
  M.gruppen.forEach(function(g){
    var k = el("div", "karte"); add(s, k);
    var hl = el("a", null, "Seite auf der Homepage"); hl.href = g.link; hl.target = "_blank"; hl.rel = "noopener";
    add(k, el("h2", null, g.name), add(el("p", "lead"), hl));
    var t = el("table", "tabelle"), tb = el("tbody");
    add(t, add(el("thead"), add(el("tr"), el("th", null, "Mannschaft"), el("th", null, "Klasse/Alter"), el("th", null, "Training"))), tb);
    g.teams.forEach(function(m){ var tr = (m.training || g.training || []).concat(m.training ? [] : (g.ort ? [g.ort] : []));
      var td = el("td"); tr.forEach(function(z, i){ if(i) add(td, el("br")); add(td, z); });
      add(tb, add(el("tr"), add(el("td"), el("b", null, m.name)), el("td", null, m.info || "–"), td)); });
    add(k, add(el("div", "scroll"), t));
  });
  var l = el("a", null, "1928schneeberg.de"); l.href = "https://1928schneeberg.de/"; l.target = "_blank"; l.rel = "noopener";
  add(s, add(el("p", "grau"), "Ergebnisse und Tabellen: ", l));
}

/* ---------- Bewässerungssteuerung (aus der iPhone-App übernommen) ----------
 * Blätter der Liegenschaften-Excel: „Kameras“, „Bewässerung“, „Bewässerung Regner“, „Bewässerung Programme“.
 * Geschaltet wird per HTTP an den Raspberry Pi: POST <Basis-URL><Ein-Pfad>?dauer=<Sekunden>, POST <Basis-URL><Aus-Pfad>,
 * GET <Basis-URL>/druck. Es läuft immer nur ein Regner; nach dem Ausschalten 10 s Sperre. Jedes Einschalten
 * schickt eine Maximaldauer mit. Programme laufen im Browser nur, solange die Seite offen ist.
 * Solange keine Basis-URL eingetragen ist, ist die Steuerung sichtbar, aber ausgegraut. */
var ANLAGEN = {alt: "Alter Sportplatz", neu: "Neuer Sportplatz"};
var BW_SPERRE = 10;
var BW_PROGRAMME_VORGABE = [ /* laut Anleitung, solange das Blatt „Bewässerung Programme“ leer ist */
  {kennung: "A", name: "Langes Programm", dauern: [90, 90, 60, 60, 60, 60, 60, 75, 75]},
  {kennung: "B", name: "Testlauf", dauern: [1, 1, 1, 1, 1, 1, 1, 1, 1]},
  {kennung: "C", name: "Tau-Programm", dauern: [5, 5, 5, 5, 5, 5, 5, 5, 5]},
  {kennung: "D", name: "30 Minuten je Regner", dauern: [30, 30, 30, 30, 30, 30, 30, 30, 30]}];
var BW = {lauf: null, sperreBis: 0, programm: null, meldung: "", fehler: "", druck: null, takt: null};
function bwBlatt(n){ return Q().blatt(n).catch(function(){ return []; }); }
function bwLaden(){
  return Promise.all(["Kameras", "Bewässerung", "Bewässerung Regner", "Bewässerung Programme"].map(bwBlatt)).then(function(b){
    var kameras = [], anlagen = [], regner = [], prog = {};
    for(var z = ERSTE; z <= b[0].length; z++){ var n = txt(zelle(b[0], z, 1)); if(n) kameras.push({name: n, anlage: txt(zelle(b[0], z, 2)), url: txt(zelle(b[0], z, 5)), typ: txt(zelle(b[0], z, 6)).toUpperCase() || "HLS", bem: txt(zelle(b[0], z, 8))}); }
    for(z = ERSTE; z <= b[1].length; z++){ var a = txt(zelle(b[1], z, 1)); if(a) anlagen.push({anlage: a, basis: txt(zelle(b[1], z, 2)), kamera: txt(zelle(b[1], z, 5)), modi: /^ja/i.test(txt(zelle(b[1], z, 6))), bem: txt(zelle(b[1], z, 7)), soll: zahl(zelle(b[1], z, 8))}); }
    for(z = ERSTE; z <= b[2].length; z++){ var r = txt(zelle(b[2], z, 2)); if(r) regner.push({anlage: txt(zelle(b[2], z, 1)), name: r, ein: txt(zelle(b[2], z, 3)), aus: txt(zelle(b[2], z, 4)), folge: zahl(zelle(b[2], z, 5)), dauer: zahl(zelle(b[2], z, 6)), soll: zahl(zelle(b[2], z, 8))}); }
    for(z = ERSTE; z <= b[3].length; z++){ var k = txt(zelle(b[3], z, 2)), an = txt(zelle(b[3], z, 1)); if(!k) continue; var id = an + "|" + k;
      (prog[id] = prog[id] || {anlage: an, kennung: k, name: txt(zelle(b[3], z, 3)), schritte: []}).schritte.push({regner: txt(zelle(b[3], z, 4)), dauer: zahl(zelle(b[3], z, 5)) || 0}); }
    /* Vorgaben wie in der App, falls Blätter fehlen */
    if(!kameras.length) kameras = [{name: "Kamera Alter Sportplatz Nord", anlage: ANLAGEN.alt, url: "", typ: "HLS"}, {name: "Kamera Alter Sportplatz Süd", anlage: ANLAGEN.alt, url: "", typ: "HLS"},
      {name: "Kamera Neuer Sportplatz Nord", anlage: ANLAGEN.neu, url: "", typ: "HLS"}, {name: "Kamera Neuer Sportplatz Süd", anlage: ANLAGEN.neu, url: "", typ: "HLS"}];
    if(!anlagen.length) anlagen = [{anlage: ANLAGEN.alt, basis: "", kamera: "Kamera Alter Sportplatz Süd", modi: false, bem: "Ein Ventil, nur ein und aus", soll: 7.5},
      {anlage: ANLAGEN.neu, basis: "", kamera: "Kamera Neuer Sportplatz Nord", modi: true, bem: "Mehrere Regner, Reihenfolge und Dauer", soll: 8}];
    if(!regner.length){ regner.push({anlage: ANLAGEN.alt, name: "Beregnung gesamt", ein: "/ventil/1/ein", aus: "/ventil/1/aus", folge: 1, dauer: 20, soll: 7.5});
      for(var i = 1; i <= 9; i++) regner.push({anlage: ANLAGEN.neu, name: "Regner " + i, ein: "/ventil/" + i + "/ein", aus: "/ventil/" + i + "/aus", folge: i, dauer: 15, soll: i <= 3 ? 9 : 8}); }
    var programme = Object.keys(prog).map(function(k){ return prog[k]; }), vorgabe = false;
    if(!programme.some(function(p){ return p.anlage === ANLAGEN.neu; })){ vorgabe = true;
      BW_PROGRAMME_VORGABE.forEach(function(p){ programme.push({anlage: ANLAGEN.neu, kennung: p.kennung, name: p.name, schritte: p.dauern.map(function(d, i){ return {regner: "Regner " + (i + 1), dauer: d}; })}); }); }
    regner.sort(function(a, b){ return (a.folge || 99) - (b.folge || 99); });
    if(TEST && /[?&]pi\b/.test(location.search)) anlagen.forEach(function(a){ a.basis = "https://pi.test"; }); /* Testmodus: Attrappe der Steuerung */
    return {kameras: kameras, anlagen: anlagen, regner: regner, programme: programme, vorgabe: vorgabe};
  });
}
function bwUrl(an, pfadR, dauerMin){ if(!an.basis || !pfadR) return null; var u = an.basis.replace(/\/$/, "") + (pfadR.charAt(0) === "/" ? "" : "/") + pfadR;
  if(dauerMin > 0) u += (u.indexOf("?") < 0 ? "?" : "&") + "dauer=" + Math.round(dauerMin * 60); return u; }
function bwSenden(u, methode, rumpf){
  if(TEST && u.indexOf("https://pi.test") === 0){ (window.__pi = window.__pi || []).push((methode || "POST") + " " + u.slice(15)); return Promise.resolve({ok: true}); }
  return fetch(u, {method: methode || "POST", headers: rumpf ? {"Content-Type": "application/json"} : undefined, body: rumpf ? JSON.stringify(rumpf) : undefined})
    .then(function(r){ if(!r.ok) throw new Error("Die Steuerung hat mit Fehler " + r.status + " geantwortet."); return r; })
    .catch(function(e){ throw new Error(e && /Fehler \d/.test(e.message) ? e.message : "Die Steuerung ist nicht erreichbar."); });
}
function bwEin(an, r, dauer){
  if(BW.lauf) return Promise.reject(new Error("Es kann nur ein Regner laufen. Erst „" + BW.lauf.name + "“ ausschalten."));
  var rest = Math.ceil((BW.sperreBis - Date.now()) / 1000); if(rest > 0) return Promise.reject(new Error("Bitte noch " + rest + " s warten, dann darf der nächste Regner starten."));
  var u = bwUrl(an, r.ein, dauer); if(!u) return Promise.reject(new Error("Für „" + r.name + "“ ist kein Ein-Pfad hinterlegt."));
  return bwSenden(u).then(function(){ BW.lauf = {name: r.name, bis: Date.now() + dauer * 60000}; BW.meldung = r.name + " eingeschaltet (" + dauer + " min)."; });
}
function bwAus(an, r){
  var u = bwUrl(an, r.aus); if(!u) return Promise.reject(new Error("Für „" + r.name + "“ ist kein Aus-Pfad hinterlegt."));
  return bwSenden(u).then(function(){ if(BW.lauf && BW.lauf.name === r.name){ BW.lauf = null; BW.sperreBis = Date.now() + BW_SPERRE * 1000; } BW.meldung = r.name + " ausgeschaltet."; });
}
function bwAllesAus(an, liste){ if(BW.programm) BW.programm.stop = true; BW.programm = null;
  return liste.reduce(function(p, r){ return p.then(function(){ return r.aus ? bwAus(an, r).catch(function(){}) : null; }); }, Promise.resolve())
    .then(function(){ BW.lauf = null; BW.meldung = "Alles ausgeschaltet."; }); }
function bwWarte(ms, P){ return new Promise(function(ok){ var t0 = Date.now(); (function w(){ if(P.stop || Date.now() - t0 >= ms) ok(); else setTimeout(w, 500); })(); }); }
function bwProgramm(an, prog, liste, neuZeichnen){
  if(BW.programm || BW.lauf){ BW.fehler = "Es läuft schon etwas – erst ausschalten."; neuZeichnen(); return; }
  var P = BW.programm = {kennung: prog.kennung, name: prog.name, schritt: 0, stop: false};
  prog.schritte.reduce(function(kette, st, i){ return kette.then(function(){ if(P.stop) return;
    var r = liste.filter(function(x){ return x.name === st.regner; })[0]; if(!r || !st.dauer) return;
    P.schritt = i; neuZeichnen();
    return bwEin(an, r, st.dauer).then(function(){ neuZeichnen(); return bwWarte(st.dauer * 60000, P); })
      .then(function(){ if(!P.stop) return bwAus(an, r); }).then(function(){ neuZeichnen(); return bwWarte(BW_SPERRE * 1000, P); }); }); }, Promise.resolve())
    .catch(function(e){ BW.fehler = e.message; })
    .then(function(){ if(BW.programm === P) BW.programm = null; if(!P.stop) BW.meldung = "Programm " + prog.kennung + " beendet."; neuZeichnen(); });
}
function seiteBewaesserung(s, platz){
  if(!platz){
    pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Bewässerungssteuerung"]]);
    s.textContent = "";
    add(s, el("h1", null, "💧 Bewässerungssteuerung"), el("p", "lead", "Sportplatz wählen: links das Kamerabild, rechts die Bewässerung."));
    add(s, add(el("div", "kacheln"), kachel("🌱", "Alter Sportplatz", "Ein Ventil: Bewässerung ein und aus", "#/verwaltung/bewaesserung/alt", true),
      kachel("🌿", "Neuer Sportplatz", "Neun Regner, einzeln oder als Programm nacheinander (z. B. Tau-Programm 5 min, langes Programm bis 90 min)", "#/verwaltung/bewaesserung/neu", true)));
    return;
  }
  var name = ANLAGEN[platz]; if(!name){ fehlerZeigen("Diesen Sportplatz gibt es nicht."); return; }
  pfad([["Start", "#/"], ["Liegenschaften", "#/verwaltung"], ["Bewässerungssteuerung", "#/verwaltung/bewaesserung"], [name]]);
  ladeHinweis(s);
  bwLaden().then(function(D){
    var an = D.anlagen.filter(function(a){ return a.anlage === name; })[0] || {anlage: name, basis: ""};
    var liste = D.regner.filter(function(r){ return r.anlage === name; }), progs = D.programme.filter(function(p){ return p.anlage === name; });
    var kams = D.kameras.filter(function(k){ return k.anlage === name; });
    var kam = kams.filter(function(k){ return k.name === an.kamera; })[0] || kams[0];
    var darfSchalten = darf("verwaltung"), verbunden = !!an.basis, httpsProblem = /^http:/i.test(an.basis) && location.protocol === "https:";
    s.textContent = "";
    add(s, el("h1", null, "💧 " + name), el("p", "lead", an.bem || ""));
    var raster = el("div", "bw"), links = el("div", "karte bw-kamera"), rechts = el("div", "karte bw-steuerung");
    add(s, add(raster, links, rechts));
    /* Kamera */
    function kameraZeigen(){ links.textContent = ""; add(links, el("h2", null, "📹 Kamera"));
      if(kams.length > 1){ var w = el("select"); kams.forEach(function(k){ var o = el("option", null, k.name); o.value = k.name; add(w, o); }); w.value = kam ? kam.name : "";
        w.addEventListener("change", function(){ kam = kams.filter(function(k){ return k.name === w.value; })[0]; kameraZeigen(); }); add(links, w); }
      var bild = el("div", "bw-bild"); add(links, bild);
      if(!kam || !kam.url){ add(bild, el("span", null, "Noch keine Kamera eingetragen. Sobald sie hängt: Adresse in der Liegenschaften-Excel, Blatt „Kameras“, Spalte F (Stream-URL), Spalte G Typ HLS, MJPEG oder Web.")); return; }
      if(/^http:/i.test(kam.url) && location.protocol === "https:"){ add(bild, el("span", null, "Die Kamera-Adresse beginnt mit http://. Das Portal läuft über https und der Browser blockiert sie – die Kamera braucht eine https-Adresse.")); return; }
      if(kam.typ === "MJPEG"){ var img = el("img"); img.alt = kam.name; img.src = kam.url; add(bild, img); }
      else if(kam.typ === "WEB"){ var fr = el("iframe"); fr.src = kam.url; fr.title = kam.name; add(bild, fr); }
      else { var v = el("video"); v.controls = true; v.muted = true; v.autoplay = true; v.playsInline = true;
        if(v.canPlayType("application/vnd.apple.mpegurl")) { v.src = kam.url; add(bild, v); }
        else add(bild, el("span", null, "Dieser Browser spielt den HLS-Stream nicht direkt ab (Safari auf iPhone und Mac kann es). Über „Groß öffnen“ geht es in einem eigenen Fenster.")); }
      var offen = el("a", "btn klein", "Groß öffnen"); offen.href = kam.url; offen.target = "_blank"; offen.rel = "noopener"; add(links, add(el("div", "reihe"), offen)); }
    kameraZeigen();
    /* Steuerung */
    var statusBox = null;
    function statusFuellen(){ var status = statusBox; if(!status || !status.isConnected) return; status.textContent = "";
      var rest = BW.lauf ? Math.max(0, Math.round((BW.lauf.bis - Date.now()) / 1000)) : 0, sperre = Math.ceil((BW.sperreBis - Date.now()) / 1000);
      add(status, add(el("span"), "Läuft: ", el("b", null, BW.lauf ? BW.lauf.name + " · noch " + Math.floor(rest / 60) + ":" + pad(rest % 60) : "nichts")));
      if(BW.programm) add(status, add(el("span"), "Programm: ", el("b", null, BW.programm.kennung + " „" + BW.programm.name + "“, Schritt " + (BW.programm.schritt + 1))));
      if(sperre > 0) add(status, add(el("span"), "Sperre: ", el("b", null, sperre + " s")));
      add(status, add(el("span"), "Druck: ", el("b", null, BW.druck != null ? String(BW.druck).replace(".", ",") + " bar" : "–"), an.soll ? " (Soll " + String(an.soll).replace(".", ",") + " bar)" : ""));
    }
    function zeichnen(){
      if(!rechts.isConnected) return;
      rechts.textContent = ""; add(rechts, el("h2", null, "🚿 Bewässerung"));
      if(!verbunden) add(rechts, el("div", "hinweis", "Noch keine Steuerung verbunden. Sobald der Raspberry Pi steht: Basis-URL in der Excel, Blatt „Bewässerung“, Spalte C. Bis dahin sind die Knöpfe ausgegraut."));
      if(httpsProblem) add(rechts, el("div", "fehler", "Die Basis-URL beginnt mit http://. Von der https-Seite des Portals blockiert der Browser das – die Steuerung braucht eine https-Adresse (z. B. über Tailscale)."));
      if(!darfSchalten) add(rechts, el("div", "hinweis", "Du kannst hier nur ansehen. Schalten kann nur die Verwaltung."));
      var aus = !verbunden || httpsProblem || !darfSchalten;
      var status = el("div", "summe"); statusBox = status; add(rechts, status); statusFuellen();
      if(BW.fehler) add(rechts, el("div", "fehler", BW.fehler));
      if(BW.meldung) add(rechts, el("p", "grau", BW.meldung));
      function tu(p){ BW.fehler = ""; p.catch(function(e){ BW.fehler = e.message; }).then(zeichnen); }
      var t = el("table", "tabelle"), tb = el("tbody");
      add(t, add(el("thead"), add(el("tr"), el("th", null, "Regner"), el("th", null, "Dauer"), el("th", null, ""))), tb);
      liste.forEach(function(r){
        var d = el("select"); [1, 5, 10, 15, 20, 30, 45, 60, 75, 90].concat(r.dauer && [1, 5, 10, 15, 20, 30, 45, 60, 75, 90].indexOf(r.dauer) < 0 ? [r.dauer] : []).sort(function(a, b){ return a - b; })
          .forEach(function(m){ var o = el("option", null, m + " min"); o.value = m; add(d, o); }); d.value = r.dauer || 15; d.disabled = aus;
        var laeuft = BW.lauf && BW.lauf.name === r.name;
        var bEin = btn("primary klein", "Ein", function(){ tu(bwEin(an, r, +d.value)); }), bAus = btn("klein", "Aus", function(){ tu(bwAus(an, r)); });
        bEin.disabled = aus || !!BW.lauf || !!BW.programm; bAus.disabled = aus;
        add(tb, add(el("tr", laeuft ? "laeuft" : null), add(el("td"), el("b", null, r.name), laeuft ? el("span", "chip ok", " läuft ") : null, r.soll ? el("div", "grau", "Solldruck " + String(r.soll).replace(".", ",") + " bar") : null), add(el("td"), d), add(el("td", "z"), bEin, " ", bAus))); });
      add(rechts, add(el("div", "scroll"), t));
      if(an.modi || progs.length){
        add(rechts, el("h2", null, "Programme"), el("p", "grau", "Ein Regner nach dem anderen, dazwischen " + BW_SPERRE + " s Pause. Das Programm läuft nur, solange diese Seite offen ist." + (D.vorgabe ? " Dauern laut Anleitung (Blatt „Bewässerung Programme“ ist noch leer)." : "")));
        var pt = el("table", "tabelle"), pb = el("tbody"); add(pt, add(el("thead"), add(el("tr"), el("th", null, "Programm"), el("th", null, "Ablauf"), el("th", null, ""))), pb);
        progs.forEach(function(p){ var ges = p.schritte.reduce(function(a, x){ return a + x.dauer; }, 0), fehlt = p.schritte.filter(function(x){ return !liste.some(function(r){ return r.name === x.regner; }); }).length;
          var ablauf = p.schritte.map(function(x){ return x.regner.replace("Regner ", "R") + " " + x.dauer; }).join(" · ") + " min";
          var b = btn("klein", "Starten", function(){ BW.fehler = ""; bwProgramm(an, p, liste, zeichnen); zeichnen(); }); b.disabled = aus || !!BW.programm || !!BW.lauf;
          add(pb, add(el("tr", BW.programm && BW.programm.kennung === p.kennung ? "laeuft" : null), add(el("td"), el("b", null, p.kennung + " · " + p.name), el("div", "grau", "gesamt " + Math.floor(ges / 60) + " h " + (ges % 60) + " min")),
            add(el("td"), ablauf, fehlt ? el("div", "grau", fehlt + " Regner nicht in der Excel – werden übersprungen") : null), add(el("td", "z"), b))); });
        add(rechts, add(el("div", "scroll"), pt));
      }
      var notaus = btn("gefahr", "Alles ausschalten", function(){ tu(bwAllesAus(an, liste)); }); notaus.disabled = !verbunden || httpsProblem || !darfSchalten;
      add(rechts, add(el("div", "reihe"), notaus));
    }
    zeichnen();
    clearInterval(BW.takt); var zaehler = 0, vorher = !!BW.lauf;
    BW.takt = setInterval(function(){ if(!rechts.isConnected){ clearInterval(BW.takt); return; }
      if(BW.lauf && BW.lauf.bis <= Date.now()){ BW.lauf = null; BW.sperreBis = Date.now() + BW_SPERRE * 1000; }
      if(verbunden && !httpsProblem && zaehler++ % 10 === 0) fetch(bwUrl(an, "/druck")).then(function(r){ return r.ok ? r.text() : null; }).then(function(t){ var v = null; try{ var j = JSON.parse(t); v = typeof j === "number" ? j : j.bar; }catch(_){ v = parseFloat(t); } BW.druck = isNaN(v) ? null : v; }).catch(function(){ BW.druck = null; });
      if(vorher && !BW.lauf) zeichnen(); else statusFuellen(); vorher = !!BW.lauf; }, 1000);
  }).catch(function(e){ fehlerZeigen(fehlerText(e)); });
}

/* ---------- Ablauf ---------- */
function route(){
  /* Im Anmelde-Popup trägt die Adresse den Antwortcode von Microsoft: dort nichts anzeigen. */
  if(/[#&](code|error|state)=/.test(location.hash)) return;
  var h = location.hash.replace(/^#\/?/, ""), teile = h.split("/").filter(Boolean), s = seite();
  s.textContent = "";
  if(!angemeldet()){ seiteWillkommen(s); return; }
  if(teile[0] === "teamwear"){ seiteTeamwear(s); }
  else if(teile[0] === "protokolle"){ seiteProtokolle(s, teile[1]); }
  else if(teile[0] === "kalender"){ seiteKalender(s, teile[1]); }
  else if(teile[0] === "veranstaltungen"){ seiteVeranstaltungen(s, teile[1], teile[2], teile[3]); }
  else if(teile[0] === "jugend" || teile[0] === "aktive"){ seiteMannschaften(s, teile[0]); }
  else if(teile[0] !== "verwaltung"){ seiteStart(s); }
  else if(teile[1] === "reinigung") seiteReinigung(s);
  else if(teile[1] === "ausleihe") seiteAusleihe(s, teile[2]);
  else if(teile[1] === "schluessel") seiteSchluessel(s);
  else if(teile[1] === "inventar") seiteInventar(s);
  else if(teile[1] === "aufgaben") seiteAufgaben(s);
  else if(teile[1] === "bewaesserung") seiteBewaesserung(s, teile[2]);
  else seiteVerwaltung(s);
  window.scrollTo(0, 0);
}
$("abmelden").addEventListener("click", abmelden);
/* Neue Fassung erkennen: GitHub Pages lässt die Startseite bis zu 10 Minuten im Browser-Zwischenspeicher.
   Weicht die Version auf dem Server ab, lädt die Seite einmal frisch nach. */
(function(){ var eigene = ((document.querySelector('script[src^="portal.js"]') || {}).getAttribute || function(){ return ""; }).call(document.querySelector('script[src^="portal.js"]'), "src") || "";
  if(TEST || location.protocol === "file:" || !/v=\d+/.test(eigene) || /[#&](code|error|state)=/.test(location.hash)) return;
  fetch(location.pathname, {cache: "reload"}).then(function(r){ return r.ok ? r.text() : ""; }).then(function(t){
    var m = /portal\.js\?v=(\d+)/.exec(t || ""), mein = /v=(\d+)/.exec(eigene)[1];
    if(m && m[1] !== mein){ var k = "spfrNeuGeladen" + m[1]; try{ if(sessionStorage.getItem(k)) return; sessionStorage.setItem(k, "1"); }catch(_){ return; }
      fetch("portal.css?v=" + m[1], {cache: "reload"}).catch(function(){}).then(function(){ location.reload(); }); } }).catch(function(){}); })();
window.addEventListener("hashchange", route);
if(!TEST && msalBereit()){
  pca.handleRedirectPromise().then(function(r){
      var acc = r && r.account ? r.account : (pca.getAllAccounts().filter(vereinskonto)[0] || pca.getAllAccounts()[0] || null);
      if(r && r.state && /^#\//.test(r.state) && location.hash !== r.state) history.replaceState(null, "", location.pathname + location.search + r.state);
      return kontoSetzen(acc); })
    .catch(function(e){ fehlerAnmeldung = "Die Anmeldung hat nicht geklappt." + (e && e.errorMessage ? " (" + e.errorMessage.split("\n")[0] + ")" : ""); return kontoSetzen(null); })
    .then(function(){ kopfzeile(); route(); });
} else { kopfzeile(); route(); }
window.__portal = {reinigungRechnen: reinigungRechnen, rabattVerteilen: rabattVerteilen};
})();
