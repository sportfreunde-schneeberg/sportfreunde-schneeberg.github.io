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
var SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All", "Calendars.ReadWrite", "Calendars.ReadWrite.Shared"];
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
    cache: { cacheLocation: "sessionStorage" }
  });
  return true;
}
function token(){
  return pca.acquireTokenSilent({scopes: SCOPES, account: konto})
    .catch(function(){ return pca.acquireTokenPopup({scopes: SCOPES, account: konto}); })
    .then(function(r){ return r.accessToken; });
}
function anmelden(){
  if(!msalBereit()){ fehlerZeigen(location.protocol === "file:" ? "Die Anmeldung funktioniert nur über die Web-Adresse, nicht als lokale Datei." : "Die Microsoft-Anmeldung konnte nicht geladen werden."); return; }
  pca.loginPopup({scopes: SCOPES, prompt: "select_account"}).then(function(r){ kontoSetzen(r.account).then(function(){ kopfzeile(); route(); }); })
    .catch(function(e){ if(e && e.errorCode === "user_cancelled") return; fehlerZeigen("Die Anmeldung hat nicht geklappt." + (e && e.errorMessage ? " (" + e.errorMessage.split("\n")[0] + ")" : "")); });
}
function abmelden(){ var k = konto; konto = null; kontoHash = ""; ids = null; sess = null; DATEN = null; kalInfo = null; prDrive = null; if(pca && k) pca.logoutPopup({account: k}).catch(function(){}); kopfzeile(); location.hash = "#/"; route(); }
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
      return ids = {wb: "/drives/" + d.id + "/items/" + beste.id + "/workbook"}; }); });
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
var fremdKonto = false;
/* Nur Vereinskonten (…@1928schneeberg.de). Gäste aus Teams (z. B. #EXT#) werden abgewiesen und abgemeldet. */
function vereinskonto(acc){ return /@1928schneeberg\.de$/i.test(String(acc && acc.username || "")); }
function kontoSetzen(acc){
  if(acc && !vereinskonto(acc)){ fremdKonto = true; if(pca) pca.logoutPopup({account: acc}).catch(function(){}); acc = null; }
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
function pfad(teile){ var p = $("pfad"); p.textContent = ""; teile.forEach(function(t, i){ if(i) add(p, el("span", null, "›")); if(t[1]){ var a = el("a", null, t[0]); a.href = t[1]; add(p, a); } else add(p, el("span", null, t[0])); }); }
function kachel(sym, titel, text, href, gross, extern){ var a = el("a", "kachel" + (gross ? " gross" : "")); a.href = href; if(extern) a.rel = "noopener"; add(a, el("span", "sym", sym), el("b", null, titel), el("span", null, text)); return a; }
function ladeHinweis(s){ s.textContent = ""; add(s, el("div", "laden", "Daten werden aus der Excel geladen …")); }

function seiteWillkommen(s){
  pfad([]);
  var w = el("section", "willkommen");
  var img = el("img"); img.src = "img/wappen.png"; img.alt = ""; img.width = 110; img.height = 110;
  add(w, img, el("h1", null, "Vereinsausschuss-Portal"),
    el("p", "lead", "Interner Bereich des Vereinsausschusses. Melde dich mit deinem Vereinskonto an (…@1928schneeberg.de)."),
    fremdKonto ? el("div", "fehler", "Dieses Konto ist kein Vereinskonto. Bitte mit einer Adresse …@1928schneeberg.de anmelden.") : null,
    btn("primary", "Mit Vereinskonto anmelden", anmelden));
  add(s, w);
}
function seiteStart(s){
  pfad([["Start"]]);
  add(s, el("h1", null, "Hallo" + (konto && konto.name ? " " + konto.name.split(" ")[0] : "") + "!"), el("p", "lead", "Wähle einen Bereich."));
  var k = el("div", "kacheln");
  add(k, kachel("🏠", "Verwaltung", "Kabinenreinigung, Ausleihliste, Schlüssel und Inventar", "#/verwaltung", true),
    kachel("🎉", "Festausschuss", "Festplaner 100 Jahre Sportfreunde (2.–5. Juni 2028)", "/festplaner/", true),
    kachel("👕", "Teamwear-Status", "Bestellungen und ihr Stand: bestellt, geliefert, ausgehändigt, bezahlt", "#/teamwear", true),
    kachel("📄", "Protokolle", "Ausschusssitzungen, Klausurtagungen und Generalversammlung nach Jahr und Monat", "#/protokolle", true),
    kachel("📅", "Vereinskalender", "Termine wie Ausschusssitzungen ansehen, eintragen, ändern und löschen", "#/kalender", true));
  add(s, k);
}
function seiteVerwaltung(s){
  pfad([["Start", "#/"], ["Verwaltung"]]);
  add(s, el("h1", null, "Verwaltung"), el("p", "lead", "Die Daten stehen in der Liegenschaften-Excel im Team „TEAM - Liegenschaften“. Änderungen hier landen direkt dort."));
  var k = el("div", "kacheln");
  add(k, kachel("🧽", "Kabinenreinigung", "Monat eintragen, Stunden und Beträge", "#/verwaltung/reinigung"),
    kachel("📦", "Ausleihliste", "Buchungen anlegen, Preise, Rabatt und Zahlung", "#/verwaltung/ausleihe"),
    kachel("🔑", "Schlüssel", "Wer hat welchen Schlüssel, Schlösser", "#/verwaltung/schluessel"),
    kachel("🗂️", "Inventar", "Gegenstände, Lagerorte und Leihgebühren", "#/verwaltung/inventar"));
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Kabinenreinigung"]]);
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Ausleihliste"]]);
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Ausleihliste", "#/verwaltung/ausleihe"], ["Neue Buchung"]]);
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Ausleihliste", "#/verwaltung/ausleihe"], ["Nr. " + b.nr]]);
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Schlüssel"]]);
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
  pfad([["Start", "#/"], ["Verwaltung", "#/verwaltung"], ["Inventar"]]);
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
var twIds = null, twSess = null;
function twFinden(){
  if(twIds) return Promise.resolve(twIds);
  return gf("/groups?$filter=" + encodeURIComponent("displayName eq '" + TW.gruppe + "'") + "&$select=id").then(function(r){
    var g = (r.value || [])[0]; if(!g){ var e = new Error("Team"); e.status = 404; throw e; }
    return gf("/groups/" + g.id + "/drive?$select=id"); }).then(function(d){
    return gf("/drives/" + d.id + "/root:/" + pfadK(TW.datei) + "?$select=id").then(function(it){ return twIds = {wb: "/drives/" + d.id + "/items/" + it.id + "/workbook"}; }); });
}
var TwLive = {
  lesen: function(){ return twFinden().then(function(I){ return gf(I.wb + "/tables('Bestellungen')/range?$select=address,values"); }).then(function(b){ return ausrichten(b.address, b.values || []); }); },
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
  schreiben: function(adresse, werte){ return Test.schreibenIn(TwTest.db, "Teamwear", adresse, werte); }
};
function TWQ(){ return TEST ? TwTest : TwLive; }
function darfTeamwearSchreiben(){ return darf("teamwear"); }
function twZeilen(w){ var out = []; for(var z = ERSTE; z <= w.length; z++){ var art = txt(zelle(w, z, 7)); if(!art) continue;
  out.push({zeile: z, eingang: zelle(w, z, 0), nr: txt(zelle(w, z, 1)), name: [txt(zelle(w, z, 3)), txt(zelle(w, z, 2))].filter(Boolean).join(" "), mannschaft: txt(zelle(w, z, 6)),
    artikel: art, groesse: txt(zelle(w, z, 8)), anzahl: zahl(zelle(w, z, 9)), initialen: txt(zelle(w, z, 10)), summe: zahl(zelle(w, z, 13)),
    stufen: STUFEN.map(function(s){ return /^ja/i.test(txt(zelle(w, z, s[1]))); }), status: txt(zelle(w, z, 19)), bem: txt(zelle(w, z, 20))}); }
  return out; }
/* Status wie die Formel in Spalte T, damit die Anzeige nach dem Klicken sofort stimmt. */
function twStatus(st){ if(st[4]) return "storniert"; if(st[0] && st[1] && st[2] && st[3]) return "abgeschlossen";
  if(!st[0]) return "zu bestellen"; if(!st[1]) return "bestellt"; if(!st[2]) return "geliefert"; if(!st[3]) return "Zahlung offen"; return "abgeschlossen"; }
var TW_FILTER = "offen";
function seiteTeamwear(s){
  pfad([["Start", "#/"], ["Teamwear-Status"]]);
  ladeHinweis(s);
  var darf = darfTeamwearSchreiben();
  TWQ().lesen().then(function(w){
    s.textContent = "";
    var zeilen = twZeilen(w);
    add(s, el("h1", null, "Teamwear-Status"),
      el("p", "lead", darf ? "Du kannst die Stufen hier abhaken. Jeder Klick wird sofort in der Excel gespeichert." : "Stand aller Bestellungen. Abhaken kann nur der Teamwear-Verantwortliche."));
    var zaehl = {}; zeilen.forEach(function(r){ zaehl[r.status] = (zaehl[r.status] || 0) + 1; });
    var c = el("div", "summe"); ["zu bestellen", "bestellt", "geliefert", "Zahlung offen", "abgeschlossen", "storniert"].forEach(function(k){ add(c, add(el("span"), k + " ", el("b", null, String(zaehl[k] || 0)))); }); add(s, c);
    var filter = el("select"); filter.style.maxWidth = "220px";
    [["offen", "Alle offenen"], ["alle", "Alle"], ["zu bestellen", "zu bestellen"], ["bestellt", "bestellt"], ["geliefert", "geliefert"], ["Zahlung offen", "Zahlung offen"], ["abgeschlossen", "abgeschlossen"], ["storniert", "storniert"]]
      .forEach(function(o){ var op = el("option", null, o[1]); op.value = o[0]; add(filter, op); });
    filter.value = TW_FILTER;
    var such = el("input", "suche"); such.type = "search"; such.placeholder = "Name, Artikel oder Mannschaft …";
    add(s, add(el("div", "reihe"), add(el("label", "fld", "Anzeigen"), filter), such));
    var box = el("div"); add(s, box);
    function zeigen(){ box.textContent = ""; TW_FILTER = filter.value; var q = such.value.trim().toLowerCase(), n = 0;
      zeilen.slice().reverse().forEach(function(r){
        if(TW_FILTER === "offen" && (r.status === "abgeschlossen" || r.status === "storniert")) return;
        if(TW_FILTER !== "offen" && TW_FILTER !== "alle" && r.status !== TW_FILTER) return;
        if(q && (r.name + " " + r.artikel + " " + r.mannschaft).toLowerCase().indexOf(q) < 0) return;
        n++; add(box, twKarte(r)); });
      if(!n) add(box, el("p", "lead", "Keine Bestellungen in dieser Ansicht.")); }
    function twKarte(r){
      var k = el("div", "buchung"), h = el("div", "kopf"), chip = el("span", "chip " + (r.status === "abgeschlossen" ? "ok" : (r.status === "storniert" ? "grau" : "offen")), r.status || "–");
      add(h, el("b", null, (r.anzahl > 1 ? r.anzahl + " × " : "") + r.artikel + (r.groesse ? " · " + r.groesse : "") + (r.initialen ? " · " + r.initialen : "")), el("span", "sp"), chip);
      var info = el("div", "meta"); add(info, el("span", null, r.name), el("span", null, r.mannschaft), el("span", null, datumText(r.eingang)), add(el("span"), "Summe ", el("b", null, euro(r.summe))));
      var st = el("div", "stufen");
      STUFEN.forEach(function(s2, i){
        var b = el("button", "stufe" + (i === 4 ? " storno" : "") + (r.stufen[i] ? " an" : "")); b.type = "button"; b.disabled = !darf;
        add(b, el("span", "kreis", i === 4 ? "✕" : "✓"), el("span", null, s2[0]));
        b.setAttribute("aria-pressed", r.stufen[i] ? "true" : "false"); b.setAttribute("aria-label", s2[0] + (r.stufen[i] ? " – erledigt" : " – offen"));
        if(darf) b.addEventListener("click", function(){ umschalten(r, i, k); });
        if(i < 3){ var sch = el("div", "schritt" + (r.stufen[i] && r.stufen[i + 1] ? " an" : "")); add(sch, b, el("span", "linie")); add(st, sch); } else add(st, b); });
      add(k, h, info, st); if(r.bem) add(k, el("p", "lead", r.bem)); return k; }
    function umschalten(r, i, karte){
      if(gesperrt(karte, "teamwear")) return;
      var neu = r.stufen.slice(); neu[i] = !neu[i];
      karte.querySelectorAll(".stufe").forEach(function(b){ b.disabled = true; });
      TWQ().lesen().then(function(w2){
        var jetzt = twZeilen(w2).filter(function(x){ return x.zeile === r.zeile; })[0];
        if(!jetzt || jetzt.nr !== r.nr || jetzt.artikel !== r.artikel) throw new Error("Die Zeile hat sich in der Excel verändert. Bitte die Seite neu laden.");
        var spalteBuchstabe = spalte(STUFEN[i][1]);
        return TWQ().schreiben(spalteBuchstabe + r.zeile + ":" + spalteBuchstabe + r.zeile, [[neu[i] ? "Ja" : ""]]);
      }).then(function(){ r.stufen = neu; r.status = twStatus(neu); zeigeToast(r.artikel + ": " + STUFEN[i][0] + (neu[i] ? " ✓" : " entfernt")); var n = twKarte(r); karte.replaceWith(n); })
        .catch(function(e){ karte.querySelectorAll(".stufe").forEach(function(b){ b.disabled = false; }); add(karte, el("div", "fehler", e && e.status ? twFehler(e) : (e.message || "Fehler beim Speichern."))); });
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
      var tit = el("td"); add(tit, el("b", null, ev.subject || "(ohne Titel)")); var ortT = (ev.location && ev.location.displayName) || ""; if(ortT) add(tit, el("div", "notiz ortm", "📍 " + ortT)); if(ev.bodyPreview) add(tit, el("div", "notiz", ev.bodyPreview));
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
  else if(teile[0] !== "verwaltung"){ seiteStart(s); }
  else if(teile[1] === "reinigung") seiteReinigung(s);
  else if(teile[1] === "ausleihe") seiteAusleihe(s, teile[2]);
  else if(teile[1] === "schluessel") seiteSchluessel(s);
  else if(teile[1] === "inventar") seiteInventar(s);
  else seiteVerwaltung(s);
  window.scrollTo(0, 0);
}
$("abmelden").addEventListener("click", abmelden);
window.addEventListener("hashchange", route);
if(!TEST && msalBereit()){
  pca.handleRedirectPromise().catch(function(){}).then(function(){ var acc = pca.getAllAccounts(); return kontoSetzen(acc.length ? acc[0] : null); }).then(function(){ kopfzeile(); route(); });
} else { kopfzeile(); route(); }
window.__portal = {reinigungRechnen: reinigungRechnen, rabattVerteilen: rabattVerteilen};
})();
