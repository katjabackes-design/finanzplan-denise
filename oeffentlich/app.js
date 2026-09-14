/*
 * Finanzplan Denise — Rechenwerk und Oberfläche.
 *
 * Kein Framework: die Seite hat einen Bildschirm und einen Datensatz. React,
 * ein Bündler und eine Klassenbibliothek wären hier mehr Werkzeug als Werk.
 *
 * Aufbau in drei Schritten:
 *   1. VORGABE — die Werte aus «Finanzplan Denise.xlsx», Zeile für Zeile.
 *   2. rechne() — die vier Rechenarten der Tabelle, in einer Formel vereinheitlicht.
 *   3. zeichne() baut das Gerüst genau einmal, aktualisiere() schreibt danach nur
 *      noch Ergebniszellen. Ein vollständiger Neuaufbau bei jedem Tastendruck
 *      würde den Fokus aus dem Feld werfen, in dem gerade getippt wird.
 */

'use strict';

/* ── 1 · Vorgabewerte aus der Tabelle ───────────────────────────────────── */

const SCHLUESSEL = 'finanzplan-denise-v1';

/* ── Konto (Supabase) ───────────────────────────────────────────────────────
 * anon-Key ist bewusst öffentlich im Frontend-Code — abgesichert wird über
 * Row-Level-Security in der Datenbank, nicht durch Geheimhaltung des Keys.
 */
const SUPABASE_URL = 'https://kdlhplgrvehnccbzwiaj.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkbGhwbGdydmVobmNjYnp3aWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODkyMzg0OTcsImV4cCI6MjEwNDgxNDQ5N30.glfCkZ_PV7xzEefuLYyW9N6WR9ZKRpVd0SAk5YwOBj8';
const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
let konto = null;

function vorgabe() {
  return {
    szenarien: ['norm'],
    mass: 'monat',
    monat: new Date().getMonth() + 1,
    jahr: new Date().getFullYear(),
    wochenProMonat: 4,
    zielUmsatz: 8000,
    kapazitaet: 20,
    fixkosten: [
      { id: 'fk1', name: 'Marketing', betrag: 0 },
      { id: 'fk2', name: 'Buchhaltung', betrag: 0 },
      { id: 'fk3', name: 'Sonstiges', betrag: 0 },
    ],
    fixzeit: [
      { id: 'fz1', name: 'Buchhaltung', stunden: 0 },
      { id: 'fz2', name: 'Marketing', stunden: 0 },
      { id: 'fz3', name: 'Strategie', stunden: 0 },
      { id: 'fz4', name: 'Sonstiges', stunden: 0 },
    ],
    bloecke: [
      {
        id: 'sonderpreise',
        titel: 'Selbstbegegnung — Sonderpreise',
        zusatz: '1.5 h im 1:1 · Preis pro Sitzung',
        art: 'sitzung',
        zeilen: [
          { id: 's1', name: 'Kennenlernen', szenario: 'sonder', aktiv: true, c: 10, e: 1, g: 111, j: 2 },
          { id: 's2', name: 'Sonderpreis',  szenario: 'sonder', aktiv: true, c: 3,  e: 4, g: 120, j: 2 },
          { id: 's3', name: 'Sonderpreis',  szenario: 'sonder', aktiv: true, c: 10, e: 2, g: 100, j: 2 },
        ],
      },
      {
        id: 'autonom15',
        titel: 'Selbstbegegnung autonom',
        zusatz: '1.5 h im 1:1 · Preis pro Sitzung',
        art: 'sitzung',
        zeilen: [
          { id: 'a1', name: 'min',  szenario: 'min',  aktiv: true, c: 1, e: 2, g: 180, j: 2 },
          { id: 'a2', name: 'norm', szenario: 'norm', aktiv: true, c: 1, e: 4, g: 200, j: 2 },
          { id: 'a3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, g: 250, j: 2 },
        ],
      },
      {
        id: 'autonom20',
        titel: 'Selbstbegegnung autonom',
        zusatz: '2 h im 1:1 · Preis pro Sitzung',
        art: 'sitzung',
        zeilen: [
          { id: 'b1', name: 'min',  szenario: 'min',  aktiv: true, c: 2, e: 1, g: 220, j: 2.5 },
          { id: 'b2', name: 'norm', szenario: 'norm', aktiv: true, c: 3, e: 2, g: 260, j: 2.5 },
          { id: 'b3', name: 'max',  szenario: 'max',  aktiv: true, c: 4, e: 4, g: 300, j: 2.5 },
        ],
      },
      {
        id: 'story1zu1',
        titel: 'Story of your life',
        zusatz: 'Paket im 1:1 + KI',
        art: 'paket',
        gruppe: false,
        laufzeit: 3,
        zeilen: [
          { id: 'c1', name: 'min',  szenario: 'min',  aktiv: true, c: 0, e: 0, f: 0, j: 0 },
          { id: 'c2', name: 'norm', szenario: 'norm', aktiv: true, c: 0, e: 0, f: 0, j: 0 },
          { id: 'c3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'storyGruppe',
        titel: 'Story of your life',
        zusatz: 'Paket im Gruppensetting + KI',
        art: 'paket',
        gruppe: true,
        laufzeit: 3,
        zeilen: [
          { id: 'd1', name: 'min',  szenario: 'min',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'd2', name: 'norm', szenario: 'norm', aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'd3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'halbjahr1zu1',
        titel: 'Halbjahresbegleitung',
        zusatz: 'Im 1:1 · Paketpreis über die Laufzeit',
        art: 'paket',
        gruppe: false,
        laufzeit: 6,
        zeilen: [
          { id: 'e1', name: 'min',  szenario: 'min',  aktiv: true, c: 1, e: 1, f: 600,  j: 2 },
          { id: 'e2', name: 'norm', szenario: 'norm', aktiv: true, c: 3, e: 2, f: 1000, j: 2 },
          { id: 'e3', name: 'max',  szenario: 'max',  aktiv: true, c: 4, e: 2, f: 2000, j: 2 },
        ],
      },
      {
        id: 'halbjahrGruppe',
        titel: 'Halbjahresbegleitung',
        zusatz: 'Im Gruppensetting · Paketpreis pro Person',
        art: 'paket',
        gruppe: true,
        laufzeit: 6,
        zeilen: [
          { id: 'f1', name: 'min',  szenario: 'min',  aktiv: true, c: 1, d: 6, e: 2, f: 1200, j: 8 },
          { id: 'f2', name: 'norm', szenario: 'norm', aktiv: true, c: 2, d: 6, e: 2, f: 2000, j: 8 },
          { id: 'f3', name: 'max',  szenario: 'max',  aktiv: true, c: 1, d: 6, e: 1, f: 600,  j: 8 },
        ],
      },
      {
        id: 'jahr1zu1',
        titel: 'Jahresbegleitung',
        zusatz: 'Im 1:1 · Paketpreis über die Laufzeit',
        art: 'paket',
        gruppe: false,
        laufzeit: 12,
        zeilen: [
          { id: 'g1', name: 'min',  szenario: 'min',  aktiv: true, c: 1, e: 0, f: 0, j: 0 },
          { id: 'g2', name: 'norm', szenario: 'norm', aktiv: true, c: 0, e: 0, f: 0, j: 0 },
          { id: 'g3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'jahrGruppe',
        titel: 'Jahresbegleitung',
        zusatz: 'Im Gruppensetting · Paketpreis pro Person',
        art: 'paket',
        gruppe: true,
        laufzeit: 12,
        zeilen: [
          { id: 'h1', name: 'min',  szenario: 'min',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'h2', name: 'norm', szenario: 'norm', aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'h3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'tagesseminar',
        titel: 'Tagesseminar',
        zusatz: 'Drei Rollen je Seminar · Preis pro Person',
        art: 'seminar',
        zeilen: [
          {
            id: 'i1', name: 'Sonderpreis', szenario: 'sonder', aktiv: true, e: 0, j: 6,
            kategorien: [
              { name: 'KlientIn',      d: 0, g: 0 },
              { name: 'Resonanz',      d: 0, g: 0 },
              { name: 'BeobachterIn',  d: 0, g: 0 },
            ],
          },
          {
            id: 'i2', name: 'norm', szenario: 'norm', aktiv: true, e: 0, j: 6,
            kategorien: [
              { name: 'KlientIn',      d: 0, g: 250 },
              { name: 'Resonanz',      d: 0, g: 120 },
              { name: 'BeobachterIn',  d: 0, g: 90 },
            ],
          },
          {
            id: 'i3', name: 'max', szenario: 'max', aktiv: true, e: 0, j: 6,
            kategorien: [
              { name: 'KlientIn',      d: 0, g: 300 },
              { name: 'Resonanz',      d: 0, g: 190 },
              { name: 'BeobachterIn',  d: 0, g: 120 },
            ],
          },
        ],
      },
      {
        id: 'ausbildungGruppe',
        titel: 'Ausbildung — Gruppenbegleitung',
        zusatz: 'Im Gruppensetting · Paketpreis pro Person',
        art: 'paket',
        gruppe: true,
        laufzeit: 6,
        zeilen: [
          { id: 'ag1', name: 'min',  szenario: 'min',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'ag2', name: 'norm', szenario: 'norm', aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'ag3', name: 'max',  szenario: 'max',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'onlinekurs1',
        titel: 'Onlinekurs 1',
        zusatz: 'Preis pro Teilnehmer:in',
        art: 'sitzung',
        zeilen: [
          { id: 'ok1a', name: 'min',  szenario: 'min',  aktiv: true, c: 0, e: 0, g: 0, j: 0 },
          { id: 'ok1b', name: 'norm', szenario: 'norm', aktiv: true, c: 0, e: 0, g: 0, j: 0 },
          { id: 'ok1c', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, g: 0, j: 0 },
        ],
      },
      {
        id: 'onlinekurs2',
        titel: 'Onlinekurs 2',
        zusatz: 'Preis pro Teilnehmer:in',
        art: 'sitzung',
        zeilen: [
          { id: 'ok2a', name: 'min',  szenario: 'min',  aktiv: true, c: 0, e: 0, g: 0, j: 0 },
          { id: 'ok2b', name: 'norm', szenario: 'norm', aktiv: true, c: 0, e: 0, g: 0, j: 0 },
          { id: 'ok2c', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, g: 0, j: 0 },
        ],
      },
      {
        id: 'neuesAngebot1',
        titel: 'Neues Angebot — 1:1',
        zusatz: 'Paketpreis über die Laufzeit',
        art: 'paket',
        gruppe: false,
        laufzeit: 1,
        zeilen: [
          { id: 'na1a', name: 'min',  szenario: 'min',  aktiv: true, c: 0, e: 0, f: 0, j: 0 },
          { id: 'na1b', name: 'norm', szenario: 'norm', aktiv: true, c: 0, e: 0, f: 0, j: 0 },
          { id: 'na1c', name: 'max',  szenario: 'max',  aktiv: true, c: 0, e: 0, f: 0, j: 0 },
        ],
      },
      {
        id: 'neuesAngebot2',
        titel: 'Neues Angebot — Gruppe',
        zusatz: 'Im Gruppensetting · Paketpreis pro Person',
        art: 'paket',
        gruppe: true,
        laufzeit: 1,
        zeilen: [
          { id: 'na2a', name: 'min',  szenario: 'min',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'na2b', name: 'norm', szenario: 'norm', aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
          { id: 'na2c', name: 'max',  szenario: 'max',  aktiv: true, c: 0, d: 0, e: 0, f: 0, j: 0 },
        ],
      },
    ],
  };
}

/* ── 2 · Rechenwerk ─────────────────────────────────────────────────────── */

/*
 * Die Tabelle kennt vier Rechenarten. Alle vier liefern dieselben zwei
 * Grundgrössen — Umsatz pro Monat und Arbeitsstunden pro Monat — und alles
 * andere folgt daraus. Genau darin liegt die Korrektur gegenüber der Datei:
 * dort war der Wochenumsatz der Halbjahresbegleitung im 1:1 zusätzlich durch
 * die Sitzungszahl geteilt und passte deshalb nicht zum Monatsumsatz.
 */
function rechne(block, z, wochen) {
  let umsatzMonat = 0;
  let stundenMonat = 0;

  if (block.art === 'sitzung') {
    umsatzMonat = z.c * z.e * z.g;
    stundenMonat = z.c * z.e * z.j;
  } else if (block.art === 'paket') {
    const proMonat = block.laufzeit > 0 ? z.f / block.laufzeit : 0;
    const koepfe = block.gruppe === true ? z.c * z.d : z.c;
    umsatzMonat = proMonat * koepfe;
    stundenMonat = z.c * z.e * z.j;
  } else {
    const proSeminar = z.kategorien.reduce((summe, k) => summe + k.d * k.g, 0);
    umsatzMonat = z.e * proSeminar;
    stundenMonat = z.e * z.j;
  }

  const w = wochen > 0 ? wochen : 1;
  return {
    umsatzWoche: umsatzMonat / w,
    umsatzMonat,
    umsatzJahr: umsatzMonat * 12,
    stundenWoche: stundenMonat / w,
    stundenMonat,
    // Ein Stundensatz ohne Stunden ist keine Null, sondern keine Aussage.
    stundensatz: stundenMonat > 0 ? umsatzMonat / stundenMonat : null,
  };
}

const LEER = {
  umsatzWoche: 0, umsatzMonat: 0, umsatzJahr: 0,
  stundenWoche: 0, stundenMonat: 0, stundensatz: null,
};

function summiere(teile) {
  const summe = teile.reduce(
    (a, t) => ({
      umsatzWoche: a.umsatzWoche + t.umsatzWoche,
      umsatzMonat: a.umsatzMonat + t.umsatzMonat,
      umsatzJahr: a.umsatzJahr + t.umsatzJahr,
      stundenWoche: a.stundenWoche + t.stundenWoche,
      stundenMonat: a.stundenMonat + t.stundenMonat,
      stundensatz: null,
    }),
    { ...LEER },
  );
  summe.stundensatz = summe.stundenMonat > 0 ? summe.umsatzMonat / summe.stundenMonat : null;
  return summe;
}

/** Zählt eine Zeile ins Total? Zwei Bedingungen: angehakt und im Szenario. */
function zaehlt(zeile, szenarien) {
  if (zeile.aktiv !== true) return false;
  return szenarien.includes(zeile.szenario);
}

/* ── Formate ────────────────────────────────────────────────────────────── */

const fGeld = new Intl.NumberFormat('de-CH', { maximumFractionDigits: 0 });
const fStd = new Intl.NumberFormat('de-CH', { maximumFractionDigits: 2 });

const geld = (n) => fGeld.format(Math.round(n));
const std = (n) => fStd.format(n);
const satz = (n) => (n === null ? '—' : fGeld.format(Math.round(n)));

function esc(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/* ── Zustand ────────────────────────────────────────────────────────────── */

let zustand = laden();

function laden() {
  const frisch = vorgabe();
  let gespeichert = null;
  try {
    const roh = localStorage.getItem(SCHLUESSEL);
    if (roh !== null) gespeichert = JSON.parse(roh);
  } catch {
    gespeichert = null;
  }
  return verschmelze(frisch, gespeichert);
}

function verschmelze(frisch, gespeichert) {
  if (gespeichert === null || gespeichert === undefined) return frisch;

  // Verschmelzen statt ersetzen: eine ältere Fassung darf keine Blöcke schlucken.
  if (Array.isArray(gespeichert.szenarien)) {
    frisch.szenarien = gespeichert.szenarien;
  } else if (typeof gespeichert.szenario === 'string') {
    // Alte Fassung vor der Mehrfachauswahl: 'alle' -> alle vier, sonst ein Wert.
    frisch.szenarien = gespeichert.szenario === 'alle' ? ['min', 'norm', 'max', 'sonder'] : [gespeichert.szenario];
  }
  frisch.mass = gespeichert.mass ?? frisch.mass;
  frisch.monat = gespeichert.monat ?? frisch.monat;
  frisch.jahr = gespeichert.jahr ?? frisch.jahr;
  frisch.wochenProMonat = gespeichert.wochenProMonat ?? frisch.wochenProMonat;
  frisch.zielUmsatz = gespeichert.zielUmsatz ?? frisch.zielUmsatz;
  frisch.kapazitaet = gespeichert.kapazitaet ?? frisch.kapazitaet;

  for (const posten of frisch.fixkosten) {
    const alt = (gespeichert.fixkosten ?? []).find((f) => f.id === posten.id);
    if (alt !== undefined && typeof alt.betrag === 'number') posten.betrag = alt.betrag;
  }

  for (const posten of frisch.fixzeit) {
    const alt = (gespeichert.fixzeit ?? []).find((f) => f.id === posten.id);
    if (alt !== undefined && typeof alt.stunden === 'number') posten.stunden = alt.stunden;
  }

  for (const block of frisch.bloecke) {
    const alt = (gespeichert.bloecke ?? []).find((b) => b.id === block.id);
    if (alt === undefined) continue;
    if (typeof alt.laufzeit === 'number') block.laufzeit = alt.laufzeit;
    for (const zeile of block.zeilen) {
      const altZ = (alt.zeilen ?? []).find((z) => z.id === zeile.id);
      if (altZ === undefined) continue;
      for (const k of ['c', 'd', 'e', 'f', 'g', 'j']) {
        if (typeof altZ[k] === 'number') zeile[k] = altZ[k];
      }
      if (typeof altZ.aktiv === 'boolean') zeile.aktiv = altZ.aktiv;
      if (Array.isArray(zeile.kategorien) && Array.isArray(altZ.kategorien)) {
        zeile.kategorien.forEach((kat, i) => {
          const altK = altZ.kategorien[i];
          if (altK === undefined) return;
          if (typeof altK.d === 'number') kat.d = altK.d;
          if (typeof altK.g === 'number') kat.g = altK.g;
        });
      }
    }
  }
  return frisch;
}

let speicherUhr = null;
function speichern() {
  clearTimeout(speicherUhr);
  speicherUhr = setTimeout(speichernJetzt, 400);
}

async function speichernJetzt() {
  try {
    localStorage.setItem(SCHLUESSEL, JSON.stringify(zustand));
  } catch { /* privates Fenster — dann eben ohne Gedächtnis */ }

  const stand = document.getElementById('stand');
  if (konto !== null) {
    const { error } = await sb.from('finanzplan_daten').upsert({
      user_id: konto.id,
      daten: zustand,
      aktualisiert_am: new Date().toISOString(),
    });
    if (error) {
      console.error(error);
      stand.textContent = 'Fehler beim Sichern';
      stand.className = 'abzeichen abzeichen-ernst';
      return;
    }
  }
  stand.textContent = 'gespeichert';
  stand.className = 'abzeichen abzeichen-gut';
  setTimeout(() => { stand.className = 'abzeichen abzeichen-neutral'; }, 1200);
}

/* ── Konto: Login, Logout, Laden vom Server ────────────────────────────────── */

let formModus = 'anmelden';

function torFormularHtml() {
  const istRegister = formModus === 'registrieren';
  return (
    '<form class="konto-formular" id="konto-formular">' +
      '<input type="email" id="konto-email" class="konto-eingabe" placeholder="deine@email.ch" autocomplete="email" required>' +
      '<input type="password" id="konto-passwort" class="konto-eingabe" placeholder="Passwort" ' +
        'autocomplete="' + (istRegister ? 'new-password' : 'current-password') + '" minlength="6" required>' +
      '<button type="submit" data-aktion="' + formModus + '" class="knopf knopf-primaer knopf-klein" style="width:100%">' +
        (istRegister ? 'Registrieren' : 'Anmelden') +
      '</button>' +
      '<button type="button" id="konto-modus-wechsel" class="konto-link">' +
        (istRegister ? 'Schon ein Konto? Anmelden' : 'Noch kein Konto? Registrieren') +
      '</button>' +
      '<button type="button" id="konto-passwort-vergessen" class="konto-link">Passwort vergessen?</button>' +
      '<p class="konto-status" id="konto-fehler"></p>' +
    '</form>'
  );
}

function wiederherstellungHtml() {
  return (
    '<form class="konto-formular" id="konto-neues-passwort-formular">' +
      '<p class="klein gedaempft">Neues Passwort festlegen (mindestens 6 Zeichen).</p>' +
      '<input type="password" id="neues-passwort" class="konto-eingabe" placeholder="Neues Passwort" minlength="6" autocomplete="new-password" required>' +
      '<button type="submit" class="knopf knopf-primaer knopf-klein">Passwort speichern</button>' +
      '<button type="button" id="passwort-form-abbrechen" class="konto-link">Abbrechen</button>' +
      '<p class="konto-status" id="konto-fehler"></p>' +
    '</form>'
  );
}

function kontoFehlertext(error, aktion) {
  const m = error.message || '';
  if (m.includes('already registered') || m.includes('already exists')) {
    return 'Diese Adresse ist schon registriert — auf «Anmelden» klicken.';
  }
  if (m.includes('Invalid login credentials')) {
    return 'E-Mail oder Passwort stimmt nicht.';
  }
  if (m.toLowerCase().includes('password')) {
    return 'Passwort zu kurz (mindestens 6 Zeichen).';
  }
  return aktion === 'registrieren' ? 'Registrierung fehlgeschlagen.' : 'Anmeldung fehlgeschlagen.';
}

// Vor dem ersten Sitzungscheck bleiben Tor und Inhalt beide verborgen —
// lieber ein leerer Moment als ein Aufblitzen des Finanzplans ohne Login.
function zeichneKonto() {
  const tor = document.getElementById('tor');
  const app = document.getElementById('app-inhalt');
  const emailAnzeige = document.getElementById('konto-email-anzeige');

  if (konto === null) {
    app.hidden = true;
    tor.hidden = false;
    const platz = document.getElementById('konto-platz');
    if (platz !== null) platz.innerHTML = torFormularHtml();
  } else {
    tor.hidden = true;
    app.hidden = false;
    if (emailAnzeige !== null) emailAnzeige.textContent = konto.email;
  }
}

// Klick auf den Link in der Passwort-Reset-Mail landet hier mit einer
// PASSWORD_RECOVERY-Sitzung — ohne neues Passwort kommt man beim nächsten
// Mal aber nicht mehr rein.
function zeichneWiederherstellung() {
  const tor = document.getElementById('tor');
  const app = document.getElementById('app-inhalt');
  app.hidden = true;
  tor.hidden = false;
  const platz = document.getElementById('konto-platz');
  if (platz !== null) platz.innerHTML = wiederherstellungHtml();
}

async function ladeVonServer() {
  if (konto === null) return;
  const { data, error } = await sb.from('finanzplan_daten').select('daten').eq('user_id', konto.id).maybeSingle();
  if (error) { console.error(error); return; }
  zustand = verschmelze(vorgabe(), data === null ? null : data.daten);
  zeichne();
  aktualisiere();
}

async function initKonto() {
  const { data: { session } } = await sb.auth.getSession();
  konto = session === null ? null : { id: session.user.id, email: session.user.email };
  zeichneKonto();
  if (konto !== null) await ladeVonServer();

  sb.auth.onAuthStateChange((ereignis, session) => {
    if (ereignis === 'PASSWORD_RECOVERY') {
      zeichneWiederherstellung();
      return;
    }
    konto = session === null ? null : { id: session.user.id, email: session.user.email };
    zeichneKonto();
    if (konto !== null) ladeVonServer();
  });
}

/* ── 3 · Gerüst ─────────────────────────────────────────────────────────── */

const FELDER = {
  sitzung: [
    { k: 'c', titel: 'Klient:innen<br>pro Monat', schritt: '1' },
    { k: 'e', titel: 'Sitzungen<br>pro Monat', schritt: '1' },
    { k: 'g', titel: 'Preis<br>pro Sitzung', schritt: '5' },
    { k: 'j', titel: 'Dauer<br>Std. brutto', schritt: '0.5' },
  ],
  paketEinzel: [
    { k: 'c', titel: 'Klient:innen', schritt: '1' },
    { k: 'e', titel: 'Sitzungen<br>pro Monat', schritt: '1' },
    { k: 'f', titel: 'Paketpreis', schritt: '50' },
    { k: 'j', titel: 'Dauer<br>Std. brutto', schritt: '0.5' },
  ],
  paketGruppe: [
    { k: 'c', titel: 'Gruppen', schritt: '1' },
    { k: 'd', titel: 'Personen<br>pro Gruppe', schritt: '1' },
    { k: 'e', titel: 'Treffen<br>pro Monat', schritt: '1' },
    { k: 'f', titel: 'Paketpreis<br>pro Person', schritt: '50' },
    { k: 'j', titel: 'Dauer<br>Std. brutto', schritt: '0.5' },
  ],
};

const ERGEBNISSE = [
  { k: 'umsatzWoche', titel: 'Umsatz<br>pro Woche', form: geld },
  { k: 'umsatzMonat', titel: 'Umsatz<br>pro Monat', form: geld, stark: true },
  { k: 'umsatzJahr', titel: 'Umsatz<br>pro Jahr', form: geld },
  { k: 'stundenWoche', titel: 'Arbeitsstd.<br>pro Woche', form: std },
  { k: 'stundensatz', titel: 'Stundensatz<br>CHF', form: satz },
];

function felderFuer(block) {
  if (block.art === 'sitzung') return FELDER.sitzung;
  return block.gruppe === true ? FELDER.paketGruppe : FELDER.paketEinzel;
}

const HAKEN = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 13l4.5 4.5L19 7"/></svg>';

function schalterHtml(zeile) {
  return (
    '<label class="schalter" title="Zeile zählt ins Total">' +
    '<input type="checkbox" data-schalter="' + zeile.id + '"' + (zeile.aktiv ? ' checked' : '') +
    ' aria-label="' + esc(zeile.name) + ' ins Total zählen">' +
    '<span>' + HAKEN + '</span></label>'
  );
}

function eingabeHtml(zeile, feld) {
  const wert = zeile[feld.k] ?? 0;
  return (
    '<td class="eingabe"><input class="zelleneingabe' + (wert === 0 ? ' leer' : '') +
    '" type="number" inputmode="decimal" min="0" step="' + feld.schritt +
    '" value="' + wert + '" data-zeile="' + zeile.id + '" data-feld="' + feld.k + '"' +
    ' aria-label="' + esc(feld.titel.replace(/<br>/g, ' ')) + '"></td>'
  );
}

function tabellenBlockHtml(block) {
  const felder = felderFuer(block);
  const spalten = 1 + felder.length;

  let kopf = '<tr><th>Variante</th>';
  felder.forEach((f) => { kopf += '<th class="rechts">' + f.titel + '</th>'; });
  ERGEBNISSE.forEach((e, i) => {
    kopf += '<th class="rechts' + (i === 0 ? ' grenze' : '') + '">' + e.titel + '</th>';
  });
  kopf += '</tr>';

  let koerper = '';
  for (const zeile of block.zeilen) {
    koerper += '<tr data-reihe="' + zeile.id + '">';
    koerper += '<td><span class="variante">' + schalterHtml(zeile) +
      '<span class="variantenname">' + esc(zeile.name) + '</span></span></td>';
    felder.forEach((f) => { koerper += eingabeHtml(zeile, f); });
    ERGEBNISSE.forEach((e, i) => {
      koerper += '<td class="ergebnis' + (i === 0 ? ' grenze' : '') + (e.stark ? ' stark' : '') +
        '" data-aus="' + zeile.id + '.' + e.k + '"></td>';
    });
    koerper += '</tr>';
  }

  let fuss = '<tr><td colspan="' + spalten + '">Zwischensumme <span class="klein gedaempft">gezählte Zeilen</span></td>';
  ERGEBNISSE.forEach((e, i) => {
    fuss += '<td class="rechts' + (i === 0 ? ' grenze' : '') + '" data-aus="' + block.id + '.summe.' + e.k + '"></td>';
  });
  fuss += '</tr>';

  return (
    '<div class="tabellenrahmen"><div class="tabellenlauf"><table class="tabelle">' +
    '<caption class="nur-vorlesen">' + esc(block.titel + ' — ' + block.zusatz) + '</caption>' +
    '<thead>' + kopf + '</thead><tbody>' + koerper + '</tbody><tfoot>' + fuss + '</tfoot>' +
    '</table></div></div>'
  );
}

function seminarBlockHtml(block) {
  let html = '<div class="seminare">';
  for (const zeile of block.zeilen) {
    let reihen = '';
    zeile.kategorien.forEach((kat, i) => {
      reihen +=
        '<tr><td>' + esc(kat.name) + '</td>' +
        '<td><input class="zelleneingabe' + (kat.d === 0 ? ' leer' : '') + '" type="number" min="0" step="1" value="' + kat.d +
        '" data-zeile="' + zeile.id + '" data-kat="' + i + '" data-feld="d" aria-label="Personen ' + esc(kat.name) + '"></td>' +
        '<td><input class="zelleneingabe' + (kat.g === 0 ? ' leer' : '') + '" type="number" min="0" step="10" value="' + kat.g +
        '" data-zeile="' + zeile.id + '" data-kat="' + i + '" data-feld="g" aria-label="Preis ' + esc(kat.name) + '"></td>' +
        '<td class="rechts" data-aus="' + zeile.id + '.kat' + i + '"></td></tr>';
    });

    html +=
      '<div class="seminar" data-reihe="' + zeile.id + '">' +
        '<div class="seminar-kopf">' +
          '<span class="variante">' + schalterHtml(zeile) +
          '<span class="variantenname">' + esc(zeile.name) + '</span></span>' +
          '<span class="abzeichen abzeichen-neutral" data-aus="' + zeile.id + '.proSeminar"></span>' +
        '</div>' +
        '<div class="seminar-felder">' +
          '<label class="seminar-feld"><span>Seminare / Monat</span>' +
            '<input class="zelleneingabe' + (zeile.e === 0 ? ' leer' : '') + '" type="number" min="0" step="1" value="' + zeile.e +
            '" data-zeile="' + zeile.id + '" data-feld="e" aria-label="Seminare pro Monat"></label>' +
          '<label class="seminar-feld"><span>Dauer Std. brutto</span>' +
            '<input class="zelleneingabe' + (zeile.j === 0 ? ' leer' : '') + '" type="number" min="0" step="0.5" value="' + zeile.j +
            '" data-zeile="' + zeile.id + '" data-feld="j" aria-label="Dauer in Stunden"></label>' +
        '</div>' +
        '<table><thead><tr><th>Rolle</th><th class="rechts">Personen</th><th class="rechts">Preis</th>' +
        '<th class="rechts">Umsatz</th></tr></thead><tbody>' + reihen + '</tbody></table>' +
        '<dl class="seminar-fuss">' +
          '<div><dt>Umsatz / Monat</dt><dd data-aus="' + zeile.id + '.umsatzMonat"></dd></div>' +
          '<div><dt>Std. / Woche</dt><dd data-aus="' + zeile.id + '.stundenWoche"></dd></div>' +
          '<div><dt>Stundensatz</dt><dd data-aus="' + zeile.id + '.stundensatz"></dd></div>' +
        '</dl>' +
      '</div>';
  }
  return html + '</div>';
}

/*
 * Vier Angebote stehen in der Tabelle vollstaendig auf null — sie sind
 * vorgesehen, aber noch nicht geplant. Ausgeklappt kosten sie zusammen die
 * halbe Seitenlaenge und zeigen dreissig Nullen. Eingeklappt bleiben sie
 * sichtbar, ohne den Blick von den Angeboten wegzuziehen, die tragen.
 */
const aufgeklappt = new Set();

function blockLeer(block) {
  return block.zeilen.every((z) => {
    const e = rechne(block, z, zustand.wochenProMonat);
    return e.umsatzMonat === 0 && e.stundenMonat === 0;
  });
}

function blockKopfHtml(block, laufzeit) {
  return (
    '<div>' +
      '<div class="block-titelzeile"><h3>' + esc(block.titel) + '</h3>' + laufzeit + '</div>' +
      '<p class="klein gedaempft" style="margin-top:4px">' + esc(block.zusatz) + '</p>' +
    '</div>'
  );
}

function blockHtml(block) {
  const laufzeit = typeof block.laufzeit === 'number'
    ? '<span class="laufzeit">Laufzeit <input type="number" min="1" max="36" step="1" value="' + block.laufzeit +
      '" data-laufzeit="' + block.id + '" aria-label="Laufzeit in Monaten"> Mt.</span>'
    : '';

  if (blockLeer(block) && !aufgeklappt.has(block.id)) {
    return (
      '<article class="karte karte-eingeklappt">' +
        '<div class="block-kopf" style="margin-bottom:0">' +
          blockKopfHtml(block, laufzeit) +
          '<div class="block-summe">' +
            '<span class="nicht-hinterlegt">noch nicht geplant</span>' +
            '<button type="button" class="knopf knopf-umriss knopf-klein" data-klapp="' + block.id + '">Planen</button>' +
          '</div>' +
        '</div>' +
      '</article>'
    );
  }

  return (
    '<article class="karte">' +
      '<div class="block-kopf">' +
        blockKopfHtml(block, laufzeit) +
        '<div class="block-summe">' +
          '<span class="block-summe-wert" data-aus="' + block.id + '.kopfsumme"></span>' +
          '<span class="klein gedaempft">CHF / Monat</span>' +
        '</div>' +
      '</div>' +
      (block.art === 'seminar' ? seminarBlockHtml(block) : tabellenBlockHtml(block)) +
    '</article>'
  );
}

/* ── Fixkosten ──────────────────────────────────────────────────────────── */

function fixkostenHtml() {
  return zustand.fixkosten
    .map(
      (f) =>
        '<label class="feld">' +
          '<span class="feld-beschriftung">' + esc(f.name) + '</span>' +
          '<input type="number" inputmode="decimal" min="0" step="10" value="' + f.betrag +
          '" data-fixkosten="' + f.id + '" aria-label="' + esc(f.name) + '">' +
        '</label>',
    )
    .join('');
}

function fixzeitHtml() {
  return zustand.fixzeit
    .map(
      (f) =>
        '<label class="feld">' +
          '<span class="feld-beschriftung">' + esc(f.name) + '</span>' +
          '<input type="number" inputmode="decimal" min="0" step="1" value="' + f.stunden +
          '" data-fixzeit="' + f.id + '" aria-label="' + esc(f.name) + '">' +
        '</label>',
    )
    .join('');
}

/* ── Kennzahlen und Diagramm ────────────────────────────────────────────── */

function kennzahlenHtml() {
  const karten = [
    { id: 'kUmsatzMonat', name: 'Umsatz pro Monat', einheit: 'CHF', hero: true },
    { id: 'kUmsatzJahr', name: 'Umsatz pro Jahr', einheit: 'CHF' },
    { id: 'kStunden', name: 'Arbeitsstunden pro Woche', einheit: 'Std.' },
    { id: 'kSatz', name: 'Ø Stundensatz', einheit: 'CHF' },
  ];
  return karten
    .map(
      (k) =>
        '<div class="kennzahl' + (k.hero ? ' kennzahl-hero' : '') + '">' +
          '<span class="kennzahl-name">' + k.name + '</span>' +
          '<span class="kennzahl-wert" data-aus="' + k.id + '">0<span class="kennzahl-einheit">' + k.einheit + '</span></span>' +
          '<span class="kennzahl-zusatz" data-aus="' + k.id + '.zusatz"></span>' +
        '</div>',
    )
    .join('');
}

const MASSE = {
  monat: { name: 'Umsatz pro Monat', wert: (e) => e.umsatzMonat, form: geld },
  stunden: { name: 'Arbeitsstunden pro Woche', wert: (e) => e.stundenWoche, form: std },
  satz: { name: 'Ø Stundensatz', wert: (e) => e.stundensatz ?? 0, form: (n) => (n === 0 ? '—' : geld(n)) },
};

function zeichneDiagramm() {
  const ziel = document.getElementById('diagramm');
  const mass = MASSE[zustand.mass];

  const reihen = zustand.bloecke
    .map((block) => {
      const teile = block.zeilen
        .filter((z) => zaehlt(z, zustand.szenarien))
        .map((z) => rechne(block, z, zustand.wochenProMonat));
      return { name: block.titel + (block.zusatz ? ' · ' + block.zusatz.split(' · ')[0] : ''), ergebnis: summiere(teile) };
    })
    .map((r) => ({ name: r.name, wert: mass.wert(r.ergebnis) }))
    .filter((r) => r.wert > 0)
    .sort((a, b) => b.wert - a.wert);

  if (reihen.length === 0) {
    ziel.innerHTML =
      '<div class="leerzustand"><span class="leerzustand-titel">Nichts zu zeigen</span>' +
      '<span class="klein">Für dieses Szenario zählt keine Zeile mit einem Wert über null.</span></div>';
    return;
  }

  const gipfel = reihen[0].wert;
  ziel.innerHTML = reihen
    .map(
      (r) =>
        '<div class="balkenzeile">' +
          '<span class="balkenname" title="' + esc(r.name) + '">' + esc(r.name) + '</span>' +
          '<span class="balkenspur"><span class="balken" style="width:' +
            (gipfel > 0 ? Math.max(2, (r.wert / gipfel) * 100) : 0) + '%"></span></span>' +
          '<span class="balkenwert">' + mass.form(r.wert) + '</span>' +
        '</div>',
    )
    .join('');
}

/* ── Aktualisieren ──────────────────────────────────────────────────────── */

const aus = {};

function setze(schluessel, text) {
  const knoten = aus[schluessel];
  if (knoten !== undefined) knoten.textContent = text;
}

function stufe(anteil, umgekehrt) {
  // umgekehrt: mehr ist schlechter (Auslastung). Sonst: mehr ist besser (Ziel).
  if (umgekehrt === true) {
    if (anteil > 1.15) return 'kritisch';
    if (anteil > 1) return 'ernst';
    if (anteil > 0.85) return 'achtung';
    return 'gut';
  }
  if (anteil >= 1) return 'gut';
  if (anteil >= 0.75) return 'achtung';
  return 'ernst';
}

function aktualisiere() {
  const w = zustand.wochenProMonat;
  const gesamtTeile = [];
  let gezaehlt = 0;
  let vorhanden = 0;

  for (const block of zustand.bloecke) {
    const blockTeile = [];
    for (const zeile of block.zeilen) {
      vorhanden += 1;
      const e = rechne(block, zeile, w);
      const zaehltMit = zaehlt(zeile, zustand.szenarien);
      if (zaehltMit) {
        gezaehlt += 1;
        blockTeile.push(e);
        gesamtTeile.push(e);
      }

      if (block.art === 'seminar') {
        const proSeminar = zeile.kategorien.reduce((s, k) => s + k.d * k.g, 0);
        setze(zeile.id + '.proSeminar', geld(proSeminar) + ' / Seminar');
        zeile.kategorien.forEach((k, i) => setze(zeile.id + '.kat' + i, geld(k.d * k.g)));
        setze(zeile.id + '.umsatzMonat', geld(e.umsatzMonat));
        setze(zeile.id + '.stundenWoche', std(e.stundenWoche));
        setze(zeile.id + '.stundensatz', satz(e.stundensatz));
      } else {
        ERGEBNISSE.forEach((sp) => setze(zeile.id + '.' + sp.k, sp.form(e[sp.k])));
      }

      const reihe = document.querySelector('[data-reihe="' + zeile.id + '"]');
      if (reihe !== null) {
        reihe.classList.toggle(block.art === 'seminar' ? 'seminar-aus' : 'zeile-aus', !zaehltMit);
      }
    }

    const summe = summiere(blockTeile);
    ERGEBNISSE.forEach((sp) => setze(block.id + '.summe.' + sp.k, blockTeile.length === 0 ? '—' : sp.form(summe[sp.k])));
    setze(block.id + '.kopfsumme', geld(summe.umsatzMonat));
  }

  const total = summiere(gesamtTeile);

  document.querySelector('[data-aus="kUmsatzMonat"]').innerHTML =
    geld(total.umsatzMonat) + '<span class="kennzahl-einheit">CHF</span>';
  document.querySelector('[data-aus="kUmsatzJahr"]').innerHTML =
    geld(total.umsatzJahr) + '<span class="kennzahl-einheit">CHF</span>';
  document.querySelector('[data-aus="kStunden"]').innerHTML =
    std(total.stundenWoche) + '<span class="kennzahl-einheit">Std.</span>';
  document.querySelector('[data-aus="kSatz"]').innerHTML =
    satz(total.stundensatz) + '<span class="kennzahl-einheit">CHF</span>';

  const szenarioWort = zustand.szenarien.length === 0 ? 'keine Auswahl' : zustand.szenarien.join(' + ');
  setze('kUmsatzMonat.zusatz', geld(total.umsatzWoche) + ' pro Woche · ' + szenarioWort);
  setze('kUmsatzJahr.zusatz', '12 × Monatsumsatz');
  setze('kStunden.zusatz', std(total.stundenMonat) + ' Std. pro Monat');
  setze('kSatz.zusatz', gezaehlt + ' von ' + vorhanden + ' Zeilen gezählt');

  /* Ziel */
  const zielAnteil = zustand.zielUmsatz > 0 ? total.umsatzMonat / zustand.zielUmsatz : 1;
  const zielStufe = stufe(zielAnteil, false);
  const zielBalken = document.getElementById('ziel-balken');
  zielBalken.style.width = Math.min(100, zielAnteil * 100) + '%';
  zielBalken.parentElement.className = 'messbalken messbalken-' + zielStufe;
  document.getElementById('ziel-wert').textContent = geld(zustand.zielUmsatz) + ' CHF';
  const luecke = zustand.zielUmsatz - total.umsatzMonat;
  document.getElementById('ziel-text').innerHTML =
    luecke <= 0
      ? '<span class="abzeichen abzeichen-gut">Ziel erreicht</span> ' + geld(-luecke) + ' CHF darüber.'
      : '<span class="abzeichen abzeichen-' + zielStufe + '">' + Math.round(zielAnteil * 100) + '&nbsp;% erreicht</span> ' +
        'Es fehlen ' + geld(luecke) + ' CHF pro Monat.';

  /* Fixkosten & Verwaltungsstunden */
  const fixkostenSumme = zustand.fixkosten.reduce((s, f) => s + (f.betrag || 0), 0);
  document.getElementById('fixkosten-summe').textContent = geld(fixkostenSumme) + ' CHF / Monat';
  const fixzeitSumme = zustand.fixzeit.reduce((s, f) => s + (f.stunden || 0), 0);
  document.getElementById('fixzeit-summe').textContent = std(fixzeitSumme) + ' Std. / Monat';
  const fixzeitWoche = fixzeitSumme / w;

  /* Kapazität — Kundenarbeit plus Verwaltungsstunden zusammen gegen die Kapazität */
  const belegteStunden = total.stundenWoche + fixzeitWoche;
  const auslastung = zustand.kapazitaet > 0 ? belegteStunden / zustand.kapazitaet : 0;
  const kapStufe = stufe(auslastung, true);
  const kapBalken = document.getElementById('kapazitaet-balken');
  kapBalken.style.width = Math.min(100, auslastung * 100) + '%';
  kapBalken.parentElement.className = 'messbalken messbalken-' + kapStufe;
  document.getElementById('kapazitaet-wert').textContent = std(zustand.kapazitaet) + ' Std.';
  const rest = zustand.kapazitaet - belegteStunden;
  document.getElementById('kapazitaet-text').innerHTML =
    '<span class="abzeichen abzeichen-' + kapStufe + '">' + Math.round(auslastung * 100) + '&nbsp;% ausgelastet</span> ' +
    '(' + std(fixzeitWoche) + ' Std./Woche davon Verwaltung) ' +
    (rest >= 0 ? std(rest) + ' Std. pro Woche frei.' : std(-rest) + ' Std. pro Woche über der Kapazität.');

  /* Eckwerte */
  document.getElementById('eckwerte').innerHTML =
    '<div><dt>Umsatz pro Woche</dt><dd>' + geld(total.umsatzWoche) + ' CHF</dd></div>' +
    '<div><dt>Arbeitsstunden pro Monat</dt><dd>' + std(total.stundenMonat) + ' Std.</dd></div>' +
    '<div><dt>Umsatz pro Arbeitstag</dt><dd>' + geld(total.umsatzWoche / 5) + ' CHF</dd></div>' +
    '<div><dt>Gezählte Zeilen</dt><dd>' + gezaehlt + ' / ' + vorhanden + '</dd></div>' +
    '<div><dt>Netto-Umsatz pro Monat (nach Fixkosten)</dt><dd>' + geld(total.umsatzMonat - fixkostenSumme) + ' CHF</dd></div>' +
    '<div><dt>Verwaltungsstunden pro Woche</dt><dd>' + std(fixzeitWoche) + ' Std.</dd></div>';

  zeichneDiagramm();
}

/* ── Verdrahtung ────────────────────────────────────────────────────────── */

function findeZeile(id) {
  for (const block of zustand.bloecke) {
    const zeile = block.zeilen.find((z) => z.id === id);
    if (zeile !== undefined) return zeile;
  }
  return null;
}

function zeichne() {
  document.getElementById('fixkosten').innerHTML = fixkostenHtml();
  document.getElementById('fixzeit').innerHTML = fixzeitHtml();
  document.getElementById('kennzahlen').innerHTML = kennzahlenHtml();

  // Befüllte Angebote nach oben, unbefüllte ("noch nicht geplant") nach unten —
  // stabile Sortierung erhält die Reihenfolge innerhalb jeder Gruppe.
  const sortierteBloecke = [...zustand.bloecke].sort(
    (a, b) => Number(blockLeer(a)) - Number(blockLeer(b)),
  );
  document.getElementById('bloecke').innerHTML = sortierteBloecke.map(blockHtml).join('');

  for (const k of Object.keys(aus)) delete aus[k];
  document.querySelectorAll('[data-aus]').forEach((knoten) => { aus[knoten.dataset.aus] = knoten; });

  document.getElementById('ziel').value = String(zustand.zielUmsatz);
  document.getElementById('kapazitaet').value = String(zustand.kapazitaet);
  document.getElementById('wochen').value = String(zustand.wochenProMonat);
  document.getElementById('monat').value = String(zustand.monat);
  document.getElementById('jahr').value = String(zustand.jahr);

  document.querySelectorAll('[data-szenario]').forEach((knopf) => {
    knopf.setAttribute('aria-pressed', String(zustand.szenarien.includes(knopf.dataset.szenario)));
  });
  document.querySelectorAll('[data-mass]').forEach((knopf) => {
    knopf.setAttribute('aria-pressed', String(knopf.dataset.mass === zustand.mass));
  });
}

function verdrahte() {
  document.body.addEventListener('input', (e) => {
    const ziel = e.target;

    if (ziel.dataset.zeile !== undefined) {
      const zeile = findeZeile(ziel.dataset.zeile);
      if (zeile === null) return;
      const wert = Number.parseFloat(ziel.value);
      const zahl = Number.isFinite(wert) && wert >= 0 ? wert : 0;
      if (ziel.dataset.kat !== undefined) {
        zeile.kategorien[Number(ziel.dataset.kat)][ziel.dataset.feld] = zahl;
      } else {
        zeile[ziel.dataset.feld] = zahl;
      }
      ziel.classList.toggle('leer', zahl === 0);
      aktualisiere();
      speichern();
      return;
    }

    if (ziel.dataset.fixkosten !== undefined) {
      const posten = zustand.fixkosten.find((f) => f.id === ziel.dataset.fixkosten);
      if (posten === undefined) return;
      const wert = Number.parseFloat(ziel.value);
      posten.betrag = Number.isFinite(wert) && wert >= 0 ? wert : 0;
      aktualisiere();
      speichern();
      return;
    }

    if (ziel.dataset.fixzeit !== undefined) {
      const posten = zustand.fixzeit.find((f) => f.id === ziel.dataset.fixzeit);
      if (posten === undefined) return;
      const wert = Number.parseFloat(ziel.value);
      posten.stunden = Number.isFinite(wert) && wert >= 0 ? wert : 0;
      aktualisiere();
      speichern();
      return;
    }

    if (ziel.dataset.laufzeit !== undefined) {
      const block = zustand.bloecke.find((b) => b.id === ziel.dataset.laufzeit);
      const wert = Number.parseFloat(ziel.value);
      if (block !== undefined) block.laufzeit = Number.isFinite(wert) && wert > 0 ? wert : 1;
      aktualisiere();
      speichern();
      return;
    }

    if (ziel.id === 'ziel') { zustand.zielUmsatz = Number(ziel.value); aktualisiere(); speichern(); return; }
    if (ziel.id === 'kapazitaet') { zustand.kapazitaet = Number(ziel.value); aktualisiere(); speichern(); return; }
    if (ziel.id === 'wochen') {
      const wert = Number.parseFloat(ziel.value);
      zustand.wochenProMonat = Number.isFinite(wert) && wert > 0 ? wert : 4;
      aktualisiere();
      speichern();
    }
  });

  document.body.addEventListener('change', (e) => {
    if (e.target.id === 'monat') { zustand.monat = Number(e.target.value); speichern(); return; }
    if (e.target.id === 'jahr') { zustand.jahr = Number(e.target.value); speichern(); return; }

    const id = e.target.dataset.schalter;
    if (id === undefined) return;
    const zeile = findeZeile(id);
    if (zeile === null) return;
    zeile.aktiv = e.target.checked;
    aktualisiere();
    speichern();
  });

  document.body.addEventListener('click', async (e) => {
    const alleKnopf = e.target.closest('[data-szenario-alle]');
    if (alleKnopf !== null) {
      zustand.szenarien = ['min', 'norm', 'max', 'sonder'];
      document.querySelectorAll('[data-szenario]').forEach((k) => {
        k.setAttribute('aria-pressed', String(zustand.szenarien.includes(k.dataset.szenario)));
      });
      aktualisiere();
      speichern();
      return;
    }

    const knopf = e.target.closest('[data-szenario], [data-mass]');
    if (knopf !== null) {
      if (knopf.dataset.szenario !== undefined) {
        // Mehrfachauswahl: jeder Knopf schaltet sich selbst ein/aus.
        const s = knopf.dataset.szenario;
        const i = zustand.szenarien.indexOf(s);
        if (i === -1) zustand.szenarien.push(s); else zustand.szenarien.splice(i, 1);
        knopf.setAttribute('aria-pressed', String(zustand.szenarien.includes(s)));
      } else {
        zustand.mass = knopf.dataset.mass;
        document.querySelectorAll('[data-mass]').forEach((k) => {
          k.setAttribute('aria-pressed', String(k.dataset.mass === zustand.mass));
        });
      }
      aktualisiere();
      speichern();
      return;
    }

    const klapp = e.target.closest('[data-klapp]');
    if (klapp !== null) {
      const blockId = klapp.dataset.klapp;
      aufgeklappt.add(blockId);
      zeichne();
      aktualisiere();
      // Der Knopf, der den Fokus trug, ist weg — er landet im ersten Feld des
      // aufgeklappten Blocks, damit die Tastaturbedienung nicht abreisst.
      const block = zustand.bloecke.find((b) => b.id === blockId);
      const erstes = block === undefined ? null
        : document.querySelector('[data-zeile="' + block.zeilen[0].id + '"]');
      if (erstes !== null) erstes.focus();
      return;
    }

    if (e.target.id === 'zuruecksetzen') {
      if (!window.confirm('Alle Eingaben auf die Werte der Excel-Datei zurücksetzen?')) return;
      try { localStorage.removeItem(SCHLUESSEL); } catch { /* egal */ }
      zustand = vorgabe();
      zeichne();
      aktualisiere();
      return;
    }

    if (e.target.id === 'jetzt-speichern') {
      clearTimeout(speicherUhr);
      await speichernJetzt();
      window.print();
      return;
    }

    if (e.target.id === 'konto-abmelden') {
      sb.auth.signOut();
      return;
    }

    if (e.target.id === 'passwort-aendern') {
      document.getElementById('app-inhalt').hidden = true;
      document.getElementById('tor').hidden = false;
      const platz = document.getElementById('konto-platz');
      if (platz !== null) platz.innerHTML = wiederherstellungHtml();
      return;
    }

    if (e.target.id === 'passwort-form-abbrechen') {
      zeichneKonto();
      return;
    }

    if (e.target.id === 'konto-modus-wechsel') {
      formModus = formModus === 'registrieren' ? 'anmelden' : 'registrieren';
      const platz = document.getElementById('konto-platz');
      if (platz !== null) platz.innerHTML = torFormularHtml();
      return;
    }

    if (e.target.id === 'konto-passwort-vergessen') {
      const email = document.getElementById('konto-email').value.trim();
      const fehlerfeld = document.getElementById('konto-fehler');
      if (email === '') { fehlerfeld.textContent = 'Zuerst E-Mail eingeben.'; return; }
      e.target.disabled = true;
      sb.auth.resetPasswordForEmail(email, {
        redirectTo: window.location.origin + window.location.pathname,
      }).then(({ error }) => {
        if (error) {
          console.error(error);
          fehlerfeld.textContent = 'Link konnte nicht verschickt werden.';
          e.target.disabled = false;
          return;
        }
        fehlerfeld.textContent = 'Link zum Passwort-Setzen wurde verschickt.';
      });
    }
  });

  document.body.addEventListener('submit', async (e) => {
    if (e.target.id === 'konto-neues-passwort-formular') {
      e.preventDefault();
      const neuesPasswort = document.getElementById('neues-passwort').value;
      const fehlerfeld = document.getElementById('konto-fehler');
      const knopf = e.target.querySelector('button');
      knopf.disabled = true;
      const { error } = await sb.auth.updateUser({ password: neuesPasswort });
      if (error) {
        console.error(error);
        knopf.disabled = false;
        fehlerfeld.textContent = 'Passwort konnte nicht gespeichert werden.';
        return;
      }
      const { data: { session } } = await sb.auth.getSession();
      konto = session === null ? null : { id: session.user.id, email: session.user.email };
      zeichneKonto();
      if (konto !== null) await ladeVonServer();
      return;
    }

    if (e.target.id !== 'konto-formular') return;
    e.preventDefault();
    const email = document.getElementById('konto-email').value.trim();
    const passwort = document.getElementById('konto-passwort').value;
    const fehlerfeld = document.getElementById('konto-fehler');
    fehlerfeld.textContent = '';
    if (email === '' || passwort === '') return;

    const aktion = formModus;
    const knopf = e.target.querySelector('button[type="submit"]');
    knopf.disabled = true;
    const textVorher = knopf.textContent;
    knopf.textContent = aktion === 'registrieren' ? 'Registrierung …' : 'Anmeldung …';

    const { data, error } = aktion === 'registrieren'
      ? await sb.auth.signUp({ email, password: passwort })
      : await sb.auth.signInWithPassword({ email, password: passwort });

    if (error) {
      console.error(error);
      knopf.disabled = false;
      knopf.textContent = textVorher;
      fehlerfeld.textContent = kontoFehlertext(error, aktion);
      return;
    }

    if (aktion === 'registrieren' && data.session === null) {
      e.target.innerHTML = '<span class="konto-status">Fast fertig — Bestätigungslink in der E-Mail anklicken, dann geht\'s los.</span>';
    }
  });

  // Pfeiltasten im Zahlenfeld sollen nicht die Seite scrollen, sondern zählen.
  document.body.addEventListener('wheel', (e) => {
    if (document.activeElement === e.target && e.target.type === 'number') e.target.blur();
  }, { passive: true });
}

zeichne();
verdrahte();
aktualisiere();
initKonto();
