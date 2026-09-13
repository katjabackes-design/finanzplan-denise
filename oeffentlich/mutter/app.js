/*
 * Расчёт дохода — рудиментарная версия на основе «доход расчет.xlsx»,
 * лист «расчет услуг». Формула унифицирована для всех блоков (доход =
 * стоимость × человек × раз в месяц; часов = раз в месяц × время) —
 * в оригинале у йоги было особое /4*3, здесь его нет по договорённости.
 *
 * Архитектура 1:1 как в Finanzplan Denise: zeichne() строит разметку один
 * раз, aktualisiere() потом только переписывает ячейки результатов.
 */

'use strict';

const SCHLUESSEL = 'dohod-raschet-v1';

// Nur diese eine Adresse darf sich hier anmelden. Das ist eine Höflichkeitssperre,
// keine Sicherheitsgrenze — die eigentliche Absicherung ist Row-Level-Security in
// Supabase: selbst mit einer fremden E-Mail sähe niemand die Daten der Mutter.
const ERLAUBTE_EMAIL = 'katja.backes@posteo.ch';

/* ── Конто (Supabase) — тот же проект, что и у Finanzplan Denise ─────────── */
const SUPABASE_URL = 'https://kdlhplgrvehnccbzwiaj.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkbGhwbGdydmVobmNjYnp3aWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODkyMzg0OTcsImV4cCI6MjEwNDgxNDQ5N30.glfCkZ_PV7xzEefuLYyW9N6WR9ZKRpVd0SAk5YwOBj8';
const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
let konto = null;

function vorgabe() {
  return {
    szenario: 'norm',
    wochenProMonat: 4,
    zielUmsatz: 5000,
    kapazitaet: 20,
    bloecke: [
      {
        id: 'yoga', titel: 'йога занятие', zusatz: 'групповое занятие · цена за занятие',
        zeilen: [
          { id: 'y1', name: 'норм', szenario: 'norm', aktiv: true, c: 5,  e: 8,  g: 25, j: 2 },
          { id: 'y2', name: 'мин',  szenario: 'min',  aktiv: true, c: 3,  e: 4,  g: 20, j: 2 },
          { id: 'y3', name: 'мах',  szenario: 'max',  aktiv: true, c: 10, e: 12, g: 30, j: 2 },
        ],
      },
      {
        id: 'gvozdi', titel: 'гвозди', zusatz: 'цена за клиента',
        zeilen: [
          { id: 'g1', name: 'норм', szenario: 'norm', aktiv: true, c: 5,  e: 1, g: 65, j: 3 },
          { id: 'g2', name: 'мин',  szenario: 'min',  aktiv: true, c: 3,  e: 1, g: 60, j: 3 },
          { id: 'g3', name: 'мах',  szenario: 'max',  aktiv: true, c: 10, e: 2, g: 70, j: 3 },
        ],
      },
      {
        id: 'vyezd', titel: 'на выезде', zusatz: 'цена за клиента',
        zeilen: [
          { id: 'v1', name: 'норм', szenario: 'norm', aktiv: true, c: 10, e: 1, g: 90, j: 7 },
          { id: 'v2', name: 'мин',  szenario: 'min',  aktiv: true, c: 5,  e: 1, g: 70, j: 7 },
          { id: 'v3', name: 'мах',  szenario: 'max',  aktiv: true, c: 15, e: 1, g: 90, j: 7 },
        ],
      },
      {
        id: 'vstrechi', titel: 'встречи', zusatz: 'цена за клиента',
        zeilen: [
          { id: 't1', name: 'норм', szenario: 'norm', aktiv: true, c: 5,  e: 1, g: 40, j: 3 },
          { id: 't2', name: 'мин',  szenario: 'min',  aktiv: true, c: 3,  e: 1, g: 20, j: 3 },
          { id: 't3', name: 'мах',  szenario: 'max',  aktiv: true, c: 10, e: 1, g: 60, j: 3 },
        ],
      },
      {
        id: 'konsultacii', titel: 'консультации', zusatz: 'цена за консультацию',
        zeilen: [
          { id: 'k1', name: 'норм', szenario: 'norm', aktiv: true, c: 1, e: 10, g: 120, j: 2 },
          { id: 'k2', name: 'мин',  szenario: 'min',  aktiv: true, c: 1, e: 10, g: 70,  j: 2 },
          { id: 'k3', name: 'мах',  szenario: 'max',  aktiv: true, c: 1, e: 20, g: 300, j: 2 },
        ],
      },
      {
        id: 'mm', titel: 'мм', zusatz: 'цена за занятие',
        zeilen: [
          { id: 'm1', name: 'норм', szenario: 'norm', aktiv: true, c: 5,  e: 4, g: 50, j: 2 },
          { id: 'm2', name: 'мин',  szenario: 'min',  aktiv: true, c: 4,  e: 4, g: 30, j: 2 },
          { id: 'm3', name: 'мах',  szenario: 'max',  aktiv: true, c: 10, e: 4, g: 80, j: 2 },
        ],
      },
    ],
  };
}

/* ── Расчёт ─────────────────────────────────────────────────────────────── */

function rechne(z, wochen) {
  const umsatzMonat = z.c * z.e * z.g;
  const stundenMonat = z.e * z.j;
  const w = wochen > 0 ? wochen : 1;
  return {
    umsatzWoche: umsatzMonat / w,
    umsatzMonat,
    umsatzJahr: umsatzMonat * 12,
    stundenWoche: stundenMonat / w,
    stundenMonat,
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

function zaehlt(zeile, szenario) {
  if (zeile.aktiv !== true) return false;
  if (szenario === 'alle') return true;
  return zeile.szenario === szenario;
}

/* ── Форматы ────────────────────────────────────────────────────────────── */

const fGeld = new Intl.NumberFormat('de-CH', { maximumFractionDigits: 0 });
const fStd = new Intl.NumberFormat('de-CH', { maximumFractionDigits: 2 });

const geld = (n) => fGeld.format(Math.round(n));
const std = (n) => fStd.format(n);
const satz = (n) => (n === null ? '—' : fGeld.format(Math.round(n)));

function esc(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/* ── Состояние ──────────────────────────────────────────────────────────── */

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

  frisch.szenario = gespeichert.szenario ?? frisch.szenario;
  frisch.wochenProMonat = gespeichert.wochenProMonat ?? frisch.wochenProMonat;
  frisch.zielUmsatz = gespeichert.zielUmsatz ?? frisch.zielUmsatz;
  frisch.kapazitaet = gespeichert.kapazitaet ?? frisch.kapazitaet;

  for (const block of frisch.bloecke) {
    const alt = (gespeichert.bloecke ?? []).find((b) => b.id === block.id);
    if (alt === undefined) continue;
    for (const zeile of block.zeilen) {
      const altZ = (alt.zeilen ?? []).find((z) => z.id === zeile.id);
      if (altZ === undefined) continue;
      for (const k of ['c', 'e', 'g', 'j']) {
        if (typeof altZ[k] === 'number') zeile[k] = altZ[k];
      }
      if (typeof altZ.aktiv === 'boolean') zeile.aktiv = altZ.aktiv;
    }
  }
  return frisch;
}

let speicherUhr = null;
function speichern() {
  clearTimeout(speicherUhr);
  speicherUhr = setTimeout(async () => {
    try {
      localStorage.setItem(SCHLUESSEL, JSON.stringify(zustand));
    } catch { /* приватное окно — тогда без сохранения */ }

    const stand = document.getElementById('stand');
    if (konto !== null) {
      const { error } = await sb.from('dohod_daten').upsert({
        user_id: konto.id,
        daten: zustand,
        aktualisiert_am: new Date().toISOString(),
      });
      if (error) {
        console.error(error);
        stand.textContent = 'Ошибка сохранения';
        stand.className = 'abzeichen abzeichen-ernst';
        return;
      }
    }
    stand.textContent = 'сохранено';
    stand.className = 'abzeichen abzeichen-gut';
    setTimeout(() => { stand.className = 'abzeichen abzeichen-neutral'; }, 1200);
  }, 400);
}

/* ── Конто: вход, выход, загрузка с сервера ────────────────────────────── */

function torFormularHtml() {
  return (
    '<form class="konto-formular" id="konto-formular">' +
      '<input type="email" id="konto-email" class="konto-eingabe" placeholder="твоя@почта.ru" autocomplete="email" required>' +
      '<input type="password" id="konto-passwort" class="konto-eingabe" placeholder="Пароль" autocomplete="current-password" minlength="6" required>' +
      '<div style="display:flex; gap:8px">' +
        '<button type="submit" data-aktion="anmelden" class="knopf knopf-primaer knopf-klein" style="flex:1">Войти</button>' +
        '<button type="submit" data-aktion="registrieren" class="knopf knopf-umriss knopf-klein" style="flex:1">Регистрация</button>' +
      '</div>' +
      '<p class="konto-status" id="konto-fehler"></p>' +
    '</form>'
  );
}

function kontoFehlertext(error, aktion) {
  const m = error.message || '';
  if (m.includes('already registered') || m.includes('already exists')) {
    return 'Этот адрес уже зарегистрирован — нажмите «Войти».';
  }
  if (m.includes('Invalid login credentials')) {
    return 'Неверный e-mail или пароль.';
  }
  if (m.includes('Password') || m.includes('password')) {
    return 'Пароль слишком короткий (минимум 6 символов).';
  }
  return aktion === 'registrieren' ? 'Регистрация не удалась.' : 'Вход не удался.';
}

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

async function ladeVonServer() {
  if (konto === null) return;
  const { data, error } = await sb.from('dohod_daten').select('daten').eq('user_id', konto.id).maybeSingle();
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

  sb.auth.onAuthStateChange((_ereignis, session) => {
    konto = session === null ? null : { id: session.user.id, email: session.user.email };
    zeichneKonto();
    if (konto !== null) ladeVonServer();
  });
}

/* ── Разметка ───────────────────────────────────────────────────────────── */

const FELDER = [
  { k: 'c', titel: 'Человек', schritt: '1' },
  { k: 'e', titel: 'Раз в<br>месяц', schritt: '1' },
  { k: 'g', titel: 'Стоимость', schritt: '5' },
  { k: 'j', titel: 'Время<br>час.', schritt: '0.5' },
];

const ERGEBNISSE = [
  { k: 'umsatzWoche', titel: 'Доход<br>в неделю', form: geld },
  { k: 'umsatzMonat', titel: 'Доход<br>в месяц', form: geld, stark: true },
  { k: 'umsatzJahr', titel: 'Доход<br>в год', form: geld },
  { k: 'stundenWoche', titel: 'Часов<br>в неделю', form: std },
  { k: 'stundensatz', titel: 'Цена<br>часа', form: satz },
];

const HAKEN = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 13l4.5 4.5L19 7"/></svg>';

function schalterHtml(zeile) {
  return (
    '<label class="schalter" title="Строка учитывается в итоге">' +
    '<input type="checkbox" data-schalter="' + zeile.id + '"' + (zeile.aktiv ? ' checked' : '') +
    ' aria-label="Учитывать ' + esc(zeile.name) + ' в итоге">' +
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
  const spalten = 1 + FELDER.length;

  let kopf = '<tr><th>Вариант</th>';
  FELDER.forEach((f) => { kopf += '<th class="rechts">' + f.titel + '</th>'; });
  ERGEBNISSE.forEach((e, i) => {
    kopf += '<th class="rechts' + (i === 0 ? ' grenze' : '') + '">' + e.titel + '</th>';
  });
  kopf += '</tr>';

  let koerper = '';
  for (const zeile of block.zeilen) {
    koerper += '<tr data-reihe="' + zeile.id + '">';
    koerper += '<td><span class="variante">' + schalterHtml(zeile) +
      '<span class="variantenname">' + esc(zeile.name) + '</span></span></td>';
    FELDER.forEach((f) => { koerper += eingabeHtml(zeile, f); });
    ERGEBNISSE.forEach((e, i) => {
      koerper += '<td class="ergebnis' + (i === 0 ? ' grenze' : '') + (e.stark ? ' stark' : '') +
        '" data-aus="' + zeile.id + '.' + e.k + '"></td>';
    });
    koerper += '</tr>';
  }

  let fuss = '<tr><td colspan="' + spalten + '">Промежуточный итог <span class="klein gedaempft">учтённые строки</span></td>';
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

function blockHtml(block) {
  return (
    '<article class="karte">' +
      '<div class="block-kopf">' +
        '<div>' +
          '<div class="block-titelzeile"><h3>' + esc(block.titel) + '</h3></div>' +
          '<p class="klein gedaempft" style="margin-top:4px">' + esc(block.zusatz) + '</p>' +
        '</div>' +
        '<div class="block-summe">' +
          '<span class="block-summe-wert" data-aus="' + block.id + '.kopfsumme"></span>' +
          '<span class="klein gedaempft">CHF / мес</span>' +
        '</div>' +
      '</div>' +
      tabellenBlockHtml(block) +
    '</article>'
  );
}

/* ── Показатели и диаграмма ─────────────────────────────────────────────── */

function kennzahlenHtml() {
  const karten = [
    { id: 'kUmsatzMonat', name: 'Доход в месяц', einheit: 'CHF', hero: true },
    { id: 'kUmsatzJahr', name: 'Доход в год', einheit: 'CHF' },
    { id: 'kStunden', name: 'Часов в неделю', einheit: 'ч.' },
    { id: 'kSatz', name: 'Средняя цена часа', einheit: 'CHF' },
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
  monat: { name: 'Доход в месяц', wert: (e) => e.umsatzMonat, form: geld },
  stunden: { name: 'Часов в неделю', wert: (e) => e.stundenWoche, form: std },
  satz: { name: 'Средняя цена часа', wert: (e) => e.stundensatz ?? 0, form: (n) => (n === 0 ? '—' : geld(n)) },
};

let mass = 'monat';

function zeichneDiagramm() {
  const ziel = document.getElementById('diagramm');
  const m = MASSE[mass];

  const reihen = zustand.bloecke
    .map((block) => {
      const teile = block.zeilen
        .filter((z) => zaehlt(z, zustand.szenario))
        .map((z) => rechne(z, zustand.wochenProMonat));
      return { name: block.titel, ergebnis: summiere(teile) };
    })
    .map((r) => ({ name: r.name, wert: m.wert(r.ergebnis) }))
    .filter((r) => r.wert > 0)
    .sort((a, b) => b.wert - a.wert);

  if (reihen.length === 0) {
    ziel.innerHTML =
      '<div class="leerzustand"><span class="leerzustand-titel">Нечего показать</span>' +
      '<span class="klein">Для этого сценария нет строк со значением больше нуля.</span></div>';
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
          '<span class="balkenwert">' + m.form(r.wert) + '</span>' +
        '</div>',
    )
    .join('');
}

/* ── Обновление ─────────────────────────────────────────────────────────── */

const aus = {};

function setze(schluessel, text) {
  const knoten = aus[schluessel];
  if (knoten !== undefined) knoten.textContent = text;
}

function stufe(anteil, umgekehrt) {
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
      const e = rechne(zeile, w);
      const zaehltMit = zaehlt(zeile, zustand.szenario);
      if (zaehltMit) {
        gezaehlt += 1;
        blockTeile.push(e);
        gesamtTeile.push(e);
      }

      ERGEBNISSE.forEach((sp) => setze(zeile.id + '.' + sp.k, sp.form(e[sp.k])));

      const reihe = document.querySelector('[data-reihe="' + zeile.id + '"]');
      if (reihe !== null) reihe.classList.toggle('zeile-aus', !zaehltMit);
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
    std(total.stundenWoche) + '<span class="kennzahl-einheit">ч.</span>';
  document.querySelector('[data-aus="kSatz"]').innerHTML =
    satz(total.stundensatz) + '<span class="kennzahl-einheit">CHF</span>';

  const szenarioWort = zustand.szenario === 'alle' ? 'все сценарии' : 'сценарий ' + zustand.szenario;
  setze('kUmsatzMonat.zusatz', geld(total.umsatzWoche) + ' в неделю · ' + szenarioWort);
  setze('kUmsatzJahr.zusatz', '12 × доход в месяц');
  setze('kStunden.zusatz', std(total.stundenMonat) + ' ч. в месяц');
  setze('kSatz.zusatz', gezaehlt + ' из ' + vorhanden + ' строк учтено');

  /* Цель */
  const zielAnteil = zustand.zielUmsatz > 0 ? total.umsatzMonat / zustand.zielUmsatz : 1;
  const zielStufe = stufe(zielAnteil, false);
  const zielBalken = document.getElementById('ziel-balken');
  zielBalken.style.width = Math.min(100, zielAnteil * 100) + '%';
  zielBalken.parentElement.className = 'messbalken messbalken-' + zielStufe;
  document.getElementById('ziel-wert').textContent = geld(zustand.zielUmsatz) + ' CHF';
  const luecke = zustand.zielUmsatz - total.umsatzMonat;
  document.getElementById('ziel-text').innerHTML =
    luecke <= 0
      ? '<span class="abzeichen abzeichen-gut">Цель достигнута</span> ' + geld(-luecke) + ' CHF сверх.'
      : '<span class="abzeichen abzeichen-' + zielStufe + '">' + Math.round(zielAnteil * 100) + '&nbsp;% достигнуто</span> ' +
        'Не хватает ' + geld(luecke) + ' CHF в месяц.';

  /* Загрузка */
  const auslastung = zustand.kapazitaet > 0 ? total.stundenWoche / zustand.kapazitaet : 0;
  const kapStufe = stufe(auslastung, true);
  const kapBalken = document.getElementById('kapazitaet-balken');
  kapBalken.style.width = Math.min(100, auslastung * 100) + '%';
  kapBalken.parentElement.className = 'messbalken messbalken-' + kapStufe;
  document.getElementById('kapazitaet-wert').textContent = std(zustand.kapazitaet) + ' ч.';
  const rest = zustand.kapazitaet - total.stundenWoche;
  document.getElementById('kapazitaet-text').innerHTML =
    '<span class="abzeichen abzeichen-' + kapStufe + '">' + Math.round(auslastung * 100) + '&nbsp;% загружено</span> ' +
    (rest >= 0 ? std(rest) + ' ч. в неделю свободно.' : std(-rest) + ' ч. в неделю сверх загрузки.');

  /* Основные показатели */
  document.getElementById('eckwerte').innerHTML =
    '<div><dt>Доход в неделю</dt><dd>' + geld(total.umsatzWoche) + ' CHF</dd></div>' +
    '<div><dt>Часов в месяц</dt><dd>' + std(total.stundenMonat) + ' ч.</dd></div>' +
    '<div><dt>Доход за рабочий день</dt><dd>' + geld(total.umsatzWoche / 5) + ' CHF</dd></div>' +
    '<div><dt>Учтено строк</dt><dd>' + gezaehlt + ' / ' + vorhanden + '</dd></div>';

  zeichneDiagramm();
}

/* ── Соединение обработчиков ────────────────────────────────────────────── */

function findeZeile(id) {
  for (const block of zustand.bloecke) {
    const zeile = block.zeilen.find((z) => z.id === id);
    if (zeile !== undefined) return zeile;
  }
  return null;
}

function zeichne() {
  document.getElementById('kennzahlen').innerHTML = kennzahlenHtml();
  document.getElementById('bloecke').innerHTML = zustand.bloecke.map(blockHtml).join('');

  for (const k of Object.keys(aus)) delete aus[k];
  document.querySelectorAll('[data-aus]').forEach((knoten) => { aus[knoten.dataset.aus] = knoten; });

  document.getElementById('ziel').value = String(zustand.zielUmsatz);
  document.getElementById('kapazitaet').value = String(zustand.kapazitaet);
  document.getElementById('wochen').value = String(zustand.wochenProMonat);

  document.querySelectorAll('[data-szenario]').forEach((knopf) => {
    knopf.setAttribute('aria-pressed', String(knopf.dataset.szenario === zustand.szenario));
  });
  document.querySelectorAll('[data-mass]').forEach((knopf) => {
    knopf.setAttribute('aria-pressed', String(knopf.dataset.mass === mass));
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
      zeile[ziel.dataset.feld] = zahl;
      ziel.classList.toggle('leer', zahl === 0);
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
    const id = e.target.dataset.schalter;
    if (id === undefined) return;
    const zeile = findeZeile(id);
    if (zeile === null) return;
    zeile.aktiv = e.target.checked;
    aktualisiere();
    speichern();
  });

  document.body.addEventListener('click', (e) => {
    const knopf = e.target.closest('[data-szenario], [data-mass]');
    if (knopf !== null) {
      if (knopf.dataset.szenario !== undefined) {
        zustand.szenario = knopf.dataset.szenario;
        document.querySelectorAll('[data-szenario]').forEach((k) => {
          k.setAttribute('aria-pressed', String(k.dataset.szenario === zustand.szenario));
        });
      } else {
        mass = knopf.dataset.mass;
        document.querySelectorAll('[data-mass]').forEach((k) => {
          k.setAttribute('aria-pressed', String(k.dataset.mass === mass));
        });
      }
      aktualisiere();
      speichern();
      return;
    }

    if (e.target.id === 'zuruecksetzen') {
      if (!window.confirm('Сбросить все значения к исходным?')) return;
      try { localStorage.removeItem(SCHLUESSEL); } catch { /* неважно */ }
      zustand = vorgabe();
      zeichne();
      aktualisiere();
      return;
    }

    if (e.target.id === 'konto-abmelden') {
      sb.auth.signOut();
    }
  });

  document.body.addEventListener('submit', async (e) => {
    if (e.target.id !== 'konto-formular') return;
    e.preventDefault();
    const email = document.getElementById('konto-email').value.trim();
    const passwort = document.getElementById('konto-passwort').value;
    const fehlerfeld = document.getElementById('konto-fehler');
    fehlerfeld.textContent = '';
    if (email === '' || passwort === '') return;
    if (email.toLowerCase() !== ERLAUBTE_EMAIL) {
      fehlerfeld.textContent = 'Доступ есть только по одному определённому адресу.';
      return;
    }

    const aktion = e.submitter && e.submitter.dataset.aktion === 'registrieren' ? 'registrieren' : 'anmelden';
    const knoepfe = e.target.querySelectorAll('button');
    knoepfe.forEach((k) => { k.disabled = true; });
    const aktivKnopf = e.submitter;
    const textVorher = aktivKnopf.textContent;
    aktivKnopf.textContent = aktion === 'registrieren' ? 'Регистрация …' : 'Вход …';

    const { data, error } = aktion === 'registrieren'
      ? await sb.auth.signUp({ email, password: passwort })
      : await sb.auth.signInWithPassword({ email, password: passwort });

    if (error) {
      console.error(error);
      knoepfe.forEach((k) => { k.disabled = false; });
      aktivKnopf.textContent = textVorher;
      fehlerfeld.textContent = kontoFehlertext(error, aktion);
      return;
    }

    // Bestätigungsmail nötig? Dann gibt es noch keine Sitzung — sonst
    // übernimmt onAuthStateChange sofort und blendet das Tor aus.
    if (aktion === 'registrieren' && data.session === null) {
      e.target.innerHTML = '<span class="konto-status">Почти готово — подтвердите адрес по ссылке в письме, потом можно войти.</span>';
    }
  });

  document.body.addEventListener('wheel', (e) => {
    if (document.activeElement === e.target && e.target.type === 'number') e.target.blur();
  }, { passive: true });
}

zeichne();
verdrahte();
aktualisiere();
initKonto();
