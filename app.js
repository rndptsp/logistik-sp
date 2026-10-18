/* Outbound Logistics — Semen Padang.
   Data: data/site.enc (AES-GCM, gzip'd JSON) built by tools/build_site.py from the master Excel.
   Rules (same as the previous logistik-sp dashboard, now logistik-sp-lama): target = SNOP (or Internal = 115% SNOP, filter); Target MTD = daily D1..D<last data
   day>; achievement % = realisasi / target of the selected incoterm(s): FRC targets per ekspeditur, FOT/FOB/FAS per distributor;
   default view FRC from CP INDARUNG. */
'use strict';

/* ===================== Small helpers ===================== */
const $ = (s, el) => (el || document).querySelector(s);
const $$ = (s, el) => [...(el || document).querySelectorAll(s)];
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => (n == null || isNaN(n)) ? '–' : Math.round(n).toLocaleString('id-ID');
const fmt1 = n => (n == null || isNaN(n)) ? '–' : n.toLocaleString('id-ID', {minimumFractionDigits:1, maximumFractionDigits:1});
const pctOf = (a, b) => b > 0 ? a / b * 100 : null;
const pctTxt = p => p == null ? '–' : Math.round(p) + '%';
const signed = n => (n >= 0 ? '+' : '−') + fmt(Math.abs(n));
const enc = encodeURIComponent;
const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const MON3 = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const DAYS = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const monthLabel = m => MONTHS[+m.slice(5) - 1] + ' ' + m.slice(0, 4);
const monthShort = m => MON3[+m.slice(5) - 1] + ' ' + m.slice(2, 4);
const dim = m => new Date(+m.slice(0, 4), +m.slice(5), 0).getDate();
const iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const parseISO = s => new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
const dayLabel = s => { const d = parseISO(s); return d.getDate() + ' ' + MON3[d.getMonth()] + ' ' + d.getFullYear(); };
/* Source district names are truncated to 20 chars in SAP; show the full name (same list as the old dashboard's DISTRIK_ALIAS). */
const DIST_FULL = {
  'KAB. HUMBANG HASUNDU': 'KAB. HUMBANG HASUNDUTAN', 'KAB. KUANTAN SINGING': 'KAB. KUANTAN SINGINGI',
  'KAB. MANDAILING NATA': 'KAB. MANDAILING NATAL', 'KAB. TANJ JABUNG TIM': 'KAB. TANJUNG JABUNG TIMUR',
  'KAB. TANJ JABUNG BRT': 'KAB. TANJUNG JABUNG BARAT', 'KAB.TAPANULI SELATAN': 'KAB. TAPANULI SELATAN',
  'KOTA PADANG SIDEMPUA': 'KOTA PADANG SIDEMPUAN', 'LABUHANBATU SELATAN': 'KAB. LABUHANBATU SELATAN'
};
const KEEP_UPPER = /\b(Pt|Cv|Sba|Sp|Cp|Gp|Pp|Smbr|Silog|Frc|Fot|So|Snop|Ud|Tbk)\b/g;
function tc(s){
  if(!s) return '—';
  s = DIST_FULL[s] || s;
  return s.toLowerCase().replace(/(^|[\s(.\/-])([a-z])/g, (m, a, b) => a + b.toUpperCase()).replace(KEEP_UPPER, w => w.toUpperCase());
}
function tierColor(p){ return p == null ? 'var(--sub)' : p >= 100 ? 'var(--green)' : p >= 85 ? 'var(--amber)' : 'var(--red)'; }
function tierClass(p){ return p == null ? '' : p >= 100 ? 'good' : p >= 85 ? 'mid' : 'bad'; }
function tierName(p){ return p >= 100 ? 'Green' : p >= 85 ? 'Amber' : 'Red'; }
function reviewText(p){
  if(p == null || isNaN(p)) return 'Tidak ada target ' + tgtLbl() + ' untuk cakupan & periode ini.';
  if(p >= 100) return 'Kinerja solid — di atas target ' + tgtLbl() + ' bulan berjalan.';
  if(p >= 85) return 'Mendekati target — perlu dorongan kecil di sisa bulan.';
  if(p >= 70) return 'Tertinggal dari target — perlu perhatian dan tindak lanjut.';
  return 'Tertinggal signifikan — perlu intervensi segera.';
}

const ICONS = {
  truck:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="7" width="13" height="9" rx="1"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/></svg>',
  clock:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  check:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.5 2.5L16 9"/></svg>',
  box:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7l9-4 9 4-9 4-9-4z"/><path d="M3 7v10l9 4 9-4V7"/><path d="M12 11v10"/></svg>',
  pin:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s7-7.1 7-12a7 7 0 10-14 0c0 4.9 7 12 7 12z"/><circle cx="12" cy="10" r="2.4"/></svg>',
  trend:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 17l6-6 4 4 8-9"/><path d="M15 6h6v6"/></svg>',
  target:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.5"/></svg>',
  gap:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 12h16"/><path d="M8 8l-4 4 4 4"/><path d="M16 8l4 4-4 4"/></svg>'
};
function kpi(icon, color, label, val, opt){
  opt = opt || {};
  const tag = opt.href ? 'a' : 'div';
  return '<' + tag + ' class="kpi"' + (opt.href ? ' href="' + opt.href + '"' : '') + '><div class="kpi-ic" style="color:' + color + '">' + ICONS[icon] + '</div><div><div class="lbl">' + label + '</div><div class="val">' + val + '</div>' + (opt.hint ? '<div class="hint">' + opt.hint + '</div>' : '') + '</div></' + tag + '>';
}
/** Bar row for achievement vs target; pass `of` (Real/SO %, may be null) to add the Real/SO value next to it. */
function rateRow(label, p, href, sub, of){
  const w = p == null ? 0 : Math.min(100, p);
  const tag = href ? 'a' : 'div';
  return '<' + tag + ' class="rate-row"' + (href ? ' href="' + href + '"' : '') + ' title="' + esc(label + (sub ? ' — ' + sub : '')) + '"><div class="rl">' + esc(label) + '</div><div class="rtrack"><div class="rfill" style="width:' + w + '%;background:' + tierColor(p) + '"></div></div><div class="rv" style="color:' + tierColor(p) + '">' + pctTxt(p) + '</div>' +
    (of !== undefined ? '<div class="rso"><small>SO</small> ' + (of == null ? '–' : pctTxt(of)) + '</div>' : '') + '</' + tag + '>';
}
/** Tooltip line for Real/SO, or '' when there is no SO. */
const soTip = o => o.so > 0 ? '|SO ' + fmt(o.so) + ' t · Real/SO ' + pctTxt(o.of) : (soOn() ? '|SO tidak tersedia' : '');

/* ===================== State ===================== */
let B = null;            // decrypted bundle
let D = null;            // derived indexes
const F = { per: null, inc: 'ALL', moda: 'ALL', area: 'ALL', srcOff: new Set(), tgt: 'SNOP', kem: 'ZAK' };   // per: {kind:'month',month} | {kind:'ytd'} | {kind:'range',from,to}
const DEFAULT_SRC = 'CP INDARUNG';   // the main plant (labels only); default view = every plant, moda & incoterm, kemasan ZAK, SNOP, area Semua
let DEFAULT_SRC_OFF = [];

function prepare(){
  const dims = B.dims;
  const dayMonth = B.days.map(s => s.slice(0, 7));
  const dayNum = B.days.map(s => +s.slice(8, 10));
  const dayDow = B.days.map(s => parseISO(s).getDay());
  const months = [...new Set(dayMonth)].sort();
  const provIdx = Object.fromEntries(dims.prov.map((p, i) => [p, i]));
  const ekspIdx = Object.fromEntries(dims.eksp.map((p, i) => [p, i]));
  const distIdx = Object.fromEntries(dims.dist.map((p, i) => [p, i]));
  const srcEmpty = dims.src.indexOf('');
  const eksEmpty = dims.eksp.indexOf('');
  const frcEksp = new Set(), liveProv = new Set(), liveDist = new Set();
  B.facts.forEach(f => { if(isFrc(f[1]) && f[5] !== eksEmpty) frcEksp.add(f[5]); liveProv.add(f[3]); liveDist.add(f[4]); });
  const tanpaEks = dims.eksp.indexOf('TANPA_EKSPEDITUR');
  B.targets.forEach(t => { if(isFrc(t[7]) && t[3] !== eksEmpty && t[3] !== tanpaEks) frcEksp.add(t[3]); });
  const distrIdx = Object.fromEntries(dims.distr.map((p, i) => [p, i]));
  const sources = [...new Set(B.facts.map(f => f[2]).concat(B.targets.map(t => t[4])))].filter(i => i !== srcEmpty).map(i => dims.src[i]).sort();
  D = { dayMonth, dayNum, dayDow, months, provIdx, ekspIdx, distIdx, srcEmpty, eksEmpty,
        ekspList: [...frcEksp].map(i => dims.eksp[i]).sort(), frcEksp, distrIdx, liveProv, liveDist, sources };
  F.per = {kind: 'month', month: months[months.length - 1]};
  D.tgtKem = new Set(B.targets.map(tKem));
  D.nfTgt = B.targets.some(t => t[7] >= 1);   // builds without sheet SNOP: FRC targets only
  DEFAULT_SRC_OFF = [];   // all plants on by default
  F.srcOff = new Set(DEFAULT_SRC_OFF);
}
const srcOk = i => i === D.srcEmpty || !F.srcOff.has(B.dims.src[i]);
const srcOkS = (i, sc) => sc && sc.src != null ? i === sc.src : srcOk(i);   // a plant page (sc.src) ignores the plant filter
/* Target basis: master holds SNOP only; Internal = 115% of SNOP. Applied wherever a target is summed. */
const TGT_INTERNAL = 1.15;
/* SLA (truck in -> out of the plant): target 4 hours, the same for every truck. */
const SLA_TARGET = 4;
const slaColor = v => v == null ? 'var(--sub)' : Math.round(v * 10) / 10 <= SLA_TARGET ? 'var(--green)' : 'var(--red)';   // judged on the shown (1-decimal) value
const tgtK = () => F.tgt === 'INTERNAL' ? TGT_INTERNAL : 1;
const tgtLbl = () => F.tgt === 'INTERNAL' ? 'Internal' : 'SNOP';
/* Incoterm codes: B.dims.inc (FRC, FOT, FOB, FAS); older builds have no list and only 0 = FRC, 1 = FOT. */
const INCS = () => B.dims.inc || ['FRC', 'FOT'];
/* Kemasan (ZAK / CURAH / JUMBO BAG): last column of facts, ship, trucks and sod rows. Default ZAK = same basis as the target (SNOP is zak only);
   targets are ignored when a non-zak kemasan is selected. Builds without B.dims.kem behave as before. */
const KEMS = () => B.dims.kem || [];
const kemOk = k => F.kem === 'ALL' || !KEMS().length || KEMS()[k] === F.kem;
/* Target kemasan: t[9]; builds without it = ZAK. Achievement counts realisasi only in kemasan that have targets. */
const tKem = t => t[9] == null ? Math.max(0, KEMS().indexOf('ZAK')) : t[9];
const kemName = () => F.kem === 'ALL' ? 'semua kemasan' : F.kem;
/* Moda: DARAT = shipments by truck (FRC, FOT); LAUT = shipments by sea (FOB, FAS, FRC KAPAL/KONTAINER, FOT KONTAINER, FOB KAPAL).
   The trucks that haul sea cargo to the port belong to LAUT as well, so Armada (darat) never mixes them in. */
const LAUT_INC = new Set(['FRC CONTAINER', 'FRC KAPAL', 'FOT CONTAINER', 'FOB', 'FAS', 'FOB KAPAL']);
const modaOf = n => LAUT_INC.has(n) ? 'LAUT' : 'DARAT';
const modaIncs = m => INCS().filter(n => m === 'ALL' || modaOf(n) === m);
const MODA_LBL = {DARAT: 'Darat', LAUT: 'Laut', ALL: 'Semua'};
const incOk = inc => { const n = INCS()[inc]; return F.inc === 'ALL' ? (F.moda === 'ALL' || modaOf(n) === F.moda) : n === F.inc; };
/* Franco family: FRC (truk), FRC CONTAINER, FRC KAPAL - each its own incoterm; only FRC TRUK has SO (Real/SO).
   The rest is picked up by the distributor. */
const INC_LBL = {'FRC': 'FRC TRUK', 'FRC CONTAINER': 'FRC KONTAINER', 'FRC KAPAL': 'FRC KAPAL', 'FOT CONTAINER': 'FOT KONTAINER'};
const incLbl = x => INC_LBL[x] || x;
const isFrc = inc => /^FRC/.test(INCS()[inc] || '');
const frcSel = () => /^FRC/.test(F.inc);
const nfOn = () => !frcSel() && F.inc !== 'ALL';   // a non-franco incoterm (distributor pick-up) is selected
const lastDayOf = m => Math.min(B.lastDay[m] || dim(m), dim(m));
/* Targets: FRC rows (ekspeditur) and FOT/FOB/FAS rows (distributor, t[8]); t[7] = incoterm, absent on older builds = FRC.
   Achievement = realisasi of the selected incoterm(s) vs their targets; Real/SO stays FRC-only (SO is FRC). */
const tInc = t => t[7] || 0;
const achOk = inc => incOk(inc) && (isFrc(inc) || D.nfTgt);   // realisasi counted against target
/* Only realisasi in a kemasan that has targets counts toward achievement (volume still shows the chosen kemasan). */
const achKem = k => !KEMS().length || D.tgtKem.has(k);
const soOn = () => F.inc === 'FRC' || (F.inc === 'ALL' && F.moda !== 'LAUT');   // SO is FRC TRUK only
const incName = () => F.inc === 'ALL' ? 'semua incoterm' + (F.moda === 'ALL' ? '' : ' ' + MODA_LBL[F.moda].toLowerCase()) : incLbl(F.inc);
const setInc = (inc, moda) => { F.inc = inc; F.moda = moda || (inc === 'ALL' ? F.moda : modaOf(inc)); };
const saveFilter = () => { try { sessionStorage.setItem('osp_f3', JSON.stringify({per: F.per, inc: F.inc, moda: F.moda, kem: F.kem, area: F.area, srcOff: [...F.srcOff], tgt: F.tgt})); } catch(e) {} };

/* ===================== Periods ===================== */
/* Every page works on a period {from, to} (ISO dates). A "month" is just one kind of period.
   Target always runs only to the last day with data (tgtTo), so a period reaching into the future stays fair. */
const pad2 = n => String(n).padStart(2, '0');
function addDays(s, n){ const d = parseISO(s); d.setDate(d.getDate() + n); return iso(d); }
const shortDate = s => { const d = parseISO(s); return d.getDate() + ' ' + MON3[d.getMonth()] + (s.slice(0, 4) !== B.asOf.slice(0, 4) ? ' ' + s.slice(0, 4) : ''); };
function makePer(kind, from, to, extra){
  if(to > B.asOf) to = B.asOf;
  if(from < B.days[0]) from = B.days[0];
  if(from > to) from = to;
  const isLatest = to === B.asOf, days = Math.round((parseISO(to) - parseISO(from)) / 864e5) + 1;
  let label, sub;
  if(kind === 'month'){
    const m = from.slice(0, 7), full = +to.slice(8) >= dim(m);
    label = monthLabel(m); sub = label + (full ? '' : ' · MTD s.d. ' + shortDate(to));
  } else if(kind === 'ytd'){
    label = 'YTD ' + to.slice(0, 4); sub = 'Tahun berjalan · ' + shortDate(from) + ' – ' + shortDate(to);
  } else {
    label = shortDate(from) + ' – ' + shortDate(to); sub = label + ' (' + days + ' hari)';
  }
  return Object.assign({kind, from, to, tgtTo: to, label, sub, isLatest, days, month: to.slice(0, 7)}, extra || {});
}
function monthPer(m, toDay){ return makePer('month', m + '-01', m + '-' + pad2(toDay != null ? Math.min(toDay, dim(m)) : lastDayOf(m))); }
function PER(){
  const p = F.per || {kind: 'month', month: D.months[D.months.length - 1]};
  if(p.kind === 'ytd') return makePer('ytd', B.asOf.slice(0, 4) + '-01-01', B.asOf);
  if(p.kind === 'range') return makePer('range', p.from, p.to);
  return monthPer(p.month);
}
const toPer = x => typeof x === 'string' ? monthPer(x) : x;
/* Previous comparable period: month -> previous month (same days when running); otherwise the same number of days just before. */
function prevPer(per){
  if(per.kind === 'month'){ const pm = prevMonth(per.month); const full = +per.to.slice(8) >= dim(per.month); return monthPer(pm, full ? dim(pm) : +per.to.slice(8)); }
  return makePer('range', addDays(per.from, -per.days), addDays(per.from, -1));
}
const inPer = (dayIdx, per) => { const s = B.days[dayIdx]; return s >= per.from && s <= per.to; };

/* ===================== Aggregation core ===================== */
/* Group key extractors work on the shared column layout: facts/targets/so all carry prov,dist,eksp
   (facts: [day,inc,src,prov,dist,eksp,...]; targets: [month,prov,dist,eksp,src,...]; so: [month,prov,dist,eksp,ton]). */
/* Provinsi non-kinerja (mis. Sumatera Selatan, hanya dari CP Indarung): ditampilkan di baris provinsi/distriknya sendiri
   (target 0 = 100%) tetapi tidak masuk total, ekspeditur, source, armada maupun capaian keseluruhan. */
const NKP = () => B._nk || (B._nk = new Set(B.dims.nonKin || []));
/* Area: AP1 = the provinces in the daily Realisasi report (all except the non-kinerja ones, e.g. Sumatera Selatan);
   NON AP1 = the non-kinerja provinces (shown without target). Default Semua (both); AP1 matches the daily report. A province's own
   page always shows its data. Non-AP1 realisasi never counts toward achievement. */
const areaOn = () => NKP().size > 0;
const areaProvOk = p => !areaOn() || F.area === 'ALL' || (F.area === 'AP1') !== NKP().has(p);
const areaName = () => F.area === 'ALL' ? 'AP1 + NON AP1' : F.area === 'AP1' ? 'AP1' : 'NON AP1';
const nkSkip = (p, group, sc) => !areaProvOk(p) && (sc || {}).prov !== p;
const FK = {prov: f => f[3], dist: f => f[4], eksp: f => f[5], src: f => f[2], all: () => 0};
const TK = {prov: t => t[1], dist: t => t[2], eksp: t => t[3], src: t => t[4], all: () => 0};
const SK = {prov: s => s[1], dist: s => s[2], eksp: s => s[3], src: () => null, all: () => 0};
function scopeOkF(f, sc){ return (sc.src == null || f[2] === sc.src) && (sc.prov == null || f[3] === sc.prov) && (sc.dist == null || f[4] === sc.dist) && (sc.eksp == null || f[5] === sc.eksp); }
function scopeOkT(t, sc){ return (sc.src == null || t[4] === sc.src) && (sc.prov == null || t[1] === sc.prov) && (sc.dist == null || t[2] === sc.dist) && (sc.eksp == null || t[3] === sc.eksp); }

/* Target of one SNOP row between two ISO dates (daily values when present, else monthly ÷ days). */
function targetBetween(t, from, to){
  const m = t[0], n = dim(m), a = m + '-01', b = m + '-' + pad2(n);
  if(b < from || a > to) return 0;
  const d0 = from > a ? +from.slice(8) : 1, d1 = to < b ? +to.slice(8) : n;
  if(t[6]){ let s = 0; for(let d = d0; d <= d1; d++) s += t[6][d - 1] || 0; return s * tgtK(); }
  return t[5] * (d1 - d0 + 1) / n * tgtK();
}
function blank(){ return {vol:0, frc:0, frc0:0, trips:0, dwS:0, dwN:0, tgt:0, tgtFull:0, so:0, hasTgt:false}; }

/** Summary per group for a period (or a 'YYYY-MM' month). vol = selected incoterm(s); frc = realisasi of the selected incoterm(s) vs their targets (achievement basis); frc0 = FRC only (Real/SO).
    tgt = target up to the last data day; tgtFull = target for the whole period (whole month for a month period). */
function summarize(group, per, sc, opt){
  sc = sc || {}; opt = opt || {};
  per = opt.toDay != null && typeof per === 'string' ? monthPer(per, opt.toDay) : toPer(per);
  const fullTo = per.kind === 'month' ? per.month + '-' + pad2(dim(per.month)) : per.to;
  const out = new Map();
  const get = k => { let o = out.get(k); if(!o){ o = blank(); o.key = k; out.set(k, o); } return o; };
  B.facts.forEach(f => {
    if(!inPer(f[0], per) || !srcOkS(f[2], sc) || !kemOk(f[10]) || !scopeOkF(f, sc) || nkSkip(f[3], group, sc)) return;
    const o = get(FK[group](f));
    if(f[1] === 0) o.frc0 += f[6];   // FRC TRUK (Real/SO)
    if(achOk(f[1]) && achKem(f[10]) && !NKP().has(f[3])) o.frc += f[6];
    if(incOk(f[1])){ o.vol += f[6]; o.trips += f[7]; o.dwS += f[8]; o.dwN += f[9]; }
  });
  const m0 = per.from.slice(0, 7), m1 = fullTo.slice(0, 7);
  B.targets.forEach(t => {
    if(!kemOk(tKem(t)) || t[0] < m0 || t[0] > m1 || !incOk(tInc(t)) || !srcOkS(t[4], sc) || !scopeOkT(t, sc) || nkSkip(t[1], group, sc)) return;
    const o = get(TK[group](t));
    o.tgt += targetBetween(t, per.from, per.tgtTo); o.tgtFull += targetBetween(t, per.from, fullTo); o.hasTgt = true;
  });
  if(soOn() && sc.src == null){   // SO has no plant: not shown for a plant scope
    // SO counts every month the period touches in full: by delivery month (TGL_KIRIM, B.sod) when the build has it,
    // otherwise by PERIODE (B.so)
    if(group !== 'src'){
      const m9 = per.to.slice(0, 7);
      B.sod.forEach(s => {
        if(s[0].slice(0, 7) < m0 || s[0].slice(0, 7) > m9 || !kemOk(s[5]) || !scopeOkT(s, sc) || nkSkip(s[1], group, sc)) return;
        get(SK[group](s)).so += s[4];
      });
      B.so.forEach(s => {
        if(s[0] < m0 || s[0] > per.to.slice(0, 7) || B.sodMonths.has(s[0]) || !scopeOkT(s, sc) || nkSkip(s[1], group, sc)) return;
        get(SK[group](s)).so += s[4];
      });
    }
  }
  out.forEach(o => {
    o.pct = o.hasTgt ? (o.tgt > 0 ? pctOf(o.frc, o.tgt) : 100) : (o.vol > 0 || o.frc > 0 ? 100 : null);   // no target but shipped = 100% (Rendy)
    o.of = soOn() && o.so > 0 ? pctOf(o.frc0, o.so) : null;
    o.gap = o.frc - o.tgt;
    o.dwell = o.dwN ? o.dwS / o.dwN : null;
  });
  return out;
}
const total = (per, sc, opt) => summarize('all', per, sc, opt).get(0) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
function ranked(map){
  const arr = [...map.values()].filter(o => o.vol > 0 || o.tgt > 0);
  arr.sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  let r = 0; arr.forEach(o => { o.rank = o.pct == null ? null : ++r; });
  arr.rankTotal = r;
  return arr;
}
/* B.trucks rows: [day, truck, inc, eksp, src, prov, trips, ton, dwellSum, dwellN] */
function trucksFor(per, sc){
  per = toPer(per); sc = sc || {};
  const out = new Map();
  B.trucks.forEach(t => {
    if(!inPer(t[0], per) || !incOk(t[2]) || !srcOkS(t[4], sc) || !kemOk(t[10])) return;
    if(sc.eksp != null && t[3] !== sc.eksp) return;
    if(sc.prov != null && t[5] !== sc.prov) return;
    if(sc.prov == null && !areaProvOk(t[5])) return;
    let o = out.get(t[1]);
    if(!o){ o = {truck:t[1], trips:0, ton:0, dwS:0, dwN:0, first:t[0], last:t[0], daySet:new Set(), eksp:new Map(), provTon:new Map(), inc:t[2]}; out.set(t[1], o); }
    o.trips += t[6]; o.ton += t[7]; o.dwS += t[8]; o.dwN += t[9]; o.first = Math.min(o.first, t[0]); o.last = Math.max(o.last, t[0]);
    o.daySet.add(t[0]); o.eksp.set(t[3], (o.eksp.get(t[3]) || 0) + t[6]); o.provTon.set(t[5], (o.provTon.get(t[5]) || 0) + t[7]);
  });
  return [...out.values()].map(o => { o.days = o.daySet.size; o.prov = [...o.provTon.entries()].sort((a, b) => b[1] - a[1])[0][0]; return o; });
}
const activeTxt = r => r.first === r.last ? shortDate(B.days[r.first]) : shortDate(B.days[r.first]) + ' – ' + shortDate(B.days[r.last]);
function prevMonth(m){ const y = +m.slice(0, 4), mo = +m.slice(5); return mo === 1 ? (y - 1) + '-12' : y + '-' + String(mo - 1).padStart(2, '0'); }

/* Range (date-based) realisasi + target, optionally grouped by prov/dist/eksp/src. */
function rangeGroup(fromISO, toISO, sc, group){
  sc = sc || {}; group = group || 'all';
  const out = new Map();
  const get = k => { let o = out.get(k); if(!o){ o = {key:k, vol:0, frc:0, trips:0, tgt:0, hasTgt:false}; out.set(k, o); } return o; };
  B.facts.forEach(f => {
    const ds = B.days[f[0]];
    if(ds < fromISO || ds > toISO || !srcOkS(f[2], sc) || !kemOk(f[10]) || !scopeOkF(f, sc) || nkSkip(f[3], group, sc)) return;
    const o = get(FK[group](f));
    if(achOk(f[1]) && achKem(f[10]) && !NKP().has(f[3])) o.frc += f[6];
    if(incOk(f[1])){ o.vol += f[6]; o.trips += f[7]; }
  });
  {
    const from = parseISO(fromISO), to = parseISO(toISO);
    B.targets.forEach(t => {
      const m = t[0];
      if(!kemOk(tKem(t)) || m + '-31' < fromISO || m + '-01' > toISO || !incOk(tInc(t)) || !srcOkS(t[4], sc) || !scopeOkT(t, sc) || nkSkip(t[1], group, sc)) return;
      const y = +m.slice(0, 4), mo = +m.slice(5) - 1, n = dim(m);
      let add = 0, any = false;
      for(let d = 1; d <= n; d++){
        const dt = new Date(y, mo, d);
        if(dt < from || dt > to) continue;
        add += t[6] ? (t[6][d - 1] || 0) : t[5] / n; any = true;
      }
      if(any){ const o = get(TK[group](t)); o.tgt += add * tgtK(); o.hasTgt = true; }
    });
  }
  out.forEach(o => { o.pct = o.hasTgt ? (o.tgt > 0 ? pctOf(o.frc, o.tgt) : 100) : (o.vol > 0 || o.frc > 0 ? 100 : null); o.gap = o.frc - o.tgt; });
  return out;
}
function rangeTotals(fromISO, toISO, sc){
  return rangeGroup(fromISO, toISO, sc, 'all').get(0) || {vol:0, frc:0, trips:0, tgt:0, hasTgt:false, pct:null, gap:0};
}

/* ===================== Trend panels: periods, compare, drill-down ===================== */
/* A bucket is one bar: {from, to, label, long}. */
function buckets(gran, n){
  const per = PER(), endISO = per.to, out = [];
  if(gran === 'mtd'){
    // cumulative from the start of the period: daily bars up to 45 days, weekly beyond that
    const step = per.days > 45 ? 7 : 1;
    let d = per.from;
    while(true){
      let to = step === 1 ? d : addDays(d, 6);
      if(to > endISO) to = endISO;
      const dd = parseISO(to);
      out.push({from: per.from, to, label: step === 1 ? String(dd.getDate()) : dd.getDate() + '/' + (dd.getMonth() + 1), long: 'Kumulatif ' + shortDate(per.from) + ' – ' + shortDate(to)});
      if(to >= endISO) break;
      d = addDays(to, 1);
    }
  } else if(gran === 'bulanan'){
    const last = endISO.slice(0, 7), idx = D.months.indexOf(last);
    const first = per.kind === 'month' ? D.months[Math.max(0, idx - (n || 6) + 1)] : per.from.slice(0, 7);
    D.months.filter(m => m >= first && m <= last).forEach(m => {
      const from = m === per.from.slice(0, 7) && per.kind !== 'month' ? per.from : m + '-01';
      const to = m === last ? endISO : m + '-' + pad2(lastDayOf(m));
      out.push({from, to, label: monthShort(m), long: monthLabel(m) + (+to.slice(8) < dim(m) || from.slice(8) !== '01' ? ' (' + shortDate(from) + ' – ' + shortDate(to) + ')' : '')});
    });
  } else if(gran === 'harian'){
    const cnt = Math.min(n || 14, Math.max(per.days, n || 14));
    for(let i = cnt - 1; i >= 0; i--){
      const s2 = addDays(endISO, -i), d = parseISO(s2);
      out.push({from: s2, to: s2, label: d.getDate() === 1 || i === cnt - 1 ? d.getDate() + '/' + (d.getMonth() + 1) : String(d.getDate()), long: DAYS[d.getDay()] + ', ' + dayLabel(s2)});
    }
  } else {
    const end = parseISO(endISO), monday = new Date(end);
    monday.setDate(end.getDate() - ((end.getDay() + 6) % 7));
    const weeks = Math.min(53, Math.max(n || 8, Math.ceil(per.days / 7)));
    for(let w = weeks - 1; w >= 0; w--){
      const a2 = new Date(monday); a2.setDate(monday.getDate() - 7 * w);
      const b2 = new Date(a2); b2.setDate(a2.getDate() + 6);
      const bb = b2 > end ? end : b2;
      out.push({from: iso(a2), to: iso(bb), label: a2.getDate() + '/' + (a2.getMonth() + 1), long: a2.getDate() + ' ' + MON3[a2.getMonth()] + ' – ' + bb.getDate() + ' ' + MON3[bb.getMonth()] + ' ' + bb.getFullYear() + (bb < b2 ? ' (berjalan)' : '')});
    }
  }
  return out;
}
/* Previous comparable period: daily -> same weekday last week; weekly -> previous week;
   monthly -> previous month (same day range when the month is still running). */
function shiftBucket(b, gran){
  if(gran === 'mtd'){ const pp = prevPer(PER()); return {from: pp.from, to: addDays(pp.from, Math.round((parseISO(b.to) - parseISO(b.from)) / 864e5))}; }
  if(gran !== 'bulanan') return {from: addDays(b.from, -7), to: addDays(b.to, -7)};
  const m = b.from.slice(0, 7), pm = prevMonth(m), full = +b.to.slice(8) >= dim(m);
  return {from: pm + '-01', to: pm + '-' + pad2(full ? dim(pm) : Math.min(+b.to.slice(8), dim(pm)))};
}
const PREV_LBL = {mtd: 'periode sebelumnya', harian: 'hari sama minggu lalu', mingguan: 'minggu sebelumnya', bulanan: 'bulan sebelumnya'};
function scName(sc){
  if(sc.dist != null) return tc(B.dims.dist[sc.dist]);
  if(sc.eksp != null) return B.dims.eksp[sc.eksp] || 'Tanpa ekspeditur';
  if(sc.prov != null) return tc(B.dims.prov[sc.prov]);
  return 'Semua provinsi';
}
function cmpScope(code){ if(!code || code === 'prev') return null; const [k, v] = code.split(':'); return {[k]: +v}; }
function cmpLabel(T){ return T.cmp === 'prev' ? PREV_LBL[T.gran] : T.cmp ? scName(cmpScope(T.cmp)) : ''; }
function withInc(T, fn){ const s = F.inc; if(T && T.incFix) F.inc = T.incFix; try { return fn(); } finally { F.inc = s; } }
/* More than one incoterm active: render a table/panel once per incoterm (build() runs with F.inc fixed to it and
   returns '' when that incoterm has nothing), each under its own incoterm header. One incoterm: build() as is. */
function incGroups(build, sub){
  const act = INCS().map((n, i) => i).filter(i => incOk(i));
  if(act.length < 2) return build();
  return act.map(i => withInc({incFix: INCS()[i]}, () => { const h = build(); return h ? '<div class="inc-grp" style="--c:' + incColor(i) + '"><span class="inc-dot" style="background:' + incColor(i) + '"></span><b>' + esc(incLbl(INCS()[i])) + '</b>' + (sub ? '<small>' + sub() + '</small>' : '') + '</div>' + h : ''; })).join('');
}
const ekspRows = (m, sc) => [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
const ekspSub = (m, sc) => () => { const t = total(m, sc); return fmt(t.vol) + ' t' + (t.hasTgt ? ' · ' + pctTxt(t.pct) + ' dari target ' + fmt(t.tgt) + ' t' : ''); };
function trendData(T){
  return withInc(T, () => buckets(T.gran).map(b => {
    const r = rangeTotals(b.from, b.to, T.sc);
    const d = Object.assign({}, b, {vol: r.vol, frc: r.frc, tgt: r.pct == null ? null : r.tgt, pct: r.pct});
    if(T.cmp){
      const cb = T.cmp === 'prev' ? shiftBucket(b, T.gran) : b;
      const c = rangeTotals(cb.from, cb.to, T.cmp === 'prev' ? T.sc : cmpScope(T.cmp));
      d.c = {vol: c.vol, frc: c.frc, tgt: c.pct == null ? null : c.tgt, pct: c.pct, from: cb.from, to: cb.to};
    }
    return d;
  }));
}
const TRENDS = {};
function deltaHTML(a, b){
  if(!b) return '<div class="bcmp">–</div>';
  const p = (a / b - 1) * 100;
  return '<div class="bcmp" style="color:' + (p >= 0 ? 'var(--green)' : 'var(--red)') + '">' + (p >= 0 ? '▲' : '▼') + Math.abs(Math.round(p)) + '%</div>';
}
const LG_NUM = '<span class="lg-num"><b>%</b> dalam batang = capaian · <b>angka</b> di atas = ton</span>';
function barsHTML(id, data, opt){
  const hasC = data.some(d => d.c);
  const max = Math.max(1, ...data.map(d => Math.max(d.vol, d.tgt || 0, d.c ? d.c.vol : 0)));
  return '<div class="bars' + (hasC ? ' pair' : '') + (data.length > 16 && !(opt && opt.labels) && !(data.length <= 31 && data.every(d => d.unit === 'jam')) ? ' dense' : '')   /* hours are short ("3,2"): always labelled */ + (data.length > 9 ? ' many' : '') + '">' + data.map((d, i) => {
    const HR = 84;   // headroom above the tallest bar for its value label
    const h = Math.max(2, d.vol / max * HR);
    const th = d.tgt != null ? Math.min(HR, d.tgt / max * HR) : null;
    const c = d.pct != null ? tierColor(d.pct) : 'var(--black)';
    const unit = d.unit === 'jam' ? fmt1(d.vol) : fmt(d.vol);
    // % sits inside the bar (white); the value floats on top of each bar, so labels follow the bar heights instead of one crowded row
    const pin = d.pct != null && h >= 30, pout = d.pct != null && !pin, pTxt = d.pct != null ? Math.round(d.pct) + '%' : '';
    return '<div class="bcol click" data-bar="' + id + ':' + i + '" title="' + esc((d.long || d.label) + (d.pct != null ? ' · ' + pTxt + ' dari target' : '')) + ' — klik untuk rincian">' +
      (d.c ? deltaHTML(d.vol, d.c.vol) : '') +
      '<div class="bwrap"><div class="b" style="height:' + (d.future ? '0%;min-height:0' : h + '%') + ';background:' + (d.color || c) + ';animation-delay:' + (i * 50) + 'ms">' + (pin ? '<span class="bin">' + pTxt + '</span>' : '') + '</div>' +
      (d.future ? '' : '<div class="btop" style="bottom:' + h + '%"><span class="bval">' + unit + (pout ? ' <b style="color:' + c + '">' + pTxt + '</b>' : '') + '</span></div>') +
      (d.c ? '<div class="b b2" style="height:' + Math.max(2, d.c.vol / max * HR) + '%;animation-delay:' + (i * 50 + 25) + 'ms"></div>' : '') +
      (th != null ? '<div class="tick" style="bottom:' + th + '%"></div>' : '') + '</div>' +
      '<div class="blabel">' + esc(d.label) + '</div></div>';
  }).join('') + '</div>';
}
function cmpOptions(T){
  const sc = T.sc, opt = [['', 'Tanpa pembanding'], ['prev', 'vs ' + PREV_LBL[T.gran]]];
  const groups = [];
  if(sc.dist != null){
    const p = B.dims.distProv[sc.dist];
    const ds = [...D.liveDist].filter(i => i !== sc.dist && B.dims.distProv[i] === p).sort((a, b) => B.dims.dist[a].localeCompare(B.dims.dist[b]));
    groups.push(['Distrik lain di ' + tc(B.dims.prov[p]), ds.map(i => ['dist:' + i, tc(B.dims.dist[i])])]);
  } else if(sc.eksp != null){
    groups.push(['Ekspeditur lain', D.ekspList.filter(e => D.ekspIdx[e] !== sc.eksp).map(e => ['eksp:' + D.ekspIdx[e], ekspLbl(D.ekspIdx[e])]).sort((a, b) => a[1].localeCompare(b[1]))]);
  } else {
    groups.push(['Provinsi', [...D.liveProv].filter(i => i !== sc.prov).sort((a, b) => B.dims.prov[a].localeCompare(B.dims.prov[b])).map(i => ['prov:' + i, tc(B.dims.prov[i])])]);
  }
  return '<select class="cmp-sel" data-cmp="' + T.id + '" aria-label="Bandingkan">' +
    opt.map(o => '<option value="' + o[0] + '"' + ((T.cmp || '') === o[0] ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') +
    groups.map(g => '<optgroup label="' + esc(g[0]) + '">' + g[1].map(o => '<option value="' + o[0] + '"' + (T.cmp === o[0] ? ' selected' : '') + '>vs ' + esc(o[1]) + '</option>').join('') + '</optgroup>').join('') + '</select>';
}
function trendPanel(sc, gran, title, opt){
  const id = 't' + Math.random().toString(36).slice(2, 8);
  TRENDS[id] = Object.assign({id, sc: sc || {}, gran: gran || 'mingguan', title: title || 'Tren realisasi', cmp: null}, opt || {});
  return '<div class="dp-panel" id="' + id + '">' + trendInner(id) + '</div>';
}
function trendInner(id){
  const T = TRENDS[id], data = trendData(T);
  T.data = data;
  const chips = [['mtd','MTD'],['harian','Harian'],['mingguan','Mingguan'],['bulanan','Bulanan']].filter(g => !T.grans || T.grans.includes(g[0])).map(g => '<span class="' + (g[0] === T.gran ? 'on' : '') + '" data-gran="' + id + ':' + g[0] + '">' + g[1] + '</span>').join('');
  const inc = T.incFix || F.inc;
  const legend = '<span><i style="background:var(--green)"></i>&ge;100% target</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target ' + tgtLbl() + (inc === 'FOT' ? ' FOT (distributor)' : '') + '</span>' +
    (T.cmp ? '<span><i style="background:#BCC3C5"></i>' + esc(cmpLabel(T)) + '</span><span>▲▼ = selisih vs pembanding</span>' : '');
  return '<div class="pt"><span>' + esc(T.title) + ' (ton)</span><div class="pt-tools">' + (T.noCmp ? '' : cmpOptions(T)) + '<div class="gran-chips">' + chips + '</div></div></div>' + barsHTML(id, data) +
    '<div class="trend-legend">' + legend + LG_NUM + '</div>' +
    '<div class="trend-detail" id="' + id + '-d">Klik batang mana pun untuk rincian per provinsi, distrik, ekspeditur &amp; source' + (T.cmp ? ', dibandingkan dengan ' + esc(cmpLabel(T)) : '') + '.</div>';
}

/* ===================== Drill-down modal ===================== */
let MODAL = null;
/* Ekspeditur = the carrier: FRC* -> the transporter (full name from the kamus); every other incoterm -> the
   distributor itself (it arranges the haul). Labels drop the legal form; links go to the right page. */
const coLbl = n => String(n || '').replace(/[\s,.]+P\.?$/i, m => /,/.test(m) ? '' : m).replace(/[\s,.]+$/, '').replace(/\s*\(PERSERO\)/i, '').replace(/[\s,.]*\bPT\.?\s*$/i, '').replace(/^\s*PT\.?\s+/i, '').trim();
const distrLbl = k => coLbl(tc(B.dims.distr[k] || '')) || tc(B.dims.distr[k] || '');   // distributor name without the legal form
const ekspLbl = k => { const c = B.dims.eksp[k]; if(!c) return 'Angkut sendiri (distributor)'; if(c === 'TANPA_EKSPEDITUR') return 'Belum dialokasikan ke ekspeditur'; const n = (B.dims.ekspNama || [])[k]; return coLbl(n || c) || c; };
const ekspHref = k => { const c = B.dims.eksp[k]; if(!c || c === 'TANPA_EKSPEDITUR') return null; if(D.frcEksp.has(k)) return '#/ekspeditur/' + enc(c); return D.distrIdx[c] != null ? '#/distributor/' + enc(c) : null; };
const DIM_LBL = {prov: 'Provinsi', dist: 'Distrik', eksp: 'Ekspeditur', src: 'Source plant'};
function dimName(dim, k){
  if(dim === 'prov') return tc(B.dims.prov[k]);
  if(dim === 'dist') return tc(B.dims.dist[k]);
  if(dim === 'eksp') return ekspLbl(k);
  return tc(B.dims.src[k] || '(tanpa plant)');
}
function dimHref(dim, k){
  if(dim === 'prov') return '#/provinsi/' + enc(B.dims.prov[k]);
  if(dim === 'dist') return '#/distrik/' + enc(B.dims.dist[k]);
  if(dim === 'eksp') return ekspHref(k);
  return null;
}
function openModal(title, sub, html){
  const old = $('#modal'); if(old) old.remove();   // keep MODAL state: openBreakdown sets it before calling us
  const el = document.createElement('div');
  el.className = 'modal-bg'; el.id = 'modal';
  el.innerHTML = '<div class="modal" role="dialog" aria-modal="true" aria-label="' + esc(title) + '"><div class="modal-head"><div><div class="modal-title">' + esc(title) + '</div><div class="modal-sub">' + sub + '</div></div><button class="modal-x" data-act="closemodal" aria-label="Tutup">&times;</button></div><div id="modalBody">' + html + '</div></div>';
  document.body.appendChild(el); document.body.style.overflow = 'hidden';
}
function closeModal(){ const m = $('#modal'); if(m){ m.remove(); document.body.style.overflow = ''; } MODAL = null; }
function openBreakdown(T, i){
  const d = T.data[i];
  const dims = ['prov', 'dist', 'eksp', 'src'].filter(k => !(k === 'prov' && (T.sc.prov != null || T.sc.dist != null)) && !(k === 'dist' && T.sc.dist != null) && !(k === 'eksp' && T.sc.eksp != null));
  const actInc = withInc(T, () => INCS().filter((n, j) => incOk(j)));
  MODAL = {T, d, dims, dim: dims[0], actInc};   // more than one incoterm: each tab's rows are grouped per incoterm
  const range = d.from === d.to ? dayLabel(d.from) : dayLabel(d.from) + ' – ' + dayLabel(d.to);
  if(!T.sc) T.sc = {};
  openModal(scName(T.sc) + ' · ' + (d.long || d.label), esc(range) + ' · ' + esc(T.incFix || incName()), modalBody());
}
function modalBody(){
  const {T, d, dims, dim, actInc} = MODAL;
  const pb = shiftBucket(d, T.gran);
  const groupRows = () => {
    const cur = rangeGroup(d.from, d.to, T.sc, dim), prv = rangeGroup(pb.from, pb.to, T.sc, dim);
    const rows = [...cur.values()].filter(o => o.vol > 0 || o.tgt > 0).map(o => Object.assign(o, {prev: (prv.get(o.key) || {}).vol || 0}));
    rows.sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
    let r = 0; rows.forEach(o => { o.rank = o.pct == null ? null : ++r; });
    return rows;
  };
  return withInc(T, () => {
    const tot = rangeTotals(d.from, d.to, T.sc), tp = rangeTotals(pb.from, pb.to, T.sc);
    const chg = tp.vol ? (tot.vol / tp.vol - 1) * 100 : null;
    let cmpCell = '';
    if(d.c && T.cmp !== 'prev'){
      const cc = d.c, dv = cc.vol ? (d.vol / cc.vol - 1) * 100 : null;
      cmpCell = '<div class="prog-cell"><div class="pl">vs ' + esc(cmpLabel(T)) + '</div><div class="pv">' + fmt(cc.vol) + ' t</div><div class="pd">capaian ' + pctTxt(cc.pct) + (dv != null ? ' · selisih ' + (dv >= 0 ? '+' : '') + Math.round(dv) + '%' : '') + '</div></div>';
    }
    const head = '<div class="prog-row" style="margin-bottom:18px">' +
      '<div class="prog-cell"><div class="pl">Realisasi</div><div class="pv">' + fmt(tot.vol) + ' t</div><div class="pd">' + fmt(tot.trips) + ' trip</div></div>' +
      '<div class="prog-cell"><div class="pl">Target ' + tgtLbl() + '</div><div class="pv">' + (tot.hasTgt ? fmt(tot.tgt) + ' t' : '–') + '</div><div class="pd">gap ' + (tot.hasTgt ? signed(tot.gap) + ' t' : '–') + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Capaian</div><div class="pv" style="color:' + tierColor(tot.pct) + '">' + pctTxt(tot.pct) + '</div><div class="pd">' + (T.incFix || incName()) + ' vs ' + tgtLbl() + '</div></div>' +
      '<div class="prog-cell"><div class="pl">vs ' + esc(PREV_LBL[T.gran]) + '</div><div class="pv" style="color:' + (chg == null ? 'inherit' : chg >= 0 ? 'var(--green)' : 'var(--red)') + '">' + (chg == null ? '–' : (chg >= 0 ? '+' : '') + Math.round(chg) + '%') + '</div><div class="pd">' + fmt(tp.vol) + ' t (' + esc(pb.from === pb.to ? dayLabel(pb.from) : dayLabel(pb.from) + ' – ' + dayLabel(pb.to)) + ')</div></div>' +
      cmpCell + '</div>';
    const tabs = '<div class="gran-chips" style="margin-bottom:14px">' + dims.map(k => '<span class="' + (k === dim ? 'on' : '') + '" data-mdim="' + k + '">Per ' + DIM_LBL[k].toLowerCase() + '</span>').join('') + '</div>';
    const cols = [C.rank,
      {h: DIM_LBL[dim], cls: 'name', val: o => dimName(dim, o.key), html: o => esc(dimName(dim, o.key)) + (dim === 'dist' ? '<small>' + esc(tc(B.dims.prov[B.dims.distProv[o.key]])) + '</small>' : '')},
      {h: 'Realisasi (t)', num: true, val: o => o.vol, html: o => fmt(o.vol)},
      {h: 'Target (t)', num: true, val: o => o.hasTgt ? o.tgt : null, html: o => o.hasTgt ? fmt(o.tgt) : '–'},
      C.pct, C.gap, C.trips,
      {h: 'Periode lalu (t)', num: true, val: o => o.prev, html: o => fmt(o.prev)},
      {h: 'Δ', num: true, val: o => o.prev ? o.vol / o.prev - 1 : null, html: o => o.prev ? '<span style="color:' + (o.vol >= o.prev ? 'var(--green)' : 'var(--red)') + '">' + (o.vol >= o.prev ? '+' : '') + Math.round((o.vol / o.prev - 1) * 100) + '%</span>' : (o.vol ? 'baru' : '–')}];
    const tbl = rows => table(cols, rows, {go: o => dimHref(dim, o.key)});
    if(actInc.length > 1){   // one group per incoterm: header with the incoterm's own totals, then its rows
      const groups = actInc.map(n => withInc({incFix: n}, () => {
        const rows = groupRows(); if(!rows.length) return '';
        const t = rangeTotals(d.from, d.to, T.sc);
        return '<div class="bd-grp"><span class="inc-dot" style="background:' + INC_COLOR[n] + '"></span><b>' + esc(incLbl(n)) + '</b>' +
          '<span>' + fmt(t.vol) + ' t' + (t.hasTgt ? ' · target ' + fmt(t.tgt) + ' t · <b style="color:' + tierColor(t.pct) + '">' + pctTxt(t.pct) + '</b> · gap ' + signed(t.gap) + ' t' : ' · tanpa target') + ' · ' + fmt(t.trips) + ' trip</span></div>' + tbl(rows);
      })).join('');
      return head + tabs + (groups || '<div class="dp-review">Tidak ada data untuk periode ini.</div>');
    }
    const rows = groupRows();
    return head + tabs + (rows.length ? tbl(rows) : '<div class="dp-review">Tidak ada data untuk periode ini.</div>');
  });
}

/* ===================== Sortable tables ===================== */
const TABLES = {};
function table(cols, rows, opt){
  opt = opt || {};
  const id = 'tb' + Math.random().toString(36).slice(2, 8);
  TABLES[id] = {cols, rows, opt, sort: opt.sort || null, dir: opt.dir || -1};
  return '<div class="tbl-wrap' + (rows.length > 15 ? ' tall' : '') + '" id="' + id + '">' + tableInner(id) + '</div>';   // tall: own scroll area so the header row stays frozen
}
function tableInner(id){
  const T = TABLES[id];
  let rows = T.rows.slice();
  if(T.sort != null){
    const c = T.cols[T.sort], v = c.sortVal || c.val;
    rows.sort((a, b) => { const x = v(a), y = v(b); if(x == null && y == null) return 0; if(x == null) return 1; if(y == null) return -1; return (x < y ? -1 : x > y ? 1 : 0) * T.dir; });
  }
  const lim = T.opt.limit || 500;
  // frozen columns: the first one, plus the next when the first is only a rank number (#)
  const nfz = /\brank\b/.test(T.cols[0] && T.cols[0].cls || '') && T.cols.length > 2 ? 2 : 1;
  const fz = i => i < nfz ? ' fz fz' + i + (i === nfz - 1 ? ' fzl' : '') : '';
  const head = '<tr>' + T.cols.map((c, i) => '<th class="' + (c.num ? 'num' : '') + fz(i) + '" data-sort="' + id + ':' + i + '">' + c.h + (T.sort === i ? '<span class="arr">' + (T.dir > 0 ? '▲' : '▼') + '</span>' : '') + '</th>').join('') + '</tr>';
  const body = rows.slice(0, lim).map((r, ri) => {
    const href = T.opt.go ? T.opt.go(r) : null;
    return '<tr' + (href ? ' class="go" data-href="' + href + '"' : '') + '>' + T.cols.map((c, i) => '<td class="' + (c.cls || '') + (c.num ? ' num' : '') + fz(i) + '"' + (c.style ? ' style="' + c.style(r) + '"' : '') + '>' + (c.html ? c.html(r, ri) : esc(c.val(r))) + '</td>').join('') + '</tr>';
  }).join('');
  const tot = T.opt.total ? '<tr class="total">' + T.cols.map((c, i) => '<td class="' + (c.num ? 'num' : '') + fz(i) + '">' + (c.total != null ? c.total : c.cls === 'name' ? 'TOTAL' : '') + '</td>').join('') + '</tr>' : '';
  const more = rows.length > lim ? '<tr><td colspan="' + T.cols.length + '" style="color:var(--sub)">… ' + (rows.length - lim) + ' baris lainnya tidak ditampilkan</td></tr>' : '';
  return '<table class="tbl"><thead>' + head + '</thead><tbody>' + body + tot + more + '</tbody></table>';
}
/* Compact tables instead of "NAME · x t + bar + %" lists: capTable (realisasi/target/capaian[/Real/SO]), volTable (tonase/porsi). */
function capTable(list, h, name, href, so){
  if(!list.length) return '';
  return '<div class="cap-tbl">' + table([{h: h, cls:'name', val: name}, C.vol('Real (t)'), Object.assign({}, C.tgt, {h:'Target (t)'}), C.pct].concat(so ? [Object.assign({}, C.of, {h:'R/SO'})] : []), list, {go: href, sort: 1}) + '</div>';
}
function volTable(list, h, name, href){
  if(!list.length) return '';
  const tot = list.reduce((a, o) => a + o.vol, 0) || 1;
  return '<div class="cap-tbl">' + table([{h: h, cls:'name', val: name}, {h:'Tonase (t)', num:true, val: o => o.vol, html: o => fmt(o.vol)}, {h:'Porsi', num:true, val: o => o.vol / tot, html: o => pctTxt(pctOf(o.vol, tot))}], list, {go: href, sort: 1}) + '</div>';
}
const C = {
  rank: {h:'Rank', cls:'rank', val: r => r.rank, html: r => r.rank ? '#' + r.rank : '–', num:false},
  vol: (h) => ({h: h || 'Realisasi (t)', num:true, val: r => r.vol, html: r => fmt(r.vol)}),
  tgt: {h:'Target MTD (t)', num:true, val: r => r.hasTgt ? r.tgt : null, html: r => r.hasTgt ? fmt(r.tgt) : '–'},
  pct: {h:'Capaian', num:true, cls:'pc', val: r => r.pct, html: r => pctTxt(r.pct), style: r => 'color:' + tierColor(r.pct)},
  gap: {h:'Gap (t)', num:true, val: r => r.hasTgt ? r.gap : null, html: r => r.hasTgt ? '<span style="color:' + (r.gap < 0 ? 'var(--red)' : 'var(--green)') + '">' + signed(r.gap) + '</span>' : '–'},
  so: {h:'SO (t)', num:true, val: r => r.so || null, html: r => r.so ? fmt(r.so) : '–'},
  of: {h:'Real/SO', num:true, val: r => r.of, html: r => pctTxt(r.of)},   // neutral: SO is a whole month, realisasi is MTD
  trips: {h:'Trip', num:true, val: r => r.trips, html: r => fmt(r.trips)},
  dwell: {h:'Waktu muat (jam)', num:true, val: r => r.dwell, html: r => r.dwell == null ? '–' : '<span style="color:' + slaColor(r.dwell) + '">' + fmt1(r.dwell) + '</span>'},
  note: {h:'Catatan', cls:'note-c', val: r => reviewText(r.pct), sortVal: r => r.pct}
};

/* ===================== Shell: nav, filter, search ===================== */
function setActiveNav(key){ $$('#mainNav .navitem').forEach(n => n.classList.toggle('on', n.dataset.nav === key)); }
function syncNavAria(){ $$('#mainNav .navitem').forEach(n => { const b = n.querySelector('[data-toggle]'); if(b) b.setAttribute('aria-expanded', n.classList.contains('open') ? 'true' : 'false'); }); }
function closeNav(){ $$('#mainNav .navitem').forEach(n => n.classList.remove('open')); syncNavAria(); }
function buildNav(){
  const pm = summarize('prov', PER());
  const provs = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  $('#navProvList').innerHTML = provs.map(o => '<a href="#/provinsi/' + enc(B.dims.prov[o.key]) + '">' + esc(tc(B.dims.prov[o.key])) + '<span>' + pctTxt(o.pct) + '</span></a>').join('') || '<span class="empty">—</span>';
  /* Company names without the legal form ("PT", ", PT", ".PT", "PT.") for compact menus. */
  const coName = n => String(n || '').replace(/\s*\(PERSERO\)/i, '').replace(/[\s,.]*\bPT\.?\s*$/i, '').replace(/^\s*PT\.?\s+/i, '').trim();
  /* Ekspeditur menu: grouped per source plant (only plants switched on in the filter), compact */
  const names = D.ekspList.slice().sort();
  $('#navEkspList').innerHTML = D.sources.filter(sn => !F.srcOff.has(sn)).map(sn => {
    const si = B.dims.src.indexOf(sn), em = summarize('eksp', PER(), {src: si});
    const items = names.map(code => {
      const o = em.get(D.ekspIdx[code]); if(!o || !(o.vol > 0 || o.tgt > 0)) return '';
      const p = o.pct;
      const full = (B.dims.ekspNama || [])[D.ekspIdx[code]], nm = full ? tc(coName(full)) : code;   // full name; code when the master has none
      return '<a class="eksp-item" href="#/ekspeditur/' + enc(code) + '" title="' + esc(code + (full ? ' · ' + full : '')) + '"><span class="en">' + esc(nm) + '</span><span class="ep" style="color:' + tierColor(p) + '">' + (o.hasTgt ? pctTxt(p) : (o.vol ? fmt(o.vol) + ' t' : '–')) + '</span></a>';
    }).join('');
    return items ? '<a class="eksp-plant" href="' + plantHref(sn) + '">' + esc(sn.toUpperCase()) + ' &rsaquo;</a><div class="eksp-grid">' + items + '</div>' : '';
  }).join('') || '<span class="empty">Tidak ada ekspeditur pada plant terpilih</span>';
  /* Distributor menu: per kemasan (all kemasan and incoterm of the chosen moda; plant/period filters apply) */
  const fk = F.kem, fi = F.inc; F.kem = 'ALL'; F.inc = 'ALL';   // moda stays (incOk honours F.moda when inc is ALL)
  let drows; try { drows = shipRows(PER()); } finally { F.kem = fk; F.inc = fi; }
  const KN = B.dims.kem || [], KORD = ['ZAK', 'CURAH', 'JUMBO BAG'];
  const byKem = new Map();
  drows.forEach(r => {
    const kn = KN.length ? (KN[r[10]] || 'ZAK') : 'ZAK';
    let kg = byKem.get(kn); if(!kg){ kg = {tot: 0, d: new Map()}; byKem.set(kn, kg); }
    kg.d.set(r[5], (kg.d.get(r[5]) || 0) + r[8]); kg.tot += r[8];
  });
  const DCAP = 15;
  const cols = [...byKem.entries()].filter(([, g]) => g.tot > 0).sort((a, b) => (KORD.indexOf(a[0]) + 1 || 9) - (KORD.indexOf(b[0]) + 1 || 9)).map(([kn, g]) => {
    const list = [...g.d.entries()].sort((a, b) => b[1] - a[1]);
    const items = list.slice(0, DCAP).map(([di, v]) => {
      const nm = tc(coName(B.dims.distr[di]));
      return '<a class="eksp-item" href="#/distributor/' + enc(B.dims.distr[di]) + '" title="' + esc(distrLbl(di)) + '"><span class="en">' + esc(nm) + '</span><span class="ep" style="color:var(--sub)">' + fmt(v) + ' t</span></a>';
    }).join('');
    const more = list.length > DCAP ? '<a class="dm-more" href="#/distributor">+' + (list.length - DCAP) + ' distributor lain &rsaquo;</a>' : '';
    return '<div class="dm-kem"><div class="dm-kh">' + esc(kn) + ' &middot; ' + fmt(g.tot) + ' t &middot; ' + list.length + ' distributor</div>' + items + more + '</div>';
  }).join('');
  $('#navDistrList').innerHTML = cols ? '<div class="dm-prov"><div class="dm-cols">' + cols + '</div></div>' : '<span class="empty">—</span>';
  /* Plant menu: every plant with shipments or targets in the period (plant filter not applied, like the plant pages) */
  const pl = plantList().map(o => Object.assign(o, total(PER(), {src: o.key}))).filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  $('#navPlantList').innerHTML = pl.map(o => '<a href="' + plantHref(o.n) + '"' + (F.srcOff.has(o.n) ? ' class="off" title="tidak dicentang di filter plant"' : '') + '>' + esc(o.n) + '<span' + (o.hasTgt ? ' style="color:' + tierColor(o.pct) + '"' : '') + '>' + (o.hasTgt ? pctTxt(o.pct) : fmt(o.vol) + ' t') + '</span></a>').join('') || '<span class="empty">—</span>';
  $('#dayLabel').textContent = 'Data s.d. ' + dayLabel(B.asOf);
}
let draft = null;
function buildFilterPanel(){
  const cur = PER();
  draft = {per: Object.assign({}, F.per), inc: F.inc, moda: F.moda, kem: F.kem, area: F.area, srcOff: new Set(F.srcOff), tgt: F.tgt,
           from: cur.from, to: cur.to, month: cur.month};
  renderFilterPanel();
}
function renderFilterPanel(){
  const k = draft.per.kind, first = B.days[0], last = B.asOf;
  const mode = [['month','Bulan'],['ytd','Tahun berjalan'],['range','Rentang tanggal']].map(x => '<span data-pmode="' + x[0] + '" class="' + (k === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  let body = '';
  if(k === 'month'){
    const monthOpts = D.months.slice().reverse().map(m => '<option value="' + m + '"' + (m === draft.per.month ? ' selected' : '') + '>' + monthLabel(m) + (m === D.months[D.months.length - 1] ? ' (berjalan)' : '') + '</option>').join('');
    body = '<select class="fp-select" id="fpMonth">' + monthOpts + '</select><div class="fp-chips" style="margin-top:8px"><span data-qm="cur">Bulan ini</span><span data-qm="prev">Bulan lalu</span></div>';
  } else if(k === 'ytd'){
    body = '<div class="fp-note">' + shortDate(last.slice(0, 4) + '-01-01') + ' – ' + shortDate(last) + ' (tahun berjalan, s.d. data terakhir)</div>';
  } else {
    body = '<div class="fp-dates"><label>Dari<input type="date" id="fpFrom" min="' + first + '" max="' + last + '" value="' + draft.from + '"></label>' +
      '<label>Sampai<input type="date" id="fpTo" min="' + first + '" max="' + last + '" value="' + draft.to + '"></label></div>' +
      '<div class="fp-chips" style="margin-top:8px"><span data-qr="7">7 hari</span><span data-qr="30">30 hari</span><span data-qr="90">90 hari</span></div>';
  }
  const moda = ['DARAT', 'LAUT', 'ALL'].map(x => '<span data-moda="' + x + '" class="' + (draft.moda === x ? 'on' : '') + '">' + MODA_LBL[x].toUpperCase() + '</span>').join('');
  const inc = modaIncs(draft.moda).map(x => [x, incLbl(x)]).concat([['ALL', draft.moda === 'ALL' ? 'Semua' : 'Semua ' + MODA_LBL[draft.moda].toLowerCase()]]).map(x => '<span data-inc="' + x[0] + '" class="' + (draft.inc === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  const kem = (KEMS().length ? KEMS().map(x => [x, x]).concat([['ALL', 'Semua']]) : []).map(x => '<span data-kem="' + x[0] + '" class="' + (draft.kem === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  const tgt = [['SNOP','SNOP'],['INTERNAL','Internal (115% SNOP)']].map(x => '<span data-tgt="' + x[0] + '" class="' + (draft.tgt === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  const src = D.sources.map(s2 => '<label class="fp-check"><input type="checkbox" data-src="' + esc(s2) + '"' + (draft.srcOff.has(s2) ? '' : ' checked') + '> ' + esc(s2.toUpperCase()) + '' + '</label>').join('');
  const tip = (t, r) => '<span class="fp-info' + (r ? ' r' : '') + '" tabindex="0" data-info="' + esc(t) + '">i</span>';
  const AREA_OPT = [['AP1', 'AP1'], ['NON', 'Non AP1'], ['ALL', 'Semua']];
  const area = areaOn() ? AREA_OPT.map(x => '<span data-area="' + x[0] + '" class="' + (draft.area === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('') : '';
  const areaTip = 'AP1 = provinsi di laporan Realisasi harian: ' + B.dims.prov.filter((p, i) => !NKP().has(i) && D.liveProv.has(i)).map(tc).join(', ') + '. NON AP1 = ' + [...NKP()].map(i => tc(B.dims.prov[i])).join(', ') + ' (tanpa target, tidak dihitung dalam capaian).';
  const kemTip = 'Default ZAK. Target tersedia untuk: ' + [...D.tgtKem].map(k => KEMS()[k]).join(', ') + '. Kemasan tanpa target tidak dihitung dalam capaian.';
  const incTip = 'DARAT = pengiriman truk (FRC, FOT). LAUT = pengiriman kapal/kontainer (FOB, FAS, FRC KAPAL, FRC KONTAINER, FOT KONTAINER); truk pengangkut ke pelabuhan ikut moda laut. % pencapaian = realisasi incoterm terpilih vs target ' + tgtLbl() + '-nya sampai hari data terakhir. Target FRC dibagi ke ekspeditur, target FOT ke distributor.';
  const srcTip = 'Berlaku untuk semua halaman: peta, grafik, tabel, ekspeditur, distributor, armada & prognosa. Pilihan tersimpan selama sesi.';
  $('#filterPanel').innerHTML =
    '<div class="fp-row"><div class="fp-group"><label class="fl">Periode</label><div class="fp-chips" style="margin-bottom:10px">' + mode + '</div>' + body + '</div>' +
    '<div class="fp-group"><label class="fl">Moda &amp; incoterm' + tip(incTip, true) + '</label><div class="fp-chips" style="margin-bottom:8px">' + moda + '</div><div class="fp-chips">' + inc + '</div></div></div>' +
    '<div class="fp-row">' + (kem ? '<div class="fp-group"><label class="fl">Kemasan' + tip(kemTip) + '</label><div class="fp-chips">' + kem + '</div></div>' : '<div class="fp-group"></div>') +
    '<div class="fp-group"><label class="fl">Target' + tip('SNOP = target resmi. Internal = 115% dari SNOP.', true) + '</label><div class="fp-chips">' + tgt + '</div></div></div>' +
    (area ? '<div class="fp-row"><div class="fp-group"><label class="fl">Area' + tip(areaTip) + '</label><div class="fp-chips">' + area + '</div></div><div class="fp-group"></div></div>' : '') +
    '<div class="fp-group" style="margin-bottom:18px"><label class="fl">Source plant' + tip(srcTip) + '<span class="fp-all"><a data-srcall="1">Centang semua</a><a data-srcall="0">Hapus semua</a></span></label><div class="fp-cols">' + src + '</div></div>' +
    '<div class="fp-foot"><div class="fp-reset" data-act="reset">Reset filter</div><div class="fp-reset" data-act="laporan">Tampilan laporan harian</div><div class="fp-apply" data-act="apply">Terapkan</div></div>';
}
function filterStrip(){
  const per = PER();
  $('#filterSummary').innerHTML = '<span>Menampilkan</span><span class="tag"><b>' + esc(per.label) + '</b>' + (per.sub !== per.label ? ' · ' + esc(per.sub.replace(per.label + ' · ', '').replace(per.label, '')) : '') + '</span>' +
    '<span class="tag qk" data-qk="moda" title="Ganti moda">moda <b>' + MODA_LBL[F.moda].toUpperCase() + '</b> &#9662;</span>' +
    '<span class="tag qk" data-qk="inc" title="Ganti incoterm"><b>' + incName() + '</b> &#9662;</span>' +
    '<span class="tag qk" data-qk="kem" title="Ganti kemasan"><b>' + (F.kem === 'ALL' ? 'SEMUA KEMASAN' : esc(F.kem)) + '</b> &#9662;</span>' +
    (areaOn() ? '<span class="tag qk" data-qk="area" title="Ganti area"><b>' + areaName() + '</b> &#9662;</span>' : '') +
    '<span class="tag">target <b>' + (F.tgt === 'INTERNAL' ? 'Internal 115%' : 'SNOP') + '</b></span>' +
    '<span class="tag">' + srcTxt() + '</span>' +
    '<span class="edit qrep' + (isReportView() ? ' on' : '') + '" data-act="laporan" title="' + (isReportView() ? 'Kembali ke tampilan lengkap (semua moda, incoterm, plant & area)' : 'Sama dengan laporan Realisasi harian: AP1 · FRC TRUK · ZAK · CP Indarung · SNOP · bulan berjalan') + '">' + (isReportView() ? '&#10003; Laporan harian' : 'Laporan harian') + '</span>' +
    '<span class="edit" data-act="openfilter">Ubah filter &#9662;</span>';
}
/* "Laporan harian" quick view = the daily Realisasi report: AP1, FRC truk, ZAK, CP Indarung, SNOP, current month. */
const REPORT_VIEW = () => ({per: {kind: 'month', month: D.months[D.months.length - 1]}, inc: 'FRC', moda: 'DARAT', kem: 'ZAK', area: 'AP1', srcOff: new Set(D.sources.filter(x => x !== DEFAULT_SRC)), tgt: 'SNOP'});
const DEFAULT_VIEW = () => ({per: {kind: 'month', month: D.months[D.months.length - 1]}, inc: 'ALL', moda: 'ALL', kem: 'ZAK', area: 'ALL', srcOff: new Set(DEFAULT_SRC_OFF), tgt: 'SNOP'});
function isReportView(){
  const r = REPORT_VIEW(), p = F.per || {kind: 'month', month: D.months[D.months.length - 1]};
  return F.inc === r.inc && F.kem === r.kem && F.area === r.area && F.tgt === r.tgt && p.kind === 'month' && p.month === r.per.month &&
    F.srcOff.size === r.srcOff.size && [...r.srcOff].every(x => F.srcOff.has(x));
}
/* Quick switches under the header (moda / incoterm / kemasan) - same filters as the filter panel, one click. */
function quickMenu(kind, anchor){
  const on = (a, b) => a === b ? ' class="on"' : '', ck = (a, b) => a === b ? '<span>&#10003;</span>' : '';
  let items = '';
  if(kind === 'moda') items = ['DARAT', 'LAUT', 'ALL'].map(x => '<a href="javascript:void(0)" data-navmoda="' + x + '"' + on(F.moda, x) + '>' + (x === 'ALL' ? 'Semua moda' : 'Moda ' + MODA_LBL[x].toLowerCase()) + ck(F.moda, x) + '</a>').join('');
  else if(kind === 'inc') items = modaIncs(F.moda).map(x => [x, incLbl(x)]).concat([['ALL', F.moda === 'ALL' ? 'Semua incoterm' : 'Semua ' + MODA_LBL[F.moda].toLowerCase()]]).map(x => '<a href="javascript:void(0)" data-navinc="' + x[0] + '"' + on(F.inc, x[0]) + '>' + x[1] + ck(F.inc, x[0]) + '</a>').join('');
  else if(kind === 'area') items = [['AP1', 'AP1'], ['NON', 'NON AP1'], ['ALL', 'AP1 + NON AP1']].map(x => '<a href="javascript:void(0)" data-navarea="' + x[0] + '"' + on(F.area, x[0]) + '>' + x[1] + ck(F.area, x[0]) + '</a>').join('');
  else items = KEMS().map(k => [k, k]).concat([['ALL', 'Semua kemasan']]).map(x => '<a href="javascript:void(0)" data-navkem="' + x[0] + '"' + on(F.kem, x[0]) + '>' + x[1] + ck(F.kem, x[0]) + '</a>').join('');
  closeQuick();
  const pop = document.createElement('div'); pop.className = 'qk-pop mega-list'; pop.id = 'qkPop'; pop.innerHTML = items;
  const st = $('#filterStrip'); st.appendChild(pop);
  pop.style.left = Math.min(anchor.getBoundingClientRect().left, window.innerWidth - 250) + 'px';
}
/* Switching moda also resets the plant filter: darat = default plant (CP Indarung), laut = all plants (sea cargo ships from other plants, e.g. Teluk Bayur). */
const modaPlants = m => { F.srcOff = m === 'DARAT' ? new Set(DEFAULT_SRC_OFF) : new Set(); };
function closeQuick(){ const p = $('#qkPop'); if(p) p.remove(); }
/* "plant CP Indarung" / "3 dari 18 plant" / "semua plant" (short enough for the filter strip) */
function srcTxt(){
  const on = D.sources.filter(s => !F.srcOff.has(s));
  return on.length === D.sources.length ? 'semua plant' : on.length <= 2 ? 'plant <b>' + on.map(s => esc(s.toUpperCase())).join(', ') + '</b>' : '<b>' + on.length + '</b> dari ' + D.sources.length + ' plant';
}
function applyFilter(){
  const k = draft.per.kind;
  if(k === 'range'){ let a = draft.from || B.days[0], b = draft.to || B.asOf; if(a > b) [a, b] = [b, a]; F.per = {kind: 'range', from: a, to: b}; }
  else if(k === 'ytd') F.per = {kind: 'ytd'};
  else F.per = {kind: 'month', month: draft.per.month || D.months[D.months.length - 1]};
  F.inc = draft.inc; F.moda = draft.moda || 'ALL'; F.kem = draft.kem || 'ZAK'; F.area = draft.area || 'ALL'; F.srcOff = new Set(draft.srcOff); F.tgt = draft.tgt || 'SNOP';
  saveFilter();
  closeNav(); closeFilter(); buildNav(); filterStrip(); route();
}
function closeFilter(){ const f = $('#filterStrip'); if(f) f.classList.remove('open'); }
function openFilter(){
  closeNav(); const f = $('#filterStrip');
  if(f.classList.contains('open')){ f.classList.remove('open'); return; }
  buildFilterPanel(); f.classList.add('open');
  const panel = $('#filterPanel');   // on narrow screens the panel is position:fixed under the strip
  panel.style.top = window.innerWidth <= 1180 ? (f.getBoundingClientRect().bottom + 4) + 'px' : '';
}
function restoreFilter(){
  try {
    const s = JSON.parse(sessionStorage.getItem('osp_f3') || 'null');
    if(s && s.per && (s.per.kind !== 'month' || D.months.includes(s.per.month))){ F.per = s.per; F.inc = s.inc; F.moda = s.moda || (s.inc === 'ALL' ? 'ALL' : modaOf(s.inc)); F.kem = s.kem || 'ZAK'; F.area = s.area || 'ALL'; F.srcOff = new Set(s.srcOff); F.tgt = s.tgt === 'INTERNAL' ? 'INTERNAL' : 'SNOP'; }
  } catch(e) {}
}

let searchIndex = [];
function buildSearch(){
  const P = B.dims.prov, Ds = B.dims.dist;
  searchIndex = [];
  [...D.liveProv].forEach(i => searchIndex.push({t: tc(P[i]), k: 'Provinsi', h: '#/provinsi/' + enc(P[i])}));
  [...D.liveDist].forEach(i => searchIndex.push({t: tc(Ds[i]), k: tc(P[B.dims.distProv[i]] || ''), h: '#/distrik/' + enc(Ds[i])}));
  D.ekspList.forEach(e => searchIndex.push({t: e, k: 'Ekspeditur', h: '#/ekspeditur/' + enc(e)}));
  { const from = B.days.length - 60, seen = new Set(); B.trucks.forEach(t => { if(t[0] >= from) seen.add(t[1]); }); seen.forEach(i => searchIndex.push({t: B.dims.truck[i], k: 'Truk', h: '#/truk/' + enc(B.dims.truck[i])})); }
  plantList().forEach(o => searchIndex.push({t: o.n, k: 'Plant', h: plantHref(o.n)}));
  const tokoDistr = new Map(), seenD = new Set();
  B.ship.forEach(r => { seenD.add(r[5]); if(B.dims.toko[r[7]] && B.dims.toko[r[7]][0] && !tokoDistr.has(r[7])) tokoDistr.set(r[7], r[5]); });
  seenD.forEach(i => { if(B.dims.distr[i]) searchIndex.push({t: distrLbl(i), k: 'Distributor', h: '#/distributor/' + enc(B.dims.distr[i])}); });
  tokoDistr.forEach((d, i) => { const t = B.dims.toko[i]; searchIndex.push({t: (t[1] || t[0]) + (t[2] ? ' — ' + t[2] : ''), k: 'Toko · ' + distrLbl(d).slice(0, 22), h: '#/distributor/' + enc(B.dims.distr[d])}); });
  [['Ringkasan','#/'],['Semua provinsi','#/provinsi'],['Peringkat ekspeditur & scorecard','#/ekspeditur'],['Distributor & toko tujuan','#/distributor'],['Armada / truk / waktu muat','#/armada'],['Kontrak angkutan (kuota, SO kontrak)','#/kontrak'],['Kualitas data (cek master)','#/kualitas'],['Prognosa minggu ini','#/prognosa/minggu-ini'],['Prognosa akhir bulan (proyeksi)','#/prognosa/akhir-bulan']]
    .forEach(x => searchIndex.push({t: x[0], k: 'Halaman', h: x[1]}));
}
function doSearch(q){
  const box = $('#searchResults');
  q = q.trim().toLowerCase();
  if(!q){ box.classList.remove('show'); return; }
  const words = q.split(/\s+/);
  const hits = searchIndex.filter(x => words.every(w => (x.t + ' ' + x.k).toLowerCase().includes(w))).slice(0, 12);
  box.innerHTML = hits.length ? hits.map((x, i) => '<a href="' + x.h + '" class="' + (i === 0 ? 'sel' : '') + '">' + esc(x.t) + '<small>' + esc(x.k) + '</small></a>').join('') : '<div class="none">Tidak ditemukan untuk “' + esc(q) + '”.</div>';
  box.classList.add('show');
}

/* ===================== Map ===================== */
/* Real province borders come from assets/sumatra.js (window.SUMATRA: {w,h,p:{PROVINSI:{d,cx,cy}}}). */
const MAP_LABEL = {
  'ACEH': ['Aceh'], 'SUMATERA UTARA': ['Sumatera Utara', 0, 10], 'SUMATERA BARAT': ['Sumatera Barat', 4, -22],
  'RIAU DARATAN': ['Riau', 22, -8], 'KEPULAUAN RIAU': ['Kep. Riau', 0, -22], 'JAMBI': ['Jambi', 4, -6],
  'SUMATERA SELATAN': ['Sumatera Selatan', 0, -8], 'BENGKULU': ['Bengkulu', -62, 6, 'out'], 'BANGKA BELITUNG': ['Babel', 0, 0],
  'LAMPUNG': ['Lampung', -8, 0]
};
/* Plant locations on the map: [label, lon, lat, label side, province, [source plant names as in the filter]].
   Plants within a few km of each other share one pin (the map is ~40 px per degree). Coordinates are approximate town/port
   locations. Not drawn (outside the Sumatra map): CP Narogong SBI and NCP Mortar Nar MKI (Java), TANPA_PLANT (no location). */
const PLANTS = [
  ['Aceh Besar', 95.40, 5.52, 'right', 'ACEH', ['PP MALAHAYATI SP', 'CP LHOKNGA SBA']],
  ['Lhokseumawe', 97.15, 5.19, 'right', 'ACEH', ['PP LHOKSEUMAWE', 'PP LHOKSEUMAWE SBA']],
  ['Belawan', 98.69, 3.78, 'right', 'SUMATERA UTARA', ['PP BELAWAN SBA']],
  ['Dumai', 101.45, 1.67, 'right', 'RIAU DARATAN', ['GP DUMAI', 'DC DUMAI SP', 'PP DUMAI SBI']],
  ['DC Pekanbaru', 101.45, 0.53, 'left', 'RIAU DARATAN', ['DC PEKANBARU SP']],
  ['Batam', 104.02, 1.12, 'right', 'RIAU KEPULAUAN', ['PP BATAM SP', 'PP BATAM SBA']],
  ['Indarung · T. Bayur', 100.42, -0.96, 'left', 'SUMATERA BARAT', ['CP INDARUNG', 'PP TELUK BAYUR', 'VP TRANSHPMNT PDG SP']],
  ['PP Bengkulu', 102.27, -3.80, 'right', 'BENGKULU', ['PP BENGKULU']],
  ['Palembang', 104.78, -2.97, 'right', 'SUMATERA SELATAN', ['PP PALEMBANG SMBR', 'DC PALEMBANG SMBR']],
  ['Baturaja', 104.17, -4.13, 'right', 'SUMATERA SELATAN', ['CP BATURAJA SMBR']]
];
const pinHref = p => { const have = p[5].filter(n => B.dims.src.includes(n)); return have.length === 1 ? plantHref(have[0]) : '#/plant'; };
const plantOn = p => p[5].some(n => !F.srcOff.has(n));
const plantTip = p => p[0] + '|' + p[5].map(n => (F.srcOff.has(n) ? '○ ' : '● ') + n).join('|');
const PROV_ALIAS = {'RIAU KEPULAUAN': 'KEPULAUAN RIAU'};   // master spelling -> map name
const mapProv = name => PROV_ALIAS[name] || name;
function mapHTML(provRows, opt){
  const G = window.SUMATRA, hero = !!(opt && opt.hero);
  const byName = Object.fromEntries(provRows.map(o => [mapProv(B.dims.prov[o.key]), o]));
  const withPct = provRows.filter(o => o.pct != null);
  const worst = withPct.length ? withPct.reduce((a, b) => a.pct < b.pct ? a : b) : null;
  const defs = '<defs>' +
    '<linearGradient id="gGreen" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2E9E5C"/><stop offset="100%" stop-color="#1A7A42"/></linearGradient>' +
    '<linearGradient id="gAmber" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#C9952E"/><stop offset="100%" stop-color="#8A5A00"/></linearGradient>' +
    '<linearGradient id="gRed" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FF6B74"/><stop offset="100%" stop-color="#E02A36"/></linearGradient>' +
    '<linearGradient id="gInk" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#5A5A5A"/><stop offset="100%" stop-color="#2B2B2B"/></linearGradient>' +
    (hero ? '<radialGradient id="gSea" cx="45%" cy="45%" r="75%"><stop offset="0%" stop-color="#22343F"/><stop offset="100%" stop-color="#0F1A21"/></radialGradient>'
          : '<linearGradient id="gSea" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#EEF4F7"/><stop offset="100%" stop-color="#E2ECF1"/></linearGradient>') +
    '<filter id="softShadow" x="-10%" y="-10%" width="120%" height="120%"><feDropShadow dx="0" dy="1.5" stdDeviation="2" flood-color="#000" flood-opacity="0.18"/></filter></defs>';
  const sea = '<rect width="' + G.w + '" height="' + G.h + '" fill="' + (hero ? 'transparent' : 'url(#gSea)') + '"/>' +
    '<text x="40" y="360" class="sea-lbl">SAMUDRA HINDIA</text><text x="290" y="70" class="sea-lbl">SELAT MALAKA</text><text x="420" y="150" class="sea-lbl">LAUT NATUNA</text>';
  let shapes = '', labels = '';
  Object.keys(G.p).forEach(id => {
    const g = G.p[id], o = byName[id], L = MAP_LABEL[id] || [tc(id)];
    const lx = g.cx + (L[1] || 0), ly = g.cy + (L[2] || 0);
    if(!o){
      shapes += '<path class="map-region" d="' + g.d + '" fill="' + (hero ? '#34424B' : '#D9DEDF') + '" stroke="' + (hero ? '#1B2730' : '#fff') + '" stroke-width="1" data-tip="' + esc(tc(id) + '|Tidak ada pengiriman darat di data') + '"/>';
      labels += '<text x="' + lx + '" y="' + ly + '" class="map-lbl dim">' + esc(L[0]) + '</text>';
      return;
    }
    const fill = o.pct != null ? 'url(#g' + tierName(o.pct) + ')' : 'url(#gInk)';
    const tip = tc(id) + '|Realisasi ' + fmt(o.vol) + ' t' + (o.hasTgt ? ' · Target MTD ' + fmt(o.tgt) + ' t' : '') + '|' + (o.pct != null ? 'Capaian ' + pctTxt(o.pct) + ' · gap ' + signed(o.gap) + ' t' : 'Tanpa target ' + tgtLbl() + '') + soTip(o);
    shapes += '<path class="map-region live' + (worst && o === worst ? ' is-attn' : '') + '" d="' + g.d + '" fill="' + fill + '" stroke="' + (hero ? '#0F1A21' : '#fff') + '" stroke-width="1.2" filter="url(#softShadow)" data-href="#/provinsi/' + enc(B.dims.prov[o.key]) + '" data-tip="' + esc(tip) + '"/>';
    const out = L[3] === 'out' ? ' out' : '';   // label placed in the sea next to a narrow province
    labels += '<text x="' + lx + '" y="' + (ly - 4) + '" class="map-lbl' + out + '">' + esc(L[0]) + '</text>' +
      '<text x="' + lx + '" y="' + (ly + 11) + '" class="map-pct' + out + '" style="' + (out ? 'fill:' + tierColor(o.pct) : '') + '">' + (o.pct != null ? pctTxt(o.pct) : fmt(o.vol) + ' t') + '</text>';
  });
  const pins = PLANTS.map(p => {
    const x = ((p[1] - 95) * 40).toFixed(1), y = ((6.2 - p[2]) * 40).toFixed(1), on = plantOn(p);
    return '<g class="map-pin' + (on ? '' : ' off') + '" data-tip="' + esc(plantTip(p)) + '" data-href="' + pinHref(p) + '"><circle cx="' + x + '" cy="' + y + '" r="5.5" fill="#fff" stroke="#000" stroke-width="1.6"/><circle cx="' + x + '" cy="' + y + '" r="2.2" fill="' + (on ? 'var(--red)' : '#9AA5AB') + '"/>' +
      '<text x="' + (p[3] === 'left' ? +x - 8 : +x + 8) + '" y="' + (+y + 3.5) + '" class="pin-lbl" text-anchor="' + (p[3] === 'left' ? 'end' : 'start') + '">' + esc(p[0]) + '</text></g>';
  }).join('');
  const sorted = provRows.slice().sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  const legend = '<div class="map-row map-head"><span class="dot"></span><span class="nm"></span><span class="tn">Realisasi</span><span class="pc">vs target</span><span class="so">Real/SO</span></div>' + sorted.map((o, i) => {
    const name = B.dims.prov[o.key];
    return '<a class="map-row" href="#/provinsi/' + enc(name) + '"><span class="dot" style="background:' + tierColor(o.pct) + '"></span><span class="nm">#' + (i + 1) + ' ' + esc(tc(name)) + '</span><span class="tn">' + fmt(o.vol) + ' t</span><span class="pc" style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</span><span class="so">' + (o.of == null ? '–' : pctTxt(o.of)) + '</span></a>';
  }).join('') + '<div class="map-row muted"><span class="dot" style="background:#D9DEDF"></span><span class="nm">Aceh, Kep. Riau, Babel, Lampung</span><span class="pc">—</span><span class="so"></span></div>' +
    '<div class="map-key"><span><i style="background:#1A7A42"></i>&ge;100%</span><span><i style="background:#8A5A00"></i>85–99%</span><span><i style="background:#E02A36"></i>&lt;85%</span><span><i class="pin"></i>plant</span></div>';
  const svg = '<svg class="map-svg" viewBox="0 0 ' + G.w + ' ' + G.h + '" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta pencapaian per provinsi di Sumatera">' + defs + sea + shapes + labels + pins + '</svg>';
  if(hero) return svg;
  return '<div class="map-wrap">' + svg + '<div class="map-legend">' + legend + '</div></div>';
}

/* ===================== District map (province & district pages) ===================== */
/* Kabupaten/kota borders come from assets/districts.js (window.DISTRICTS: {PROVINSI:{vb,d:[{n,d,cx,cy,w,h}]}}),
   loaded on first use. SAP district names are cut at 20 characters and spelled differently from the map
   ("KAB. KUANTAN SINGING", "KOTA PADANG SIDEMPUA"), so names are matched per province by prefix, then loosely. */
let distMapLoading = false, distMapPending = null;
function loadDistricts(){
  if(distMapLoading) return;
  distMapLoading = true;
  const own = document.querySelector('script[src*="app.js"]'), v = own && own.src.split('?v=')[1];
  const s = document.createElement('script');
  s.src = 'assets/districts.js' + (v ? '?v=' + v : '');
  s.onload = () => { const slot = $('#dmapSlot'); if(slot && distMapPending) slot.outerHTML = distMapPending(); distMapPending = null; };
  s.onerror = () => { const slot = $('#dmapSlot'); if(slot) slot.innerHTML = '<div class="empty">Peta distrik tidak bisa dimuat.</div>'; distMapLoading = false; };
  document.head.appendChild(s);
}
const DIST_ABBR = {TANJ: 'TANJUNG', TJ: 'TANJUNG', BRT: 'BARAT', TIM: 'TIMUR', TMR: 'TIMUR', SEL: 'SELATAN', SLT: 'SELATAN',
  UTR: 'UTARA', TGH: 'TENGAH', TNG: 'TENGAH', KEP: 'KEPULAUAN'};   // SAP shortenings, e.g. "KAB. TANJ JABUNG BRT"
function distKey(s){
  s = String(s || '').toUpperCase().trim().replace(/^KAB(UPATEN)?\b\.?\s*/, '');
  s = s.replace(/[A-Z]+/g, w => DIST_ABBR[w] || w);
  let kota = /^KOTA\b/.test(s) ? true : null;
  if(kota) s = s.replace(/^KOTA\b\.?\s*/, '');
  return {kota, k: s.replace(/[^A-Z]/g, '')};
}
function lev(a, b){
  const row = Array.from({length: b.length + 1}, (_, j) => j);
  for(let i = 1; i <= a.length; i++){
    let prev = row[0]; row[0] = i;
    for(let j = 1; j <= b.length; j++){ const t = row[j]; row[j] = Math.min(row[j] + 1, row[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1)); prev = t; }
  }
  return row[b.length];
}
/** Index of the map shape for a data district name, or -1. */
function matchDistrict(name, shapes){
  const q = distKey(name), isKab = /^KAB/i.test(String(name).trim());
  let best = -1, bestScore = Infinity;
  shapes.forEach((g, i) => {
    const t = distKey(g.n);
    t.kota = /^Kota /.test(g.n);
    if((q.kota && !t.kota) || (isKab && t.kota)) return;
    const score = q.k === t.k ? 0 : t.k.startsWith(q.k) && q.k.length >= 5 ? 1 : q.k.length >= 6 && lev(q.k, t.k.slice(0, q.k.length)) <= Math.max(1, q.k.length >> 3) ? 2 : Infinity;
    if(score < bestScore || (score === bestScore && best >= 0 && t.k.length < distKey(shapes[best].n).k.length)){ best = i; bestScore = score; }
  });
  return bestScore === Infinity ? -1 : best;
}
function distMapSection(pname, m, sel){
  const map = distMapHTML(pname, m, {sel});
  return map ? (sel != null ? '<div class="section-head"><div class="section-title">Peta capaian per distrik</div><a class="section-link" href="#/provinsi/' + enc(pname) + '">Halaman provinsi &rsaquo;</a></div>' : '') + map : '';
}
/** Province map coloured by district achievement. opt.sel = data district index to highlight. */
function distMapHTML(pname, m, opt){
  opt = opt || {};
  if(!window.DISTRICTS){
    distMapPending = () => distMapHTML(pname, m, opt);
    loadDistricts();
    return '<div id="dmapSlot" class="map-wrap dmap"><div class="empty">Memuat peta distrik…</div></div>';
  }
  const G = window.DISTRICTS[mapProv(pname)];
  if(!G) return '';
  const pi = D.provIdx[pname], rows = [...summarize('dist', m, {prov: pi}).values()].filter(o => o.vol > 0 || o.tgt > 0);
  // data districts -> map shapes (several SAP spellings can land on one shape)
  const agg = new Map(), missing = [];
  rows.forEach(o => {
    const name = B.dims.dist[o.key], i = matchDistrict(name, G.d);
    if(i < 0){ missing.push(o); return; }
    let a = agg.get(i);
    if(!a){ a = {vol:0, frc:0, frc0:0, tgt:0, so:0, trips:0, hasTgt:false, keys:[]}; agg.set(i, a); }
    a.vol += o.vol; a.frc += o.frc; a.frc0 += o.frc0; a.tgt += o.tgt; a.so += o.so; a.trips += o.trips; a.hasTgt = a.hasTgt || o.hasTgt; a.keys.push(o.key);
  });
  agg.forEach(a => { a.pct = a.hasTgt ? (a.tgt > 0 ? pctOf(a.frc, a.tgt) : 100) : (a.vol > 0 || a.frc > 0 ? 100 : null); a.gap = a.frc - a.tgt; a.of = soOn() && a.so > 0 ? pctOf(a.frc0, a.so) : null; });   // same rules as summarize()
  const [vx, vy, vw, vh] = G.vb, fs = Math.max(vw, vh) / 52;
  const defs = '<defs>' + [['Green', '#2E9E5C', '#1A7A42'], ['Amber', '#C9952E', '#8A5A00'], ['Red', '#FF6B74', '#E02A36'], ['Ink', '#7A8A93', '#4E5D66']].map(c =>
    '<linearGradient id="d' + c[0] + '" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="' + c[1] + '"/><stop offset="100%" stop-color="' + c[2] + '"/></linearGradient>').join('') + '</defs>';
  let shapes = '', labels = '', selShape = '';
  G.d.forEach((g, i) => {
    const a = agg.get(i), sel = a && opt.sel != null && a.keys.includes(opt.sel);
    if(!a){
      shapes += '<path class="map-region" d="' + g.d + '" fill="#34424B" stroke="#1B2730" stroke-width="1" vector-effect="non-scaling-stroke" data-tip="' + esc(g.n + '|Tidak ada pengiriman di periode ini') + '"/>';
      return;
    }
    const fill = a.pct != null ? 'url(#d' + tierName(a.pct) + ')' : 'url(#dInk)';
    const tip = g.n + '|Realisasi ' + fmt(a.vol) + ' t · ' + fmt(a.trips) + ' trip' + (a.hasTgt ? ' · Target MTD ' + fmt(a.tgt) + ' t' : '') + '|' + (a.pct != null ? 'Capaian ' + pctTxt(a.pct) + ' · gap ' + signed(a.gap) + ' t' : 'Tanpa target ' + tgtLbl() + '') + soTip(a);
    const href = '#/distrik/' + enc(B.dims.dist[a.keys[0]]);
    const path = '<path class="map-region live' + (sel ? ' is-sel' : '') + '" d="' + g.d + '" fill="' + fill + '" stroke="' + (sel ? '#fff' : '#0F1A21') + '" stroke-width="' + (sel ? 3 : 1.2) + '" vector-effect="non-scaling-stroke" data-href="' + href + '" data-tip="' + esc(tip) + '"/>';
    if(sel) selShape = path; else shapes += path;
    if(Math.min(g.w, g.h) >= Math.max(vw, vh) * 0.09){
      const short = g.n.replace(/^Kota /, 'Kt. ');
      labels += '<text x="' + g.cx + '" y="' + (g.cy - fs * 0.3) + '" class="map-lbl" style="font-size:' + fs.toFixed(1) + 'px">' + esc(short) + '</text>' +
        '<text x="' + g.cx + '" y="' + (g.cy + fs * 1.05) + '" class="map-pct" style="font-size:' + (fs * 1.15).toFixed(1) + 'px">' + (a.pct != null ? pctTxt(a.pct) : fmt(a.vol) + ' t') + '</text>';
    }
  });
  const pins = PLANTS.filter(p => mapProv(p[4]) === mapProv(pname)).map(p => {
    const x = (p[1] - 95) * 200, y = (6.2 - p[2]) * 200, on = plantOn(p);
    return '<g class="map-pin' + (on ? '' : ' off') + '" data-tip="' + esc(plantTip(p)) + '" data-href="' + pinHref(p) + '"><circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + (fs * 0.55).toFixed(1) + '" fill="#fff" stroke="#000" stroke-width="1.6" vector-effect="non-scaling-stroke"/><circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + (fs * 0.22).toFixed(1) + '" fill="' + (on ? 'var(--red)' : '#9AA5AB') + '"/>' +
      '<text x="' + (x + fs * 0.9).toFixed(1) + '" y="' + (y + fs * 0.35).toFixed(1) + '" class="pin-lbl" style="font-size:' + (fs * 0.95).toFixed(1) + 'px">' + esc(p[0]) + '</text></g>';
  }).join('');
  const svg = '<svg class="map-svg" viewBox="' + G.vb.join(' ') + '" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta capaian per distrik di ' + esc(tc(pname)) + '">' +
    defs + shapes + selShape + labels + pins + '</svg>';
  const list = [...agg.entries()].map(([i, a]) => Object.assign({n: G.d[i].n}, a))
    .sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  // headline numbers: the same totals as the province KPIs
  const t = total(m, {prov: pi}), withT = list.filter(a => a.pct != null), ok = withT.filter(a => a.pct >= 100).length;
  const worst = withT.length ? withT.reduce((x, y) => x.pct < y.pct ? x : y) : null, best = withT.length ? withT.reduce((x, y) => x.pct > y.pct ? x : y) : null;
  const sentP = withT.length ? '<b>' + ok + ' dari ' + withT.length + '</b> distrik bertarget di ' + esc(tc(pname)) + ' sudah &ge;100%. ' +
    (worst && worst.pct < 100 ? 'Terberat: <b class="bad">' + esc(worst.n) + ' ' + pctTxt(worst.pct) + '</b> (gap ' + signed(worst.gap) + ' t).' : 'Tertinggi: <b>' + esc(best.n) + ' ' + pctTxt(best.pct) + '</b>.')
    : esc(fmt(t.vol)) + ' t terkirim ke ' + list.length + ' distrik di ' + esc(tc(pname)) + '.';
  const maxV = Math.max(1, ...list.map(a => a.vol));
  const rankRows = list.map((a, r) => '<a class="mh-row' + (opt.sel != null && a.keys.includes(opt.sel) ? ' is-sel' : '') + '" href="#/distrik/' + enc(B.dims.dist[a.keys[0]]) + '" data-hl="#/distrik/' + enc(B.dims.dist[a.keys[0]]) + '"><span class="mh-rk">' + (r + 1) + '</span><span class="mh-nm">' + esc(a.n) + '<i style="width:' + (a.vol / maxV * 100) + '%;background:' + tierColor(a.pct) + '"></i></span><span class="mh-tn">' + fmt(a.vol) + ' t</span><span class="mh-pc" style="color:' + heroTier(a.pct) + '">' + pctTxt(a.pct) + '</span><span class="mh-so">' + (a.of == null ? '–' : pctTxt(a.of)) + '</span></a>').join('') +
    (missing.length ? '<div class="mh-miss">Tidak ada di peta: ' + missing.map(o => esc(tc(B.dims.dist[o.key]))).join(', ') + '</div>' : '');
  return heroPanel({photo: pname, eyebrow: 'Peta capaian per distrik' + (opt.sel != null ? ' · <a href="#/provinsi/' + enc(pname) + '">' + esc(tc(pname)) + ' &rsaquo;</a>' : ''), head: sentP,
    chips: kpiChips(t).concat(heroChip('Distrik terkirim', list.length + ' <small>/ ' + G.d.length + '</small>', 'distrik di peta')),
    listHead: 'Rank distrik', listCols: ['Realisasi', 'Target', 'Real/SO'], rows: rankRows,
    key: '<span><i style="background:#2E9E5C"></i>&ge;100%</span><span><i style="background:#C9952E"></i>85–99%</span><span><i style="background:#FF6B74"></i>&lt;85%</span><span><i style="background:#34424B"></i>tanpa kiriman</span><span><i class="pin"></i>plant</span>',
    map: svg, hint: 'Arahkan kursor untuk angka · klik distrik untuk rincian'});
}
/* Map panels in the home-hero style: dark band with the province photo, headline + chips + ranking on the left. */
/** The page KPIs as hero chips: realisasi, target MTD, sisa/lebih, capaian, Real/SO (t = total()). */
function kpiChips(t){
  return [heroChip('Realisasi', fmt(t.vol) + ' <small>t</small>', fmt(t.trips) + ' trip'),
    heroChip('Target ' + tgtLbl() + ' MTD', t.hasTgt ? fmt(t.tgt) + ' <small>t</small>' : '–', t.hasTgt ? 'bulan penuh ' + fmt(t.tgtFull) + ' t' : 'tidak ada target'),
    heroChip(t.gap < 0 ? 'Sisa ke target' : 'Lebih dari target', t.hasTgt ? fmt(Math.abs(t.gap)) + ' <small>t</small>' : '–', t.hasTgt ? (t.gap < 0 ? 'belum terkirim' : 'di atas target') : '', t.hasTgt ? (t.gap < 0 ? '#FF6B74' : '#5FD08F') : null),
    heroChip('Capaian', pctTxt(t.pct), incName(), heroTier(t.pct)),
    heroChip('Real / SO', t.so ? pctTxt(t.of) : '–', t.so ? 'SO ' + fmt(t.so) + ' t' : 'SO tidak tersedia')];
}
const heroTier = p => p == null ? '#C9D3D9' : p >= 100 ? '#5FD08F' : p >= 85 ? '#E8B04A' : '#FF6B74';
const heroChip = (l, v, sub, color) => '<div class="mh-chip"><div class="l">' + l + '</div><div class="v"' + (color ? ' style="color:' + color + '"' : '') + '>' + v + '</div>' + (sub ? '<div class="s">' + sub + '</div>' : '') + '</div>';
function heroPanel(o){
  const bg = (window.BG || {})[provBg(o.photo)];
  return '<section class="map-hero mh-page' + (o.cls ? ' ' + o.cls : '') + '"><div class="mh-bg"' + (bg ? ' style="background-image:linear-gradient(90deg,rgba(15,26,33,.95) 0%,rgba(15,26,33,.8) 42%,rgba(15,26,33,.38) 100%),url(&quot;assets/bg/' + bg.f + '&quot;)"' : '') + '></div><div class="mh-inner">' +
    '<div class="mh-left"><div class="eyebrow">' + o.eyebrow + '</div><h2 class="mh-h">' + o.head + '</h2>' +
      '<div class="mh-chips">' + o.chips.join('') + '</div>' +
      '<div class="mh-list"><div class="mh-lh"><span>' + o.listHead + '</span></div><div class="mh-row mh-head' + (o.listCols.length === 2 ? ' mh-row2' : '') + '"><span></span><span></span>' + o.listCols.map((c, i) => '<span class="' + ['mh-tn', 'mh-pc', 'mh-so'][i] + '">' + c + '</span>').join('') + '</div><div class="mh-scroll">' + o.rows + '</div></div>' +
      '<div class="mh-key">' + o.key + '</div>' + (o.note || '') +
    '</div>' +
    '<div class="mh-map">' + o.map + '<div class="mh-hint">' + o.hint + '</div></div>' +
  '</div></section>';
}

let KMODE = null;
function kmapMode(){ if(!KMODE){ try { KMODE = sessionStorage.getItem('kmode'); } catch(e){} } return ['distr', 'eksp', 'toko'].includes(KMODE) ? KMODE : 'distr'; }
/* ===================== Kecamatan map (district page) ===================== */
/* Kecamatan borders come from assets/kec/<province>.js (window.KEC[PROVINSI][district shape name] = {vb, d:[{n,d,cx,cy,w,h}]}),
   loaded when a district page opens. Kecamatan are shaded by tonnage (there are no targets below district level).
   Ship-to location, best first: coordinates uploaded on the admin page (B.dims.toko[i][4..5] = lat, lon); else the centre
   of its kecamatan (B.dims.toko[i][3] -> B.dims.kec, or a kecamatan name found in the ship-to name/address), spread out
   when several share one kecamatan. */
const kecLoading = {};   // file -> 'loading' | 'done'
let kecPending = null;     // renders the section for the district page currently waiting on a file
function loadKec(prov, done){
  const fn = 'assets/kec/' + slug(prov) + '.js';
  if(kecLoading[fn]) return;
  kecLoading[fn] = 'loading';
  const own = document.querySelector('script[src*="app.js"]'), v = own && own.src.split('?v=')[1];
  const s = document.createElement('script');
  s.src = fn + (v ? '?v=' + v : '');
  s.onload = s.onerror = () => { kecLoading[fn] = 'done'; done(); };   // no file: the section falls back to the province map
  document.head.appendChild(s);
}
const kecKey = s => String(s || '').toUpperCase().replace(/^KEC(AMATAN)?\b\.?\s*/, '').replace(/[A-Z]+/g, w => DIST_ABBR[w] || w).replace(/[^A-Z]/g, '');
function matchKec(name, shapes){
  const q = kecKey(name);
  if(q.length < 3) return -1;
  let best = -1, bestScore = Infinity;
  shapes.forEach((g, i) => {
    const t = kecKey(g.n);
    const score = q === t ? 0 : t.startsWith(q) && q.length >= 5 ? 1 : q.length >= 6 && lev(q, t.slice(0, q.length)) <= Math.max(1, q.length >> 3) ? 2 : Infinity;
    if(score < bestScore){ best = i; bestScore = score; }
  });
  return best;
}
/** Is (x, y) inside an SVG path made of "M x,y L x,y ... Z" rings (even-odd)? */
function inPath(x, y, d){
  let hit = false;
  d.split('M').forEach(ring => {
    const p = ring.replace(/Z/g, '').split('L').filter(Boolean).map(q => q.split(',').map(Number));
    for(let i = 0, j = p.length - 1; i < p.length; j = i++)
      if((p[i][1] > y) !== (p[j][1] > y) && x < (p[j][0] - p[i][0]) * (y - p[i][1]) / (p[j][1] - p[i][1]) + p[i][0]) hit = !hit;
  });
  return hit;
}
function kecMapSection(name, pname, m, di){
  const prov = mapProv(pname), back = () => distMapSection(pname, m, di);
  if(!window.DISTRICTS){   // districts.js first (its loader fills #dmapSlot), then the kecamatan file
    distMapPending = () => kecMapSection(name, pname, PER(), di);
    loadDistricts();
    return '<div id="dmapSlot" class="map-wrap dmap"><div class="empty">Memuat peta…</div></div>';
  }
  if(!(window.KEC && window.KEC[prov]) && kecLoading['assets/kec/' + slug(prov) + '.js'] !== 'done'){
    kecPending = () => kecMapSection(name, pname, PER(), di);
    loadKec(prov, () => { const slot = $('#kmapSlot'); if(slot && kecPending) slot.outerHTML = kecPending(); kecPending = null; });
    return '<div id="kmapSlot" class="map-wrap dmap"><div class="empty">Memuat peta kecamatan…</div></div>';
  }
  const G = window.DISTRICTS[prov], gi = G ? matchDistrict(name, G.d) : -1;
  const K = gi >= 0 && window.KEC && window.KEC[prov] && window.KEC[prov][G.d[gi].n];
  if(!K) return back();
  // ship-to totals in this district
  const per = new Map();
  shipRows(m, {dist: di}).forEach(r => {
    let o = per.get(r[7]);
    if(!o){ o = {ti: r[7], vol: 0, trips: 0, distr: new Set(), byD: new Map(), byE: new Map()}; per.set(r[7], o); }
    o.vol += r[8]; o.trips += r[9]; o.distr.add(r[5]);
    o.byD.set(r[5], (o.byD.get(r[5]) || 0) + r[8]); o.byE.set(r[6], (o.byE.get(r[6]) || 0) + r[8]);
  });
  // list mode (Rendy): distributor (default) | ekspeditur | toko; pins take the colour of their main distributor / ekspeditur
  const mode = kmapMode(), top = (mp, ok) => { let k = null, v = -1; mp.forEach((x, key) => { if(x > v && (!ok || ok(key))) { v = x; k = key; } }); return k; };
  const sc = {dist: di};
  let grp = [];
  if(mode === 'distr'){
    grp = annotateDistrTargets(buildTree(shipRows(m, sc), ['distr']), m, sc).filter(n => B.dims.distr[n.key]).sort((a, b) => b.vol - a.vol)
      .map(n => ({key: n.key, n: distrLbl(n.key), vol: n.vol, pct: n.t && n.t.hasTgt ? n.t.pct : null, href: '#/distributor/' + enc(B.dims.distr[n.key])}));
  } else if(mode === 'eksp'){
    grp = ekspRows(m, sc).filter(o => D.frcEksp.has(o.key)).map(o => ({key: o.key, n: ekspLbl(o.key), vol: o.vol, pct: o.hasTgt ? o.pct : (o.vol ? 100 : null), of: o.of, href: ekspHref(o.key)}));
  }
  const PAL = ['#4E9BF5', '#F2A93B', '#B07CF0', '#3CC7B4', '#F06EA9'], colOf = new Map(grp.slice(0, 5).map((g, i) => [g.key, PAL[i]]));
  [...per.values()].forEach(o => { o.gk = mode === 'distr' ? top(o.byD) : mode === 'eksp' ? top(o.byE, k => D.frcEksp.has(k)) : null; });
  const kecNames = B.dims.kec || [], byKec = new Map(), exact = [];
  const words = s => ' ' + String(s || '').toUpperCase().replace(/[^A-Z]+/g, ' ') + ' ';
  const byLen = K.d.map((g, i) => [words(g.n), i]).sort((a, b) => b[0].length - a[0].length);
  [...per.values()].forEach(o => {
    const t = B.dims.toko[o.ti] || [];
    o.kec = -1;
    if(!t[0]) return;
    let ki = t[3] != null && t[3] >= 0 ? matchKec(kecNames[t[3]], K.d) : -1;
    if(ki < 0){ const txt = words(t[1] + ' ' + t[2]), hit = byLen.find(([w]) => txt.includes(w)); if(hit) ki = hit[1]; }
    o.kec = ki;
    if(t[4] != null && t[5] != null){
      o.x = (t[5] - 95) * 200; o.y = (6.2 - t[4]) * 200; exact.push(o);
      const at = K.d.findIndex(g => inPath(o.x, o.y, g.d));   // the coordinate decides the kecamatan
      if(at >= 0) ki = o.kec = at;
    }
    if(ki >= 0) (byKec.get(ki) || byKec.set(ki, []).get(ki)).push(o);
  });
  const [vx, vy, vw, vh] = K.vb, fs = Math.max(vw, vh) / 55, maxK = Math.max(1, ...[...byKec.values()].map(l => l.reduce((a, o) => a + o.vol, 0)));
  const shade = v => { const t = v > 0 ? 0.25 + 0.75 * Math.sqrt(v / maxK) : 0; const c = (a, b) => Math.round(a + (b - a) * t); return 'rgb(' + c(38, 95) + ',' + c(92, 208) + ',' + c(62, 143) + ')'; };   // green: deeper = less, brighter = more tonnage
  let shapes = '', labels = '', pins = '';
  K.d.forEach((g, i) => {
    const l = byKec.get(i), v = l ? l.reduce((a, o) => a + o.vol, 0) : 0;
    const tip = 'Kec. ' + g.n + '|' + (l ? fmt(v) + ' t ke ' + l.length + ' toko/gudang' : 'Tidak ada toko/gudang tujuan tercatat');
    shapes += '<path class="map-region live' + (l ? ' has' : '') + '" d="' + g.d + '" fill="' + shade(v) + '" stroke="#0F1A21" stroke-width="1.2" vector-effect="non-scaling-stroke" data-tip="' + esc(tip) + '"/>';
    if(Math.min(g.w, g.h) >= Math.max(vw, vh) * 0.08)
      labels += '<text x="' + g.cx + '" y="' + (g.cy - fs * 0.9) + '" class="map-lbl kec-lbl" style="font-size:' + (fs * 0.85).toFixed(1) + 'px">' + esc(g.n) + '</text>';
  });
  const maxT = Math.max(1, ...[...per.values()].map(o => o.vol));
  const pin = (o, x, y, approx) => {
    const t = B.dims.toko[o.ti], r = fs * (0.32 + 0.38 * Math.sqrt(o.vol / maxT));
    const tip = (t[1] || t[0]) + '|' + (t[2] ? t[2] + ' · ' : '') + (o.kec >= 0 ? 'Kec. ' + K.d[o.kec].n : '') + (approx ? ' (lokasi perkiraan)' : '') + '|' + fmt(o.vol) + ' t · ' + fmt(o.trips) + ' trip|' + [...o.distr].map(d => distrLbl(d)).filter(Boolean).join(', ');
    const fill = mode === 'toko' ? '' : ' style="fill:' + (colOf.get(o.gk) || '#8C9AA2') + '"';
    return '<circle class="toko-pin' + (approx ? ' approx' : '') + '"' + fill + (o.gk != null ? ' data-k="' + o.gk + '"' : '') + ' cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + r.toFixed(2) + '" stroke-width="1.2" vector-effect="non-scaling-stroke" data-tip="' + esc(tip) + '"/>';
  };
  byKec.forEach((l, ki) => {
    const g = K.d[ki];
    l.filter(o => o.x == null).sort((a, b) => b.vol - a.vol).forEach((o, j) => {
      const a = j * 2.4, rr = j ? fs * 0.75 * Math.sqrt(j) : 0;   // sunflower spread around the kecamatan centre
      pins += pin(o, g.cx + rr * Math.cos(a), g.cy + rr * Math.sin(a), true);
    });
  });
  /* Shared warehouse (gudang bersama): ship-tos on the same coordinate become one pin with a count. */
  const spot = new Map();
  exact.forEach(o => { const k = o.x.toFixed(2) + ',' + o.y.toFixed(2); (spot.get(k) || spot.set(k, []).get(k)).push(o); });
  [...spot.values()].sort((a, b) => b.reduce((x, o) => x + o.vol, 0) - a.reduce((x, o) => x + o.vol, 0)).forEach(l => {
    if(l.length === 1) return pins += pin(l[0], l[0].x, l[0].y, false);
    l.sort((a, b) => b.vol - a.vol); l.forEach(o => o.shared = l.length);
    const v = l.reduce((x, o) => x + o.vol, 0), tr = l.reduce((x, o) => x + o.trips, 0), r = fs * (0.42 + 0.38 * Math.sqrt(Math.min(1, v / maxT))), o = l[0];
    const tip = 'Gudang bersama · ' + l.length + ' ship-to|' + (o.kec >= 0 ? 'Kec. ' + K.d[o.kec].n + '|' : '') + fmt(v) + ' t · ' + fmt(tr) + ' trip|' + l.map(q => { const t = B.dims.toko[q.ti]; return (t[1] || t[0]) + ': ' + fmt(q.vol) + ' t'; }).join('|');
    pins += '<g class="toko-shared" data-tip="' + esc(tip) + '"><circle class="toko-pin"' + (mode === 'toko' ? '' : ' style="fill:' + (colOf.get(o.gk) || '#8C9AA2') + '"') + (o.gk != null ? ' data-k="' + o.gk + '"' : '') + ' cx="' + o.x.toFixed(1) + '" cy="' + o.y.toFixed(1) + '" r="' + r.toFixed(2) + '" stroke-width="1.6" vector-effect="non-scaling-stroke"/>' +
      '<text x="' + o.x.toFixed(1) + '" y="' + o.y.toFixed(1) + '" class="toko-n" style="font-size:' + (r * 1.1).toFixed(2) + 'px">' + l.length + '</text></g>';
  });
  const svg = '<svg class="map-svg" viewBox="' + K.vb.join(' ') + '" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta kecamatan dan toko/gudang tujuan di ' + esc(tc(name)) + '">' +
    shapes + labels + pins + '</svg>';
  const list = [...per.values()].filter(o => (B.dims.toko[o.ti] || [])[0]).sort((a, b) => b.vol - a.vol);
  const tot = list.reduce((a, o) => a + o.vol, 0), maxV = Math.max(1, ...list.map(o => o.vol));
  const nEx = list.filter(o => o.x != null).length, nAp = list.filter(o => o.x == null && o.kec >= 0).length, nNo = list.length - nEx - nAp;
  const top3 = list.slice(0, 3).reduce((a, o) => a + o.vol, 0);
  const kecTop = [...byKec.entries()].map(([ki, l]) => [K.d[ki].n, l.reduce((a, o) => a + o.vol, 0)]).sort((a, b) => b[1] - a[1])[0];
  const head = list.length ? '<b>' + list.length + ' toko/gudang</b> menerima ' + fmt(tot) + ' t di ' + esc(tc(name)) + '. 3 teratas menyerap <b>' + pctTxt(pctOf(top3, tot)) + '</b> tonase' +
    (kecTop ? '; kecamatan terpadat: <b>' + esc(kecTop[0]) + '</b>.' : '.') : 'Belum ada toko/gudang tujuan tercatat di ' + esc(tc(name)) + ' untuk filter ini.';
  const dot = o => o.x != null ? 'var(--red)' : o.kec >= 0 ? '#F08A90' : '#5B6B74';
  const rows = list.map((o, r) => { const t = B.dims.toko[o.ti]; return '<div class="mh-row mh-row2"><span class="mh-rk">' + (r + 1) + '</span><span class="mh-nm"><b class="mh-dot" style="background:' + dot(o) + '"></b>' + esc(t[1] || t[0]) + '<small>' + esc(o.kec >= 0 ? 'Kec. ' + K.d[o.kec].n : (t[2] || 'lokasi belum diketahui')) + (o.shared ? ' · gudang bersama (' + o.shared + ')' : '') + '</small><i style="width:' + (o.vol / maxV * 100) + '%;background:#5FD08F"></i></span><span class="mh-tn">' + fmt(o.vol) + ' t</span><span class="mh-pc">' + pctTxt(pctOf(o.vol, tot)) + '</span></div>'; }).join('');
  const gMax = Math.max(1, ...grp.map(g => g.vol)), gTot = grp.reduce((a, g) => a + g.vol, 0);
  const nToko = k => list.filter(o => o.gk === k).length;
  const gRows = grp.map((g, r) => '<' + (g.href ? 'a href="' + g.href + '"' : 'div') + ' class="mh-row" data-kd="' + g.key + '"><span class="mh-rk">' + (r + 1) + '</span><span class="mh-nm"><b class="mh-dot" style="background:' + (colOf.get(g.key) || '#8C9AA2') + '"></b>' + esc(g.n) +
    '<small>' + (mode === 'distr' ? nToko(g.key) + ' toko utama · ' : '') + pctTxt(pctOf(g.vol, gTot)) + ' tonase</small><i style="width:' + (g.vol / gMax * 100) + '%;background:' + tierColor(g.pct) + '"></i></span>' +
    '<span class="mh-tn">' + fmt(g.vol) + ' t</span><span class="mh-pc" style="color:' + heroTier(g.pct) + '">' + pctTxt(g.pct) + '</span><span class="mh-so">' + (mode === 'eksp' ? (g.of == null ? '–' : pctTxt(g.of)) : '') + '</span></' + (g.href ? 'a' : 'div') + '>').join('');
  const modeBar = '<span class="kmode">' + [['distr', 'Distributor'], ['eksp', 'Ekspeditur'], ['toko', 'Toko']].map(([k, l]) => '<button class="' + (k === mode ? 'on' : '') + '" data-kmode="' + k + '">' + l + '</button>').join('') + '</span>';
  const gKey = (grp.length > 5 || mode === 'eksp' ? '<span><i style="background:#8C9AA2;border-radius:50%"></i>' + (mode === 'eksp' ? (grp.length > 5 ? 'lainnya / ' : '') + 'tanpa FRC (FOT, ambil sendiri)' : 'lainnya') + '</span>' : '');
  const note = list.length ? '<div class="mh-note">' + nEx + ' toko/gudang berkoordinat · ' + nAp + ' di pusat kecamatan (perkiraan)' + (nNo ? ' · ' + nNo + ' belum terpetakan' : '') + '. Koordinat bisa diupload di halaman admin, slot <i>Lokasi gudang / toko</i>.</div>' : '';
  return heroPanel({photo: pname, cls: 'kmap', eyebrow: 'Peta kecamatan &amp; toko/gudang tujuan · <a href="#/provinsi/' + enc(pname) + '">peta ' + esc(tc(pname)) + ' &rsaquo;</a>', head,
      chips: kpiChips(total(m, {dist: di})).concat(heroChip('Toko / gudang', fmt(list.length) + ' <small>di ' + byKec.size + '/' + K.d.length + ' kec.</small>', (list.length ? pctTxt(pctOf(nEx + nAp, list.length)) : '–') + ' terpetakan · ' + nEx + ' berkoordinat')),
      listHead: modeBar, listCols: mode === 'toko' ? ['Realisasi', 'Porsi'] : mode === 'eksp' ? ['Realisasi', 'Capaian', 'Real/SO'] : ['Realisasi', 'Capaian', ''],
      rows: mode === 'toko' ? rows : gRows || '<div class="mh-note">Tidak ada ' + (mode === 'eksp' ? 'ekspeditur FRC' : 'distributor') + ' untuk filter ini.</div>',
      key: (mode === 'toko' ? '<span><i style="background:var(--red);border-radius:50%"></i>koordinat gudang</span><span><i style="background:#F08A90;border-radius:50%"></i>perkiraan</span>' : gKey + (nAp ? '<span>titik pudar = lokasi perkiraan</span>' : '')) +
        '<span><i style="background:linear-gradient(90deg,#265C3E,#5FD08F)"></i>tonase kecamatan (makin terang makin banyak)</span>',
      note, map: svg, hint: 'Arahkan kursor ke titik untuk nama, alamat & tonase'});
}

/* ===================== Shipping destinations: distributor & ship-to (toko/gudang) ===================== */
/* B.ship rows: [day, inc, src, prov, dist, distributor, eksp, toko, ton, trips] */
const SHK = {prov: r => r[3], dist: r => r[4], distr: r => r[5], eksp: r => r[6], toko: r => r[7]};
function shipRows(m, sc){
  sc = sc || {};
  const per = m == null ? null : toPer(m);
  return B.ship.filter(r => (per == null || inPer(r[0], per)) && incOk(r[1]) && srcOkS(r[2], sc) && kemOk(r[10]) && (sc.prov != null || areaProvOk(r[3])) &&
    (sc.prov == null || r[3] === sc.prov) && (sc.dist == null || r[4] === sc.dist) &&
    (sc.distr == null || r[5] === sc.distr) && (sc.eksp == null || r[6] === sc.eksp));
}
function buildTree(rows, levels, depth){
  depth = depth || 0;
  const lv = levels[depth], groups = new Map();
  rows.forEach(r => { const k = SHK[lv](r); let g = groups.get(k); if(!g){ g = {lv, key: k, vol: 0, az: 0, trips: 0, last: -1, rows: []}; groups.set(k, g); } g.vol += r[8]; if(achKem(r[10])) g.az += r[8]; g.trips += r[9]; g.last = Math.max(g.last, r[0]); g.rows.push(r); });
  const nodes = [...groups.values()].sort((a, b) => b.vol - a.vol);
  nodes.forEach(n => { if(depth + 1 < levels.length) n.children = buildTree(n.rows, levels, depth + 1); n.rows = null; });
  return nodes;
}
const LV_LBL = {prov: 'Provinsi', dist: 'Distrik', distr: 'Distributor', eksp: 'Ekspeditur', toko: 'Toko / gudang tujuan'};
function nodeLabel(n){
  if(n.lv === 'prov') return esc(tc(B.dims.prov[n.key]));
  if(n.lv === 'dist') return esc(tc(B.dims.dist[n.key]));
  if(n.lv === 'distr') return esc(distrLbl(n.key) || '(tanpa distributor)');
  if(n.lv === 'eksp') return esc(ekspLbl(n.key));
  const t = B.dims.toko[n.key];
  if(!t || !t[0]) return '<span class="t-unk">Toko belum tercatat</span><small>data tujuan tidak tersedia untuk periode ini</small>';
  return esc(t[1] || t[0]) + '<small>' + esc([t[2], t[0]].filter(Boolean).join(' · ')) + '</small>';
}
function nodeHref(n){
  if(n.lv === 'prov') return '#/provinsi/' + enc(B.dims.prov[n.key]);
  if(n.lv === 'dist') return '#/distrik/' + enc(B.dims.dist[n.key]);
  if(n.lv === 'distr') return B.dims.distr[n.key] ? '#/distributor/' + enc(B.dims.distr[n.key]) : null;
  if(n.lv === 'eksp') return ekspHref(n.key);
  return null;
}
/* Expandable table: click a row to open the next level; "›" opens that item's own page.
   opt.tgt = {prov: Map, dist: Map} summaries (from summarize) to show target & achievement at those levels. */
/** Put distributor targets on tree nodes (levels distr / prov / dist), within scope `base` ({dist} or {distr}). */
function annotateDistrTargets(nodes, m, base, ctx, only){
  ctx = ctx || {};
  nodes.forEach(n => {
    const here = Object.assign({}, ctx);
    if(n.lv === 'distr' || n.lv === 'prov' || n.lv === 'dist') here[n.lv] = n.key;
    if((n.lv === 'distr' || n.lv === 'prov' || n.lv === 'dist') && (!only || only.includes(n.lv))){
      const sc = Object.assign({}, base, here), tg = distrTargets(m, sc);
      const t = sc.distr != null ? (tg.get(sc.distr) || 0) : [...tg.values()].reduce((a, v) => a + v, 0);
      n.t = t > 0 ? {hasTgt: true, tgt: t, pct: pctOf(n.az, t)} : {hasTgt: false};
    }
    if(n.children) annotateDistrTargets(n.children, m, base, here, only);
  });
  return nodes;
}
function treeHTML(nodes, levels, opt){
  opt = opt || {};
  const id = 'tr' + Math.random().toString(36).slice(2, 8);
  const hasT = !!(opt.tgt || opt.distrTgt), soCols = !opt.distrTgt;   // distributor targets: no SO per distributor
  let html = '<div class="tree-tools"><span>' + levels.map(l => LV_LBL[l]).join(' &rsaquo; ') + '</span><span><a data-tree-all="' + id + ':1">Buka semua</a> · <a data-tree-all="' + id + ':0">Tutup semua</a></span></div>' +
    '<div class="tbl-wrap' + (nodes.length > 15 ? ' tall' : '') + '"><table class="tbl tree" id="' + id + '"><thead><tr><th class="fz fz0 fzl">Tujuan</th><th class="num">Realisasi (t)</th>' + (hasT ? '<th class="num">Target MTD (t)</th><th class="num">Capaian</th>' + (soCols ? '<th class="num">SO (t)</th><th class="num">Real/SO</th>' : '') : '') + '<th class="num">Trip</th><th class="num">Kirim terakhir</th><th></th></tr></thead><tbody>';
  const walk = (list, depth, parent) => list.forEach((n, i) => {
    const path = parent ? parent + '-' + i : String(i), kids = n.children && n.children.length;
    const t = n.t || (hasT && opt.tgt && opt.tgt[n.lv] ? opt.tgt[n.lv].get(n.key) : null);
    const href = nodeHref(n);
    html += '<tr class="lv' + depth + (depth ? ' tr-hide' : '') + (kids ? ' has-kids' : '') + '" data-path="' + path + '" data-parent="' + (parent || '') + '">' +
      '<td class="name fz fz0 fzl" style="--d:' + depth + '"><span class="caret-t">' + (kids ? '▸' : '') + '</span><span class="t-lbl">' + nodeLabel(n) + '</span>' + (kids ? '<span class="t-cnt">' + n.children.length + ' ' + LV_LBL[levels[depth + 1]].toLowerCase() + '</span>' : '') + '</td>' +
      '<td class="num">' + fmt(n.vol) + '</td>' +
      (hasT ? (t && t.hasTgt ? '<td class="num">' + fmt(t.tgt) + '</td><td class="num pc" style="color:' + tierColor(t.pct) + '">' + pctTxt(t.pct) + '</td>' : '<td class="num">–</td><td class="num">–</td>') +
        (!soCols ? '' : t && t.so ? '<td class="num">' + fmt(t.so) + '</td><td class="num pc">' + pctTxt(t.of) + '</td>' : '<td class="num">–</td><td class="num">–</td>') : '') +
      '<td class="num">' + fmt(n.trips) + '</td><td class="num">' + (n.last >= 0 ? dayLabel(B.days[n.last]) : '–') + '</td>' +
      '<td class="num">' + (href ? '<a class="t-go" href="' + href + '" title="Buka halaman">&rsaquo;</a>' : '') + '</td></tr>';
    if(kids) walk(n.children, depth + 1, path);
  });
  walk(nodes, 0, '');
  return html + '</tbody></table></div>';
}
function treeToggle(row){
  const tb = row.closest('tbody'), path = row.dataset.path, open = !row.classList.contains('open');
  row.classList.toggle('open', open);
  $('.caret-t', row).textContent = open ? '▾' : '▸';
  $$('tr', tb).forEach(r => {
    if(open ? r.dataset.parent === path : r.dataset.path.startsWith(path + '-')){
      r.classList.toggle('tr-hide', !open);
      if(!open && r.classList.contains('open')){ r.classList.remove('open'); $('.caret-t', r).textContent = '▸'; }
    }
  });
}
function treeAll(id, open){
  $$('#' + id + ' tbody tr').forEach(r => {
    if(r.dataset.parent) r.classList.toggle('tr-hide', !open);
    if(r.classList.contains('has-kids') && !(r.dataset.lazy && !r.dataset.loaded)){ r.classList.toggle('open', open); $('.caret-t', r).textContent = open ? '▾' : '▸'; }
  });
}
function distrSummary(rows){
  const per = new Map();
  rows.forEach(r => { let o = per.get(r[5]); if(!o){ o = {key: r[5], frc: 0, fc: 0, nf: 0, nfz: 0, vol: 0, trips: 0, prov: new Set(), dist: new Set(), toko: new Set(), last: -1}; per.set(r[5], o); }
    if(r[1] === 0) o.frc += r[8]; else if(isFrc(r[1])) o.fc += r[8]; else o.nf += r[8];
    if(achKem(r[10])) o.nfz += r[8];   // realisasi counted against the distributor target (kemasan with targets, every selected incoterm)
    o.vol += r[8]; o.trips += r[9]; o.prov.add(r[3]); o.dist.add(r[4]); if(B.dims.toko[r[7]] && B.dims.toko[r[7]][0]) o.toko.add(r[7]); o.last = Math.max(o.last, r[0]); });
  return [...per.values()];
}

/* Distributor target for the selected incoterm(s), up to the last data day of the period: FOT/FOB/FAS rows of B.targets
   (t[8]) plus FRC rows of B.tdistr (Tactical, per distributor; FRC in B.targets is split per ekspeditur instead). */
function distrTargets(per, sc){
  per = toPer(per); sc = sc || {};
  const out = new Map(), m0 = per.from.slice(0, 7), m1 = per.to.slice(0, 7);
  B.targets.concat(B.tdistr || []).forEach(t => {
    if(!kemOk(tKem(t)) || !areaProvOk(t[1]) || !(t[8] >= 0) || t[0] < m0 || t[0] > m1 || !incOk(tInc(t)) || !srcOkS(t[4], sc) || (sc.distr != null && t[8] !== sc.distr) || (sc.prov != null && t[1] !== sc.prov) || (sc.dist != null && t[2] !== sc.dist)) return;
    out.set(t[8], (out.get(t[8]) || 0) + targetBetween(t, per.from, per.tgtTo));
  });
  return out;
}
const blankDistr = k => ({key: k, frc: 0, fc: 0, nf: 0, nfz: 0, vol: 0, trips: 0, prov: new Set(), dist: new Set(), toko: new Set(), last: -1});

/* Distributor table (realisasi per incoterm, target distributor & capaian, coverage) for a scope: {} = all, {prov}, {src}. */
/* Distributor table, one block per kemasan (ZAK / CURAH / JUMBO BAG) in the kemasan filter; list/rows/tot stay overall. */
function distrTable(m, sc, limit){
  const all = distrTableOne(m, sc, limit);
  if(!KEMS().length) return all;
  const ks = F.kem === 'ALL' ? KEMS().slice() : [F.kem], sv = F.kem;
  const html = ks.map(k => { F.kem = k; try { const o = distrTableOne(m, sc, limit); return o.list.length ? '<div class="inc-grp" style="--c:var(--black)"><b>' + esc(k) + '</b><small>' + fmt(o.list.length) + ' distributor · ' + fmt(o.tot) + ' t</small></div>' + o.html : ''; } finally { F.kem = sv; } }).join('');
  return Object.assign(all, {html: html || all.html});
}
function distrTableOne(m, sc, limit){
  const rows = shipRows(m, sc), list = distrSummary(rows), tg = distrTargets(m, sc);
  tg.forEach((v, k) => { if(B.dims.distr[k] && !list.some(o => o.key === k)) list.push(blankDistr(k)); });
  list.forEach(o => { o.tgt = tg.get(o.key); o.pct = o.tgt ? pctOf(o.nfz, o.tgt) : (o.vol > 0 ? 100 : null); });   // no target but shipped = 100%
  list.sort((a, b) => b.vol - a.vol || (b.tgt || 0) - (a.tgt || 0));
  list.forEach((o, i) => o.rank = i + 1);
  const tot = list.reduce((a, o) => a + o.vol, 0), unk = rows.filter(r => !(B.dims.toko[r[7]] && B.dims.toko[r[7]][0])).reduce((a, r) => a + r[8], 0);
  const cols = [{h:'#', cls:'rank', val: o => o.rank, html: o => '#' + o.rank}, {h:'Distributor', cls:'name', val: o => distrLbl(o.key)},
    {h:'FRC TRUK (t)', num:true, val: o => o.frc, html: o => o.frc ? fmt(o.frc) : '–'}].concat(list.some(o => o.fc) ? [
    {h:'FRC KONTAINER/KAPAL (t)', num:true, val: o => o.fc, html: o => o.fc ? fmt(o.fc) : '–'}] : [], [{h:'Ambil sendiri (t)', num:true, val: o => o.nf, html: o => o.nf ? fmt(o.nf) : '–'},
    {h:'Total (t)', num:true, val: o => o.vol, html: o => '<b>' + fmt(o.vol) + '</b>'}, {h:'Share', num:true, val: o => o.vol / (tot || 1), html: o => pctTxt(pctOf(o.vol, tot))}].concat(tg.size ? [
    {h:'Target distributor ' + tgtLbl() + ' (t)', num:true, val: o => o.tgt || 0, html: o => o.tgt ? fmt(o.tgt) : '–'},
    {h:'Capaian', num:true, val: o => o.pct == null ? -1 : o.pct, html: o => o.pct == null ? '–' : '<b style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</b>'}] : [], [
    {h:'Trip', num:true, val: o => o.trips}, {h:'Provinsi', num:true, val: o => o.prov.size}, {h:'Distrik', num:true, val: o => o.dist.size},
    {h:'Toko / gudang', num:true, val: o => o.toko.size}, {h:'Kirim terakhir', num:true, val: o => o.last, html: o => o.last >= 0 ? dayLabel(B.days[o.last]) : '–'}]));
      const sum = f => list.reduce((a, o) => a + (f(o) || 0), 0), uni = k => { const u = new Set(); list.forEach(o => o[k].forEach(x => u.add(x))); return u.size; };
      const wt = list.filter(o => o.tgt), tTgt = sum(o => o.tgt), tNfz = wt.reduce((a, o) => a + o.nfz, 0);
      const T = {'FRC TRUK (t)': fmt(sum(o => o.frc)), 'FRC KONTAINER/KAPAL (t)': fmt(sum(o => o.fc)), 'Ambil sendiri (t)': fmt(sum(o => o.nf)), 'Total (t)': '<b>' + fmt(tot) + '</b>', 'Share': '100%',
        'Trip': fmt(sum(o => o.trips)), 'Provinsi': fmt(uni('prov')), 'Distrik': fmt(uni('dist')), 'Toko / gudang': fmt(uni('toko')), 'Kirim terakhir': list.some(o => o.last >= 0) ? dayLabel(B.days[Math.max(...list.map(o => o.last))]) : '–'};
      T['Target distributor ' + tgtLbl() + ' (t)'] = tTgt ? fmt(tTgt) : '–'; T['Capaian'] = tTgt ? '<b style="color:' + tierColor(pctOf(tNfz, tTgt)) + '">' + pctTxt(pctOf(tNfz, tTgt)) + '</b>' : '–';
      cols.forEach(c => { if(T[c.h] != null) c.total = T[c.h]; });
  return {list, rows, tot, unk, html: list.length ? table(cols, list, {go: o => '#/distributor/' + enc(B.dims.distr[o.key]), limit: limit || 400, total: true}) : '<div class="dp-review">Tidak ada pengiriman untuk filter ini.</div>'};
}
function pageDistrList(){
  const m = PER(), D2 = distrTable(m, {}), list = D2.list, tot = D2.tot, unk = D2.unk;
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], ['Distributor']]) +
    '<div class="dp-head"><div><div class="nm">Distributor &amp; Tujuan Kirim</div><div class="sub">' + fmt(list.length) + ' distributor · ' + incName() + ' · ' + periodSub() + '</div></div></div>' +
    (unk > 0 ? '<div class="dp-review">' + pctTxt(pctOf(unk, tot)) + ' tonase periode ini belum punya data toko/gudang tujuan (' + fmt(unk) + ' t). Distributor tetap tercatat; toko tujuan akan muncul begitu kolom KODE_TOKO/NAMA_TOKO diisi di master Excel.</div>' : '') +
    D2.html +
    '</div>', 'wilayah');
}

function pageDistr(name){
  const di = B.dims.distr.indexOf(name);
  if(di < 0) return notFound('Distributor ' + name);
  const m = PER(), sc = {distr: di}, rows = shipRows(m, sc);
  const s = distrSummary(rows)[0] || blankDistr(di), tgt = distrTargets(m, sc).get(di), pct = tgt ? pctOf(s.nfz, tgt) : null;
  const ekspTree = buildTree(rows, ['eksp']);
  // per incoterm allowed by the incoterm/moda filter: realisasi vs distributor target (FRC from Tactical, others SNOP/Tactical)
  const incRows = INCS().filter((n, i) => incOk(i)).map(n => withInc({incFix: n}, () => {
    const z = distrSummary(shipRows(m, sc))[0] || blankDistr(di), t = distrTargets(m, sc).get(di) || 0;
    return {n, vol: z.vol, az: z.nfz, trips: z.trips, tgt: t, pct: t ? pctOf(z.nfz, t) : null, gap: z.nfz - t};
  })).filter(o => o.vol > 0 || o.tgt > 0);
  const incT = incRows.reduce((a, o) => ({vol: a.vol + o.vol, az: a.az + (o.tgt ? o.az : 0), tgt: a.tgt + o.tgt, trips: a.trips + o.trips}), {vol: 0, az: 0, tgt: 0, trips: 0});
  const incCols = [{h:'Incoterm', cls:'name', val: o => incLbl(o.n), html: o => esc(incLbl(o.n))},
    {h:'Realisasi (t)', num:true, val: o => o.vol, html: o => fmt(o.vol)},
    {h:'Target ' + tgtLbl() + ' MTD (t)', num:true, val: o => o.tgt, html: o => o.tgt ? fmt(o.tgt) : '–'},
    {h:'Capaian', num:true, val: o => o.pct == null ? -1 : o.pct, html: o => o.pct == null ? '–' : '<b style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</b>'},
    {h:'Gap (t)', num:true, val: o => o.tgt ? o.gap : 0, html: o => o.tgt ? '<span style="color:' + (o.gap < 0 ? 'var(--red)' : 'var(--green)') + '">' + signed(o.gap) + '</span>' : '–'},
    {h:'Trip', num:true, val: o => o.trips}];
  if(incRows.length > 1){
    const tp = incT.tgt ? pctOf(incT.az, incT.tgt) : null;
    incCols[0].total = '<b>Total</b>'; incCols[1].total = '<b>' + fmt(incT.vol) + '</b>'; incCols[2].total = incT.tgt ? '<b>' + fmt(incT.tgt) + '</b>' : '–';
    incCols[3].total = tp == null ? '–' : '<b style="color:' + tierColor(tp) + '">' + pctTxt(tp) + '</b>'; incCols[4].total = incT.tgt ? signed(incT.az - incT.tgt) : '–'; incCols[5].total = fmt(incT.trips);
  }
  // monthly volume, last 6 months; bar click lists that month's ship-to points
  const idx = D.months.indexOf(m.month), months = m.kind === 'month' ? D.months.slice(Math.max(0, idx - 5), idx + 1) : D.months.filter(x => x >= m.from.slice(0, 7) && x <= m.month);
  const all = shipRows(null, sc);
  const data = months.map(mm => {
    const rs = all.filter(r => D.dayMonth[r[0]] === mm), t = distrTargets(mm, sc).get(di) || null;
    const az = distrSummary(rs.filter(r => incOk(r[1]) && kemOk(r[10])))[0];
    return {label: monthShort(mm), long: monthLabel(mm), vol: rs.reduce((a, r) => a + r[8], 0), tgt: t, pct: t ? pctOf(az ? az.nfz : 0, t) : null, month: mm};
  });
  TRENDS.dm = {id: 'dm', data, click: i => {
    const mm = data[i].month, nodes = buildTree(shipRows(mm, sc), ['prov', 'dist', 'toko']);
    openModal(tc(name) + ' · ' + monthLabel(mm), 'Tujuan kirim per provinsi › distrik › toko', nodes.length ? treeHTML(annotateDistrTargets(nodes, mm, {distr: di}), ['prov', 'dist', 'toko'], {distrTgt: true}) : '<div class="dp-review">Tidak ada pengiriman di bulan ini.</div>');
  }};
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], ['Distributor', '#/distributor'], [coLbl(tc(name)) || tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(coLbl(tc(name)) || tc(name)) + '</div><div class="sub">Distributor · ' + incName() + ' · ' + periodSub() + '</div></div></div>' +

    '<div class="dp-grid4">' +
      kpi('box', 'var(--black)', 'Realisasi', fmt(s.vol) + ' <small>ton</small>', {hint: 'FRC TRUK ' + fmt(s.frc) + (s.fc ? ' · FRC KONTAINER/KAPAL ' + fmt(s.fc) : '') + ' · ambil sendiri ' + fmt(s.nf)}) +
      kpi('target', tgt ? tierColor(pct) : 'var(--black)', 'Target ' + tgtLbl() + ' MTD', tgt ? fmt(tgt) + ' <small>ton</small>' : '–', {hint: tgt ? 'capaian <b style="color:' + tierColor(pct) + '">' + pctTxt(pct) + '</b> · ' + (s.nfz - tgt < 0 ? 'sisa ' + fmt(tgt - s.nfz) : 'lebih ' + fmt(s.nfz - tgt)) + ' t' : 'tidak ada target'}) +
      kpi('trend', 'var(--black)', 'Trip', fmt(s.trips), {hint: s.last >= 0 ? 'terakhir ' + dayLabel(B.days[s.last]) : ''}) +
      kpi('pin', 'var(--amber)', 'Toko / gudang', fmt(s.toko.size), {hint: s.prov.size + ' provinsi · ' + s.dist.size + ' distrik'}) +
    '</div>' +
    tabsHTML('distr', [{k: 'ringkas', label: 'Incoterm & pengangkut', html:
    '<div class="section-head"><div class="section-title">Target &amp; realisasi per incoterm</div><div class="section-sub">' + incName() + ' · ' + periodSub() + '</div></div>' +
    (incRows.length ? table(incCols, incRows, {total: incRows.length > 1}) : '<div class="dp-review">Tidak ada realisasi maupun target untuk periode ini.</div>') +
    '<div class="dp-cols"><div class="dp-panel"><div class="pt">Volume bulanan (ton)<span class="pt-sub">' + (data.some(d => d.tgt) ? 'garis = target distributor' : '') + '</span></div>' + barsHTML('dm', data) + '<div class="trend-detail">Klik batang untuk tujuan kirim bulan itu.</div></div>' +
    '<div class="dp-panel"><div class="pt">Diangkut oleh<span class="pt-sub">' + fmt(ekspTree.filter(n => B.dims.eksp[n.key]).length) + ' ekspeditur</span></div>' + (incGroups(() => volTable(buildTree(shipRows(m, sc), ['eksp']), 'Ekspeditur', n => ekspLbl(n.key), n => nodeHref(n)), () => fmt(shipRows(m, sc).reduce((a, r) => a + r[8], 0)) + ' t') || '<div class="empty">Tidak ada pengiriman untuk filter ini.</div>') + '</div></div>' },
    {k: 'tujuan', label: 'Tujuan pengiriman', html: (rows.length ? treeHTML(annotateDistrTargets(buildTree(rows, ['prov', 'dist', 'toko']), m, {distr: di}), ['prov', 'dist', 'toko'], {distrTgt: true}) : '<div class="dp-review">Tidak ada pengiriman untuk filter ini.</div>')}]) +
    '</div>', 'wilayah');
}

/* New version check: a long-open tab keeps running the code it loaded. Every 5 minutes and whenever the tab comes back
   into view, index.html is re-read; when its app.js version differs from the running one, a bar offers a reload. */
function watchNewVersion(){
  const own = document.querySelector('script[src*="app.js"]'), mine = own && own.src.split('?v=')[1];
  if(!mine) return;
  const check = () => fetch('index.html', {cache: 'no-store'}).then(r => r.text()).then(t => {
    const m = t.match(/app\.js\?v=([^"'&]+)/);
    if(m && m[1] !== mine && !$('#newVersion')){
      document.body.insertAdjacentHTML('beforeend', '<div id="newVersion" class="new-ver">Versi dashboard atau data terbaru tersedia. <a href="#" onclick="location.reload();return false">Muat ulang</a></div>');
    }
  }).catch(() => {});
  setInterval(check, 5 * 60 * 1000);
  document.addEventListener('visibilitychange', () => { if(document.visibilityState === 'visible') check(); });
}
/* ===================== Pages ===================== */
const view = () => $('#view');
/* Page header photo: assets/bg/bg.js sets window.BG = {key: {f: file, t: title, a: author, l: licence, u: source page}}.
   Keys: prov-<province>, plant-<plant>, moda-darat / moda-laut / moda-kontainer. No entry = plain header. */
const slug = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const PROV_BG = {'RIAU KEPULAUAN': 'kepulauan-riau', 'KEPULAUAN RIAU': 'kepulauan-riau'};
const provBg = n => 'prov-' + (PROV_BG[n] || slug(n));
function plantBg(n){
  const u = String(n || '').toUpperCase();
  if(/INDARUNG/.test(u)) return 'plant-indarung';
  if(/TELUK BAYUR|TRANSHPMNT|BATAM|BELAWAN|MALAHAYATI|LHOKSEUMAWE/.test(u)) return 'moda-laut';
  return 'moda-darat';
}
function pageBgKey(){
  const [a, b] = (location.hash.replace(/^#\/?/, '') || '').split('/').map(decodeURIComponent);
  if(a === 'provinsi') return b ? provBg(b) : 'moda-darat';
  if(a === 'distrik' && b){ const di = D.distIdx[b]; return di == null ? null : provBg(B.dims.prov[B.dims.distProv[di]]); }
  if(a === 'plant') return b ? plantBg(b) : 'plant-indarung';
  if(a === 'kapal') return 'moda-laut';
  if(a === 'distributor') return F.moda === 'LAUT' ? 'moda-laut' : 'moda-darat';
  if(a === 'ekspeditur' || a === 'armada' || a === 'kontrak' || a === 'prognosa') return 'moda-darat';
  return null;
}
/* Region pages (provinsi / distrik) show the regional emblem in the header instead of a photo:
   assets/logo/logo.js sets window.LOGO = {key: {f, t, a, l, u, p: province, n: district shape name or null}}. */
function regionLogo(){
  const L = window.LOGO; if(!L) return null;
  const [a, b] = (location.hash.replace(/^#\/?/, '') || '').split('/').map(decodeURIComponent);
  if(!b || (a !== 'provinsi' && a !== 'distrik')) return null;
  const pname = a === 'provinsi' ? b : (D.distIdx[b] == null ? null : B.dims.prov[B.dims.distProv[D.distIdx[b]]]);
  if(!pname) return null;
  const prov = mapProv(pname), provLogo = L['prov-' + slug(prov)];
  if(a === 'provinsi') return provLogo || null;
  const list = Object.values(L).filter(o => o.p === prov && o.n), i = matchDistrict(b, list);
  return i >= 0 ? list[i] : provLogo || null;
}
// focal point per header photo (the subject, not the sky): wide headers show only a band of the photo
const BG_POS = {'moda-darat': 'center 92%', 'moda-laut': 'center 88%', 'moda-kontainer': 'center 62%', 'plant-indarung': 'center 68%', 'prov-sumatera-utara': 'center 60%', 'prov-sumatera-barat': 'center 75%'};
function applyPageBg(nav){
  const h = $('#view .dp-head'), lg = h && regionLogo();
  if(lg){
    h.classList.add('has-logo');
    h.style.setProperty('--logo', 'url("assets/logo/' + lg.f + '")');
    h.insertAdjacentHTML('afterbegin', '<img class="dp-logo" src="assets/logo/' + lg.f + '" alt="' + esc(lg.t) + '">');
    h.insertAdjacentHTML('beforeend', '<a class="bg-credit" href="#/kredit" title="' + esc(lg.t + ' (' + lg.l + ')') + '">' + esc(lg.t) + '</a>');
    return;
  }
  const bg = (window.BG || {})[pageBgKey()];
  if(!bg || !h) return;
  h.classList.add('has-bg');
  h.style.setProperty('--bg', 'url("assets/bg/' + bg.f + '")');
  h.style.setProperty('--bgpos', BG_POS[pageBgKey()] || 'center 45%');
  // eyebrow = the menu this page belongs to, so the title reads as "section > page"
  const nb = nav && $('#mainNav [data-nav="' + nav + '"] .navbtn, #mainNav a[data-nav="' + nav + '"]'), eb = nb ? nb.textContent.replace(/[▼\s]+$/, '').trim() : '';
  if(eb && h.firstElementChild) h.firstElementChild.insertAdjacentHTML('afterbegin', '<div class="dp-eyebrow">' + esc(eb) + '</div>');
  h.insertAdjacentHTML('beforeend', '<a class="bg-credit" href="#/kredit" title="' + esc(bg.t + ' — ' + bg.a + ' (' + bg.l + ')') + '">Foto: ' + esc(bg.t) + '</a>');
}
/* Visit log (admin page "Kunjungan"): one beacon per page view after unlock. Anonymous: random visitor id kept in
   this browser, session id per tab, page path, device class. No name, no IP (the server keeps only city). */
const rid = () => Math.random().toString(36).slice(2, 12).padEnd(10, '0');
let lastView = '';
function trackView(){
  const p = location.hash.replace(/^#\/?/, '').split('/').slice(0, 2).map(x => { try { return decodeURIComponent(x); } catch(e){ return x; } }).join('/').slice(0, 120);
  if(p === lastView || location.hostname === 'localhost') return; lastView = p;
  let v = null, s = null;
  try { v = localStorage.getItem('osp_vid'); if(!v){ v = rid(); localStorage.setItem('osp_vid', v); } } catch(e){ v = rid(); }
  try { s = sessionStorage.getItem('osp_sid'); if(!s){ s = rid(); sessionStorage.setItem('osp_sid', s); } } catch(e){ s = rid(); }
  const w = Math.min(screen.width, innerWidth), d = w < 700 ? 'hp' : w < 1100 ? 'tablet' : 'desktop';
  try { navigator.sendBeacon('/api/hit', new Blob([JSON.stringify({v, s, p, d})], {type: 'text/plain'})); } catch(e){}
}
function page(html, nav){ view().innerHTML = '<div class="page-fade">' + html + '</div>'; applyPageBg(nav); setActiveNav(nav); setStick(); window.scrollTo(0, 0); }
/* Photo credits (licence requirement for CC BY / CC BY-SA photos). */
function pageCredits(){
  const rows = Object.entries(Object.assign({}, window.BG || {}, window.LOGO || {})).map(([k, v]) => '<tr><td class="name">' + esc(v.t) + '<small>' + esc(k) + '</small></td><td>' + esc(v.a) + '</td><td>' + esc(v.l) + '</td><td><a href="' + esc(v.u) + '" target="_blank" rel="noopener">Sumber &rsaquo;</a></td></tr>').join('');
  page('<div class="wrap">' + crumb([['Kredit foto']]) + '<div class="dp-head"><div><div class="nm">Kredit foto</div><div class="sub">Foto dan lambang daerah di header halaman, berlisensi bebas / domain publik (Wikimedia Commons)</div></div></div>' +
    (rows ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Foto</th><th>Fotografer</th><th>Lisensi</th><th></th></tr></thead><tbody>' + rows + '</tbody></table></div>' : '<div class="dp-review">Belum ada foto.</div>') + '</div>', '');
}
function crumb(parts){ return '<div class="crumb"><a href="#/">Home</a>' + parts.map(p => '<span class="sep">/</span>' + (p[1] ? '<a href="' + p[1] + '">' + esc(p[0]) + '</a>' : esc(p[0]))).join('') + '</div>'; }
function periodSub(){ return PER().sub; }
function pageHome(){
  const m = PER(), t = total(m);
  const pm = summarize('prov', m);
  const provRows = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0);
  const pp = prevPer(m), prev = pp.to >= B.days[0] && pp.from >= B.days[0] ? total(pp) : null;
  const chg = prev && prev.vol ? (t.vol / prev.vol - 1) * 100 : null;
  const isLatest = m.isLatest && m.kind === 'month';
  const incTxt = incName();
  const prevTxt = m.kind === 'month' ? 'periode sama bulan lalu' : 'periode ' + m.days + ' hari sebelumnya';
  const head = 'Volume ' + incTxt + ' ' + esc(m.label) + ' <b>' + fmt(t.vol) + ' ton</b>' +
    (t.pct != null ? (', ') + (isLatest ? 'sudah ' : '') + 'mencapai <b>' + pctTxt(t.pct) + '</b> dari target ' + tgtLbl() + (isLatest ? ' MTD' : '') : '') +
    (chg != null ? ' — ' + (chg >= 0 ? 'naik ' : 'turun ') + Math.abs(Math.round(chg)) + '% dari ' + prevTxt + '.' : '.');
  const withPct = provRows.filter(o => o.pct != null);
  const worst = withPct.length ? withPct.reduce((a, b) => a.gap < b.gap ? a : b) : null;
  const attn = worst && worst.gap < 0 ? '<div class="attn"><div><span class="dot"></span></div><div><div class="t1">Perlu perhatian — ' + esc(tc(B.dims.prov[worst.key])) + '</div><div class="t2">Capaian ' + pctTxt(worst.pct) + ' dari target ' + tgtLbl() + ' MTD — gap ' + fmt(-worst.gap) + ' ton, terbesar di antara semua provinsi. Realisasi ' + fmt(worst.frc) + ' t vs target ' + fmt(worst.tgt) + ' t.</div></div><a class="cta" href="#/provinsi/' + enc(B.dims.prov[worst.key]) + '">Lihat rincian &rsaquo;</a></div>' : '';
  const tRows = B.trucks.filter(x => inPer(x[0], m) && truckOk(x)), tTrips = tRows.reduce((a, x) => a + x[6], 0), tTrucks = new Set(tRows.map(x => x[1])).size;   // same basis as Armada
  const prog = B.prog;
  page(
    mapHero(provRows, t, head, chg, isLatest) +
    '<div class="wrap"><div class="data-asof">Data realisasi s.d. ' + dayLabel(B.asOf) + (prog.date ? ' · snapshot Prognosa ' + dayLabel(prog.date) + ' ' + esc(prog.time) : '') + ' · ' + esc(B.build || '') + '</div>' +
    attn + incPerf(m, {}) +
    '<div class="section-head"><div class="section-title">Capaian ' + esc(m.label) + '</div></div>' +
    '<div class="main-chart">' + trendPanel({}, 'harian', 'Realisasi vs target ' + tgtLbl(), {grans: ['harian', 'mingguan', 'bulanan'], noCmp: true}) + '</div>' +
    '<div class="section-head"><div class="section-title">Ringkasan operasional</div><a class="section-link" href="#/armada">Armada &rsaquo;</a></div>' +
    '<div class="kpi-row">' +
      kpi('trend', 'var(--black)', 'Trip truk', fmt(tTrips), {href:'#/armada', hint: incTxt + ' · ' + esc(m.label)}) +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(tTrucks), {href:'#/armada', hint: fmt1(tTrips / Math.max(1, tTrucks)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata waktu muat', t.dwell == null ? '–' : fmt1(t.dwell) + ' <small>jam</small>', {href:'#/armada', hint:'masuk → keluar pabrik'}) +
      kpi('check', 'var(--green)', 'Realisasi / SO', pctTxt(t.of), {href:'#/ekspeditur', hint: t.so ? 'SO ' + fmt(t.so) + ' t' + ' (bulan penuh yang tercakup)' : 'SO tidak tersedia'}) +
    '</div>' +
    '</div>', 'home');
}

/* Home hero: the interactive Sumatra map is the centerpiece, headline + ranking beside it. */
function mapHero(provRows, t, head, chg, isLatest){
  const m = PER();
  const rank = provRows.slice().sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  const maxV = Math.max(1, ...rank.map(o => o.vol));
  const list = rank.map((o, i) => {
    const name = B.dims.prov[o.key];
    return '<a class="mh-row" href="#/provinsi/' + enc(name) + '" data-prov="' + esc(name) + '"><span class="mh-rk">' + (i + 1) + '</span><span class="mh-nm">' + esc(tc(name)) + '<i style="width:' + (o.vol / maxV * 100) + '%;background:' + tierColor(o.pct) + '"></i></span><span class="mh-tn">' + fmt(o.vol) + ' t</span><span class="mh-pc" style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</span><span class="mh-so">' + (o.of == null ? '–' : pctTxt(o.of)) + '</span></a>';
  }).join('');
  const chip = (l, v, sub, color) => '<div class="mh-chip"><div class="l">' + l + '</div><div class="v"' + (color ? ' style="color:' + color + '"' : '') + '>' + v + '</div>' + (sub ? '<div class="s">' + sub + '</div>' : '') + '</div>';
  return '<section class="map-hero"><div class="mh-bg"></div><div class="mh-inner">' +
    '<div class="mh-left"><div class="eyebrow">Peta capaian · ' + esc(m.sub) + '</div>' +
      '<h1>' + head + '</h1>' +
      '<div class="mh-chips">' +
        chip('Realisasi', fmt(t.vol) + ' <small>t</small>', fmt(t.trips) + ' trip') +
        chip('Target ' + tgtLbl() + (isLatest ? ' MTD' : ''), t.hasTgt ? fmt(t.tgt) + ' <small>t</small>' : '–', t.hasTgt ? 'gap ' + signed(t.gap) + ' t' : '') +
        chip('Capaian', pctTxt(t.pct), incName() + ' vs ' + tgtLbl(), t.pct == null ? null : t.pct >= 100 ? '#5FD08F' : t.pct >= 85 ? '#E8B04A' : '#FF6B74') +
        chip(m.kind === 'month' ? 'vs bulan lalu' : 'vs periode lalu', chg == null ? '–' : (chg >= 0 ? '▲ ' : '▼ ') + Math.abs(Math.round(chg)) + '%', 'periode sama', chg == null ? null : chg >= 0 ? '#5FD08F' : '#FF6B74') +
      '</div>' +
      '<div class="mh-list"><div class="mh-lh"><span>Rank provinsi</span><a href="#/provinsi">Lihat tabel &rsaquo;</a></div><div class="mh-row mh-head"><span></span><span></span><span class="mh-tn">Realisasi</span><span class="mh-pc">Target</span><span class="mh-so">Real/SO</span></div>' + list + '</div>' +
      '<div class="mh-key"><span><i style="background:#2E9E5C"></i>&ge;100%</span><span><i style="background:#C9952E"></i>85–99%</span><span><i style="background:#FF6B74"></i>&lt;85%</span><span><i style="background:#34424B"></i>tidak dilayani darat</span><span><i class="pin"></i>plant</span></div>' +
    '</div>' +
    '<div class="mh-map">' + mapHTML(provRows, {hero: true}) + '<div class="mh-hint">Arahkan kursor untuk angka · klik / ketuk provinsi untuk rincian</div></div>' +
  '</div></section>';
}

/* Per-incoterm performance (cards) for a scope, shown whenever the filter has more than one incoterm active. */
function incPerf(m, sc){
  const act = INCS().map((n, i) => i).filter(i => incOk(i));
  if(act.length < 2) return '';
  const rows = act.map(i => withInc({incFix: INCS()[i]}, () => Object.assign({i}, total(m, sc || {})))).filter(o => o.vol > 0 || o.tgt > 0);
  if(rows.length < 2) return '';
  const cards = rows.map(o => '<div class="ip-card" style="--c:' + incColor(o.i) + '"><div class="ip-n">' + esc(incLbl(INCS()[o.i])) + '</div>' +
    '<div class="ip-v">' + fmt(o.vol) + ' <small>t</small></div>' +
    '<div class="ip-bar"><i style="width:' + (o.pct == null ? 0 : Math.min(100, o.pct)) + '%;background:' + tierColor(o.pct) + '"></i></div>' +
    '<div class="ip-m">' + (o.hasTgt ? '<b style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</b> dari target ' + fmt(o.tgt) + ' t · <span style="color:' + (o.gap < 0 ? 'var(--red)' : 'var(--green)') + '">' + signed(o.gap) + '</span>' : 'tanpa target ' + tgtLbl()) + '</div>' +
    '<div class="ip-s">' + fmt(o.trips) + ' trip' + (o.of != null ? ' · Real/SO ' + pctTxt(o.of) : '') + '</div></div>').join('');
  return '<div class="section-head"><div class="section-title">Kinerja per incoterm</div><div class="section-sub">' + esc(incName()) + ' · ' + periodSub() + '</div></div><div class="ip-grid">' + cards + '</div>';
}
function pageProvList(){
  const rows = ranked(summarize('prov', PER()));
  const t = total(PER());
  const cols = [C.rank, {h:'Provinsi', cls:'name', val: r => tc(B.dims.prov[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.note];
  cols[2].total = fmt(t.vol); cols[3].total = fmt(t.tgt); cols[4].total = pctTxt(t.pct); cols[5].total = signed(t.gap); cols[6].total = fmt(t.so); cols[7].total = pctTxt(t.of); cols[8].total = fmt(t.trips);
  page('<div class="wrap">' + crumb([['Wilayah'], ['Semua provinsi']]) +
    '<div class="dp-head"><div><div class="nm">Semua Provinsi</div><div class="sub">Capaian tonase &amp; rank · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Total ' + pctTxt(t.pct) + ' dari target</div></div></div>' +
    incPerf(PER(), {}) +
    table(cols, rows, {go: r => '#/provinsi/' + enc(B.dims.prov[r.key]), total: true}) +
    '<div class="dp-cols even">' + trendPanel({}, 'mingguan', 'Tren seluruh provinsi') + sourcePanel({}) + '</div>' +
    '</div>', 'wilayah');
}
function sourcePanel(sc){
  const rows = [...summarize('src', PER(), sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  return '<div class="dp-panel"><div class="pt">Per source plant</div>' + (capTable(rows, 'Source plant', o => tc(B.dims.src[o.key]), o => B.dims.src[o.key] ? plantHref(B.dims.src[o.key]) : null) || '<div class="empty">Tidak ada data.</div>') +
    '<div class="trend-detail">Baris tanpa atribusi plant pada target dihitung di semua source. ' + 'Ditampilkan: ' + srcTxt().replace(/<\/?b>/g, '') + '.' + '</div></div>';
}
function progCells(rows, label){
  const s = k => rows.reduce((a, r) => a + (r[k] || 0), 0);
  const tgt = s('snopTarget');
  return '<div class="prog-row">' +
    '<div class="prog-cell"><div class="pl">Sudah rilis</div><div class="pv">' + fmt(s('sudahRilis')) + ' t</div><div class="pd">' + pctTxt(pctOf(s('sudahRilis'), tgt)) + ' dari target hari ini</div></div>' +
    '<div class="prog-cell"><div class="pl">Potensi</div><div class="pv">' + fmt(s('potensi')) + ' t</div><div class="pd">rilis + antri + timbang</div></div>' +
    '<div class="prog-cell"><div class="pl">Prognose</div><div class="pv" style="color:' + tierColor(pctOf(s('prognose'), tgt)) + '">' + fmt(s('prognose')) + ' t</div><div class="pd">' + pctTxt(pctOf(s('prognose'), tgt)) + ' · potensi + booking</div></div>' +
    '<div class="prog-cell"><div class="pl">Target ' + tgtLbl() + ' hari ini</div><div class="pv">' + fmt(tgt) + ' t</div><div class="pd">' + esc(label) + '</div></div></div>';
}
function progRows(filter){
  const P = B.prog, fi = Object.fromEntries(P.fields.map((k, i) => [k, i + 3]));
  return P.rows.filter(r => r[2] !== 'TOTAL' && !F.srcOff.has(r[0]) && (!filter || filter(r))).map(r => {
    const o = {plant: r[0], inc: r[1], prov: r[2]};
    P.fields.forEach(k => o[k] = /[Tt]arget/.test(k) ? r[fi[k]] * tgtK() : r[fi[k]]);
    return o;
  });
}

function pageProv(name){
  const pi = D.provIdx[name];
  if(pi == null) return notFound('Provinsi ' + name);
  const m = PER(), sc = {prov: pi};
  const all = ranked(summarize('prov', m));
  const me = all.find(o => o.key === pi) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
  const dist = ranked(summarize('dist', m, sc));
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.note];
  const mapS = distMapSection(name, m);   // the map panel carries the KPIs; the KPI row only shows when there is no map
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">Capaian tonase · ' + periodSub() + '</div></div><div class="badges">' + (me.rank ? '<div class="dp-badge">Rank ' + me.rank + ' / ' + all.rankTotal + '</div>' : '') + '<div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    (mapS || (
    '<div class="dp-review">' + reviewText(me.pct) + (me.hasTgt ? ' Realisasi ' + incName() + ' ' + fmt(me.frc) + ' t vs target MTD ' + fmt(me.tgt) + ' t (gap ' + signed(me.gap) + ' t).' : '') + (me.so ? ' Realisasi/SO ' + pctTxt(me.of) + '.' : '') + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–', {hint: me.hasTgt ? 'bulan penuh ' + fmt(me.tgtFull) + ' t' : 'tidak ada target ' + tgtLbl() + ''}) +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', me.gap < 0 ? 'Sisa ke target' : 'Lebih dari target', me.hasTgt ? fmt(Math.abs(me.gap)) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>'
    )) + tabsHTML('prov', [{k: 'tren', label: 'Tren & incoterm', html: incPerf(m, sc) +
    '<div class="dp-cols">' + trendPanel(sc, 'mingguan') + '</div>' +
    '<div class="dp-cols even">' + sourcePanel(sc) + '<div class="dp-panel"><div class="pt">Bulan penuh</div>' + ('<div class="tgt-panel" style="margin:0;border:0;padding:0"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi ' + incName() + '</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div>') + '</div></div>'},
    {k: 'distrik', label: 'Distrik', n: dist.filter(o => o.vol > 0).length, html: treeHTML(annotateDistrTargets(buildTree(shipRows(m, sc), ['dist', 'distr', 'toko']), m, {prov: pi}, {}, ['distr']), ['dist', 'distr', 'toko'], {tgt: {dist: summarize('dist', m, sc)}})},
    {k: 'eksp', label: 'Ekspeditur', html:
      '<div class="dp-panel"><div class="pt">Ekspeditur di provinsi ini<span class="pt-sub">capaian target · Real/SO</span></div>' + (incGroups(() => capTable(ekspRows(m, sc), 'Ekspeditur', o => ekspLbl(o.key), o => ekspHref(o.key), true), ekspSub(m, sc)) || '<div class="empty">Tidak ada data.</div>') + '</div>'
    },
    {k: 'distr', label: 'Distributor', html: distrTable(m, sc, 60).html}]) +
    '</div>', 'wilayah');
}

/* ===================== Plant pages ===================== */
/* A plant page (sc.src) always shows its own plant, even when it is not ticked in the plant filter.
   Moda/incoterm/kemasan/area/period filters still apply; the per-incoterm table shows every incoterm. */
const plantHref = n => '#/plant/' + enc(n);
const plantList = () => B.dims.src.map((n, i) => ({key: i, n})).filter(o => o.n && o.n !== 'TANPA_PLANT');
function plantIncRows(m, sc){
  return INCS().map(n => withInc({incFix: n}, () => Object.assign({n}, total(m, sc)))).filter(o => o.vol > 0 || o.tgt > 0);
}
function pagePlantList(){
  const m = PER();
  const rows = plantList().map(o => Object.assign(o, total(m, {src: o.key}))).filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  rows.forEach((o, i) => o.rank = i + 1);
  const cols = [C.rank, {h:'Plant', cls:'name', val: r => r.n, html: r => esc(r.n) + (F.srcOff.has(r.n) ? '<small>tidak dicentang di filter plant</small>' : '')}, C.vol(), C.tgt, C.pct, C.gap, C.trips];
  page('<div class="wrap">' + crumb([['Plant']]) +
    '<div class="dp-head"><div><div class="nm">Semua Plant</div><div class="sub">Plant / packing plant asal pengiriman · ' + incName() + ' · ' + periodSub() + '</div></div></div>' +
    '<div class="dp-review">Semua plant ditampilkan, termasuk yang tidak dicentang di filter plant. Klik baris untuk rincian plant.</div>' +
    (rows.length ? table(cols, rows, {go: r => plantHref(r.n)}) : '<div class="dp-review">Tidak ada data untuk filter ini.</div>') +
    '</div>', 'plant');
}
/* ===================== Plant review (kinerja rilis) ===================== */
/* Incoterm colours: fixed per incoterm (colour follows the entity), validated categorical set (CVD-safe adjacent pairs). */
const INC_COLOR = {'FOT': '#2a78d6', 'FRC': '#eb6834', 'FOB': '#6250d6', 'FAS': '#1baf7a', 'FRC CONTAINER': '#eda100', 'FRC KAPAL': '#e87ba4', 'FOT CONTAINER': '#008300'};
const incColor = i => INC_COLOR[INCS()[i]] || '#888';
let plantAllInc = false;   // plant page: all incoterms instead of the incoterm filter (toggle on the page)
/** Daily (or monthly, for long periods) release, targets, SLA and trucks of one plant, per incoterm. */
function plantReview(si, per, all){
  const incs = INCS().map((n, i) => i).filter(i => all || incOk(i)), inSet = new Set(incs);
  const monthly = per.days > 62;
  const bk = s => monthly ? s.slice(0, 7) : s;
  const keys = [];
  if(monthly){ for(let mm = per.from.slice(0, 7); mm <= per.to.slice(0, 7); mm = nextMonth(mm)) keys.push(mm); }
  else { const end = per.kind === 'month' ? per.month + '-' + pad2(dim(per.month)) : per.to; for(let s = per.from; s <= end; s = addDays(s, 1)) keys.push(s); }
  const blankB = () => ({inc: new Map(), tot: 0, trips: 0, dwS: 0, dwN: 0, tgt: 0, prev: null, trucks: new Set()});
  const Bk = new Map(keys.map(k => [k, blankB()]));
  const I = new Map(incs.map(i => [i, {i, vol: 0, az: 0, trips: 0, dwS: 0, dwN: 0, tgt: 0, tgtFull: 0, prev: 0}]));
  const pp = prevPer(per);
  B.facts.forEach(f => {
    if(f[2] !== si || !inSet.has(f[1]) || !kemOk(f[10]) || !areaProvOk(f[3])) return;
    const ds = B.days[f[0]], o = I.get(f[1]);
    if(ds >= per.from && ds <= per.to){
      const b = Bk.get(bk(ds)); if(!b) return;
      b.inc.set(f[1], (b.inc.get(f[1]) || 0) + f[6]); b.tot += f[6]; b.trips += f[7]; b.dwS += f[8]; b.dwN += f[9];
      o.vol += f[6]; o.trips += f[7]; o.dwS += f[8]; o.dwN += f[9];
      if(achKem(f[10]) && !NKP().has(f[3])) o.az += f[6];
    } else if(ds >= pp.from && ds <= pp.to){
      o.prev += f[6];
      if(!monthly && per.kind === 'month'){ const b = Bk.get(per.month + ds.slice(7)); if(b) b.prev = (b.prev || 0) + f[6]; }
    }
  });
  B.trucks.forEach(t => {
    if(t[4] !== si || !inSet.has(t[2]) || !kemOk(t[10]) || !areaProvOk(t[5])) return;
    const ds = B.days[t[0]]; if(ds < per.from || ds > per.to) return;
    const b = Bk.get(bk(ds)); if(b) b.trucks.add(t[1]);
  });
  const lastKey = keys[keys.length - 1], chartTo = monthly ? lastKey + '-' + pad2(dim(lastKey)) : lastKey;
  B.targets.forEach(t => {
    if(t[4] !== si || !inSet.has(tInc(t)) || !kemOk(tKem(t)) || !areaProvOk(t[1])) return;
    const o = I.get(tInc(t));
    o.tgt += targetBetween(t, per.from, per.tgtTo) / tgtK();   // SNOP basis; Internal = x TGT_INTERNAL
    o.tgtFull += targetBetween(t, per.from, chartTo) / tgtK();
    const n = dim(t[0]);
    for(let d = 1; d <= n; d++){
      const ds = t[0] + '-' + pad2(d); if(ds < per.from || ds > chartTo) continue;
      const b = Bk.get(bk(ds)); if(b) b.tgt += t[6] ? (t[6][d - 1] || 0) : t[5] / n;
    }
  });
  const list = [...I.values()].filter(o => o.vol > 0 || o.tgt > 0 || o.prev > 0);
  return {incs: list.map(o => o.i), keys, Bk, I: list, monthly, pp};
}
/** Stacked release bars per incoterm with the SNOP target line, the average and last month (same day) as a line. */
function plantChart(R){
  const W = 1100, H = 330, L = 52, Rm = 12, T = 26, Bm = 28, n = R.keys.length;
  const bs = R.keys.map(k => R.Bk.get(k)), withData = bs.filter(b => b.tot > 0);
  let avg = withData.length ? withData.reduce((a, b) => a + b.tot, 0) / withData.length : 0;
  const max = Math.max(1, ...bs.map(b => Math.max(b.tot, b.tgt * tgtK(), b.prev || 0))) * 1.1;
  const step = (W - L - Rm) / n, bw = Math.min(30, step * 0.68), y = v => T + (H - T - Bm) * (1 - v / max);
  const nice = v => { const p = Math.pow(10, Math.floor(Math.log10(v))); return [1, 2, 2.5, 5, 10].map(x => x * p).find(x => v / x <= 5); };
  const gs = nice(max), grid = [];
  for(let g = 0; g <= max; g += gs) grid.push('<line x1="' + L + '" x2="' + (W - Rm) + '" y1="' + y(g).toFixed(1) + '" y2="' + y(g).toFixed(1) + '" class="pc-grid"/><text x="' + (L - 6) + '" y="' + (y(g) + 3.5).toFixed(1) + '" class="pc-ax" text-anchor="end">' + fmt(g) + '</text>');
  let bars = '', labels = '', hits = '', xl = '';
  const order = R.incs.slice().sort((a, b) => ['FOT', 'FOT CONTAINER', 'FRC', 'FRC CONTAINER', 'FRC KAPAL', 'FOB', 'FAS'].indexOf(INCS()[a]) - ['FOT', 'FOT CONTAINER', 'FRC', 'FRC CONTAINER', 'FRC KAPAL', 'FOB', 'FAS'].indexOf(INCS()[b]));
  bs.forEach((b, i) => {
    const cx = L + step * (i + 0.5), x0 = cx - bw / 2;
    let acc = 0;
    const segs = order.filter(ii => b.inc.get(ii) > 0);
    segs.forEach((ii, j) => {
      const v = b.inc.get(ii), y1 = y(acc + v), y0 = y(acc), top = j === segs.length - 1;
      bars += top ? '<path d="M' + x0.toFixed(1) + ',' + y0.toFixed(1) + 'V' + (y1 + 3).toFixed(1) + 'q0,-3 3,-3H' + (x0 + bw - 3).toFixed(1) + 'q3,0 3,3V' + y0.toFixed(1) + 'Z" fill="' + incColor(ii) + '"/>'
                  : '<rect x="' + x0.toFixed(1) + '" y="' + y1.toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + Math.max(0, y0 - y1 - 1.5).toFixed(1) + '" fill="' + incColor(ii) + '"/>';
      acc += v;
    });
    if(b.tot > 0 && (n <= 16 || i % 1 === 0)) labels += '<text x="' + cx.toFixed(1) + '" y="' + (y(b.tot) - 5).toFixed(1) + '" class="pc-val" text-anchor="middle">' + (b.tot >= 1000 ? (b.tot / 1000).toFixed(1).replace('.', ',') + 'k' : fmt(b.tot)) + '</text>';
    const k = R.keys[i], lab = R.tick ? R.tick(k) : R.monthly ? monthShort(k) : String(+k.slice(8));
    xl += '<text x="' + cx.toFixed(1) + '" y="' + (H - 10) + '" class="pc-ax" text-anchor="middle">' + lab + '</text>';
    const tip = (R.title ? R.title(k) : R.monthly ? monthLabel(k) : dayLabel(k)) + '|' + (b.tot ? 'Rilis ' + fmt(b.tot) + ' t · ' + fmt(b.trips) + ' trip' + (R.hourly ? '' : ' · ' + b.trucks.size + ' truk') : 'Belum ada rilis') +
      segs.slice().reverse().map(ii => '|' + incLbl(INCS()[ii]) + ': ' + fmt(b.inc.get(ii)) + ' t').join('') +
      (b.tgt ? '|Target ' + tgtLbl() + ' ' + fmt(b.tgt * tgtK()) + ' t' + (b.tot ? ' · capaian ' + pctTxt(pctOf(b.tot, b.tgt * tgtK())) : '') : '') + (b.prev != null ? '|' + (R.prevLbl || 'Bulan lalu (tgl sama)') + ' ' + fmt(b.prev) + ' t' : '') + (b.dwN ? '|Waktu muat ' + fmt1(b.dwS / b.dwN) + ' jam' : '');
    hits += '<rect x="' + (cx - step / 2).toFixed(1) + '" y="' + T + '" width="' + step.toFixed(1) + '" height="' + (H - T - Bm) + '" class="pc-hit" data-tip="' + esc(tip) + '"/>';
  });
  const pts = (f) => bs.map((b, i) => f(b) == null ? null : [L + step * (i + 0.5), y(f(b))]);
  const path = ps => { let d = '', pen = false; ps.forEach(p => { if(!p){ pen = false; return; } d += (pen ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1); pen = true; }); return d; };
  const tgtLine = '<path d="' + path(pts(b => b.tgt ? b.tgt * tgtK() : null)) + '" class="pc-tgt"/>';
  const prevPts = pts(b => b.prev != null && b.prev > 0 ? b.prev : null);
  const prevLine = R.monthly || R.noPrev ? '' : '<path d="' + path(prevPts) + '" class="pc-prev"/>' + prevPts.filter(Boolean).map(p => '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="2.6" class="pc-prevdot"/>').join('');
  if(R.avgAll) avg = bs.reduce((a, b) => a + b.tot, 0) / n;   // per hour: the day's total spread over 24 hours
  const avgLine = avg ? '<line x1="' + L + '" x2="' + (W - Rm) + '" y1="' + y(avg).toFixed(1) + '" y2="' + y(avg).toFixed(1) + '" class="pc-avg"/><text x="' + (L + 4) + '" y="' + (y(avg) - 4).toFixed(1) + '" class="pc-avgl">rata-rata ' + fmt(avg) + ' t</text>' : '';
  return '<svg class="pc-svg" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Rilis per ' + (R.monthly ? 'bulan' : 'hari') + ' per incoterm">' + grid.join('') + bars + avgLine + tgtLine + prevLine + labels + xl + hits + '</svg>';
}
/** SLA (masuk -> keluar pabrik) per day/month as bars in the trend style: green = within the 4-hour target. */
function plantSla(R, per){
  const keys = R.keys.filter(k => (R.monthly ? k + '-01' : k) <= per.to);
  const data = keys.map(k => { const b = R.Bk.get(k), v = b.dwN ? b.dwS / b.dwN : null;
    return {label: R.monthly ? monthShort(k) : String(+k.slice(8)), long: (R.monthly ? monthLabel(k) : dayLabel(k)) + (v == null ? ' · tidak ada data SLA' : ' · SLA ' + fmt1(v) + ' jam'), vol: v || 0, tgt: SLA_TARGET, pct: null, unit: 'jam', color: slaColor(v)}; });
  if(!data.some(d => d.vol)) return '';
  TRENDS.psla = {id: 'psla', data, click: () => {}};
  return barsHTML('psla', data, {labels: data.length <= 31});
}
const nextMonth = m => { const y = +m.slice(0, 4), mo = +m.slice(5); return mo === 12 ? (y + 1) + '-01' : y + '-' + pad2(mo + 1); };
/* Release per hour (truck exit time, bundle.hours: last 62 days) for one plant and day. */
let plantHourDay = null, plantHourCtx = null;
function plantHours(si, day, all){
  const incs = INCS().map((n, i) => i).filter(i => all || incOk(i)), inSet = new Set(incs);
  const H = (B.hours || []).filter(h => h[1] === si && inSet.has(h[2]) && kemOk(h[6]));
  const days = [...new Set(H.map(h => h[0]))].sort((a, b) => a - b);
  if(!days.length) return null;
  const per = PER(), inPerDays = days.filter(d => inPer(d, per));
  let di = day != null && days.includes(day) ? day : (inPerDays.length ? inPerDays[inPerDays.length - 1] : days[days.length - 1]);
  const pos = days.indexOf(di), prevDay = pos > 0 ? days[pos - 1] : null;
  const keys = Array.from({length: 24}, (_, h) => pad2(h));
  const Bk = new Map(keys.map(k => [k, {inc: new Map(), tot: 0, trips: 0, dwS: 0, dwN: 0, tgt: 0, prev: prevDay == null ? null : 0, trucks: new Set()}]));
  H.forEach(h => {
    const b = Bk.get(pad2(h[3]));
    if(h[0] === di){ b.inc.set(h[2], (b.inc.get(h[2]) || 0) + h[4]); b.tot += h[4]; b.trips += h[5]; }
    else if(h[0] === prevDay) b.prev += h[4];
  });
  const ds = B.days[di], m = ds.slice(0, 7), d = +ds.slice(8);
  let dayTgt = 0;
  B.targets.forEach(t => {
    if(t[0] !== m || t[4] !== si || !inSet.has(tInc(t)) || !kemOk(tKem(t)) || !areaProvOk(t[1])) return;
    dayTgt += t[6] ? (t[6][d - 1] || 0) : t[5] / dim(m);
  });
  Bk.forEach(b => b.tgt = dayTgt / 24);
  const used = [...new Set(H.filter(h => h[0] === di).map(h => h[2]))];
  return {keys, Bk, incs: incs.filter(i => used.includes(i)), monthly: false, hourly: true, avgAll: true, noPrev: prevDay == null,
    tick: k => k, title: k => dayLabel(ds) + ' · jam ' + k + '.00–' + k + '.59', prevLbl: prevDay != null ? 'Hari sebelumnya (' + shortDate(B.days[prevDay]) + ')' : '',
    day: di, days, dayTgt, prevDay};
}
function plantHourSection(){
  const c = plantHourCtx; if(!c) return '';
  const R = plantHours(c.si, plantHourDay, plantAllInc);
  if(!R) return '<div id="plantHours"><div class="section-head"><div class="section-title">Rilis per jam (ton)</div></div><div class="dp-review">Belum ada data jam keluar truk untuk plant ini (kolom WAKTU_KELUAR di Rilis; 62 hari terakhir).</div></div>';
  const bs = R.keys.map(k => R.Bk.get(k)), tot = bs.reduce((a, b) => a + b.tot, 0), trips = bs.reduce((a, b) => a + b.trips, 0);
  const peak = bs.reduce((a, b, i) => b.tot > a.v ? {v: b.tot, i} : a, {v: 0, i: 0}), idle = bs.filter(b => !b.tot).length;
  const ds = B.days[R.day], pos = R.days.indexOf(R.day), prv = pos > 0 ? R.days[pos - 1] : null, nxt = pos < R.days.length - 1 ? R.days[pos + 1] : null;
  const tgtDay = R.dayTgt * tgtK();
  const nav = '<div class="ph-nav">' + (prv != null ? '<a href="#" data-plantday="' + prv + '">&lsaquo; ' + shortDate(B.days[prv]) + '</a>' : '<span></span>') +
    '<b>' + dayLabel(ds) + '</b>' + (nxt != null ? '<a href="#" data-plantday="' + nxt + '">' + shortDate(B.days[nxt]) + ' &rsaquo;</a>' : '<span></span>') + '</div>';
  const stats = '<div class="ph-stats"><span>Total <b>' + fmt(tot) + ' t</b> · ' + fmt(trips) + ' trip</span><span>Rata-rata <b>' + fmt(tot / 24) + ' t/jam</b></span>' +
    (tgtDay ? '<span>Target ' + tgtLbl() + ' <b>' + fmt(tgtDay / 24) + ' t/jam</b> · capaian hari ini <b style="color:' + tierColor(pctOf(tot, tgtDay)) + '">' + pctTxt(pctOf(tot, tgtDay)) + '</b></span>' : '') +
    '<span>Puncak <b>jam ' + pad2(peak.i) + '</b> (' + fmt(peak.v) + ' t)</span><span>Jam tanpa rilis <b>' + idle + '</b></span></div>';
  const hData = bs.map((b, h) => { const t = tgtDay ? tgtDay / 24 : null; return {label: pad2(h), long: dayLabel(ds) + ' · jam ' + pad2(h) + '.00–' + pad2(h) + '.59 · ' + fmt(b.trips) + ' trip' + [...b.inc.entries()].map(([i, v]) => ' · ' + incLbl(INCS()[i]) + ' ' + fmt(v) + ' t').join(''), vol: b.tot, tgt: t, pct: t ? pctOf(b.tot, t) : null}; });
  TRENDS.ph = {id: 'ph', data: hData, click: () => {}};
  return '<div id="plantHours"><div class="section-head"><div class="section-title">Rilis per jam (ton) · jam truk keluar</div>' + nav + '</div>' +
    '<div class="dp-panel pc-panel">' + stats + barsHTML('ph', hData, {labels: true}) + '<div class="trend-legend">' + LG_NUM + '<span><i style="background:var(--green)"></i>&ge;100% target/jam</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span>' + (tgtDay ? '<span><i class="dash"></i>target ' + tgtLbl() + ' / jam</span>' : '') + '</div></div></div>';
}
function pagePlant(name){
  const si = B.dims.src.indexOf(name);
  if(si < 0 || !name) return notFound('Plant ' + name);
  const m = PER(), sc = {src: si};
  const R = plantReview(si, m, plantAllInc);
  const sum = f => R.I.reduce((a, o) => a + f(o), 0);
  const vol = sum(o => o.vol), az = sum(o => o.tgt ? o.az : 0), tgt = sum(o => o.tgt), prev = sum(o => o.prev), trips = sum(o => o.trips);
  const dwS = sum(o => o.dwS), dwN = sum(o => o.dwN), tI = tgt * TGT_INTERNAL;
  const pS = tgt ? pctOf(az, tgt) : null, pI = tgt ? pctOf(az, tI) : null, chg = prev ? (vol / prev - 1) * 100 : null;
  const dayKeys = R.keys.filter(k => R.Bk.get(k).tot > 0), lastK = dayKeys[dayKeys.length - 1];
  const avg = dayKeys.length ? vol / dayKeys.length : 0;
  const scopeTxt = plantAllInc ? 'semua incoterm' : incName();
  // headline
  const head = vol ? 'Rilis <b>' + fmt(vol) + ' t</b> ' + esc(scopeTxt) + (tgt ? ', <b class="' + (pS >= 100 ? 'ok' : 'bad') + '">' + pctTxt(pS) + '</b> dari target ' + tgtLbl() + ' (' + signed(az - tgt) + ' t)' : '') +
    (avg && !R.monthly ? '; rata-rata ' + fmt(avg) + ' t/hari' : '') + (chg != null ? ', <b class="' + (chg >= 0 ? 'ok' : 'bad') + '">' + (chg >= 0 ? '▲ ' : '▼ ') + Math.abs(Math.round(chg * 10) / 10).toString().replace('.', ',') + '%</b> vs periode sama ' + (m.kind === 'month' ? 'bulan lalu' : 'sebelumnya') : '') + '.'
    : 'Belum ada rilis dari ' + esc(name) + ' untuk filter ini.';
  const chips = [
    heroChip('Realisasi', fmt(vol) + ' <small>t</small>', fmt(trips) + ' trip'),
    heroChip('Target SNOP', tgt ? fmt(tgt) + ' <small>t</small>' : '–', tgt ? 'capaian <b style="color:' + heroTier(pctOf(az, tgt)) + '">' + pctTxt(pctOf(az, tgt)) + '</b> · ' + signed(az - tgt) + ' t' : 'tidak ada target'),
    heroChip('Target Internal', tgt ? fmt(tI) + ' <small>t</small>' : '–', tgt ? 'capaian <b style="color:' + heroTier(pI) + '">' + pctTxt(pI) + '</b> · ' + signed(az - tI) + ' t' : '115% SNOP'),
    heroChip(R.monthly ? 'Rata-rata / bulan' : 'Rata-rata / hari', fmt(R.monthly ? vol / Math.max(1, dayKeys.length) : avg) + ' <small>t</small>', lastK ? (R.monthly ? monthShort(lastK) : shortDate(lastK)) + ': ' + fmt(R.Bk.get(lastK).tot) + ' t' : ''),
    heroChip(m.kind === 'month' ? 'vs bulan lalu' : 'vs periode lalu', chg == null ? '–' : (chg >= 0 ? '▲ ' : '▼ ') + fmt1(Math.abs(chg)) + '%', prev ? fmt(prev) + ' t · ' + shortDate(R.pp.from) + '–' + shortDate(R.pp.to) : '', chg == null ? null : chg >= 0 ? '#5FD08F' : '#FF6B74'),
    heroChip('Waktu muat rata-rata', dwN ? fmt1(dwS / dwN) + ' <small>jam</small>' : '–', 'target ' + SLA_TARGET + ' jam · masuk → keluar', dwN ? (dwS / dwN <= SLA_TARGET ? '#5FD08F' : '#FF6B74') : null)];
  const cards = R.I.map(o => {
    const n = INCS()[o.i], ps = o.tgt ? pctOf(o.az, o.tgt) : null, pi = o.tgt ? pctOf(o.az, o.tgt * TGT_INTERNAL) : null;
    return '<div class="ph-card" style="--c:' + incColor(o.i) + '"><div class="ph-n">' + esc(incLbl(n)) + '</div><div class="ph-v">' + fmt(o.vol) + ' <small>t</small></div>' +
      '<div class="ph-k"><span>SNOP</span><b style="color:' + heroTier(ps) + '">' + (o.tgt ? pctTxt(ps) : '–') + '</b><small>' + (o.tgt ? 'T ' + fmt(o.tgt) : 'tanpa target') + '</small></div>' +
      '<div class="ph-k"><span>Internal</span><b style="color:' + heroTier(pi) + '">' + (o.tgt ? pctTxt(pi) : '–') + '</b><small>' + (o.tgt ? 'T ' + fmt(o.tgt * TGT_INTERNAL) : '') + '</small></div></div>';
  }).join('') || '<div class="mh-miss">Tidak ada incoterm dengan rilis atau target.</div>';
  const bg = (window.BG || {})[plantBg(name)];
  const hero = '<section class="map-hero mh-page plant-hero"><div class="mh-bg"' + (bg ? ' style="background-image:linear-gradient(90deg,rgba(15,26,33,.95) 0%,rgba(15,26,33,.82) 48%,rgba(15,26,33,.5) 100%),url(&quot;assets/bg/' + bg.f + '&quot;)"' : '') + '></div><div class="mh-inner">' +
    '<div class="mh-left"><div class="eyebrow">Kinerja rilis · ' + esc(m.sub) + '</div><h2 class="mh-h">' + head + '</h2><div class="mh-chips">' + chips.join('') + '</div></div>' +
    '<div class="ph-cards"><div class="mh-lh"><span>KPI per incoterm · SNOP &amp; Internal</span></div>' + cards + '</div></div></section>';
  // daily trend: the standard trend panel (tier-coloured bars, % and value, target ticks, MTD/harian/mingguan/bulanan)
  // loading port (most tonnage on sea incoterms): truck SLA and release-per-hour say nothing about ships -> ship timeline
  const seaVol = R.I.reduce((a, o) => a + (modaOf(INCS()[o.i]) === 'LAUT' ? o.vol : 0), 0), isSea = vol > 0 && seaVol / vol >= 0.5;
  const sla = isSea || !dwN ? '' : plantSla(R, m), slaAvg = dwN ? dwS / dwN : null;
  const chart = '<div class="section-head"><div class="section-title">Tren realisasi</div></div>' +
    '<div class="main-chart">' + trendPanel(sc, 'harian', 'Tren realisasi ' + name) + '</div>' +
    (sla ? '<div class="dp-panel pc-panel" style="margin-top:14px"><div class="pt">SLA per ' + (R.monthly ? 'bulan' : 'hari') + ' (jam, masuk → keluar pabrik)<span class="pt-sub">rata-rata <b style="color:' + slaColor(slaAvg) + '">' + fmt1(slaAvg) + ' jam</b> · target ' + SLA_TARGET + ' jam</span></div>' + sla +
      '<div class="trend-legend"><span><i style="background:var(--green)"></i>&le;' + SLA_TARGET + ' jam (memenuhi target)</span><span><i style="background:var(--red)"></i>&gt;' + SLA_TARGET + ' jam</span><span><i class="dash"></i>target ' + SLA_TARGET + ' jam</span></div></div>' : '');
  // comparison table per incoterm
  const rows = R.I.map(o => ({n: o.i, vol: o.vol, tgt: o.tgt, ps: o.tgt ? pctOf(o.az, o.tgt) : null, pi: o.tgt ? pctOf(o.az, o.tgt * TGT_INTERNAL) : null, gap: o.az - o.tgt, prev: o.prev, chg: o.prev ? (o.vol / o.prev - 1) * 100 : null, trips: o.trips, sla: o.dwN ? o.dwS / o.dwN : null}));
  const pc = p => p == null ? '–' : '<b style="color:' + tierColor(p) + '">' + pctTxt(p) + '</b>';
  const cols = [{h:'Incoterm', cls:'name', val: r => incLbl(INCS()[r.n]), html: r => '<span class="inc-dot" style="background:' + incColor(r.n) + '"></span>' + esc(incLbl(INCS()[r.n]))},
    {h:'Realisasi (t)', num:true, val: r => r.vol, html: r => fmt(r.vol)},
    {h:'Target SNOP (t)', num:true, val: r => r.tgt, html: r => r.tgt ? fmt(r.tgt) : '–'},
    {h:'vs SNOP', num:true, val: r => r.ps == null ? -1 : r.ps, html: r => pc(r.ps)},
    {h:'vs Internal', num:true, val: r => r.pi == null ? -1 : r.pi, html: r => pc(r.pi)},
    {h:'Gap SNOP (t)', num:true, val: r => r.tgt ? r.gap : 0, html: r => r.tgt ? '<span style="color:' + (r.gap < 0 ? 'var(--red)' : 'var(--green)') + '">' + signed(r.gap) + '</span>' : '–'},
    {h:(m.kind === 'month' ? 'Bulan lalu' : 'Periode lalu') + ' (t)', num:true, val: r => r.prev, html: r => r.prev ? fmt(r.prev) : '–'},
    {h:'Δ', num:true, val: r => r.chg == null ? -999 : r.chg, html: r => r.chg == null ? '–' : '<span style="color:' + (r.chg >= 0 ? 'var(--green)' : 'var(--red)') + '">' + (r.chg >= 0 ? '+' : '−') + fmt1(Math.abs(r.chg)) + '%</span>'},
    {h:'Trip', num:true, val: r => r.trips}, {h:'Waktu muat (jam)', num:true, val: r => r.sla == null ? -1 : r.sla, html: r => r.sla == null ? '–' : fmt1(r.sla)}];
  if(rows.length > 1){
    cols[0].total = '<b>Total</b>'; cols[1].total = '<b>' + fmt(vol) + '</b>'; cols[2].total = tgt ? '<b>' + fmt(tgt) + '</b>' : '–'; cols[3].total = pc(pS); cols[4].total = pc(pI);
    cols[5].total = tgt ? signed(az - tgt) : '–'; cols[6].total = prev ? fmt(prev) : '–'; cols[7].total = chg == null ? '–' : (chg >= 0 ? '+' : '−') + fmt1(Math.abs(chg)) + '%'; cols[8].total = fmt(trips); cols[9].total = dwN ? fmt1(dwS / dwN) : '–';
  }
  if(!plantHourCtx || plantHourCtx.si !== si) plantHourDay = null;   // another plant: start at its latest day
  plantHourCtx = {si};
  const shipTl = shipGantt(m, null, {src: si, bare: true});   // plants that also load ships (GP Dumai) get the timeline too
  const shipSec = isSea || shipTl ? '<div class="section-head"><div class="section-title">Kapal muat · 30 hari</div><a class="section-link" href="#/kapal">Semua kapal &rsaquo;</a></div>' +
    (shipTl || '<div class="dp-review">Tidak ada kapal bernama yang muat di ' + esc(name) + ' dalam 30 hari terakhir.</div>') : '';
  const hourSec = shipSec;   // release-per-hour panel removed (Rendy); plantHourSection kept for a possible return
  const cmp = '<div class="section-head"><div class="section-title">Perbandingan per incoterm</div><div class="section-sub">' + esc(m.sub) + ' vs ' + shortDate(R.pp.from) + '–' + shortDate(R.pp.to) + '</div></div>' +
    (rows.length ? table(cols, rows, {total: rows.length > 1}) : '<div class="dp-review">Tidak ada data untuk filter ini.</div>');
  // destinations (filter-driven, as before)
  const provs = [...summarize('prov', m, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const DT = distrTable(m, sc, 60);
  page('<div class="wrap">' + crumb([['Plant', '#/plant'], [name]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(name) + '</div><div class="sub">Plant / packing plant asal pengiriman · ' + esc(scopeTxt) + ' · ' + periodSub() + '</div></div></div>' +
    (F.srcOff.has(name) ? '<div class="dp-review">Plant ini tidak dicentang di filter plant: halaman ini tetap menampilkan datanya, tetapi Ringkasan dan halaman lain tidak menghitungnya sampai dicentang.</div>' : '') +
    hero + tabsHTML('plant', [
    {k: 'kinerja', label: 'Tren & incoterm', html: chart + cmp},
    ...(hourSec ? [{k: 'kapal', label: 'Kapal', html: hourSec}] : []),
    {k: 'tujuan', label: 'Tujuan & ekspeditur', html: '<div class="dp-cols">' +
      '<div class="dp-panel"><div class="pt">Provinsi tujuan<span class="pt-sub">capaian target · ' + esc(incName()) + '</span></div>' + (capTable(provs, 'Provinsi', o => tc(B.dims.prov[o.key]), o => '#/provinsi/' + enc(B.dims.prov[o.key])) || '<div class="empty">Tidak ada data.</div>') + '</div>' +
      '<div class="dp-panel"><div class="pt">Ekspeditur / pengangkut<span class="pt-sub">' + esc(incName()) + '</span></div>' + (incGroups(() => capTable(ekspRows(m, sc), 'Ekspeditur', o => ekspLbl(o.key), o => ekspHref(o.key)), ekspSub(m, sc)) || '<div class="empty">Tidak ada data.</div>') + '</div>' +
    '</div>'},
    {k: 'distr', label: 'Distributor', html: DT.html}]) +
    '</div>', 'plant');
}

function pageDist(name){
  const di = D.distIdx[name];
  if(di == null) return notFound('Distrik ' + name);
  const m = PER(), sc = {dist: di}, pname = B.dims.prov[B.dims.distProv[di]] || '';
  const me = total(m, sc);
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const ekCols = [{h:'Ekspeditur', cls:'name', val: r => ekspLbl(r.key)}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.dwell];
  const mapS = kecMapSection(name, pname, m, di);   // kecamatan panel carries the KPIs; fallback (province map) keeps the KPI row
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], [tc(pname), '#/provinsi/' + enc(pname)], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">' + esc(tc(pname)) + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    (/kmap|Slot/.test(mapS) ? mapS : (
    '<div class="dp-review">' + reviewText(me.pct) + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–') +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', 'Gap', me.hasTgt ? signed(me.gap) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>' +
    mapS)) + tabsHTML('dist', [
    {k: 'tren', label: 'Tren & incoterm', html: incPerf(m, sc) + '<div class="dp-cols">' + trendPanel(sc, 'harian') + '</div>'},
    {k: 'tujuan', label: 'Distributor & toko', html: treeHTML(annotateDistrTargets(buildTree(shipRows(m, sc), ['distr', 'toko']), m, {dist: di}), ['distr', 'toko'], {distrTgt: true})},
    {k: 'eksp', label: 'Ekspeditur', html: incGroups(() => { const l = ekspRows(m, sc); return l.length ? table(ekCols, l, {go: r => ekspHref(r.key), sort: 1}) : ''; }, ekspSub(m, sc)) || '<div class="dp-review">Tidak ada ekspeditur untuk filter ini.</div>'}]) +
    '</div>', 'wilayah');
}

function pageEkspList(){
  const m = PER(), t = total(m);
  const tr = trucksFor(m);
  const multi = tr.length ? pctOf(tr.filter(x => x.trips > 1).length, tr.length) : null;
  /* ranking table for the incoterm(s) currently in F.inc (run per incoterm by incGroups) */
  const frcNow = () => { const fi = INCS().indexOf(F.inc); return fi < 0 || isFrc(fi); };   // only franco incoterms have transporters
  const build = () => {
    if(!frcNow()) return '';
    const list = ranked(summarize('eksp', m)).filter(r => D.frcEksp.has(r.key) && (r.vol > 0 || r.tgt > 0));   // transporters only (FRC)
    if(!list.length) return '';
    const t = total(m), tr = trucksFor(m);
    const perE = new Map(); tr.forEach(x => x.eksp.forEach((n, e) => { const o = perE.get(e) || {n: 0, one: 0}; o.n++; if(x.trips === 1) o.one++; perE.set(e, o); }));
    const tk = r => (perE.get(r.key) || {n: 0, one: 0});
    const cols = [C.rank, {h:'Ekspeditur', cls:'name', val: r => ekspLbl(r.key)}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips,
      {h:'Truk', num:true, val: r => tk(r).n, html: r => fmt(tk(r).n)},
      {h:'Trip / truk', num:true, val: r => r.trips / Math.max(1, tk(r).n), html: r => fmt1(r.trips / Math.max(1, tk(r).n))},
      {h:'Truk 1 trip', num:true, val: r => tk(r).one}, C.dwell, C.note];
    cols[2].total = fmt(t.vol); cols[3].total = fmt(t.tgt); cols[4].total = pctTxt(t.pct); cols[5].total = signed(t.gap); cols[6].total = fmt(t.so); cols[7].total = pctTxt(t.of); cols[8].total = fmt(t.trips); cols[9].total = fmt(tr.length); cols[10].total = fmt1(t.trips / Math.max(1, tr.length));
    return table(cols, list, {go: r => ekspHref(r.key), total: true});
  };
  const rates = () => !frcNow() ? '' : capTable(ranked(summarize('eksp', m)).filter(r => D.frcEksp.has(r.key) && r.pct != null), 'Ekspeditur', r => ekspLbl(r.key), r => ekspHref(r.key), true);
  page('<div class="wrap">' + crumb([['Ekspeditur']]) +
    '<div class="dp-head"><div><div class="nm">Peringkat Ekspeditur</div><div class="sub">Capaian target ' + tgtLbl() + ' &amp; scorecard operasional · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Total ' + pctTxt(t.pct) + ' dari target</div></div></div>' +
    (nfOn() ? '<div class="dp-review">' + incName() + ' diambil langsung oleh distributor (tanpa ekspeditur). Ganti filter ke FRC, atau buka <a href="#/distributor" style="color:var(--red);font-weight:700">Distributor</a>.</div>' : '') +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Scorecard keseluruhan</div>' +
      rateRow('Capaian target ' + tgtLbl(), t.pct) + rateRow('Realisasi / SO', t.of) + rateRow('Truk aktif > 1 trip', multi) +
    '</div><div class="dp-panel"><div class="pt">Capaian per ekspeditur<span class="pt-sub">capaian target · Real/SO</span></div>' + (incGroups(rates) || '<div class="empty">Tidak ada target untuk filter ini.</div>') + '</div></div>' +
    (incGroups(build, () => { const t = total(m); return fmt(t.vol) + ' t' + (t.hasTgt ? ' · ' + pctTxt(t.pct) + ' dari target ' + fmt(t.tgt) + ' t' : ''); }) || '<div class="dp-review">Tidak ada ekspeditur FRC untuk filter ini.</div>') +
    '</div>', 'ekspeditur');
}

/* Nama lengkap ekspeditur (kamus master baru) bila kode hanya singkatan. */
const ekspFull = c => { const n = coLbl((B.dims.ekspNama || [])[D.ekspIdx[c]] || ''); return n && n !== c ? esc(n) + ' · ' : ''; };
function pageEksp(code){
  const ei = D.ekspIdx[code];
  if(ei == null) return notFound('Ekspeditur ' + code);
  const m = PER(), sc = {eksp: ei};
  const all = ranked(new Map([...summarize('eksp', m)].filter(([k]) => D.frcEksp.has(k))));   // rank among FRC ekspeditur only (non-FRC 'eksp' = distributor)
  const me = all.find(o => o.key === ei) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
  const provs = [...summarize('prov', m, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  const dist = ranked(summarize('dist', m, sc));
  const trucks = trucksFor(m, sc).sort((a, b) => b.trips - a.trips);
  const one = trucks.filter(t => t.trips === 1).length;
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key]), html: r => esc(tc(B.dims.dist[r.key])) + '<small>' + esc(tc(B.dims.prov[B.dims.distProv[r.key]])) + '</small>'}, C.vol(), C.tgt, C.pct, C.gap, C.trips, C.note];
  const truckCols = [{h:'Nopol', cls:'name', val: r => B.dims.truck[r.truck]}, {h:'Trip', num:true, val: r => r.trips}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)},
    {h:'Hari aktif', num:true, val: r => r.days}, {h:'Waktu muat (jam)', num:true, val: r => r.dwN ? r.dwS / r.dwN : null, html: r => r.dwN ? fmt1(r.dwS / r.dwN) : '–'},
    {h:'Aktif', val: r => r.first, html: r => activeTxt(r)}, {h:'Provinsi utama', val: r => tc(B.dims.prov[r.prov])}];
  page('<div class="wrap">' + crumb([['Ekspeditur', '#/ekspeditur'], [code]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(code) + '</div><div class="sub">' + ekspFull(code) + 'Mitra ekspeditur FRC · ' + periodSub() + '</div></div><div class="badges">' + (me.rank ? '<div class="dp-badge">Rank ' + me.rank + ' / ' + all.filter(o => D.frcEksp.has(o.key) && o.rank).length + '</div>' : '') + '<div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + (me.hasTgt ? ' Realisasi ' + fmt(me.frc) + ' t vs target MTD ' + fmt(me.tgt) + ' t (gap ' + signed(me.gap) + ' t).' : '') + (one ? ' ' + one + ' truk baru 1 trip bulan ini.' : '') + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–', {hint: me.hasTgt ? 'bulan penuh ' + fmt(me.tgtFull) + ' t' : ''}) +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {hint: fmt1(me.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata waktu muat', me.dwell == null ? '–' : fmt1(me.dwell) + ' <small>jam</small>', {hint: me.so ? 'Real/SO ' + pctTxt(me.of) : ''}) +
    '</div>' + tabsHTML('eksp', [{k: 'kinerja', label: 'Kinerja & provinsi', html: incPerf(m, sc) +
    '<div class="dp-cols">' + trendPanel(sc, 'mingguan') +
      '<div class="dp-panel"><div class="pt">Capaian per provinsi<span class="pt-sub">capaian target · Real/SO</span></div>' + (provs.length ? incGroups(() => capTable([...summarize('prov', m, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol), 'Provinsi', o => tc(B.dims.prov[o.key]), o => '#/provinsi/' + enc(B.dims.prov[o.key]), true), ekspSub(m, sc)) : '<div class="empty">Tidak ada data untuk filter ini.</div>') + '</div>' +
    '</div>' +
    '<div class="section-head"><div class="section-title">Target bulan berjalan</div></div>' +
    '<div class="tgt-panel"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Target MTD</div><div class="tv">' + (me.hasTgt ? fmt(me.tgt) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa ke target bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div>'},
    {k: 'tujuan', label: 'Tujuan pengiriman', html: treeHTML(buildTree(shipRows(m, sc), ['prov', 'dist', 'distr', 'toko']), ['prov', 'dist', 'distr', 'toko'], {tgt: {prov: summarize('prov', m, sc), dist: summarize('dist', m, sc)}})},
    {k: 'truk', label: 'Truk', n: trucks.length, html: '<div class="tab-note"><a class="section-link" href="#/armada">Semua armada &rsaquo;</a></div>' + table(truckCols, trucks, {limit: 300})}]) +
    '</div>', 'ekspeditur');
}

function pageForecast(gran){
  gran = ['mingguan', 'bulanan'].includes(gran) ? gran : 'mingguan';   // 'Hari ini' (snapshot monitoring) tidak ada di master baru
  const chips = '<div class="gran-chips" style="margin-bottom:24px">' + [['mingguan','Minggu ini'],['bulanan','Akhir bulan']].map(g => '<a href="#/prognosa/' + PROG_SLUG[g[0]] + '" class="' + (g[0] === gran ? 'on' : '') + '">' + g[1] + '</a>').join('') + '</div>';
  let body = '';
  if(gran === 'harian'){
    const P = B.prog;
    const incs = F.inc === 'ALL' ? ['FRC', 'FOT'] : [F.inc];
    body = incs.map(inc => {
      const rows = progRows(r => r[1] === inc);
      const byProv = new Map();
      rows.forEach(r => { const o = byProv.get(r.prov) || {prov: r.prov}; P.fields.forEach(k => o[k] = (o[k] || 0) + (r[k] || 0)); byProv.set(r.prov, o); });
      const list = [...byProv.values()].map(o => Object.assign(o, {pctP: pctOf(o.prognose, o.snopTarget)}));
      const cols = [{h:'Provinsi', cls:'name', val: r => tc(r.prov)}, {h:'Antri', num:true, val: r => r.antriTon, html: r => fmt(r.antriTon) + ' <small style="color:var(--sub)">(' + fmt(r.antriUnit) + ')</small>'},
        {h:'Timbang', num:true, val: r => r.timbangTon, html: r => fmt(r.timbangTon) + ' <small style="color:var(--sub)">(' + fmt(r.timbangUnit) + ')</small>'},
        {h:'Booking', num:true, val: r => r.bookTon, html: r => fmt(r.bookTon)}, {h:'Sudah rilis', num:true, val: r => r.sudahRilis, html: r => fmt(r.sudahRilis)},
        {h:'Potensi', num:true, val: r => r.potensi, html: r => fmt(r.potensi)}, {h:'Prognose', num:true, val: r => r.prognose, html: r => fmt(r.prognose)},
        {h:'Target hari ini', num:true, val: r => r.snopTarget || null, html: r => r.snopTarget ? fmt(r.snopTarget) : '–'},
        {h:'Prognose %', num:true, cls:'pc', val: r => r.pctP, html: r => pctTxt(r.pctP), style: r => 'color:' + tierColor(r.pctP)}];
      const hCols = inc === 'FRC' ? [{h:'Provinsi', cls:'name', val: r => tc(r.prov)}].concat([1, 2, 3].flatMap(h => [
        {h:'H+' + h + ' target', num:true, val: r => r['h' + h + 'TargetSnop'] || null, html: r => r['h' + h + 'TargetSnop'] ? fmt(r['h' + h + 'TargetSnop']) : '–'},
        {h:'H+' + h + ' SO', num:true, val: r => r['h' + h + 'So'] || null, html: r => r['h' + h + 'So'] ? fmt(r['h' + h + 'So']) : '–'}])) : null;
      const hasH = list.some(r => r.h1TargetSnop || r.h1So);
      return '<div class="section-head"><div class="section-title">' + inc + ' — snapshot ' + esc(P.date ? dayLabel(P.date) + ' pukul ' + P.time : '–') + '</div></div>' +
        progCells(rows, inc === 'FOT' ? 'FOT: target internal' : 'jumlah seluruh provinsi') +
        table(cols, list, {go: r => D.provIdx[r.prov] != null ? '#/provinsi/' + enc(r.prov) : null, sort: 6}) +
        (hCols && hasH ? '<div class="section-head"><div class="section-title">Target ' + tgtLbl() + ' vs SO, H+1 s.d. H+3</div></div>' + table(hCols, list, {sort: 1}) : '');
    }).join('');
  } else if(gran === 'mingguan'){
    const end = parseISO(PER().to);
    const mon = new Date(end); mon.setDate(end.getDate() - ((end.getDay() + 6) % 7));
    const days = [...Array(7)].map((_, i) => { const d = new Date(mon); d.setDate(mon.getDate() + i); return d; });
    const data = days.map(d => {
      const s = iso(d), future = d > end, r = future ? rangeTotals(s, s) : rangeTotals(s, s);
      return {from: s, to: s, label: DAYS[d.getDay()].slice(0, 3) + ' ' + d.getDate(), long: DAYS[d.getDay()] + ', ' + dayLabel(s) + (future ? ' (belum terjadi)' : ''), vol: future ? 0 : r.vol, frc: future ? 0 : r.frc, tgt: r.pct == null && !future ? null : r.tgt, pct: future ? null : r.pct, future};
    });
    const sofar = data.filter((d, i) => days[i] <= end);
    const sReal = sofar.reduce((a, d) => a + d.frc, 0), sTgt = sofar.reduce((a, d) => a + (d.tgt || 0), 0), wTgt = data.reduce((a, d) => a + (d.tgt || 0), 0);
    const runrate = sofar.length ? sReal / sofar.length : 0, proj = sReal + runrate * (7 - sofar.length);
    const id = 'tw'; TRENDS[id] = {id, data, sc: {}, gran: 'harian'};
    body = '<div class="prog-row">' +
      '<div class="prog-cell"><div class="pl">Realisasi minggu ini</div><div class="pv">' + fmt(sReal) + ' t</div><div class="pd">' + sofar.length + ' dari 7 hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target s.d. hari ini</div><div class="pv">' + fmt(sTgt) + ' t</div><div class="pd">capaian ' + pctTxt(pctOf(sReal, sTgt)) + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Proyeksi minggu</div><div class="pv" style="color:' + tierColor(pctOf(proj, wTgt)) + '">' + fmt(proj) + ' t</div><div class="pd">laju ' + fmt(runrate) + ' t/hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target minggu penuh</div><div class="pv">' + fmt(wTgt) + ' t</div><div class="pd">proyeksi ' + pctTxt(pctOf(proj, wTgt)) + '</div></div></div>' +
      '<div class="dp-panel" style="margin-bottom:34px"><div class="pt">Realisasi ' + incName() + ' per hari vs target ' + tgtLbl() + ', ' + days[0].getDate() + ' ' + MON3[days[0].getMonth()] + ' – ' + days[6].getDate() + ' ' + MON3[days[6].getMonth()] + '</div>' + barsHTML(id, data) +
      '<div class="trend-legend"><span><i style="background:var(--green)"></i>&ge;100%</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target ' + tgtLbl() + '</span>' + LG_NUM + '</div><div class="trend-detail" id="tw-d">Klik batang untuk rincian hari itu per provinsi, distrik, ekspeditur &amp; source.</div></div>';
  } else {
    const m = PER().month, ld = lastDayOf(m), n = dim(m);
    const rows = [...summarize('prov', m).values()].filter(o => o.hasTgt || o.vol > 0).map(o => {
      const rr = o.frc / ld, proj = o.frc + rr * (n - ld);
      return Object.assign(o, {rr, proj, projPct: o.hasTgt ? pctOf(proj, o.tgtFull) : null});
    }).sort((a, b) => (a.projPct == null ? 999 : a.projPct) - (b.projPct == null ? 999 : b.projPct));
    const t = total(m), rr = t.frc / ld, proj = t.frc + rr * (n - ld);
    const cols = [{h:'Provinsi', cls:'name', val: r => tc(B.dims.prov[r.key])}, {h:'Realisasi ' + incName(), num:true, val: r => r.frc, html: r => fmt(r.frc)},
      {h:'Laju / hari', num:true, val: r => r.rr, html: r => fmt(r.rr)}, {h:'Proyeksi akhir bulan', num:true, val: r => r.proj, html: r => fmt(r.proj)},
      {h:'Target bulan', num:true, val: r => r.hasTgt ? r.tgtFull : null, html: r => r.hasTgt ? fmt(r.tgtFull) : '–'},
      {h:'Proyeksi %', num:true, cls:'pc', val: r => r.projPct, html: r => pctTxt(r.projPct), style: r => 'color:' + tierColor(r.projPct)},
      {h:'Perlu / hari', num:true, val: r => r.hasTgt && n > ld ? (r.tgtFull - r.frc) / (n - ld) : null, html: r => r.hasTgt && n > ld ? fmt(Math.max(0, (r.tgtFull - r.frc) / (n - ld))) : '–'}];
    body = '<div class="prog-row">' +
      '<div class="prog-cell"><div class="pl">Realisasi ' + incName() + '</div><div class="pv">' + fmt(t.frc) + ' t</div><div class="pd">tgl 1–' + ld + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Proyeksi akhir bulan</div><div class="pv" style="color:' + tierColor(pctOf(proj, t.tgtFull)) + '">' + fmt(proj) + ' t</div><div class="pd">laju ' + fmt(rr) + ' t/hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target bulan</div><div class="pv">' + fmt(t.tgtFull) + ' t</div><div class="pd">proyeksi ' + pctTxt(pctOf(proj, t.tgtFull)) + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Perlu per hari</div><div class="pv">' + (n > ld ? fmt(Math.max(0, (t.tgtFull - t.frc) / (n - ld))) : '–') + ' t</div><div class="pd">sisa ' + (n - ld) + ' hari</div></div></div>' +
      table(cols, rows, {go: r => '#/provinsi/' + enc(B.dims.prov[r.key])});
  }
  const pname = {harian: 'Hari ini', mingguan: 'Minggu ini', bulanan: 'Akhir bulan'}[gran];
  page('<div class="wrap">' + crumb([['Prognosa'], [pname]]) +
    '<div class="dp-head"><div><div class="nm">Prognosa — ' + pname + '</div><div class="sub">' + (gran === 'mingguan' ? 'Minggu berjalan, Senin–Minggu' : 'Proyeksi akhir ' + monthLabel(PER().month)) + '</div></div></div>' +
    chips + body + '</div>', 'prognosa');
}

/* Moda laut: one row per vessel (shipName: bundle column 11). Sea rows without a vessel name
   (FAS, containers, ...) are shown as a separate line so the gap is visible instead of silently missing. */
/* "NAMA" for one value, "NAMA +2" for several (full list in the tooltip) */
function setTxt(set, dims, unit){
  const nm = [...set].map(i => tc(dims[i]) || '–'); if(!nm.length) return '–';
  return '<span title="' + esc(nm.join(', ')) + '">' + esc(nm[0]) + (nm.length > 1 ? ' <small>+' + (nm.length - 1) + ' ' + unit + '</small>' : '') + '</span>';
}
/* Voyages: the source has no voyage number. Two rules (agreed with Rendy):
   - regular/feeder ships (e.g. Padang–Mentawai), average load < VOY_SMALL t per loading day over all data: every
     loading day is its own voyage;
   - cargo ships: loading days of one ship at one plant with gaps of at most VOY_GAP days form one voyage. */
const VOY_GAP = 3, VOY_SMALL = 150;
/* Vessel of a ship row: bundle column 11 (B.dims.kapal) since the ekspeditur of FOB/FAS/FOT is the distributor;
   older bundles carried the vessel as ekspeditur "KAPAL <nama>". */
const shipName = r => B.dims.kapal ? (r[11] != null && r[11] >= 0 ? 'KAPAL ' + B.dims.kapal[r[11]] : null)
  : (/^KAPAL /.test(B.dims.eksp[r[6]] || '') ? B.dims.eksp[r[6]] : null);
const shipKind = nm => {   // 'reguler' | 'kargo', from the ship's whole history (stable across filters)
  const c = B._shipKind || (B._shipKind = new Map());
  if(!c.has(nm)){ const d = new Set(); let t = 0; B.ship.forEach(r => { if(shipName(r) === nm){ d.add(r[0]); t += r[8]; } }); c.set(nm, d.size && t / d.size < VOY_SMALL ? 'reguler' : 'kargo'); }
  return c.get(nm);
};
function voyagesOf(rows){
  const gap = rows.length && shipKind(shipName(rows[0])) === 'reguler' ? 0 : VOY_GAP;
  const bySrc = new Map();
  rows.forEach(r => { (bySrc.get(r[2]) || bySrc.set(r[2], []).get(r[2])).push(r); });
  const out = [];
  bySrc.forEach((rs, src) => {
    rs.sort((a, b) => a[0] - b[0]);
    let v = null;
    rs.forEach(r => {
      const ds = B.days[r[0]];
      if(!v || (parseISO(ds) - parseISO(B.days[v.d1])) / 864e5 > gap){
        v = {src, d0: r[0], d1: r[0], vol: 0, trips: 0, days: new Set(), inc: new Set(), distr: new Map(), prov: new Set(), dist: new Set()}; out.push(v);
      }
      v.d1 = Math.max(v.d1, r[0]); v.vol += r[8]; v.trips += r[9]; v.days.add(r[0]); v.inc.add(INCS()[r[1]]);
      v.distr.set(r[5], (v.distr.get(r[5]) || 0) + r[8]); v.prov.add(r[3]); v.dist.add(r[4]);
    });
  });
  return out.sort((a, b) => a.d0 - b.d0);
}
const kapalRows = (m, name) => {   // sea shipments (FOB/FAS/kontainer/kapal), any incoterm of moda LAUT
  const sv = [F.inc, F.moda]; F.inc = 'ALL'; F.moda = 'LAUT';
  try { return shipRows(m).filter(r => name == null || shipName(r) === name); } finally { F.inc = sv[0]; F.moda = sv[1]; }
};
/* Vessel master (admin slot "kapal", public registry facts) for a ship name "KAPAL X"; null when not filled in. */
const kapalInfo = name => { const K = B.dims.kapal || [], i = K.indexOf(String(name).replace(/^KAPAL /, '')); return i >= 0 ? (B.dims.kapalInfo || [])[i] || null : null; };
/* Commons file page -> direct thumbnail (Special:FilePath redirects to upload.wikimedia.org). */
const commonsImg = (u, w) => { const m = /\/wiki\/File:(.+)$/.exec(u || ''); return m ? 'https://commons.wikimedia.org/wiki/Special:FilePath/' + m[1] + '?width=' + (w || 900) : (u || ''); };
const SHIP_SVG = '<svg viewBox="0 0 320 150" class="ks-art" aria-hidden="true"><path d="M18 96h284l-30 34H52z" fill="currentColor" opacity=".9"/><rect x="70" y="62" width="150" height="34" rx="3" fill="currentColor" opacity=".55"/><rect x="190" y="30" width="56" height="66" rx="4" fill="currentColor" opacity=".75"/><rect x="200" y="40" width="36" height="10" rx="2" fill="#0f1a21" opacity=".5"/><rect x="226" y="12" width="10" height="20" fill="currentColor" opacity=".6"/><path d="M0 140c26-8 54-8 80 0s54 8 80 0 54-8 80 0 54 8 80 0" stroke="currentColor" stroke-width="3" fill="none" opacity=".35"/></svg>';
/* 30-day loading timeline per dermaga (loading point = source plant): one row per vessel in order of arrival,
   a light bar spans each voyage, solid blocks mark the loading days. name: only that vessel. */
const dermagaLbl = si => tc(B.dims.src[si] || '(tanpa dermaga)').replace(/^Vp Transhpmnt Pdg Sp$/i, 'VP Transhipment Padang');
function shipGantt(m, name, opt){
  opt = opt || {};
  const end = Math.max(0, B.days.indexOf(m.to) >= 0 ? B.days.indexOf(m.to) : B.days.length - 1), start = Math.max(0, end - 29), N = end - start + 1;
  const rows = kapalRows(makePer('range', B.days[start], B.days[end]), name == null ? null : name).filter(r => shipName(r));
  if(!rows.length) return '';
  const bySrc = new Map();
  rows.forEach(r => { if(opt.src != null && r[2] !== opt.src) return; const g = bySrc.get(r[2]) || bySrc.set(r[2], new Map()).get(r[2]), nm = shipName(r); (g.get(nm) || g.set(nm, []).get(nm)).push(r); });
  const pc = d => ((d - start) / N * 100).toFixed(3) + '%', wd = n => (n / N * 100).toFixed(3) + '%';
  let head = '';
  for(let d = start; d <= end; d++){ const ds = B.days[d], dw = D.dayDow[d], first = d === start || ds.slice(8) === '01';
    head += '<span class="kg-d' + (dw === 0 ? ' sun' : '') + '"' + ' title="' + esc(dayLabel(ds)) + '">' + (first ? '<b>' + MON3[+ds.slice(5, 7) - 1] + '</b>' : '') + +ds.slice(8) + '</span>'; }
  const blocks = [...bySrc.entries()].map(([si, ships]) => {
    const list = [...ships.entries()].map(([nm, rs]) => ({nm, rs, vs: voyagesOf(rs), vol: rs.reduce((a, r) => a + r[8], 0), kind: shipKind(nm)}));
    list.forEach(o => o.first = Math.min(...o.vs.map(v => v.d0)));
    list.sort((a, b) => a.first - b.first || b.vol - a.vol);
    const vol = list.reduce((a, o) => a + o.vol, 0), nv = list.reduce((a, o) => a + o.vs.length, 0);
    const lines = list.map(o => {
      const dayT = new Map(); o.rs.forEach(r => dayT.set(r[0], (dayT.get(r[0]) || 0) + r[8]));
      const bars = o.vs.map(v => '<a class="kg-v ' + o.kind + '" href="#/kapal/' + enc(o.nm) + '" style="left:' + pc(v.d0) + ';width:' + wd(v.d1 - v.d0 + 1) + '" title="' + esc(o.nm.replace(/^KAPAL /, '') + ' · ' + shortDate(B.days[v.d0]) + (v.d1 !== v.d0 ? '–' + shortDate(B.days[v.d1]) : '') + ' · ' + fmt(v.vol) + ' t · ' + v.days.size + ' hari muat') + '">' +
        '</a>' + (v.d1 - v.d0 + 1 >= 2 ? '<span class="kg-lb" style="left:' + pc(v.d0) + ';width:' + wd(v.d1 - v.d0 + 1) + '">' + fmt(v.vol) + ' t</span>' : '')).join('') +
        [...dayT.entries()].map(([d, t]) => '<i class="kg-day ' + o.kind + '" style="left:' + pc(d) + ';width:' + wd(1) + '" title="' + esc(dayLabel(B.days[d]) + ' · ' + fmt(t) + ' t') + '"></i>').join('');
      return '<div class="kg-row"><a class="kg-nm" href="#/kapal/' + enc(o.nm) + '">' + esc(o.nm.replace(/^KAPAL /, '')) + '<small>' + o.kind + ' · ' + fmt(o.vol) + ' t</small></a><div class="kg-track">' + bars + '</div></div>';
    }).join('');
    return {vol, html: (name == null && !opt.bare ? '<div class="inc-grp" style="--c:var(--black)"><b>' + esc(dermagaLbl(si)) + '</b><small>' + fmt(list.length) + ' kapal · ' + fmt(nv) + ' voyage · ' + fmt(vol) + ' t</small></div>' : '') +
      '<div class="kgantt"><div class="kg-row kg-head"><span class="kg-nm">' + (name == null ? 'Kapal (urut kedatangan)' : esc(dermagaLbl(si))) + '</span><div class="kg-days">' + head + '</div></div>' + lines + '</div>'};
  }).sort((a, b) => b.vol - a.vol);
  if(!blocks.length) return '';
  return blocks.map(b => b.html).join('') + (opt.noKey ? '' : shipKey(start, end));
}
const shipKey = (start, end) => '<div class="kg-key"><span><i class="kargo"></i>kapal kargo</span><span><i class="reguler"></i>kapal reguler</span><span><i class="span"></i>rentang voyage</span><span>blok pekat = hari muat · ' + shortDate(B.days[start]) + ' – ' + shortDate(B.days[end]) + '</span></div>';
function pageKapalDetail(name){
  if(!/^KAPAL /.test(name) && !B.ship.some(r => shipName(r) === name) && B.ship.some(r => shipName(r) === 'KAPAL ' + name)) name = 'KAPAL ' + name;   // bare vessel name (search, typed link)
  const m = PER(), rows = kapalRows(m, name);
  if(!rows.length && !B.ship.some(r => shipName(r) === name)) return notFound('Kapal ' + name);
  const vs = voyagesOf(rows), vol = rows.reduce((a, r) => a + r[8], 0), trips = rows.reduce((a, r) => a + r[9], 0);
  const days = new Set(rows.map(r => r[0])), distr = new Map(), prov = new Map(), src = new Map();
  rows.forEach(r => { distr.set(r[5], (distr.get(r[5]) || 0) + r[8]); prov.set(r[3], (prov.get(r[3]) || 0) + r[8]); src.set(r[2], (src.get(r[2]) || 0) + r[8]); });
  vs.forEach((v, i) => v.n = i + 1);
  const gaps = vs.slice(1).map((v, i) => v.d0 - vs[i].d1).filter(g => g > 0).sort((a, b) => a - b);
  const medGap = gaps.length ? gaps[gaps.length >> 1] : null;
  const kind = rows.length ? shipKind(shipName(rows[0])) : '', I = kapalInfo(name) || {};
  const nm = name.replace(/^KAPAL /, ''), title = I.nama_resmi || nm;
  const facts = [['Jenis', I.jenis], ['IMO', I.imo], ['MMSI', I.mmsi], ['GT', I.gt && fmt(+I.gt) || I.gt], ['DWT', I.dwt && fmt(+I.dwt) || I.dwt], ['Panjang', I.panjang_m && I.panjang_m + ' m'], ['Tahun buat', I.tahun_buat], ['Operator / pemilik', I.operator], ['Pelabuhan asal', I.pelabuhan_asal]].filter(f => f[1]);
  const conf = I.keyakinan ? '<span class="ks-conf ks-' + esc(I.keyakinan.toLowerCase()) + '">keyakinan ' + esc(I.keyakinan.toLowerCase()) + '</span>' : '';
  const srcLinks = String(I.sumber || '').split(/[\s;,]+/).filter(u => /^https:\/\//.test(u)).slice(0, 3);
  const photo = I.foto_url ? '<img class="ks-photo" src="' + esc(commonsImg(I.foto_url)) + '" alt="' + esc(title) + '" loading="lazy" referrerpolicy="no-referrer">' + (I.foto_kredit ? '<a class="bg-credit" href="' + esc(I.foto_url) + '" target="_blank" rel="noopener">Foto: ' + esc(I.foto_kredit) + '</a>' : '') : '<div class="ks-noart">' + SHIP_SVG + '<div>Foto belum tersedia</div></div>';
  const rr = (mp, lbl, href, h) => volTable([...mp.entries()].sort((a, b) => b[1] - a[1]).map(([k, t]) => ({key: k, vol: t})), h, o => lbl(o.key), o => href(o.key));
  const cols = [{h:'Voyage', cls:'rank', val: v => v.n, html: v => '#' + v.n},
    {h:'Muat', cls:'name', val: v => v.d0, html: v => shortDate(B.days[v.d0]) + (v.d1 !== v.d0 ? ' – ' + shortDate(B.days[v.d1]) : '') + '<small>' + v.days.size + ' hari muat · ' + esc(B.dims.src[v.src] || '') + '</small>'},
    {h:'Tonase (t)', num:true, val: v => v.vol, html: v => '<b>' + fmt(v.vol) + '</b>'}, {h:'SPJ', num:true, val: v => v.trips},
    {h:'Incoterm', val: v => [...v.inc].join(', '), html: v => esc([...v.inc].map(incLbl).join(', '))},
    {h:'Distributor', val: v => v.distr.size, html: v => [...v.distr.entries()].sort((a, b) => b[1] - a[1]).map(([d, t]) => '<a href="#/distributor/' + enc(B.dims.distr[d]) + '">' + esc(distrLbl(d) || '–') + '</a> <small>' + fmt(t) + ' t</small>').join('<br>')},
    {h:'Tujuan', val: v => v.prov.size, html: v => setTxt(v.dist, B.dims.dist, 'distrik') + '<small>' + esc([...v.prov].map(p => tc(B.dims.prov[p])).join(', ')) + '</small>'}];
  page('<div class="wrap">' + crumb([['Kapal', '#/kapal'], [nm]]) +
    '<section class="ks-hero"><div class="ks-left"><div class="eyebrow">Kapal ' + esc(kind) + ' · moda laut · ' + esc(periodSub()) + '</div><h2 class="mh-h">' + esc(title) + '</h2>' +
      (I.nama_resmi && I.nama_resmi.toUpperCase() !== nm ? '<div class="ks-alias">di Rilis: ' + esc(nm) + '</div>' : '') +
      '<div class="mh-chips">' + heroChip('Voyage', fmt(vs.length), kind === 'reguler' ? 'tiap hari muat = 1 voyage' : 'jeda muat ≤ ' + VOY_GAP + ' hari = 1') + heroChip('Tonase', fmt(vol) + ' t', vs.length ? fmt(vol / vs.length) + ' t / voyage' : '') +
        heroChip('Jeda antar voyage', medGap == null ? '–' : fmt(medGap) + ' hari', medGap == null ? '' : 'median') + heroChip('Distributor', fmt(distr.size), fmt(days.size) + ' hari muat · ' + fmt(trips) + ' SPJ') + '</div>' +
      (facts.length ? '<dl class="ks-facts">' + facts.map(f => '<div><dt>' + f[0] + '</dt><dd>' + esc(String(f[1])) + '</dd></div>').join('') + '</dl>' : '<div class="ks-empty">Data registri kapal (IMO, GT, operator) belum ada di master kapal.</div>') +
      (srcLinks.length ? '<div class="ks-src">' + conf + (I.catatan && facts.length ? esc(I.catatan) + ' · ' : '') + 'Sumber: ' + srcLinks.map(u => '<a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(u.replace(/^https:\/\/(www\.)?/, '').split('/')[0]) + '</a>').join(' · ') + '</div>' : '') +
    '</div><div class="ks-right">' + photo + '</div></section>' +
    '<div class="section-head"><div class="section-title">Timeline 30 hari</div></div>' + (shipGantt(m, name) || '<div class="dp-review">Tidak ada muatan kapal ini dalam 30 hari terakhir periode.</div>') +
    (vs.length ? '' : '<div class="dp-review">Tidak ada muatan kapal ini pada filter (plant, kemasan, periode).' + (() => { let l = -1; B.ship.forEach(r => { if(shipName(r) === name) l = Math.max(l, r[0]); }); return l >= 0 ? ' Muat terakhir di data: <b>' + dayLabel(B.days[l]) + '</b> — ubah periode untuk melihat riwayatnya.' : ''; })() + '</div>') +
    (rows.length ? '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Muat dari</div>' + rr(src, k => tc(B.dims.src[k] || '(tanpa plant)'), k => B.dims.src[k] ? '#/plant/' + enc(B.dims.src[k]) : null, 'Plant / dermaga') + '</div>' +
      '<div class="dp-panel"><div class="pt">Tujuan provinsi</div>' + rr(prov, k => tc(B.dims.prov[k]), k => '#/provinsi/' + enc(B.dims.prov[k]), 'Provinsi') + '</div></div>' : '') +
    (vs.length ? '<div class="section-head"><div class="section-title">Riwayat voyage</div></div>' + table(cols, vs.slice().reverse(), {limit: 200}) : '') +
    (rows.length ? '<div class="section-head"><div class="section-title">Tujuan pengiriman</div></div>' + treeHTML(buildTree(rows, ['prov', 'dist', 'distr']), ['prov', 'dist', 'distr']) : '') +
    '</div>', 'armada');
}
function pageKapal(){
  const sv = [F.inc, F.moda]; F.inc = 'ALL'; F.moda = 'LAUT';
  try {
    const m = PER(), rows = shipRows(m), tot = rows.reduce((a, r) => a + r[8], 0);
    const agg = rs => {
      const by = new Map();
      rs.forEach(r => {
        const nm = shipName(r), k = nm || '\u0000';
        let o = by.get(k); if(!o){ o = {k, nm: k === '\u0000' ? '(kapal tidak tercatat di sumber)' : nm, vol: 0, trips: 0, days: new Set(), inc: new Set(), distr: new Set(), prov: new Set(), dist: new Set(), last: -1, rows: []}; by.set(k, o); }
        o.rows.push(r); o.vol += r[8]; o.trips += r[9]; o.days.add(r[0]); o.inc.add(INCS()[r[1]]); o.distr.add(r[5]); o.prov.add(r[3]); o.dist.add(r[4]); o.last = Math.max(o.last, r[0]);
      });
      const list = [...by.values()].sort((a, b) => (a.k === '\u0000') - (b.k === '\u0000') || b.vol - a.vol);
      list.forEach(o => { o.voy = o.k === '\u0000' ? null : voyagesOf(o.rows).length; o.rows = null; });
      list.forEach((o, i) => o.rank = o.k === '\u0000' ? null : i + 1);
      return list;
    };
    const cols = [{h:'#', cls:'rank', val: o => o.rank || 999, html: o => o.rank ? '#' + o.rank : '–'}, {h:'Kapal', cls:'name', val: o => o.nm, html: o => { const I = o.k === '\u0000' ? null : kapalInfo(o.nm); return esc(o.nm.replace(/^KAPAL /, '')) + '<small>' + esc([o.k === '\u0000' ? '' : 'kapal ' + shipKind(o.nm), I && I.jenis, I && I.gt && 'GT ' + I.gt].filter(Boolean).join(' · ')) + '</small>'; }},
      {h:'Voyage', num:true, val: o => o.voy == null ? -1 : o.voy, html: o => o.voy == null ? '–' : fmt(o.voy)}, {h:'Tonase (t)', num:true, val: o => o.vol, html: o => '<b>' + fmt(o.vol) + '</b>'},
      {h:'t / voyage', num:true, val: o => o.voy ? o.vol / o.voy : 0, html: o => o.voy ? fmt(o.vol / o.voy) : '–'}, {h:'Hari muat', num:true, val: o => o.days.size},
      {h:'Muat terakhir', val: o => o.last, html: o => o.last >= 0 ? shortDate(B.days[o.last]) : '–'},
      {h:'Distributor', val: o => o.distr.size, html: o => setTxt(o.distr, B.dims.distr.map((x, i) => distrLbl(i)), 'distributor')}, {h:'Provinsi', val: o => o.prov.size, html: o => setTxt(o.prov, B.dims.prov, 'provinsi')},
      {h:'Distrik', val: o => o.dist.size, html: o => setTxt(o.dist, B.dims.dist, 'distrik')}];
    const all = agg(rows), named = all.filter(o => o.k !== '\u0000'), unk = all.find(o => o.k === '\u0000');
    // grouping: per dermaga (loading point = source plant); vessel names exist for FOB only, so an incoterm split adds nothing
    const srcs = [...new Set(rows.map(r => r[2]))].map(si => ({si, rs: rows.filter(r => r[2] === si)})).map(g => Object.assign(g, {vol: g.rs.reduce((a, r) => a + r[8], 0)})).sort((a, b) => b.vol - a.vol);
    const groups = tabsHTML('kapal', srcs.map(g => {
      const list = agg(g.rs), c = cols.map(x => Object.assign({}, x)), n = list.filter(o => o.k !== '\u0000');
      c[2].total = fmt(list.reduce((a, o) => a + (o.voy || 0), 0)); c[3].total = fmt(g.vol);
      const tl = shipGantt(m, null, {src: g.si, bare: true, noKey: true});
      const h = '<section class="krm-grp"><div class="krm-h"><b>' + esc(dermagaLbl(g.si)) + '</b><small>' + fmt(n.length) + ' kapal · ' + fmt(g.vol) + ' t · ' + pctTxt(pctOf(g.vol, tot)) + ' tonase laut · ' + esc(periodSub()) + '</small></div>' +
        '<div class="krm-sub">Timeline 30 hari</div>' + (tl || '<div class="dp-review">Tidak ada kapal bernama dalam 30 hari terakhir.</div>') +
        '<div class="krm-sub">Detail kapal</div>' + table(c, list, {total: true, limit: 400, go: o => o.k === '\u0000' ? null : '#/kapal/' + enc(o.nm)}) + '</section>';
      return {k: String(g.si), label: dermagaLbl(g.si), n: n.length, html: h};
    }));
    const e30 = Math.max(0, B.days.indexOf(m.to) >= 0 ? B.days.indexOf(m.to) : B.days.length - 1);
    page('<div class="wrap">' + crumb([['Kapal']]) +
      '<div class="dp-head"><div><div class="nm">Kapal · moda laut</div><div class="sub">' + fmt(named.length) + ' kapal di ' + fmt(srcs.length) + ' dermaga · voyage: kapal reguler (rata-rata < ' + VOY_SMALL + ' t/hari muat) = tiap hari muat; kapal kargo = hari muat berjeda ≤ ' + VOY_GAP + ' hari · ' + periodSub() + '</div></div></div>' +
      '<div class="dp-hint" style="font-size:12px;color:var(--sub);margin:-6px 0 14px">Filter aktif: ' + srcTxt() + ' · kemasan <b>' + (F.kem === 'ALL' ? 'semua' : esc(F.kem)) + '</b> · <a class="section-link" data-act="openfilter" style="cursor:pointer;color:var(--red);font-weight:700">ubah filter</a></div>' +
      (unk && tot ? '<div class="dp-review">' + pctTxt(pctOf(unk.vol, tot)) + ' tonase laut periode ini (' + fmt(unk.vol) + ' t) belum punya nama kapal di sumber Rilis. Yang tercatat nama kapalnya baru FOB; FAS, kontainer, FRC KAPAL, FRC KONTAINER dan FOB KAPAL (polysling) menunggu data tambahan.</div>' : '') +
      '<div class="section-head"><div class="section-title">Kapal per dermaga</div><div class="section-sub">timeline: kapal masuk berurutan dari atas · klik bar untuk detail kapal</div></div>' + shipKey(Math.max(0, e30 - 29), e30) +
      (rows.length ? groups : '<div class="dp-review">Tidak ada pengiriman laut pada filter ini (plant, kemasan, periode).</div>') + '</div>', 'armada');
  } finally { F.inc = sv[0]; F.moda = sv[1]; }
}
/* ===================== Armada =====================
   Three questions, per the agreed concept:
   1. Kinerja rilis: trucks loaded per day, time in the plant (SLA masuk -> keluar) and queue (antri -> masuk), per plant.
   2. Koridor (plant -> tujuan): cycle time (days between two trips of the same truck), t/trip and fleet need = trip/hari x siklus.
      Destination = province; Sumatera Barat from Indarung per district (Padang vs Pasaman/Dharmasraya differ too much).
   3. Scorecard ekspeditur: judged against the median cycle of the corridors it serves, not the overall average.
   Truck rows: [day, truck, inc, eksp, src, prov, trips, ton, dwS, dwN, kem, dist, waitS, waitN] */
// trucks only: sea incoterms (vessel loads in NOPOL) and loads above 45 t/trip are not truck trips
const truckOk = t => incOk(t[2]) && srcOk(t[4]) && kemOk(t[10]) && areaProvOk(t[5]) && (F.moda === 'LAUT' || modaOf(INCS()[t[2]]) !== 'LAUT') && t[7] <= 45 * Math.max(1, t[6]);
const corrKey = t => { const sn = B.dims.src[t[4]] || '', pn = B.dims.prov[t[5]] || ''; return t[4] + '|' + t[5] + (/INDARUNG/.test(sn) && pn === 'SUMATERA BARAT' && t[11] != null ? '|' + t[11] : ''); };
const corrLbl = k => { const [si, pi, di] = k.split('|'); return tc(B.dims.src[si] || '?').replace(/^(Cp|Pp|Gp|Dc|Vp) /, '') + ' → ' + (di != null ? tc(B.dims.dist[di]) : tc(B.dims.prov[pi])); };
const median = a => { if(!a.length) return null; const b = a.slice().sort((x, y) => x - y), m = b.length >> 1; return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2; };
/* cycle gaps over the 60 days up to the period end: gap between consecutive loading days of a truck (1..12 days),
   booked on the corridor and ekspeditur of the earlier trip */
function cycleGaps(per){
  const endI = B.days.indexOf(per.to) >= 0 ? B.days.indexOf(per.to) : B.days.length - 1, startI = Math.max(0, endI - 59);
  const by = new Map();
  B.trucks.forEach(t => { if(t[0] < startI || t[0] > endI || !truckOk(t)) return; (by.get(t[1]) || by.set(t[1], []).get(t[1])).push(t); });
  const corr = new Map(), ek = new Map(), sub = new Map();
  by.forEach(rs => {
    rs.sort((a, b) => a[0] - b[0]);
    for(let i = 1; i < rs.length; i++){
      const g = rs[i][0] - rs[i - 1][0]; if(g < 1 || g > 12) continue;
      const k = corrKey(rs[i - 1]); (corr.get(k) || corr.set(k, []).get(k)).push(g);
      const e = rs[i - 1][3]; (ek.get(e) || ek.set(e, []).get(e)).push([k, g]);
      [k + '#e' + e, k + '#t' + rs[i - 1][1], 's' + rs[i - 1][4]].forEach(x => (sub.get(x) || sub.set(x, []).get(x)).push(g));
    }
  });
  const med = new Map([...corr.entries()].map(([k, a]) => [k, median(a)]));
  const avg = new Map([...corr.entries()].map(([k, a]) => [k, a.reduce((x, y) => x + y, 0) / a.length]));
  return {med, avg, n: new Map([...corr.entries()].map(([k, a]) => [k, a.length])), ek, sub};
}
/* ===================== Tabs (long pages) =====================
   tabsHTML(id, [{k, label, html, n}]): one bar + panes; empty panes are dropped; the chosen tab is remembered per page
   (sessionStorage, best effort). The bar sticks under the header + filter strip (--stick, set by setStick). */
function tabsHTML(id, list){
  list = list.filter(t => t && t.html);
  if(list.length < 2) return list.map(t => t.html).join('');
  let act = null; try { act = sessionStorage.getItem('tab:' + id); } catch(e){}
  if(!list.some(t => t.k === act)) act = list[0].k;
  return '<div class="tabbar" role="tablist" data-tabs="' + id + '">' + list.map(t => '<button class="tab' + (t.k === act ? ' on' : '') + '" role="tab" data-tab="' + id + ':' + t.k + '">' + esc(t.label) + (t.n != null ? '<small>' + t.n + '</small>' : '') + '</button>').join('') + '</div>' +
    list.map(t => '<div class="tabpane" data-pane="' + id + ':' + t.k + '"' + (t.k === act ? '' : ' hidden') + '>' + t.html + '</div>').join('');
}
function tabClick(btn){
  const [id, k] = btn.dataset.tab.split(':');
  $$('[data-tab^="' + id + ':"]').forEach(b => b.classList.toggle('on', b === btn));
  $$('[data-pane^="' + id + ':"]').forEach(p => { p.hidden = p.dataset.pane !== id + ':' + k; });
  try { sessionStorage.setItem('tab:' + id, k); } catch(e){}
  const bar = btn.closest('.tabbar'), top = bar.getBoundingClientRect().top, stick = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--stick')) || 0;
  if(top < stick - 1) window.scrollBy(0, top - stick);   // bar is stuck: bring the new pane's start into view
}
// tab bar sticks right under whatever stays on screen: the header, plus the filter strip only where it is sticky (desktop; on phones it scrolls away)
function setStick(){
  const stuck = el => el && /sticky|fixed/.test(getComputedStyle(el).position) ? el.getBoundingClientRect().height : 0;
  document.documentElement.style.setProperty('--stick', Math.round(stuck($('header')) + stuck($('#filterStrip'))) + 'px');
}
window.addEventListener('resize', setStick);

/* ===================== Kontrak angkutan (slot "kontrak") =====================
   B.kontrak: [eksp, src, [dist..], sub, kuota, volEksThn, amendVol, amendFrom, from, to, lead, cycle, dedRute, toa, distName]
   Contract = jasa transportasi darat semen SAK (ZAK + jumbo bag), FRC only, and only for the awarded ekspeditur x route.
   Estimasi bulanan = volume tahunan ekspeditur / 12 (amandemen dari bulannya); bulan berjalan pro-rata s.d. data terakhir. */
function kontrakModel(){
  const K = B.kontrak || []; if(!K.length) return null;
  const frc = INCS().indexOf('FRC'), sak = k => (KEMS()[k] || 'ZAK') !== 'CURAH';
  const asM = B.asOf.slice(0, 7), first = K.map(r => r[8]).filter(Boolean).sort()[0] || asM;
  const months = []; for(let m = first; m <= asM; m = nextMonth(m)) months.push(m);
  const estM = (r, m) => { if((r[8] && m < r[8]) || (r[9] && m > r[9])) return 0; const v = (r[6] && r[7] && m >= r[7] ? r[6] : r[5]) || 0; return v / 12 * (m === asM ? +B.asOf.slice(8) / dim(m) : 1); };
  const aw = new Map(); K.forEach((r, i) => r[2].forEach(di => aw.set(r[0] + '|' + r[1] + '|' + di, i)));
  const awD = new Set(); K.forEach(r => r[2].forEach(di => awD.add(r[0] + '|' + di)));
  const eks = [...new Set(K.map(r => r[0]))], M = new Map();
  const cell = (e, m) => { const k = e + '|' + m; let o = M.get(k); if(!o){ o = {e, m, est: 0, real: 0, luar: 0, so: 0, okDay: 0, soDay: 0, rt: new Map()}; M.set(k, o); } return o; };
  K.forEach((r, i) => months.forEach(m => { const o = cell(r[0], m), v = estM(r, m); o.est += v; const x = o.rt.get(i) || {est: 0, real: 0, so: 0}; x.est += v; o.rt.set(i, x); }));
  const eSet = new Set(eks), dayReal = new Map();
  B.facts.forEach(f => {
    if(f[1] !== frc || !sak(f[10]) || !eSet.has(f[5])) return;
    const m = B.days[f[0]].slice(0, 7); if(m < first) return;
    const ri = aw.get(f[5] + '|' + f[2] + '|' + f[4]), o = cell(f[5], m);
    if(ri == null){ o.luar += f[6]; return; }
    o.real += f[6]; const x = o.rt.get(ri) || {est: 0, real: 0, so: 0}; x.real += f[6]; o.rt.set(ri, x);
    const dk = B.days[f[0]] + '|' + f[4] + '|' + f[5]; dayReal.set(dk, (dayReal.get(dk) || 0) + f[6]);
  });
  /* SO per contract proportion (Rendy): all FRC sak SO of a contracted district, whoever it was booked to, is split over
     that district's winners by their contract volume for the month (= kuota; amendments and released routes included).
     Real vs this "SO kontrak" shows who carries more or less than its share. */
  const byDist = new Map();   // dist -> [{i, e}]
  K.forEach((r, i) => r[2].forEach(di => { const g = byDist.get(di) || []; g.push(i); byDist.set(di, g); }));
  const shareCache = new Map();
  const shares = (di, m) => { const k = di + '|' + m; if(shareCache.has(k)) return shareCache.get(k);
    const ix = byDist.get(di) || [], v = ix.map(i => estM(K[i], m)), t = v.reduce((a, x) => a + x, 0);
    const out = t > 0 ? ix.map((i, j) => [i, v[j] / t]) : []; shareCache.set(k, out); return out; };
  const soDay = new Map();
  (B.sod || []).forEach(r => {   // [date, prov, dist, eksp, ton, kem]
    const m = r[0].slice(0, 7); if(m < first || !byDist.has(r[2]) || !sak(r[5])) return;
    shares(r[2], m).forEach(([i, f]) => {
      const e = K[i][0], o = cell(e, m), v = r[4] * f; o.so += v;
      const x = o.rt.get(i) || {est: 0, real: 0, so: 0}; x.so += v; o.rt.set(i, x);
      const k = r[0] + '|' + r[2] + '|' + e; soDay.set(k, (soDay.get(k) || 0) + v);
    });
  });
  soDay.forEach((v, k) => { const [day, , e] = k.split('|'), o = cell(+e, day.slice(0, 7)); const rl = dayReal.get(k) || 0; o.soDay += v; o.okDay += Math.min(rl, v); });
  // route share on shared routes (kuota vs aktual), per month
  return {K, months, eks, cell, asM};
}
function kontrakSection(){
  const KM = kontrakModel(); if(!KM) return '';
  const {months, eks, cell} = KM, mSel = PER().month && months.includes(PER().month) ? PER().month : KM.asM;
  const rows = eks.map(e => { const o = {key: e, per: months.map(m => cell(e, m))}; o.cur = cell(e, mSel); return o; }).sort((a, b) => b.cur.est - a.cur.est);
  const pc = (a, b) => b > 0 ? a / b * 100 : null;
  const cols = [{h:'Ekspeditur', cls:'name', val: r => ekspLbl(r.key)}].concat(months.map((m, i) => ({h: monthShort(m) + (m === KM.asM ? ' *' : ''), num:true, val: r => pc(r.per[i].real, r.per[i].est),
      html: r => { const o = r.per[i], p = pc(o.real, o.est); return p == null ? '–' : '<b style="color:' + tierColor(p) + '">' + pctTxt(p) + '</b><small>R/SO ' + pctTxt(pc(o.real, o.so)) + '</small>'; }})),
    [{h:'OF harian ' + monthShort(mSel), num:true, val: r => pc(r.cur.okDay, r.cur.soDay), html: r => { const p = pc(r.cur.okDay, r.cur.soDay); return p == null ? '–' : '<span style="color:' + tierColor(p) + '">' + pctTxt(p) + '</span>'; }},
     {h:'Di luar rute ' + monthShort(mSel), num:true, val: r => r.cur.luar, html: r => r.cur.luar ? fmt(r.cur.luar) + ' t' : '–'}]);
  TRENDS.kontrak = {rows, mSel, KM};
  // shared routes: kuota vs aktual (selected month)
  const byRoute = new Map();
  KM.K.forEach((r, i) => { const k = r[1] + '|' + r[14]; const g = byRoute.get(k) || []; g.push(i); byRoute.set(k, g); });
  const shared = [...byRoute.entries()].filter(([, ix]) => ix.length > 1).map(([k, ix]) => {
    const parts = ix.map(i => { const r = KM.K[i], x = cell(r[0], mSel).rt.get(i) || {real: 0, est: 0}; return {e: r[0], kuota: r[4], real: x.real, est: x.est}; });
    const tot = parts.reduce((a, p) => a + p.real, 0); return {k, n: tc(KM.K[ix[0]][14]), parts, tot};
  }).filter(r => r.tot > 0).sort((a, b) => b.tot - a.tot);
  const shTbl = shared.length ? table([{h:'Rute (dari ' + tc(B.dims.src[KM.K[0][1]] || 'plant') + ')', cls:'name', val: r => r.n, html: r => esc(r.n) + '<small>' + fmt(r.tot) + ' t</small>'},
    {h:'Kuota kontrak → porsi aktual', val: r => r.tot, html: r => r.parts.sort((a, b) => b.real - a.real).map(p => { const a = pctOf(p.real, r.tot) || 0, q = (p.kuota || 0) * 100, off = Math.abs(a - q) > 15;
      return '<span class="kq' + (off ? ' off' : '') + '"><b>' + esc(B.dims.eksp[p.e]) + '</b> ' + pctTxt(q) + ' → ' + pctTxt(a) + '</span>'; }).join('')}], shared, {sort: 0}) : '';
  // whole contract period to date: who carries more / less than its contract share of the SO
  const sum = rows.map(r => r.per.reduce((a, o) => ({key: r.key, est: a.est + o.est, real: a.real + o.real, so: a.so + o.so, ok: a.ok + o.okDay, sd: a.sd + o.soDay, luar: a.luar + o.luar}), {est: 0, real: 0, so: 0, ok: 0, sd: 0, luar: 0}));
  const pcC = (v, h) => v == null ? '–' : '<b style="color:' + tierColor(v) + '">' + pctTxt(v) + '</b>' + (h || '');
  const soCols = [{h:'Ekspeditur', cls:'name', val: r => ekspLbl(r.key)}, {h:'Estimasi (t)', num:true, val: r => r.est, html: r => fmt(r.est)}, {h:'Realisasi (t)', num:true, val: r => r.real, html: r => fmt(r.real)},
    {h:'Capaian', num:true, val: r => pc(r.real, r.est), html: r => pcC(pc(r.real, r.est))}, {h:'SO kontrak (t)', num:true, val: r => r.so, html: r => fmt(r.so)},
    {h:'Real / SO', num:true, val: r => pc(r.real, r.so), html: r => { const v = pc(r.real, r.so); return pcC(v, v == null ? '' : '<small>' + (v >= 105 ? 'lebih dari porsi' : v >= 90 ? 'sesuai porsi' : v >= 50 ? 'kurang dari porsi' : 'kritis') + '</small>'); }},
    {h:'OF harian', num:true, val: r => pc(r.ok, r.sd), html: r => pcC(pc(r.ok, r.sd))},
    {h:'Di luar rute (t)', num:true, val: r => r.luar, html: r => r.luar ? fmt(r.luar) : '–'}];
  ['est', 'real', 'so', 'luar'].forEach((k, i) => soCols[[1, 2, 4, 7][i]].total = fmt(sum.reduce((a, r) => a + r[k], 0)));
  soCols[0].total = '<b>Total</b>';
  const T = sum.reduce((a, r) => ({est: a.est + r.est, real: a.real + r.real, so: a.so + r.so, ok: a.ok + r.ok, sd: a.sd + r.sd}), {est: 0, real: 0, so: 0, ok: 0, sd: 0});
  soCols[3].total = pctTxt(pc(T.real, T.est)); soCols[5].total = pctTxt(pc(T.real, T.so)); soCols[6].total = pctTxt(pc(T.ok, T.sd));
  return {mSel, months, cur: rows.map(r => r.cur), sum, T,
    note: '<div class="tab-note">realisasi rute pemenang vs estimasi kontrak (volume tahunan ÷ 12; amandemen dari bulannya) · R/SO = realisasi vs SO rute yang dibagi sesuai proporsi kontrak · semen sak, FRC · * bulan berjalan pro-rata s.d. ' + shortDate(B.asOf) + ' · klik baris untuk rincian rute ' + monthShort(mSel) + '</div>',
    monthly: table(cols, rows, {go: r => 'kontrak:' + r.key, sort: months.length}),
    soTbl: '<div class="tab-note">' + monthShort(months[0]) + '–' + monthShort(months[months.length - 1]) + ' · SO kontrak = seluruh SO FRC sak di distrik kontrak, dibagi ke pemenang sesuai porsi volume kontrak · Real/SO &gt; 100% = mengangkut porsi ekspeditur lain · OF harian = SO terpenuhi di hari yang sama</div>' +
      table(soCols, sum, {total: true, sort: 5, go: r => 'kontrak:' + r.key}),
    shared: shTbl ? '<div class="tab-note">' + monthLabel(mSel) + ' · merah = porsi aktual selisih lebih dari 15 poin dari kuota</div>' + shTbl : ''};
}
function pageKontrak(){
  const S = kontrakSection();
  if(!S){ page('<div class="wrap">' + crumb([['Kontrak']]) + '<div class="dp-review">Data kontrak angkutan belum diupload (slot kontrak).</div></div>', 'kontrak'); return; }
  const c = S.cur.reduce((a, o) => ({est: a.est + o.est, real: a.real + o.real, so: a.so + o.so, ok: a.ok + o.okDay, sd: a.sd + o.soDay, luar: a.luar + o.luar}), {est: 0, real: 0, so: 0, ok: 0, sd: 0, luar: 0});
  const pc = (a, b) => b > 0 ? a / b * 100 : null, ml = monthLabel(S.mSel);
  const low = S.sum.filter(r => pc(r.real, r.so) != null && pc(r.real, r.so) < 50).map(r => B.dims.eksp[r.key]);
  page('<div class="wrap">' + crumb([['Kontrak']]) +
    '<div class="dp-head"><div><div class="nm">Kontrak angkutan</div><div class="sub">' + fmt(S.sum.length) + ' ekspeditur pemenang · FRC truk, semen sak · rute sesuai pengumuman pemenang</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('target', tierColor(pc(c.real, c.est) || 0), 'Capaian kontrak · ' + ml, pctTxt(pc(c.real, c.est)), {hint: fmt(c.real) + ' dari estimasi ' + fmt(c.est) + ' t'}) +
      kpi('box', tierColor(pc(c.real, c.so) || 0), 'Real / SO kontrak', pctTxt(pc(c.real, c.so)), {hint: 'SO ' + fmt(c.so) + ' t'}) +
      kpi('clock', tierColor(pc(c.ok, c.sd) || 0), 'OF harian', pctTxt(pc(c.ok, c.sd)), {hint: 'SO terpenuhi di hari yang sama'}) +
      kpi('pin', low.length ? 'var(--red)' : 'var(--black)', 'Kritis (Real/SO &lt; 50%)', fmt(low.length) + ' <small>ekspeditur</small>', {hint: low.length ? esc(low.join(', ')) + ' · sejak ' + monthShort(S.months[0]) : 'tidak ada'}) +
    '</div>' + tabsHTML('kontrak', [
      {k: 'so', label: 'Kinerja vs SO kontrak', html: S.soTbl},
      {k: 'bulan', label: 'Capaian per bulan', html: S.note + S.monthly},
      {k: 'rute', label: 'Rute bersama', html: S.shared}]) +
    '</div>', 'kontrak');
}
function kontrakDetail(e){
  const T = TRENDS.kontrak; if(!T) return;
  const o = T.KM.cell(e, T.mSel), K = T.KM.K;
  const rs = [...o.rt.entries()].map(([i, x]) => ({r: K[i], ...x})).filter(x => x.est > 0 || x.real > 0).sort((a, b) => b.est - a.est);
  openModal(ekspLbl(e) + ' · kontrak ' + monthLabel(T.mSel), 'per rute pemenang', table([
    {h:'Rute', cls:'name', val: x => x.r[14], html: x => esc(tc(x.r[14])) + '<small>' + esc(tc(B.dims.src[x.r[1]] || '')) + (x.r[4] != null && x.r[4] < 1 ? ' · kuota ' + pctTxt(x.r[4] * 100) : '') + (x.r[9] ? ' · s.d. ' + monthShort(x.r[9]) : '') + (x.r[6] ? ' · amandemen' : '') + '</small>'},
    {h:'Estimasi (t)', num:true, val: x => x.est, html: x => fmt(x.est)}, {h:'Realisasi (t)', num:true, val: x => x.real, html: x => fmt(x.real)},
    {h:'Capaian', num:true, cls:'pc', val: x => x.est ? x.real / x.est * 100 : null, html: x => x.est ? pctTxt(x.real / x.est * 100) : '–', style: x => 'color:' + tierColor(x.est ? x.real / x.est * 100 : null)},
    {h:'SO kontrak (t)', num:true, val: x => x.so, html: x => x.so ? fmt(x.so) : '–'},
    {h:'Real / SO', num:true, cls:'pc', val: x => x.so ? x.real / x.so * 100 : null, html: x => x.so ? pctTxt(x.real / x.so * 100) : '–', style: x => 'color:' + tierColor(x.so ? x.real / x.so * 100 : null)},
    {h:'Lead / siklus (hari)', num:true, val: x => x.r[10], html: x => (x.r[10] != null ? fmt1(x.r[10]) : '–') + ' / ' + (x.r[11] != null ? fmt1(x.r[11]) : '–')},
    {h:'TOA (Rp/t)', num:true, val: x => x.r[13], html: x => x.r[13] ? fmt(x.r[13]) : '–'}], rs, {sort: 1}) +
    (o.luar ? '<div class="dp-review">Di luar rute pemenang: ' + fmt(o.luar) + ' t pada bulan ini.</div>' : ''));
}
/* Koridor (Rendy): one table per plant, rows = tujuan (provinsi; distrik for Sumbar from Indarung).
   Click a tujuan = its ekspeditur, click an ekspeditur = its trucks, click a truck = the truck page (#/truk/NOPOL).
   Kebutuhan armada at the actual corridor cycle: plan = contract t/day ÷ t/trip; real = trips/day; selisih over contract ekspeditur only. */
// contract plan per day keyed corridor#eksp: FRC awarded route volume / 12 / days in month, split evenly over the route's districts
function kontrakPlanDay(mon){
  const out = new Map(), K = B.kontrak || [];
  K.forEach(r => {
    if((r[8] && mon < r[8]) || (r[9] && mon > r[9]) || !r[2].length) return;
    const v = ((r[6] && r[7] && mon >= r[7] ? r[6] : r[5]) || 0) / 12 / dim(mon) / r[2].length;
    const sn = B.dims.src[r[1]] || '';
    r[2].forEach(di => { const pv = B.dims.distProv[di], k = r[1] + '|' + pv + (/INDARUNG/.test(sn) && B.dims.prov[pv] === 'SUMATERA BARAT' ? '|' + di : '') + '#' + r[0]; out.set(k, (out.get(k) || 0) + v); });
  });
  return out;
}
let KOR = null;
const korSum = (rs, i) => rs.reduce((a, t) => a + t[i], 0);
function korStat(rs, cyc){ const tr = korSum(rs, 6), ton = korSum(rs, 7), n = korSum(rs, 9), wn = korSum(rs, 13);
  return {trips: tr, ton, tpt: ton / Math.max(1, tr), cyc, plan: null, real: null, rk: null, trucks: new Set(rs.map(t => t[1])).size, sla: n ? korSum(rs, 8) / n : null, wait: wn ? korSum(rs, 12) / wn : null}; }
function korEksp(c){
  const K = KOR, byE = new Map(); K.byC.get(c.k).forEach(t => (byE.get(t[3]) || byE.set(t[3], []).get(t[3])).push(t));
  return [...byE.entries()].sort((a, b) => korSum(b[1], 7) - korSum(a[1], 7)).map(([e, rs]) => {
    const o = korStat(rs, K.med(c.k + '#e' + e)), pd = K.PL.get(c.k + '#' + e);
    o.real = c.cyc != null ? o.trips / K.nDays * c.cyc : null; o.plan = c.cyc != null && pd ? pd / (o.tpt || c.tpt) * c.cyc : null; o.rk = o.plan != null ? o.real : null;
    return Object.assign(o, {key: e, rs, kontrak: !!pd});
  });
}
const korDiff = o => o.plan != null && o.rk != null ? o.rk - o.plan : null;
function korCols(first, extra){
  const n1 = (h, f, fm) => ({h, num: true, val: r => f(r) == null ? -1e9 : f(r), html: r => { const v = f(r); return v == null ? '–' : fm ? fm(v) : fmt1(v); }});
  return [first].concat(extra || [], [n1('Trip', r => r.trips, fmt), n1('t / trip', r => r.tpt), n1('Siklus aktual (hari)', r => r.cyc, v => '<b>' + fmt1(v) + '</b>'),
    n1('Kebutuhan armada · plan kontrak', r => r.plan, fmt), n1('Kebutuhan armada · real', r => r.real, fmt),
    n1('Selisih (real − plan)', korDiff, v => Math.abs(v) < 0.5 ? '0' : '<b style="color:' + (v < 0 ? 'var(--red)' : 'var(--green)') + '">' + (v > 0 ? '+' : '−') + fmt(Math.abs(v)) + '</b>'),
    {h: 'Waktu muat (jam)', num: true, val: r => r.sla == null ? -1 : r.sla, html: r => r.sla == null ? '–' : '<b style="color:' + slaColor(r.sla) + '">' + fmt1(r.sla) + '</b>'},
    n1('Antri (jam)', r => r.wait)]);
}
function korTables(byC, corrRows, CG, nDays){
  const med = k => { const a = CG.sub.get(k); return a && a.length >= 3 ? median(a) : null; };
  KOR = {byC, CG, nDays, med, PL: kontrakPlanDay(PER().to.slice(0, 7)), corr: new Map()};
  const add = (a, b) => a == null ? b : b == null ? a : a + b;
  const plants = new Map();
  corrRows.forEach(c => {
    c.cyc = c.cn >= 10 ? c.cyc : null;
    const o = korStat(byC.get(c.k), c.cyc), es = korEksp(c);
    o.real = c.cyc != null ? c.tpd * c.cyc : null; es.forEach(x => { o.plan = add(o.plan, x.plan); o.rk = add(o.rk, x.rk); });
    const [si, pv, di] = c.k.split('|');
    Object.assign(o, {key: c.k, c, lbl: di != null ? tc(B.dims.dist[di]) : tc(B.dims.prov[pv]), sub: di != null ? 'Sumatera Barat' : ''});
    KOR.corr.set(c.k, o); (plants.get(+si) || plants.set(+si, []).get(+si)).push(o);
  });
  return [...plants.entries()].map(([si, list]) => ({si, list, ton: list.reduce((a, o) => a + o.ton, 0)})).sort((a, b) => b.ton - a.ton).map(P => {
    const T = P.list.reduce((a, o) => ({trips: a.trips + o.trips, plan: add(a.plan, o.plan), real: add(a.real, o.real), rk: add(a.rk, o.rk)}), {trips: 0, plan: null, real: null, rk: null});
    const cols = korCols({h: 'Tujuan', cls: 'name', val: r => r.lbl, html: r => esc(r.lbl) + (r.sub ? '<small>' + r.sub + '</small>' : '')});
    cols[0].total = '<b>Total</b>'; cols[1].total = fmt(T.trips); cols[4].total = T.plan == null ? '–' : fmt(T.plan); cols[5].total = T.real == null ? '–' : fmt(T.real);
    const d = korDiff(T); cols[6].total = d == null ? '–' : (d > 0 ? '+' : d < 0 ? '−' : '') + fmt(Math.abs(d));
    return '<div class="inc-grp" style="--c:var(--red)"><b>' + esc(tc(B.dims.src[P.si] || '?')) + '</b><small>' + fmt(T.trips) + ' trip · ' + P.list.length + ' tujuan</small></div>' +
      table(cols, P.list, {total: P.list.length > 1, sort: 1, go: r => 'kor:' + r.key});
  }).join('');
}
function korEkspModal(k){
  const c = KOR && KOR.corr.get(k); if(!c) return;
  const es = korEksp(c.c); KOR.cur = {k, es};
  const cols = korCols({h: 'Ekspeditur', cls: 'name', val: r => ekspLbl(r.key), html: r => esc(ekspLbl(r.key)) + (r.kontrak ? '<small>pemenang kontrak</small>' : '')}, [{h: 'Truk', num: true, val: r => r.trucks}]);
  openModal(tc(B.dims.src[k.split('|')[0]] || '') + ' → ' + c.lbl, 'Ekspeditur di koridor ini · ' + periodSub() + ' · klik ekspeditur untuk daftar truk',
    table(cols, es, {sort: 2, go: r => 'kort:' + es.indexOf(r)}));
}
function korTruckModal(i){
  const K = KOR, x = K && K.cur && K.cur.es[i]; if(!x) return;
  const c = K.corr.get(K.cur.k), byT = new Map(); x.rs.forEach(t => (byT.get(t[1]) || byT.set(t[1], []).get(t[1])).push(t));
  const list = [...byT.entries()].map(([tk, rs]) => Object.assign(korStat(rs, K.med(K.cur.k + '#t' + tk)), {key: tk, days: new Set(rs.map(t => t[0])).size, last: Math.max(...rs.map(t => t[0]))}));
  const cols = [{h: 'Nopol', cls: 'name', val: r => B.dims.truck[r.key]}, {h: 'Hari muat', num: true, val: r => r.days}, {h: 'Trip', num: true, val: r => r.trips}, {h: 'Tonase', num: true, val: r => r.ton, html: r => fmt(r.ton)},
    {h: 't / trip', num: true, val: r => r.tpt, html: r => fmt1(r.tpt)}, {h: 'Siklus truk (hari)', num: true, val: r => r.cyc == null ? -1 : r.cyc, html: r => r.cyc == null ? '–' : fmt1(r.cyc)},
    {h: 'Waktu muat (jam)', num: true, val: r => r.sla == null ? -1 : r.sla, html: r => r.sla == null ? '–' : '<b style="color:' + slaColor(r.sla) + '">' + fmt1(r.sla) + '</b>'},
    {h: 'Antri (jam)', num: true, val: r => r.wait == null ? -1 : r.wait, html: r => r.wait == null ? '–' : fmt1(r.wait)}, {h: 'Muat terakhir', val: r => r.last, html: r => shortDate(B.days[r.last])}];
  openModal(ekspLbl(x.key) + ' · ' + c.lbl, list.length + ' truk · ' + periodSub() + ' · klik truk untuk detail',
    '<div class="tab-note"><a class="section-link" data-href="kor:' + esc(K.cur.k) + '">&lsaquo; semua ekspeditur di koridor ini</a></div>' + table(cols, list, {sort: 2, go: r => '#/truk/' + enc(B.dims.truck[r.key])}));
}
/* Truck page: everything the data holds about one nopol (Rilis has no driver or owner columns). */
const PLATE = {BL: 'Aceh', BB: 'Sumatera Utara (Tapanuli)', BK: 'Sumatera Utara', BA: 'Sumatera Barat', BM: 'Riau', BP: 'Kepulauan Riau', BH: 'Jambi', BD: 'Bengkulu', BG: 'Sumatera Selatan', BN: 'Bangka Belitung', BE: 'Lampung', B: 'DKI Jakarta & sekitarnya', D: 'Bandung', F: 'Bogor', T: 'Karawang / Purwakarta', E: 'Cirebon', Z: 'Priangan Timur', A: 'Banten', H: 'Semarang', AA: 'Kedu', AB: 'Yogyakarta', AD: 'Surakarta', K: 'Pati', R: 'Banyumas', G: 'Pekalongan', L: 'Surabaya', W: 'Sidoarjo / Gresik', N: 'Malang', S: 'Bojonegoro', P: 'Besuki', AE: 'Madiun', AG: 'Kediri', M: 'Madura', DK: 'Bali', KB: 'Kalimantan Barat', DA: 'Kalimantan Selatan', KT: 'Kalimantan Timur', KH: 'Kalimantan Tengah', DD: 'Sulawesi Selatan', DB: 'Sulawesi Utara'};
const hhmm = m => { if(m == null) return '–'; const d = Math.floor(m / 1440), r = ((m % 1440) + 1440) % 1440; return String(Math.floor(r / 60)).padStart(2, '0') + ':' + String(r % 60).padStart(2, '0') + (d ? '<small>' + (d > 0 ? '+' : '') + d + ' hr</small>' : ''); };
function pageTruk(nopol){
  const ti = B.dims.truck.indexOf(nopol);
  if(ti < 0) return notFound('Truk ' + nopol);
  const m = PER(), all = B.trucks.filter(t => t[1] === ti), rows = all.filter(t => inPer(t[0], m));
  const sp = (B.spj || []).filter(r => r[1] === ti).sort((a, b) => b[0] - a[0] || (b[12] || 0) - (a[12] || 0));
  const st = korStat(rows, null), ek = new Map(); rows.forEach(t => ek.set(t[3], (ek.get(t[3]) || 0) + t[6]));
  const days = [...new Set(all.map(t => t[0]))].sort((a, b) => a - b), endI = B.days.indexOf(m.to) >= 0 ? B.days.indexOf(m.to) : B.days.length - 1;
  const d60 = days.filter(d => d > endI - 60 && d <= endI), gaps = d60.slice(1).map((d, i) => d - d60[i]).filter(g => g >= 1 && g <= 12);
  const pre = (nopol.match(/^[A-Z]{1,2}/) || [''])[0], reg = PLATE[pre];
  const tares = sp.map(r => r[14]).filter(v => v > 1000), tare = tares.length ? median(tares) : null;
  // routes: plant -> tujuan
  const rt = new Map(); rows.forEach(t => { const k = t[4] + '|' + (t[11] != null ? 'd' + t[11] : 'p' + t[5]); const o = rt.get(k) || {src: t[4], t, trips: 0, ton: 0, last: -1, eks: new Set()}; o.trips += t[6]; o.ton += t[7]; o.last = Math.max(o.last, t[0]); o.eks.add(t[3]); rt.set(k, o); });
  const tuj = t => t[11] != null ? tc(B.dims.dist[t[11]]) + '<small>' + esc(tc(B.dims.prov[t[5]])) + '</small>' : esc(tc(B.dims.prov[t[5]]));
  const rtTbl = table([{h: 'Plant', cls: 'name', val: r => tc(B.dims.src[r.src] || '')}, {h: 'Tujuan', val: r => r.t[11] != null ? B.dims.dist[r.t[11]] : B.dims.prov[r.t[5]], html: r => tuj(r.t)},
    {h: 'Ekspeditur', val: r => [...r.eks].map(ekspLbl).join(', '), html: r => esc([...r.eks].map(ekspLbl).join(', '))}, {h: 'Trip', num: true, val: r => r.trips}, {h: 'Tonase', num: true, val: r => r.ton, html: r => fmt(r.ton)},
    {h: 'Terakhir', val: r => r.last, html: r => shortDate(B.days[r.last])}], [...rt.values()], {sort: 3});
  // 30-day strip
  const s30 = Array.from({length: 30}, (_, i) => endI - 29 + i).filter(d => d >= 0), dset = new Set(days);
  const strip = '<div class="tk-strip">' + s30.map(d => '<span class="' + (dset.has(d) ? 'on' : '') + '" title="' + esc(dayLabel(B.days[d])) + '"><i>' + (+B.days[d].slice(8)) + '</i></span>').join('') + '</div>';
  const spCols = [{h: 'Tanggal', cls: 'name', val: r => r[0] * 1e5 + (r[12] || 0), html: r => shortDate(B.days[r[0]]) + '<small>SPJ ' + esc(r[2] || '–') + '</small>'},
    {h: 'Plant', val: r => tc(B.dims.src[r[3]] || ''), html: r => esc(tc(B.dims.src[r[3]] || '')) + '<small>' + esc(incLbl(INCS()[r[4]] || '')) + '</small>'},
    {h: 'Tujuan', val: r => B.dims.dist[r[7]] || '', html: r => esc(tc(B.dims.dist[r[7]] || '–')) + '<small>' + esc(tc(B.dims.prov[B.dims.distProv[r[7]]] || '')) + '</small>'},
    {h: 'Distributor · toko', val: r => distrLbl(r[6]), html: r => esc(distrLbl(r[6]) || '–') + (B.dims.toko[r[8]] ? '<small>' + esc(tc(B.dims.toko[r[8]][1] || '')) + '</small>' : '')},
    {h: 'Ekspeditur', val: r => ekspLbl(r[5]), html: r => esc(ekspLbl(r[5]))},
    {h: 'Ton', num: true, val: r => r[9], html: r => fmt1(r[9]) + (r[16] ? '<small>' + esc(tc(r[16])) + '</small>' : '')},
    {h: 'Antri', num: true, val: r => r[11] == null ? -1e9 : r[11], html: r => hhmm(r[11])}, {h: 'Masuk', num: true, val: r => r[12] == null ? -1e9 : r[12], html: r => hhmm(r[12])},
    {h: 'Keluar', num: true, val: r => r[13] == null ? -1e9 : r[13], html: r => hhmm(r[13])},
    {h: 'Waktu muat (jam)', num: true, val: r => r[12] != null && r[13] != null ? r[13] - r[12] : -1, html: r => { if(r[12] == null || r[13] == null || r[13] <= r[12]) return '–'; const h = (r[13] - r[12]) / 60; return '<b style="color:' + slaColor(h) + '">' + fmt1(h) + '</b>'; }},
    {h: 'Berat kosong (kg)', num: true, val: r => r[14] || 0, html: r => r[14] ? fmt(r[14]) : '–'}];
  page('<div class="wrap">' + crumb([['Armada', '#/armada'], ['Truk ' + nopol]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(nopol) + '</div><div class="sub">' + esc([...ek.keys()].map(ekspLbl).join(' · ') || 'tidak muat di periode ini') + (reg ? ' · plat ' + esc(pre) + ' = ' + esc(reg) : '') + ' · ' + periodSub() + '</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('trend', 'var(--black)', 'Trip', fmt(st.trips), {hint: fmt(st.ton) + ' t · ' + fmt1(st.tpt) + ' t / trip'}) +
      kpi('truck', 'var(--black)', 'Hari muat', fmt(new Set(rows.map(t => t[0])).size), {hint: gaps.length ? 'siklus median ' + fmt1(median(gaps)) + ' hari (60 hari)' : 'siklus: data belum cukup'}) +
      kpi('clock', st.sla == null ? 'var(--black)' : slaColor(st.sla), 'Waktu muat', st.sla == null ? '–' : fmt1(st.sla) + ' <small>jam</small>', {hint: st.wait == null ? '' : 'antri ' + fmt1(st.wait) + ' jam'}) +
      kpi('box', 'var(--black)', 'Berat kosong', tare ? fmt1(tare / 1000) + ' <small>t</small>' : '–', {hint: tare ? 'median timbangan · muatan ' + fmt1(st.tpt) + ' t/trip' : 'tersedia setelah build data berikutnya'}) +
    '</div>' +
    '<div class="section-head"><div class="section-title">Hari muat · 30 hari</div></div>' + strip +
    tabsHTML('truk', [
      {k: 'spj', label: 'Riwayat SPJ', n: sp.length, html: sp.length ? '<div class="tab-note">45 hari terakhir · satu baris per SPJ · jam relatif tanggal SPJ</div>' + table(spCols, sp, {sort: 0, limit: 300}) : '<div class="dp-review">Rincian per SPJ tersedia setelah build data berikutnya.</div>'},
      {k: 'rute', label: 'Rute', n: rt.size, html: rt.size ? rtTbl : '<div class="dp-review">Tidak ada trip di periode ini.</div>'}]) +
    '<div class="dp-review" style="margin-top:16px">Nama supir dan pemilik kendaraan tidak ada di data Rilis. Kolom yang tersedia: nopol, SPJ, plant, tujuan, distributor, toko, ekspeditur, tonase, jam antri/masuk/keluar dan timbangan.</div>' +
    '</div>', 'armada');
}
function pageArmada(){
  const m = PER(), rows = B.trucks.filter(t => inPer(t[0], m) && truckOk(t));
  const dayIdx = [...new Set(rows.map(t => t[0]))].sort((a, b) => a - b), nDays = Math.max(1, dayIdx.length);
  const sum = (rs, i) => rs.reduce((a, t) => a + t[i], 0);
  const trips = sum(rows, 6), ton = sum(rows, 7), dwS = sum(rows, 8), dwN = sum(rows, 9), wS = sum(rows, 12) || 0, wN = sum(rows, 13) || 0;
  const perDayTrucks = dayIdx.map(d => new Set(rows.filter(t => t[0] === d).map(t => t[1])).size);
  const truckDays = new Map(); rows.forEach(t => { const k = t[0] + ':' + t[1]; const o = truckDays.get(k) || {s: 0, n: 0}; o.s += t[8]; o.n += t[9]; truckDays.set(k, o); });
  const over4 = rs => { const td = new Map(); rs.forEach(t => { if(!t[9]) return; const k = t[0] + ':' + t[1]; const o = td.get(k) || {s: 0, n: 0}; o.s += t[8]; o.n += t[9]; td.set(k, o); }); const v = [...td.values()]; return v.length ? pctOf(v.filter(o => o.s / o.n > SLA_TARGET).length, v.length) : null; };
  const pctCol = (h, f) => ({h, num:true, val: f, html: r => { const v = f(r); return v == null ? '–' : '<span style="color:' + (v > 25 ? 'var(--red)' : v > 10 ? 'var(--amber)' : 'var(--green)') + '">' + pctTxt(v) + '</span>'; }});
  const jam = (h, f, colored) => ({h, num:true, val: f, html: r => { const v = f(r); return v == null ? '–' : colored ? '<b style="color:' + slaColor(v) + '">' + fmt1(v) + '</b>' : fmt1(v); }});

  // 1. Kinerja muat di plant
  const bySrc = new Map(); rows.forEach(t => (bySrc.get(t[4]) || bySrc.set(t[4], []).get(t[4])).push(t));
  const plantRows = [...bySrc.entries()].map(([si, rs]) => { const ds = [...new Set(rs.map(t => t[0]))];
    return {key: si, n: tc(B.dims.src[si] || '(tanpa plant)'), tpd: median(ds.map(d => new Set(rs.filter(t => t[0] === d).map(t => t[1])).size)), trips: sum(rs, 6), ton: sum(rs, 7), tpt: sum(rs, 7) / Math.max(1, sum(rs, 6)),
      sla: sum(rs, 9) ? sum(rs, 8) / sum(rs, 9) : null, o4: over4(rs), wait: sum(rs, 13) ? sum(rs, 12) / sum(rs, 13) : null}; }).sort((a, b) => b.ton - a.ton);
  const plantTbl = plantRows.length ? table([{h:'Plant', cls:'name', val: r => r.n}, {h:'Truk / hari', num:true, val: r => r.tpd, html: r => fmt(r.tpd)}, {h:'Trip', num:true, val: r => r.trips, html: r => fmt(r.trips)},
    {h:'t / trip', num:true, val: r => r.tpt, html: r => fmt1(r.tpt)}, jam('Antri (jam)', r => r.wait), jam('Waktu muat (jam)', r => r.sla, true), pctCol('> ' + SLA_TARGET + ' jam', r => r.o4)], plantRows, {sort: 2, go: r => B.dims.src[r.key] ? plantHref(B.dims.src[r.key]) : null}) : '';
  // daily: time in plant per day (all plants in the filter), click = per ekspeditur that day
  const dData = dayIdx.map(d => { const rs = rows.filter(t => t[0] === d), n = sum(rs, 9), v = n ? sum(rs, 8) / n : 0;
    return {label: dayIdx.length > 20 ? String(+B.days[d].slice(8)) : shortDate(B.days[d]).replace(/ \w+$/, ''), long: dayLabel(B.days[d]), vol: v, unit: 'jam', color: slaColor(v), tgt: SLA_TARGET, pct: null, d}; });
  TRENDS.ad = {id: 'ad', data: dData, click: i => {
    const d = dData[i].d, per = new Map();
    rows.forEach(t => { if(t[0] !== d) return; const o = per.get(t[3]) || {key: t[3], s: 0, n: 0, trips: 0, vol: 0}; o.s += t[8]; o.n += t[9]; o.trips += t[6]; o.vol += t[7]; per.set(t[3], o); });
    const rs = [...per.values()].filter(o => o.n).map(o => Object.assign(o, {dw: o.s / o.n}));
    openModal('Waktu muat · ' + dayLabel(B.days[d]), 'per ekspeditur', rs.length ? table([{h:'Ekspeditur', cls:'name', val: o => ekspLbl(o.key)}, {h:'Rata-rata (jam)', num:true, cls:'pc', val: o => o.dw, html: o => fmt1(o.dw), style: o => 'color:' + slaColor(o.dw)}, {h:'Trip', num:true, val: o => o.trips}, {h:'Tonase', num:true, val: o => o.vol, html: o => fmt(o.vol)}], rs, {go: o => ekspHref(o.key), sort: 1}) : '<div class="dp-review">Tidak ada data.</div>');
  }};

  // 2. Koridor
  const CG = cycleGaps(m);
  const byC = new Map(); rows.forEach(t => { const k = corrKey(t); (byC.get(k) || byC.set(k, []).get(k)).push(t); });
  const corrRows = [...byC.entries()].map(([k, rs]) => { const tr = sum(rs, 6), cyc = CG.med.get(k), tpd = tr / nDays;
    return {k, n: corrLbl(k), trips: tr, ton: sum(rs, 7), tpt: sum(rs, 7) / Math.max(1, tr), cyc, cn: CG.n.get(k) || 0, tpd, need: cyc != null ? tpd * cyc : null, used: new Set(rs.map(t => t[1])).size,
      sla: sum(rs, 9) ? sum(rs, 8) / sum(rs, 9) : null}; }).filter(r => r.trips >= 5).sort((a, b) => b.ton - a.ton);
  const corrTbl = corrRows.length ? korTables(byC, corrRows, CG, nDays) : '';

  // 3. Scorecard ekspeditur (per incoterm when several are active)
  const scoreTbl = () => {
    const rs = rows.filter(t => incOk(t[2])), be = new Map();
    rs.forEach(t => { if(t[3] === D.eksEmpty) return; (be.get(t[3]) || be.set(t[3], []).get(t[3])).push(t); });
    if(!be.size) return '';
    const S = summarize('eksp', m);
    const list = [...be.entries()].map(([e, ts]) => {
      const trk = new Map(); ts.forEach(t => trk.set(t[1], (trk.get(t[1]) || 0) + t[6]));
      const ds = [...new Set(ts.map(t => t[0]))];
      const gaps = (CG.ek.get(e) || []).filter(([k]) => CG.avg.get(k) && (CG.n.get(k) || 0) >= 10).map(([k, g]) => [g, CG.avg.get(k)]);
      const o = S.get(e) || {};
      return {key: e, trucks: trk.size, tpd: median(ds.map(d => new Set(ts.filter(t => t[0] === d).map(t => t[1])).size)), trips: sum(ts, 6), ton: sum(ts, 7),
        cyc: gaps.length >= 5 ? gaps.reduce((a, g) => a + Math.min(g[0], 3 * g[1]), 0) / gaps.reduce((a, g) => a + g[1], 0) * 100 : null,   // mean gap vs the mean gap of the same corridors (capped at 3x)
        one: pctOf([...trk.values()].filter(v => v === 1).length, trk.size),
        sla: sum(ts, 9) ? sum(ts, 8) / sum(ts, 9) : null, o4: over4(ts), wait: sum(ts, 13) ? sum(ts, 12) / sum(ts, 13) : null, pct: o.pct, hasTgt: o.hasTgt, tgt: o.tgt};
    }).sort((a, b) => b.ton - a.ton);
    return table([{h:'Ekspeditur', cls:'name', val: r => ekspLbl(r.key)}, {h:'Truk', num:true, val: r => r.trucks}, {h:'Truk / hari', num:true, val: r => r.tpd, html: r => fmt(r.tpd)},
      {h:'Trip', num:true, val: r => r.trips, html: r => fmt(r.trips)}, {h:'Ton', num:true, val: r => r.ton, html: r => fmt(r.ton)},
      {h:'Siklus vs koridor', num:true, val: r => r.cyc, html: r => r.cyc == null ? '–' : '<b style="color:' + (r.cyc <= 100 ? 'var(--green)' : r.cyc <= 120 ? 'var(--amber)' : 'var(--red)') + '">' + pctTxt(r.cyc) + '</b>'},
      {h:'Truk 1 trip', num:true, val: r => r.one, html: r => pctTxt(r.one)}, jam('Antri (jam)', r => r.wait), jam('Waktu muat (jam)', r => r.sla, true), pctCol('> ' + SLA_TARGET + ' jam', r => r.o4),
      C.pct], list, {go: r => ekspHref(r.key), sort: 4});
  };

  const tpdAll = median(perDayTrucks);
  page('<div class="wrap">' + crumb([['Armada']]) +
    '<div class="dp-head"><div><div class="nm">Armada' + (F.moda === 'LAUT' ? ' · truk ke pelabuhan' : '') + '</div><div class="sub">Kinerja rilis, koridor &amp; ekspeditur · ' + incName() + ' · ' + periodSub() + '</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('truck', 'var(--black)', 'Truk dimuat / hari', fmt(tpdAll), {hint: 'median · ' + fmt(new Set(rows.map(t => t[1])).size) + ' truk berbeda'}) +
      kpi('trend', 'var(--black)', 'Trip', fmt(trips), {hint: fmt1(ton / Math.max(1, trips)) + ' t / trip'}) +
      kpi('clock', 'var(--amber)', 'Waktu muat', dwN ? fmt1(dwS / dwN) + ' <small>jam</small>' : '–', {hint: 'masuk → keluar · ' + pctTxt(over4(rows)) + ' truk > ' + SLA_TARGET + ' jam'}) +
      kpi('clock', 'var(--black)', 'Antri', wN ? fmt1(wS / wN) + ' <small>jam</small>' : '–', {hint: 'antri → masuk pabrik'}) +
    '</div>' + tabsHTML('armada', [{k: 'rilis', label: 'Kinerja rilis', html:
    '<div class="dp-panel" style="margin-bottom:14px"><div class="pt">Waktu muat per hari (jam)</div>' + barsHTML('ad', dData) +
      '<div class="trend-legend"><span><i style="background:var(--green)"></i>&le;' + SLA_TARGET + ' jam</span><span><i style="background:var(--red)"></i>&gt;' + SLA_TARGET + ' jam</span><span><i class="dash"></i>target ' + SLA_TARGET + ' jam</span><span>klik batang = per ekspeditur</span></div></div>' +
    (plantTbl || '<div class="dp-review">Tidak ada data truk untuk filter ini.</div>')},
    {k: 'koridor', label: 'Koridor', html: '<div class="tab-note">Kebutuhan armada = trip per hari × siklus aktual (median jeda antar muat truk yang sama, 60 hari) · <b>plan kontrak</b> = volume kontrak per hari ÷ t/trip aktual (ekspeditur pemenang rute) · <b>real</b> = trip aktual · selisih atas ekspeditur kontrak, merah = kurang dari kebutuhan kontrak · klik tujuan → ekspeditur → truk</div>' +
    (corrTbl || '<div class="dp-review">Tidak ada data.</div>')},
    {k: 'eksp', label: 'Ekspeditur', html: '<div class="tab-note">Siklus vs koridor: jeda rata-rata antar muat truk ekspeditur dibanding rata-rata koridor yang sama; 100% = sama, &lt;100% lebih cepat</div>' +
    (incGroups(scoreTbl) || '<div class="dp-review">Tidak ada data.</div>')}]) +
    '</div>', 'armada');
}
function qTable(head, rows){
  return '<div class="tbl-wrap"><table class="tbl"><thead><tr>' + head.map((h, i) => '<th' + (i ? ' class="num"' : '') + '>' + h + '</th>').join('') + '</tr></thead><tbody>' +
    (rows.length ? rows.map(r => '<tr>' + r.map((c, i) => '<td' + (i ? ' class="num"' : '') + '>' + c + '</td>').join('') + '</tr>').join('') : '<tr><td colspan="' + head.length + '">Tidak ada.</td></tr>') + '</tbody></table></div>';
}
/** Data districts per province that do not match a kabupaten/kota on the map (needs assets/districts.js). */
function unmatchedHTML(){
  const rows = [];
  B.dims.dist.forEach((name, di) => {
    const pn = B.dims.prov[B.dims.distProv[di]], G = window.DISTRICTS[mapProv(pn)];
    if(name && G && matchDistrict(name, G.d) < 0) rows.push([esc(tc(pn)), esc(name)]);
  });
  return '<div id="qDist">' + qTable(['Provinsi', 'Nama distrik di data'], rows) + '</div>';
}
function pageQuality(){
  const Q = B.quality || {checks: [], toko: []};
  // days per month with FRC / FOT realisasi
  const seen = {};
  B.facts.forEach(f => { if(f[1] > 1) return; const d = B.days[f[0]], m = d.slice(0, 7); (seen[m] = seen[m] || [new Set(), new Set()])[f[1]].add(+d.slice(8)); });
  const months = Object.keys(seen).sort().reverse().slice(0, 6);
  const gaps = months.map(m => {
    const last = m === B.asOf.slice(0, 7) ? +B.asOf.slice(8) : dim(m);
    const miss = k => { const out = []; for(let d = 1; d <= last; d++) if(!seen[m][k].has(d)) out.push(d); return out; };
    const fmtDays = a => a.length ? (a.length > 12 ? a.length + ' hari' : a.join(', ')) : '<span style="color:var(--green)">lengkap</span>';
    return [esc(monthLabel(m)) + (last < dim(m) ? ' <small>(s.d. ' + last + ')</small>' : ''), fmtDays(miss(0)), fmtDays(miss(1))];
  });
  const toko = {};
  Q.toko.forEach(([m, inc, blank, n]) => { (toko[m] = toko[m] || {})[inc] = [blank, n]; });
  const tokoCell = x => !x ? '–' : x[0] ? '<b style="color:var(--red)">' + fmt(x[0]) + '</b> / ' + fmt(x[1]) + ' (' + pctTxt(x[0] / x[1] * 100) + ')' : '<span style="color:var(--green)">0</span> / ' + fmt(x[1]);
  const tokoRows = Object.keys(toko).sort().reverse().slice(0, 6).map(m => [esc(monthLabel(m)), tokoCell(toko[m].FRC), tokoCell(toko[m].FOT)]);
  let dist;
  if(window.DISTRICTS) dist = unmatchedHTML();
  else { distMapPending = unmatchedHTML; loadDistricts(); dist = '<div id="dmapSlot"><div class="empty">Memuat peta distrik…</div></div>'; }
  page('<div class="wrap">' + crumb([['Kualitas data']]) +
    '<div class="dp-head"><div><div class="nm">Kualitas data</div><div class="sub">Cek master Excel · build ' + esc(B.build || '') + ' · data s.d. ' + dayLabel(B.asOf) + '</div></div></div>' +
    '<div class="dp-review">Halaman ini untuk pengelola master Excel: apa yang kosong, ganda atau tidak terbaca pada build terakhir. Angka dashboard hanya sebaik data di sini.</div>' +
    '<div class="section-head"><div class="section-title">Tanggal tanpa realisasi</div></div>' +
    qTable(['Bulan', 'FRC: tanggal kosong', 'FOT: tanggal kosong'], gaps) +
    '<div class="section-head"><div class="section-title">Baris tanpa toko tujuan (ship-to)</div></div>' +
    qTable(['Bulan', 'FRC: kosong / baris', 'FOT: kosong / baris'], tokoRows) +
    '<div class="section-head"><div class="section-title">Hasil cek build</div></div>' +
    qTable(['Cek', 'Jumlah'], Q.checks.map(([k, n]) => [esc(k), fmt(n)])) +
    '<div class="section-head"><div class="section-title">Distrik yang tidak cocok dengan peta</div></div>' + dist +
    '</div>', '');
}

function notFound(what){ page('<div class="wrap">' + crumb([['Tidak ditemukan']]) + '<div class="dp-head"><div><div class="nm">Tidak ditemukan</div><div class="sub">' + esc(what) + ' tidak ada di data.</div></div></div><a class="section-link" href="#/">Kembali ke Home &rsaquo;</a></div>', ''); }

/* ===================== Router ===================== */
const PROG_SLUG = {harian: 'hari-ini', mingguan: 'minggu-ini', bulanan: 'akhir-bulan'};
const PROG_GRAN = {'hari-ini': 'harian', 'minggu-ini': 'mingguan', 'akhir-bulan': 'bulanan'};
function route(){
  if(!B) return;
  closeNav(); $('#searchResults').classList.remove('show');
  closeModal();
  for(const k in TRENDS) delete TRENDS[k];
  for(const k in TABLES) delete TABLES[k];
  const parts = (location.hash.replace(/^#\/?/, '') || '').split('/').map(decodeURIComponent);
  const [a, b] = parts;
  try {
    if(!a) pageHome();
    else if(a === 'provinsi') b ? pageProv(b) : pageProvList();
    else if(a === 'plant') b ? pagePlant(b) : pagePlantList();
    else if(a === 'distrik' && b) pageDist(b);
    else if(a === 'ekspeditur') b ? pageEksp(b) : pageEkspList();
    else if(a === 'distributor') b ? pageDistr(b) : pageDistrList();
    else if(a === 'fot') location.replace('#/distributor');
    else if(a === 'armada') pageArmada();
    else if(a === 'kontrak') pageKontrak();
    else if(a === 'kapal') b ? pageKapalDetail(b) : pageKapal();
    else if(a === 'truk' && b) pageTruk(b);
    else if(a === 'kualitas') pageQuality();
    else if(a === 'kredit') pageCredits();
    else if(a === 'prognosa') pageForecast(PROG_GRAN[b] || 'harian');
    // old addresses -> their single new home
    else if(a === 'forecast') location.replace('#/prognosa/' + PROG_SLUG[b in PROG_SLUG ? b : 'harian']);
    else if(a === 'incoterm') location.replace(b && b !== 'FRC' ? '#/distributor' : '#/ekspeditur');
    else if(a === 'scorecard') location.replace('#/ekspeditur');
    else if(a === 'tren') location.replace('#/');
    else notFound('Halaman ' + a);
  } catch(err) {
    console.error(err);
    view().innerHTML = '<div class="wrap"><div class="dp-review" style="border-left:4px solid var(--red)">Terjadi kesalahan saat menampilkan halaman ini: ' + esc(err.message) + '. <a href="#/" style="color:var(--red);font-weight:700">Kembali ke Home</a></div></div>';
  }
}

/* ===================== Events (delegated) ===================== */
function bindEvents(){
  document.addEventListener('click', e => {
    const qk = e.target.closest('[data-qk]');
    if(qk){ e.stopPropagation(); const was = $('#qkPop') && $('#qkPop').dataset.k === qk.dataset.qk; closeQuick(); closeNav(); closeFilter(); if(!was){ quickMenu(qk.dataset.qk, qk); $('#qkPop').dataset.k = qk.dataset.qk; } return; }
    if(!e.target.closest('#qkPop')) closeQuick();
    const ni = e.target.closest('[data-navinc],[data-navkem],[data-navmoda],[data-navarea]');
    if(ni){
      e.preventDefault();
      if(ni.dataset.navmoda != null){ F.moda = ni.dataset.navmoda; F.inc = F.moda === 'DARAT' ? 'FRC' : 'ALL'; modaPlants(F.moda); }
      else if(ni.dataset.navinc != null) setInc(ni.dataset.navinc); else if(ni.dataset.navarea != null) F.area = ni.dataset.navarea; else F.kem = ni.dataset.navkem;
      saveFilter();
      closeQuick(); closeNav(); closeFilter(); buildNav(); filterStrip(); route(); return;
    }
    if(e.target.closest('#navBurger')){ const hd = $('header'), o = hd.classList.toggle('nav-open'); $('#navBurger').setAttribute('aria-expanded', o ? 'true' : 'false'); if(!o) closeNav(); return; }
    if(window.innerWidth <= 1180 && e.target.closest('#mainNav a[href]')) $('header').classList.remove('nav-open');
    const tgl = e.target.closest('[data-toggle]');
    if(tgl){
      const item = tgl.parentElement, was = item.classList.contains('open');
      closeNav();
      if(!was){
        closeFilter(); item.classList.add('open'); syncNavAria();
        const panel = item.querySelector('.dropdown,.mega,.filter-panel');   // on narrow screens panels are position:fixed under the header
        if(panel) panel.style.top = window.innerWidth <= 1180 ? ($('header').getBoundingClientRect().bottom + 4) + 'px' : '';
      }
      return;
    }
    const gr = e.target.closest('[data-grp]');
    if(gr){ e.preventDefault(); $$('details.trk-grp').forEach(d => { d.open = gr.dataset.grp === '1'; }); return; }
    const act = e.target.closest('[data-act]');
    if(act){
      const a = act.dataset.act;
      if(a === 'apply') applyFilter();
      if(a === 'reset'){ draft = DEFAULT_VIEW(); applyFilter(); }
      if(a === 'laporan'){ e.stopPropagation(); draft = isReportView() ? DEFAULT_VIEW() : REPORT_VIEW(); applyFilter(); }
      if(a === 'openfilter'){ e.stopPropagation(); openFilter(); }
      if(a === 'logout'){ logout(); }
      if(a === 'closemodal'){ closeModal(); }
      return;
    }
    const tg = e.target.closest('[data-tgt]');
    if(tg){ draft.tgt = tg.dataset.tgt; $$('[data-tgt]').forEach(x => x.classList.toggle('on', x === tg)); return; }
    const arEl = e.target.closest('[data-area]');
    if(arEl){ draft.area = arEl.dataset.area; $$('[data-area]').forEach(x => x.classList.toggle('on', x === arEl)); return; }
    const kemEl = e.target.closest('[data-kem]');
    if(kemEl){ draft.kem = kemEl.dataset.kem; $$('[data-kem]').forEach(x => x.classList.toggle('on', x === kemEl)); return; }
    const inc = e.target.closest('[data-inc]');
    const mo = e.target.closest('[data-moda]');
    if(mo){ draft.moda = mo.dataset.moda; draft.inc = draft.moda === 'DARAT' ? 'FRC' : 'ALL'; draft.srcOff = new Set(draft.moda === 'DARAT' ? DEFAULT_SRC_OFF : []); renderFilterPanel(); return; }
    if(inc){ draft.inc = inc.dataset.inc; $$('[data-inc]').forEach(x => x.classList.toggle('on', x === inc)); return; }
    const qm = e.target.closest('[data-qm]');
    if(qm){ const last = D.months[D.months.length - 1]; draft.per.month = qm.dataset.qm === 'cur' ? last : (D.months.includes(prevMonth(last)) ? prevMonth(last) : last); $('#fpMonth').value = draft.per.month; return; }
    const pm = e.target.closest('[data-pmode]');
    if(pm){ draft.per = {kind: pm.dataset.pmode, month: draft.per.month || draft.month}; renderFilterPanel(); return; }
    const qr = e.target.closest('[data-qr]');
    if(qr){ draft.to = B.asOf; draft.from = addDays(B.asOf, -(+qr.dataset.qr) + 1); $('#fpFrom').value = draft.from; $('#fpTo').value = draft.to; return; }
    const sa = e.target.closest('[data-srcall]');
    if(sa){ const on = sa.dataset.srcall === '1'; draft.srcOff = on ? new Set() : new Set(D.sources); $$('#filterPanel [data-src]').forEach(c => { c.checked = on; }); return; }
    const src = e.target.closest('[data-src]');
    if(src){ src.checked ? draft.srcOff.delete(src.dataset.src) : draft.srcOff.add(src.dataset.src); return; }
    const gran = e.target.closest('[data-gran]');
    if(gran){ const [id, g] = gran.dataset.gran.split(':'); TRENDS[id].gran = g; $('#' + id).innerHTML = trendInner(id); return; }
    const mdim = e.target.closest('[data-mdim]');
    if(mdim && MODAL){ MODAL.dim = mdim.dataset.mdim; $('#modalBody').innerHTML = modalBody(); return; }
    if(e.target.id === 'modal'){ closeModal(); return; }
    const bar = e.target.closest('[data-bar]');
    if(bar){
      const [id, i] = bar.dataset.bar.split(':'); const T = TRENDS[id];
      $$('[data-bar^="' + id + ':"]').forEach(x => x.classList.toggle('sel', x === bar));
      if(T.click) T.click(+i); else if(T.data[+i].from) openBreakdown(T, +i);
      return;
    }
    const tall = e.target.closest('[data-tree-all]');
    if(tall){ const [id, o] = tall.dataset.treeAll.split(':'); treeAll(id, o === '1'); return; }
    const trow = e.target.closest('table.tree tr.has-kids');
    if(trow && !e.target.closest('.t-go')){ treeToggle(trow); return; }
    const sort = e.target.closest('[data-sort]');
    if(sort){ const [id, i] = sort.dataset.sort.split(':'); const T = TABLES[id]; if(T.sort === +i) T.dir = -T.dir; else { T.sort = +i; T.dir = T.cols[+i].num ? -1 : 1; } $('#' + id).innerHTML = tableInner(id); return; }
    if(e.target.closest('[data-plantinc]')){ e.preventDefault(); plantAllInc = !plantAllInc; route(); return; }
    const pd = e.target.closest('[data-plantday]');
    if(pd){ e.preventDefault(); plantHourDay = +pd.dataset.plantday; const sec = $('#plantHours'); if(sec) sec.outerHTML = plantHourSection(); return; }
    const go = e.target.closest('[data-href]');
    const km = e.target.closest('[data-kmode]');
    if(km){ KMODE = km.dataset.kmode; try { sessionStorage.setItem('kmode', KMODE); } catch(x){} const y = scrollY; route(); scrollTo(0, y); return; }
    const tb = e.target.closest('[data-tab]');
    if(tb){ tabClick(tb); return; }
    if(go && /^kontrak:/.test(go.dataset.href || '')){ kontrakDetail(+go.dataset.href.slice(8)); return; }
    if(go && /^kor:/.test(go.dataset.href || '')){ korEkspModal(go.dataset.href.slice(4)); return; }
    if(go && /^kort:/.test(go.dataset.href || '')){ korTruckModal(+go.dataset.href.slice(5)); return; }
    if(go && go.dataset.href && go.dataset.href !== 'null'){ location.hash = go.dataset.href; return; }
    if(!e.target.closest('#mainNav')) closeNav();
    if(!e.target.closest('#filterStrip')) closeFilter();
    if(!e.target.closest('.search')) $('#searchResults').classList.remove('show');
  });
  document.addEventListener('change', e => {
    if(e.target.dataset && e.target.dataset.cmp){ const T = TRENDS[e.target.dataset.cmp]; T.cmp = e.target.value || null; $('#' + T.id).innerHTML = trendInner(T.id); return; }
    if(e.target.id === 'fpMonth') draft.per.month = e.target.value;
    if(e.target.id === 'fpFrom') draft.from = e.target.value;
    if(e.target.id === 'fpTo') draft.to = e.target.value; });
  $$('#mainNav [data-toggle]').forEach(b => { b.setAttribute('role', 'button'); b.setAttribute('tabindex', '0'); b.setAttribute('aria-haspopup', 'true'); b.setAttribute('aria-expanded', 'false'); });
  document.addEventListener('keydown', e => {
    const t = e.target;
    if(t && t.matches && t.matches('#mainNav [data-toggle]') && (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown')){ e.preventDefault(); if(!(e.key === 'ArrowDown' && t.parentElement.classList.contains('open'))) t.click(); const f = t.parentElement.classList.contains('open') && t.parentElement.querySelector('a'); if(f && e.key === 'ArrowDown') f.focus(); }
    if(t && t.closest && t.closest('#mainNav .dropdown,#mainNav .mega') && (e.key === 'ArrowDown' || e.key === 'ArrowUp')){ const l = [...t.closest('.navitem').querySelectorAll('.dropdown a,.mega a')], i = l.indexOf(t); if(i >= 0){ e.preventDefault(); l[(i + (e.key === 'ArrowDown' ? 1 : -1) + l.length) % l.length].focus(); } }
    if(e.key === '/' && !/^(INPUT|TEXTAREA|SELECT)$/.test((t && t.tagName) || '')){ const sbx = $('#searchBox'); if(sbx && sbx.offsetParent){ e.preventDefault(); sbx.focus(); } }
    if(e.key === 'Escape'){ const hd = $('header'); hd.classList.remove('nav-open'); $('#navBurger').setAttribute('aria-expanded', 'false');  closeNav(); closeFilter(); closeModal(); $('#searchResults').classList.remove('show'); } });
  /* desktop: open menus on hover with small intent delay */
  if(window.matchMedia('(hover:hover)').matches){ let ht; $$('#mainNav .navitem').forEach(n => { if(!n.querySelector('[data-toggle]')) return; n.addEventListener('mouseenter', () => { if(window.innerWidth <= 1180) return; clearTimeout(ht); ht = setTimeout(() => { const b = n.querySelector('[data-toggle]'); if(!n.classList.contains('open')) b.click(); }, 120); }); n.addEventListener('mouseleave', () => { clearTimeout(ht); if(window.innerWidth > 1180) ht = setTimeout(() => { n.classList.remove('open'); syncNavAria(); }, 200); }); }); }
  const sb = $('#searchBox');
  sb.addEventListener('input', () => doSearch(sb.value));
  sb.addEventListener('focus', () => doSearch(sb.value));
  sb.addEventListener('keydown', e => {
    const links = $$('#searchResults a'); if(!links.length) return;
    let i = links.findIndex(a => a.classList.contains('sel'));
    if(e.key === 'ArrowDown' || e.key === 'ArrowUp'){ e.preventDefault(); i = (i + (e.key === 'ArrowDown' ? 1 : -1) + links.length) % links.length; links.forEach((a, j) => a.classList.toggle('sel', j === i)); }
    if(e.key === 'Enter'){ e.preventDefault(); location.hash = links[Math.max(0, i)].getAttribute('href'); sb.value = ''; sb.blur(); $('#searchResults').classList.remove('show'); }
  });
  // ranking list <-> map highlight
  document.addEventListener('mouseover', e => {
    const r = e.target.closest && e.target.closest('.mh-row');
    $$('.map-hero .map-region.hl').forEach(x => x.classList.remove('hl'));
    const want = r && (r.dataset.hl || (r.dataset.prov ? '#/provinsi/' + enc(r.dataset.prov) : null));
    if(want) $$('.map-hero .map-region.live').forEach(x => { if(x.dataset.href === want) x.classList.add('hl'); });
    const kd = r && r.dataset.kd; $$('.map-hero .toko-pin').forEach(x => x.classList.toggle('dim', kd != null && x.dataset.k !== kd));
  });
  // map tooltip
  const tip = document.createElement('div'); tip.className = 'map-tip'; document.body.appendChild(tip);
  document.addEventListener('mousemove', e => {
    const t = e.target.closest && e.target.closest('[data-tip]');
    if(!t){ tip.style.display = 'none'; return; }
    const [a, ...rest] = t.dataset.tip.split('|');
    tip.innerHTML = '<b>' + esc(a) + '</b>' + rest.map(x => '<br>' + esc(x)).join('');
    tip.style.display = 'block'; tip.style.left = (e.clientX + 14) + 'px'; tip.style.top = (e.clientY + 14) + 'px';
  });
  window.addEventListener('hashchange', () => { route(); trackView(); });
  watchNewVersion();
}

/* ===================== Crypto, gate, session ===================== */
const IDLE_MS = 15 * 60 * 1000;
const REMEMBER_MS = 30 * 24 * 3600 * 1000;   // opt-in 'tetap masuk' (checkbox, off by default)
function rememberKey(){ try { const k = localStorage.getItem('osp_rk'), e = +(localStorage.getItem('osp_re') || 0); if(k && Date.now() < e) return k; } catch(_) {} return null; }
function forgetRemember(){ try { localStorage.removeItem('osp_rk'); localStorage.removeItem('osp_re'); } catch(_) {} }
const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const toB64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
let META = null;
async function getMeta(){
  if(META) return META;
  const r = await fetch('data/meta.json?_=' + Date.now(), {cache: 'no-store'});
  if(!r.ok) throw new Error('meta.json tidak ditemukan');
  return META = await r.json();
}
async function deriveKey(pw, meta){
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({name: 'PBKDF2', salt: b64(meta.salt), iterations: meta.iterations, hash: 'SHA-256'}, base, {name: 'AES-GCM', length: 256}, true, ['decrypt']);
}
/** v2 (public-key lock): password key -> RSA private key (meta.kpriv) -> this build's data key (meta.dk).
    v1: the password key decrypts the data directly. */
async function dataKey(key, meta){
  if(meta.v !== 2) return key;
  const der = await crypto.subtle.decrypt({name: 'AES-GCM', iv: b64(meta.kiv)}, key, b64(meta.kpriv));   // throws on wrong password
  const priv = await crypto.subtle.importKey('pkcs8', der, {name: 'RSA-OAEP', hash: 'SHA-256'}, false, ['decrypt']);
  const raw = await crypto.subtle.decrypt({name: 'RSA-OAEP'}, priv, b64(meta.dk));
  return crypto.subtle.importKey('raw', raw, {name: 'AES-GCM'}, false, ['decrypt']);
}
async function loadBundle(key, meta){
  const dk = await dataKey(key, meta);
  const r = await fetch('data/site.enc?b=' + enc(meta.build), {cache: 'no-cache'});
  if(!r.ok) throw new Error('data/site.enc tidak ditemukan');
  const ct = await r.arrayBuffer();
  const gz = await crypto.subtle.decrypt({name: 'AES-GCM', iv: b64(meta.iv)}, dk, ct);   // throws on wrong key
  const stream = new Blob([gz]).stream().pipeThrough(new DecompressionStream('gzip'));
  return JSON.parse(await new Response(stream).text());
}
/** The site shows plan vs realisasi up to H-1: drop days from today on (a build can carry a partial day H),
    so Target MTD never counts a full day against half a day's realisasi. B.days is sorted. */
function trimToYesterday(b){
  const now = new Date(), today = now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate());
  const n = b.days.findIndex(d => d >= today);
  if(n <= 0) return;   // nothing from today, or nothing before today (keep the data rather than show an empty site)
  b.days = b.days.slice(0, n);
  ['facts', 'ship', 'trucks'].forEach(k => { b[k] = b[k].filter(r => r[0] < n); });
  b.asOf = b.days[n - 1];
  b.lastDay = {};
  b.days.forEach(d => { const m = d.slice(0, 7); b.lastDay[m] = Math.max(b.lastDay[m] || 0, +d.slice(8)); });
}
function touch(){ try { sessionStorage.setItem('osp_t', String(Date.now())); } catch(e) {} }
function logout(){ forgetRemember(); try { sessionStorage.removeItem('osp_k'); sessionStorage.removeItem('osp_t'); } catch(e) {} location.reload(); }
async function start(key, meta, remember){
  B = await loadBundle(key, meta);
  B.build = meta.build;
  trimToYesterday(B);
  B.sod = B.sod || [];   // bundles built before SOCC-by-delivery-date
  B.sodMonths = new Set(B.sod.map(s => s[0].slice(0, 7)));
  prepare(); restoreFilter();
  try { sessionStorage.setItem('osp_k', toB64(await crypto.subtle.exportKey('raw', key))); } catch(e) {}
  if(remember){ try { localStorage.setItem('osp_rk', toB64(await crypto.subtle.exportKey('raw', key))); localStorage.setItem('osp_re', String(Date.now() + REMEMBER_MS)); } catch(e) {} }
  touch();
  $('#gate').hidden = true; $('#app').hidden = false;
  buildNav(); buildSearch(); filterStrip();
  $('#foot').innerHTML = '<span>PT Semen Padang · Outbound Logistics · ' + esc(meta.build) + ' · data s.d. ' + dayLabel(B.asOf) + '</span><span><a href="#/kualitas">Kualitas data</a> · <a href="#/kredit">Kredit foto</a> · ' + (rememberKey() ? 'Tetap masuk di perangkat ini sampai ' + new Date(+localStorage.getItem('osp_re')).toLocaleDateString('id-ID', {day: 'numeric', month: 'short', year: 'numeric'}) : 'Sesi otomatis keluar setelah 15 menit tanpa aktivitas') + ' · <a data-act="logout">Keluar</a></span>';
  bindEvents(); route(); trackView();
  ['click', 'keydown', 'scroll', 'mousemove', 'touchstart'].forEach(ev => window.addEventListener(ev, throttleTouch, {passive: true}));
  setInterval(() => { const t = +(sessionStorage.getItem('osp_t') || 0); if(Date.now() - t > IDLE_MS && !rememberKey()) logout(); }, 30000);
}
let lastTouch = 0;
function throttleTouch(){ const n = Date.now(); if(n - lastTouch > 5000){ lastTouch = n; touch(); } }

(async function init(){
  const msg = $('#gateMsg'), btn = $('#gateBtn');
  try {
    const rk = rememberKey();
    const k = sessionStorage.getItem('osp_k') || rk, t = +(sessionStorage.getItem('osp_t') || 0);
    if(k && (rk || Date.now() - t < IDLE_MS)){
      msg.style.color = 'var(--sub)'; msg.textContent = 'Memulihkan sesi…';
      const meta = await getMeta();
      const key = await crypto.subtle.importKey('raw', b64(k), {name: 'AES-GCM'}, true, ['decrypt']);
      await start(key, meta); return;
    }
  } catch(e) { try { sessionStorage.removeItem('osp_k'); } catch(_) {} if(e && e.name === 'OperationError') forgetRemember(); msg.textContent = ''; }
  $('#gatePw').focus();
  $('#gateForm').addEventListener('submit', async e => {
    e.preventDefault();
    const pw = $('#gatePw').value; if(!pw) return;
    btn.disabled = true; msg.style.color = 'var(--sub)'; msg.textContent = 'Membuka data…';
    try {
      const meta = await getMeta();
      await start(await deriveKey(pw, meta), meta, $('#gateRemember').checked);
    } catch(err) {
      btn.disabled = false; msg.style.color = 'var(--red)';
      msg.textContent = err && err.name === 'OperationError' ? 'Kata sandi salah.' : 'Gagal memuat data: ' + (err.message || err);
    }
  });
})();
