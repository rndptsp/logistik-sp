/* Outbound Logistics — Semen Padang.
   Data: data/site.enc (AES-GCM, gzip'd JSON) built by tools/build_site.py from the master Excel.
   Rules (same as the old logistik-sp dashboard): target = SNOP only; Target MTD = daily D1..D<last data
   day>; achievement % = FRC realisasi / SNOP; FOT has no target; PP BELAWAN SBA excluded by default. */
'use strict';

/* ===================== Small helpers ===================== */
const $ = (s, el) => (el || document).querySelector(s);
const $$ = (s, el) => [...(el || document).querySelectorAll(s)];
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => (n == null || isNaN(n)) ? '–' : Math.round(n).toLocaleString('id-ID');
const fmt1 = n => (n == null || isNaN(n)) ? '–' : n.toLocaleString('id-ID', {minimumFractionDigits:1, maximumFractionDigits:1});
const pctOf = (a, b) => b > 0 ? a / b * 100 : null;
const pctTxt = p => p == null ? 'N/A' : Math.round(p) + '%';
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
  if(p == null || isNaN(p)) return 'Tidak ada target SNOP untuk cakupan & periode ini.';
  if(p >= 100) return 'Kinerja solid — di atas target SNOP bulan berjalan.';
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
const soTip = o => o.so > 0 ? '|SO ' + fmt(o.so) + ' t · Real/SO ' + pctTxt(o.of) : (frcOn() ? '|SO tidak tersedia' : '');

/* ===================== State ===================== */
let B = null;            // decrypted bundle
let D = null;            // derived indexes
const F = { per: null, inc: 'FRC', srcOff: new Set(['PP BELAWAN SBA']) };   // per: {kind:'month',month} | {kind:'ytd'} | {kind:'range',from,to}
const DEFAULT_SRC_OFF = ['PP BELAWAN SBA'];

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
  B.facts.forEach(f => { if(f[1] === 0 && f[5] !== eksEmpty) frcEksp.add(f[5]); liveProv.add(f[3]); liveDist.add(f[4]); });
  const sources = [...new Set(B.facts.map(f => f[2]))].filter(i => i !== srcEmpty).map(i => dims.src[i]).sort();
  D = { dayMonth, dayNum, dayDow, months, provIdx, ekspIdx, distIdx, srcEmpty, eksEmpty,
        ekspList: [...frcEksp].map(i => dims.eksp[i]).sort(), liveProv, liveDist, sources };
  F.per = {kind: 'month', month: months[months.length - 1]};
}
const srcOk = i => i === D.srcEmpty || !F.srcOff.has(B.dims.src[i]);
const incOk = inc => F.inc === 'ALL' || (F.inc === 'FRC' ? inc === 0 : inc === 1);
const lastDayOf = m => Math.min(B.lastDay[m] || dim(m), dim(m));
const frcOn = () => F.inc !== 'FOT';

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
const FK = {prov: f => f[3], dist: f => f[4], eksp: f => f[5], src: f => f[2], all: () => 0};
const TK = {prov: t => t[1], dist: t => t[2], eksp: t => t[3], src: t => t[4], all: () => 0};
const SK = {prov: s => s[1], dist: s => s[2], eksp: s => s[3], src: () => null, all: () => 0};
function scopeOkF(f, sc){ return (sc.prov == null || f[3] === sc.prov) && (sc.dist == null || f[4] === sc.dist) && (sc.eksp == null || f[5] === sc.eksp); }
function scopeOkT(t, sc){ return (sc.prov == null || t[1] === sc.prov) && (sc.dist == null || t[2] === sc.dist) && (sc.eksp == null || t[3] === sc.eksp); }

/* Target of one SNOP row between two ISO dates (daily values when present, else monthly ÷ days). */
function targetBetween(t, from, to){
  const m = t[0], n = dim(m), a = m + '-01', b = m + '-' + pad2(n);
  if(b < from || a > to) return 0;
  const d0 = from > a ? +from.slice(8) : 1, d1 = to < b ? +to.slice(8) : n;
  if(t[6]){ let s = 0; for(let d = d0; d <= d1; d++) s += t[6][d - 1] || 0; return s; }
  return t[5] * (d1 - d0 + 1) / n;
}
function blank(){ return {vol:0, frc:0, trips:0, dwS:0, dwN:0, tgt:0, tgtFull:0, so:0, hasTgt:false}; }

/** Summary per group for a period (or a 'YYYY-MM' month). vol = selected incoterm(s); frc = FRC-only realisasi (achievement basis).
    tgt = target up to the last data day; tgtFull = target for the whole period (whole month for a month period). */
function summarize(group, per, sc, opt){
  sc = sc || {}; opt = opt || {};
  per = opt.toDay != null && typeof per === 'string' ? monthPer(per, opt.toDay) : toPer(per);
  const fullTo = per.kind === 'month' ? per.month + '-' + pad2(dim(per.month)) : per.to;
  const out = new Map();
  const get = k => { let o = out.get(k); if(!o){ o = blank(); o.key = k; out.set(k, o); } return o; };
  B.facts.forEach(f => {
    if(!inPer(f[0], per) || !srcOk(f[2]) || !scopeOkF(f, sc)) return;
    const o = get(FK[group](f));
    if(f[1] === 0) o.frc += f[6];
    if(incOk(f[1])){ o.vol += f[6]; o.trips += f[7]; o.dwS += f[8]; o.dwN += f[9]; }
  });
  if(frcOn()){
    const m0 = per.from.slice(0, 7), m1 = fullTo.slice(0, 7);
    B.targets.forEach(t => {
      if(t[0] < m0 || t[0] > m1 || !srcOk(t[4]) || !scopeOkT(t, sc)) return;
      const o = get(TK[group](t));
      o.tgt += targetBetween(t, per.from, per.tgtTo); o.tgtFull += targetBetween(t, per.from, fullTo); o.hasTgt = true;
    });
    // SO counts every month the period touches in full: by delivery month (TGL_KIRIM, B.sod) when the build has it,
    // otherwise by PERIODE (B.so)
    if(group !== 'src'){
      const m9 = per.to.slice(0, 7);
      B.sod.forEach(s => {
        if(s[0].slice(0, 7) < m0 || s[0].slice(0, 7) > m9 || !scopeOkT(s, sc)) return;
        get(SK[group](s)).so += s[4];
      });
      B.so.forEach(s => {
        if(s[0] < m0 || s[0] > per.to.slice(0, 7) || B.sodMonths.has(s[0]) || !scopeOkT(s, sc)) return;
        get(SK[group](s)).so += s[4];
      });
    }
  }
  out.forEach(o => {
    o.pct = o.hasTgt && frcOn() ? (o.tgt > 0 ? pctOf(o.frc, o.tgt) : 100) : null;
    o.of = frcOn() && o.so > 0 ? pctOf(o.frc, o.so) : null;
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
    if(!inPer(t[0], per) || !incOk(t[2]) || !srcOk(t[4])) return;
    if(sc.eksp != null && t[3] !== sc.eksp) return;
    if(sc.prov != null && t[5] !== sc.prov) return;
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
    if(ds < fromISO || ds > toISO || !srcOk(f[2]) || !scopeOkF(f, sc)) return;
    const o = get(FK[group](f));
    if(f[1] === 0) o.frc += f[6];
    if(incOk(f[1])){ o.vol += f[6]; o.trips += f[7]; }
  });
  if(frcOn()){
    const from = parseISO(fromISO), to = parseISO(toISO);
    B.targets.forEach(t => {
      const m = t[0];
      if(m + '-31' < fromISO || m + '-01' > toISO || !srcOk(t[4]) || !scopeOkT(t, sc)) return;
      const y = +m.slice(0, 4), mo = +m.slice(5) - 1, n = dim(m);
      let add = 0, any = false;
      for(let d = 1; d <= n; d++){
        const dt = new Date(y, mo, d);
        if(dt < from || dt > to) continue;
        add += t[6] ? (t[6][d - 1] || 0) : t[5] / n; any = true;
      }
      if(any){ const o = get(TK[group](t)); o.tgt += add; o.hasTgt = true; }
    });
  }
  out.forEach(o => { o.pct = o.hasTgt && frcOn() ? (o.tgt > 0 ? pctOf(o.frc, o.tgt) : 100) : null; o.gap = o.frc - o.tgt; });
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
function barsHTML(id, data){
  const hasC = data.some(d => d.c);
  const max = Math.max(1, ...data.map(d => Math.max(d.vol, d.tgt || 0, d.c ? d.c.vol : 0)));
  return '<div class="bars' + (hasC ? ' pair' : '') + (data.length > 16 ? ' dense' : '') + '">' + data.map((d, i) => {
    const h = Math.max(2, d.vol / max * 100);
    const th = d.tgt != null ? Math.min(100, d.tgt / max * 100) : null;
    const c = d.pct != null ? tierColor(d.pct) : 'var(--black)';
    const unit = d.unit === 'jam' ? fmt1(d.vol) : fmt(d.vol);
    return '<div class="bcol click" data-bar="' + id + ':' + i + '" title="' + esc(d.long || d.label) + ' — klik untuk rincian">' +
      (d.pct != null ? '<div class="bpct" style="color:' + c + '">' + Math.round(d.pct) + '%</div>' : '') +
      '<div class="bval">' + unit + '</div>' + (d.c ? deltaHTML(d.vol, d.c.vol) : '') +
      '<div class="bwrap"><div class="b" style="height:' + h + '%;background:' + (d.color || c) + ';animation-delay:' + (i * 50) + 'ms"></div>' +
      (d.c ? '<div class="b b2" style="height:' + Math.max(2, d.c.vol / max * 100) + '%;animation-delay:' + (i * 50 + 25) + 'ms"></div>' : '') +
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
    groups.push(['Ekspeditur lain', D.ekspList.filter(e => D.ekspIdx[e] !== sc.eksp).map(e => ['eksp:' + D.ekspIdx[e], e])]);
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
  const legend = (inc !== 'FOT' ? '<span><i style="background:var(--green)"></i>&ge;100% target</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target SNOP</span>' : '<span><i style="background:var(--black)"></i>volume FOT (tanpa target)</span>') +
    (T.cmp ? '<span><i style="background:#BCC3C5"></i>' + esc(cmpLabel(T)) + '</span><span>▲▼ = selisih vs pembanding</span>' : '');
  return '<div class="pt"><span>' + esc(T.title) + ' (ton)</span><div class="pt-tools">' + (T.noCmp ? '' : cmpOptions(T)) + '<div class="gran-chips">' + chips + '</div></div></div>' + barsHTML(id, data) +
    '<div class="trend-legend">' + legend + '</div>' +
    '<div class="trend-detail" id="' + id + '-d">Klik batang mana pun untuk rincian per provinsi, distrik, ekspeditur &amp; source' + (T.cmp ? ', dibandingkan dengan ' + esc(cmpLabel(T)) : '') + '.</div>';
}

/* ===================== Drill-down modal ===================== */
let MODAL = null;
const DIM_LBL = {prov: 'Provinsi', dist: 'Distrik', eksp: 'Ekspeditur', src: 'Source plant'};
function dimName(dim, k){
  if(dim === 'prov') return tc(B.dims.prov[k]);
  if(dim === 'dist') return tc(B.dims.dist[k]);
  if(dim === 'eksp') return B.dims.eksp[k] || '(tanpa ekspeditur / FOT)';
  return tc(B.dims.src[k] || '(tanpa plant)');
}
function dimHref(dim, k){
  if(dim === 'prov') return '#/provinsi/' + enc(B.dims.prov[k]);
  if(dim === 'dist') return '#/distrik/' + enc(B.dims.dist[k]);
  if(dim === 'eksp') return B.dims.eksp[k] ? '#/ekspeditur/' + enc(B.dims.eksp[k]) : null;
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
  MODAL = {T, d, dims, dim: dims[0]};
  const range = d.from === d.to ? dayLabel(d.from) : dayLabel(d.from) + ' – ' + dayLabel(d.to);
  if(!T.sc) T.sc = {};
  openModal(scName(T.sc) + ' · ' + (d.long || d.label), esc(range) + ' · ' + esc(T.incFix || (F.inc === 'ALL' ? 'FRC + FOT' : F.inc)), modalBody());
}
function modalBody(){
  const {T, d, dims, dim} = MODAL;
  const pb = shiftBucket(d, T.gran);
  return withInc(T, () => {
    const cur = rangeGroup(d.from, d.to, T.sc, dim), prv = rangeGroup(pb.from, pb.to, T.sc, dim);
    const rows = [...cur.values()].filter(o => o.vol > 0 || o.tgt > 0).map(o => Object.assign(o, {prev: (prv.get(o.key) || {}).vol || 0}));
    rows.sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
    let r = 0; rows.forEach(o => { o.rank = o.pct == null ? null : ++r; });
    const tot = rangeTotals(d.from, d.to, T.sc), tp = rangeTotals(pb.from, pb.to, T.sc);
    const chg = tp.vol ? (tot.vol / tp.vol - 1) * 100 : null;
    let cmpCell = '';
    if(d.c && T.cmp !== 'prev'){
      const cc = d.c, dv = cc.vol ? (d.vol / cc.vol - 1) * 100 : null;
      cmpCell = '<div class="prog-cell"><div class="pl">vs ' + esc(cmpLabel(T)) + '</div><div class="pv">' + fmt(cc.vol) + ' t</div><div class="pd">capaian ' + pctTxt(cc.pct) + (dv != null ? ' · selisih ' + (dv >= 0 ? '+' : '') + Math.round(dv) + '%' : '') + '</div></div>';
    }
    const head = '<div class="prog-row" style="margin-bottom:18px">' +
      '<div class="prog-cell"><div class="pl">Realisasi</div><div class="pv">' + fmt(tot.vol) + ' t</div><div class="pd">' + fmt(tot.trips) + ' trip</div></div>' +
      '<div class="prog-cell"><div class="pl">Target SNOP</div><div class="pv">' + (tot.hasTgt ? fmt(tot.tgt) + ' t' : '–') + '</div><div class="pd">gap ' + (tot.hasTgt ? signed(tot.gap) + ' t' : '–') + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Capaian</div><div class="pv" style="color:' + tierColor(tot.pct) + '">' + pctTxt(tot.pct) + '</div><div class="pd">FRC vs SNOP</div></div>' +
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
    return head + tabs + (rows.length ? table(cols, rows, {go: o => dimHref(dim, o.key)}) : '<div class="dp-review">Tidak ada data untuk periode ini.</div>') +
      '<div class="note" style="margin-top:8px">“Periode lalu” = ' + esc(PREV_LBL[T.gran]) + '. Klik baris untuk membuka halamannya, klik judul kolom untuk mengurutkan.</div>';
  });
}

/* ===================== Sortable tables ===================== */
const TABLES = {};
function table(cols, rows, opt){
  opt = opt || {};
  const id = 'tb' + Math.random().toString(36).slice(2, 8);
  TABLES[id] = {cols, rows, opt, sort: opt.sort || null, dir: opt.dir || -1};
  return '<div class="tbl-wrap" id="' + id + '">' + tableInner(id) + '</div>';
}
function tableInner(id){
  const T = TABLES[id];
  let rows = T.rows.slice();
  if(T.sort != null){
    const c = T.cols[T.sort], v = c.sortVal || c.val;
    rows.sort((a, b) => { const x = v(a), y = v(b); if(x == null && y == null) return 0; if(x == null) return 1; if(y == null) return -1; return (x < y ? -1 : x > y ? 1 : 0) * T.dir; });
  }
  const lim = T.opt.limit || 500;
  const head = '<tr>' + T.cols.map((c, i) => '<th class="' + (c.num ? 'num' : '') + '" data-sort="' + id + ':' + i + '">' + c.h + (T.sort === i ? '<span class="arr">' + (T.dir > 0 ? '▲' : '▼') + '</span>' : '') + '</th>').join('') + '</tr>';
  const body = rows.slice(0, lim).map((r, ri) => {
    const href = T.opt.go ? T.opt.go(r) : null;
    return '<tr' + (href ? ' class="go" data-href="' + href + '"' : '') + '>' + T.cols.map(c => '<td class="' + (c.cls || '') + (c.num ? ' num' : '') + '"' + (c.style ? ' style="' + c.style(r) + '"' : '') + '>' + (c.html ? c.html(r, ri) : esc(c.val(r))) + '</td>').join('') + '</tr>';
  }).join('');
  const tot = T.opt.total ? '<tr class="total">' + T.cols.map(c => '<td class="' + (c.num ? 'num' : '') + '">' + (c.total != null ? c.total : '') + '</td>').join('') + '</tr>' : '';
  const more = rows.length > lim ? '<tr><td colspan="' + T.cols.length + '" style="color:var(--sub)">… ' + (rows.length - lim) + ' baris lainnya tidak ditampilkan</td></tr>' : '';
  return '<table class="tbl"><thead>' + head + '</thead><tbody>' + body + tot + more + '</tbody></table>';
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
  dwell: {h:'Dwell (jam)', num:true, val: r => r.dwell, html: r => r.dwell == null ? '–' : fmt1(r.dwell)},
  note: {h:'Catatan', cls:'note-c', val: r => reviewText(r.pct), sortVal: r => r.pct}
};

/* ===================== Shell: nav, filter, search ===================== */
function setActiveNav(key){ $$('#mainNav .navitem').forEach(n => n.classList.toggle('on', n.dataset.nav === key)); }
function closeNav(){ $$('#mainNav .navitem').forEach(n => n.classList.remove('open')); }
function buildNav(){
  const pm = summarize('prov', PER());
  const provs = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  $('#navProvList').innerHTML = provs.map(o => '<a href="#/provinsi/' + enc(B.dims.prov[o.key]) + '">' + esc(tc(B.dims.prov[o.key])) + '<span>' + pctTxt(o.pct) + '</span></a>').join('') || '<span class="empty">—</span>';
  const em = summarize('eksp', PER());
  $('#navEkspList').innerHTML = D.ekspList.map(code => {
    const o = em.get(D.ekspIdx[code]);
    const p = o ? o.pct : null;
    return '<a class="eksp-item" href="#/ekspeditur/' + enc(code) + '"><span class="en">' + esc(code) + '</span><span class="ep" style="color:' + tierColor(p) + '">' + (o && o.hasTgt ? pctTxt(p) : (o && o.vol ? fmt(o.vol) + ' t' : '–')) + '</span></a>';
  }).join('');
  $('#dayLabel').textContent = 'Data s.d. ' + dayLabel(B.asOf);
}
let draft = null;
function buildFilterPanel(){
  const cur = PER();
  draft = {per: Object.assign({}, F.per), inc: F.inc, srcOff: new Set(F.srcOff),
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
  const inc = [['FRC','FRC'],['FOT','FOT'],['ALL','FRC + FOT']].map(x => '<span data-inc="' + x[0] + '" class="' + (draft.inc === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  const src = D.sources.map(s2 => '<label class="fp-check"><input type="checkbox" data-src="' + esc(s2) + '"' + (draft.srcOff.has(s2) ? '' : ' checked') + '> ' + esc(tc(s2)) + (DEFAULT_SRC_OFF.includes(s2) ? ' <small>(default tidak dicentang)</small>' : '') + '</label>').join('');
  $('#filterPanel').innerHTML =
    '<div class="fp-row"><div class="fp-group"><label class="fl">Periode</label><div class="fp-chips" style="margin-bottom:10px">' + mode + '</div>' + body + '</div>' +
    '<div class="fp-group"><label class="fl">Incoterm</label><div class="fp-chips">' + inc + '</div><div style="font-size:11.5px;color:var(--sub);margin-top:8px;line-height:1.5">% pencapaian selalu dihitung dari FRC vs target SNOP sampai hari data terakhir. FOT tidak punya target.</div></div></div>' +
    '<div class="fp-row"><div class="fp-group"><label class="fl">Source plant</label>' + src + '</div>' +
    '<div class="fp-group"><label class="fl">Berlaku untuk</label><div style="font-size:12px;color:var(--gray-d);line-height:1.6">Semua halaman: peta, grafik, tabel, ekspeditur, distributor, armada &amp; prognosa. Pilihan tersimpan selama sesi.</div></div></div>' +
    '<div class="fp-foot"><div class="fp-reset" data-act="reset">Reset filter</div><div class="fp-apply" data-act="apply">Terapkan</div></div>';
}
function filterStrip(){
  const offs = [...F.srcOff];
  const per = PER();
  $('#filterSummary').innerHTML = '<span>Menampilkan</span><span class="tag"><b>' + esc(per.label) + '</b>' + (per.sub !== per.label ? ' · ' + esc(per.sub.replace(per.label + ' · ', '').replace(per.label, '')) : '') + '</span>' +
    '<span class="tag"><b>' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + '</b></span>' +
    (offs.length ? '<span class="tag">tanpa ' + offs.map(s => esc(tc(s))).join(', ') + '</span>' : '<span class="tag">semua source</span>') +
    '<span class="edit" data-act="openfilter">Ubah filter &#9662;</span>';
}
function applyFilter(){
  const k = draft.per.kind;
  if(k === 'range'){ let a = draft.from || B.days[0], b = draft.to || B.asOf; if(a > b) [a, b] = [b, a]; F.per = {kind: 'range', from: a, to: b}; }
  else if(k === 'ytd') F.per = {kind: 'ytd'};
  else F.per = {kind: 'month', month: draft.per.month || D.months[D.months.length - 1]};
  F.inc = draft.inc; F.srcOff = new Set(draft.srcOff);
  try { sessionStorage.setItem('osp_f', JSON.stringify({per: F.per, inc: F.inc, srcOff: [...F.srcOff]})); } catch(e) {}
  closeNav(); closeFilter(); buildNav(); filterStrip(); route();
}
function closeFilter(){ const f = $('#filterStrip'); if(f) f.classList.remove('open'); }
function openFilter(){
  closeNav(); const f = $('#filterStrip');
  if(f.classList.contains('open')){ f.classList.remove('open'); return; }
  buildFilterPanel(); f.classList.add('open');
  const panel = $('#filterPanel');   // on narrow screens the panel is position:fixed under the strip
  panel.style.top = window.innerWidth <= 1100 ? (f.getBoundingClientRect().bottom + 4) + 'px' : '';
}
function restoreFilter(){
  try {
    const s = JSON.parse(sessionStorage.getItem('osp_f') || 'null');
    if(s && s.per && (s.per.kind !== 'month' || D.months.includes(s.per.month))){ F.per = s.per; F.inc = s.inc; F.srcOff = new Set(s.srcOff); }
  } catch(e) {}
}

let searchIndex = [];
function buildSearch(){
  const P = B.dims.prov, Ds = B.dims.dist;
  searchIndex = [];
  [...D.liveProv].forEach(i => searchIndex.push({t: tc(P[i]), k: 'Provinsi', h: '#/provinsi/' + enc(P[i])}));
  [...D.liveDist].forEach(i => searchIndex.push({t: tc(Ds[i]), k: tc(P[B.dims.distProv[i]] || ''), h: '#/distrik/' + enc(Ds[i])}));
  D.ekspList.forEach(e => searchIndex.push({t: e, k: 'Ekspeditur', h: '#/ekspeditur/' + enc(e)}));
  const tokoDistr = new Map(), seenD = new Set();
  B.ship.forEach(r => { seenD.add(r[5]); if(B.dims.toko[r[7]] && B.dims.toko[r[7]][0] && !tokoDistr.has(r[7])) tokoDistr.set(r[7], r[5]); });
  seenD.forEach(i => { if(B.dims.distr[i]) searchIndex.push({t: tc(B.dims.distr[i]), k: 'Distributor', h: '#/distributor/' + enc(B.dims.distr[i])}); });
  tokoDistr.forEach((d, i) => { const t = B.dims.toko[i]; searchIndex.push({t: (t[1] || t[0]) + (t[2] ? ' — ' + t[2] : ''), k: 'Toko · ' + tc(B.dims.distr[d]).slice(0, 22), h: '#/distributor/' + enc(B.dims.distr[d])}); });
  [['Ringkasan','#/'],['Semua provinsi','#/provinsi'],['Peringkat ekspeditur & scorecard','#/ekspeditur'],['Distributor & toko tujuan','#/distributor'],['Armada / truk / dwell','#/armada'],['Prognosa hari ini','#/prognosa/hari-ini'],['Prognosa minggu ini','#/prognosa/minggu-ini'],['Prognosa akhir bulan (proyeksi)','#/prognosa/akhir-bulan']]
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
  'RIAU DARATAN': ['Riau', 0, -8], 'KEPULAUAN RIAU': ['Kep. Riau', 0, -22], 'JAMBI': ['Jambi', 4, -6],
  'SUMATERA SELATAN': ['Sumatera Selatan', 0, -8], 'BENGKULU': ['Bengkulu', -62, 6, 'out'], 'BANGKA BELITUNG': ['Babel', 0, 0],
  'LAMPUNG': ['Lampung', -8, 0]
};
const PLANTS = [  // [label, lon, lat, label side]
  ['CP Indarung', 100.47, -0.95, 'left'], ['GP Dumai', 101.45, 1.67, 'right'], ['PP Bengkulu', 102.27, -3.80, 'right'], ['PP Belawan', 98.69, 3.78, 'right']
];
function mapHTML(provRows, opt){
  const G = window.SUMATRA, hero = !!(opt && opt.hero);
  const byName = Object.fromEntries(provRows.map(o => [B.dims.prov[o.key], o]));
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
    const tip = tc(id) + '|Realisasi ' + fmt(o.vol) + ' t' + (o.hasTgt ? ' · Target MTD ' + fmt(o.tgt) + ' t' : '') + '|' + (o.pct != null ? 'Capaian ' + pctTxt(o.pct) + ' · gap ' + signed(o.gap) + ' t' : 'Tanpa target SNOP') + soTip(o);
    shapes += '<path class="map-region live' + (worst && o === worst ? ' is-attn' : '') + '" d="' + g.d + '" fill="' + fill + '" stroke="' + (hero ? '#0F1A21' : '#fff') + '" stroke-width="1.2" filter="url(#softShadow)" data-href="#/provinsi/' + enc(id) + '" data-tip="' + esc(tip) + '"/>';
    const out = L[3] === 'out' ? ' out' : '';   // label placed in the sea next to a narrow province
    labels += '<text x="' + lx + '" y="' + (ly - 4) + '" class="map-lbl' + out + '">' + esc(L[0]) + '</text>' +
      '<text x="' + lx + '" y="' + (ly + 11) + '" class="map-pct' + out + '" style="' + (out ? 'fill:' + tierColor(o.pct) : '') + '">' + (o.pct != null ? pctTxt(o.pct) : fmt(o.vol) + ' t') + '</text>';
  });
  const pins = PLANTS.filter(p => !F.srcOff.has(p[0] === 'PP Belawan' ? 'PP BELAWAN SBA' : p[0].toUpperCase())).map(p => {
    const x = ((p[1] - 95) * 40).toFixed(1), y = ((6.2 - p[2]) * 40).toFixed(1);
    return '<g class="map-pin" data-tip="' + esc(p[0] + '|Plant / packing plant asal pengiriman') + '"><circle cx="' + x + '" cy="' + y + '" r="5.5" fill="#fff" stroke="#000" stroke-width="1.6"/><circle cx="' + x + '" cy="' + y + '" r="2.2" fill="var(--red)"/>' +
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
function distKey(s){
  s = String(s || '').toUpperCase().trim().replace(/^KAB(UPATEN)?\b\.?\s*/, '');
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
const PLANT_PROV = {'CP Indarung': 'SUMATERA BARAT', 'GP Dumai': 'RIAU DARATAN', 'PP Bengkulu': 'BENGKULU', 'PP Belawan': 'SUMATERA UTARA'};
function distMapSection(pname, m, sel){
  const map = distMapHTML(pname, m, {sel});
  return map ? '<div class="section-head"><div class="section-title">Peta capaian per distrik' + (sel != null ? ' — ' + esc(tc(pname)) : '') + '</div>' +
    (sel != null ? '<a class="section-link" href="#/provinsi/' + enc(pname) + '">Halaman provinsi &rsaquo;</a>' : '') + '</div>' + map : '';
}
/** Province map coloured by district achievement. opt.sel = data district index to highlight. */
function distMapHTML(pname, m, opt){
  opt = opt || {};
  if(!window.DISTRICTS){
    distMapPending = () => distMapHTML(pname, m, opt);
    loadDistricts();
    return '<div id="dmapSlot" class="map-wrap dmap"><div class="empty">Memuat peta distrik…</div></div>';
  }
  const G = window.DISTRICTS[pname];
  if(!G) return '';
  const pi = D.provIdx[pname], rows = [...summarize('dist', m, {prov: pi}).values()].filter(o => o.vol > 0 || o.tgt > 0);
  // data districts -> map shapes (several SAP spellings can land on one shape)
  const agg = new Map(), missing = [];
  rows.forEach(o => {
    const name = B.dims.dist[o.key], i = matchDistrict(name, G.d);
    if(i < 0){ missing.push(o); return; }
    let a = agg.get(i);
    if(!a){ a = {vol:0, frc:0, tgt:0, so:0, trips:0, hasTgt:false, keys:[]}; agg.set(i, a); }
    a.vol += o.vol; a.frc += o.frc; a.tgt += o.tgt; a.so += o.so; a.trips += o.trips; a.hasTgt = a.hasTgt || o.hasTgt; a.keys.push(o.key);
  });
  agg.forEach(a => { a.pct = a.hasTgt && frcOn() ? (a.tgt > 0 ? pctOf(a.frc, a.tgt) : 100) : null; a.gap = a.frc - a.tgt; a.of = frcOn() && a.so > 0 ? pctOf(a.frc, a.so) : null; });   // same rules as summarize()
  const [vx, vy, vw, vh] = G.vb, fs = Math.max(vw, vh) / 52;
  const defs = '<defs>' + [['Green', '#2E9E5C', '#1A7A42'], ['Amber', '#C9952E', '#8A5A00'], ['Red', '#FF6B74', '#E02A36'], ['Ink', '#5A5A5A', '#2B2B2B']].map(c =>
    '<linearGradient id="d' + c[0] + '" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="' + c[1] + '"/><stop offset="100%" stop-color="' + c[2] + '"/></linearGradient>').join('') + '</defs>';
  let shapes = '', labels = '', selShape = '';
  G.d.forEach((g, i) => {
    const a = agg.get(i), sel = a && opt.sel != null && a.keys.includes(opt.sel);
    if(!a){
      shapes += '<path class="map-region" d="' + g.d + '" fill="#D9DEDF" stroke="#fff" stroke-width="1" vector-effect="non-scaling-stroke" data-tip="' + esc(g.n + '|Tidak ada pengiriman di periode ini') + '"/>';
      return;
    }
    const fill = a.pct != null ? 'url(#d' + tierName(a.pct) + ')' : 'url(#dInk)';
    const tip = g.n + '|Realisasi ' + fmt(a.vol) + ' t · ' + fmt(a.trips) + ' trip' + (a.hasTgt ? ' · Target MTD ' + fmt(a.tgt) + ' t' : '') + '|' + (a.pct != null ? 'Capaian ' + pctTxt(a.pct) + ' · gap ' + signed(a.gap) + ' t' : 'Tanpa target SNOP') + soTip(a);
    const path = '<path class="map-region live' + (sel ? ' is-sel' : '') + '" d="' + g.d + '" fill="' + fill + '" stroke="' + (sel ? '#000' : '#fff') + '" stroke-width="' + (sel ? 3 : 1.2) + '" vector-effect="non-scaling-stroke" data-href="#/distrik/' + enc(B.dims.dist[a.keys[0]]) + '" data-tip="' + esc(tip) + '"/>';
    if(sel) selShape = path; else shapes += path;
    if(Math.min(g.w, g.h) >= Math.max(vw, vh) * 0.09){
      const short = g.n.replace(/^Kota /, 'Kt. ');
      labels += '<text x="' + g.cx + '" y="' + (g.cy - fs * 0.3) + '" class="map-lbl" style="font-size:' + fs.toFixed(1) + 'px">' + esc(short) + '</text>' +
        '<text x="' + g.cx + '" y="' + (g.cy + fs * 1.05) + '" class="map-pct" style="font-size:' + (fs * 1.15).toFixed(1) + 'px">' + (a.pct != null ? pctTxt(a.pct) : fmt(a.vol) + ' t') + '</text>';
    }
  });
  const pins = PLANTS.filter(p => PLANT_PROV[p[0]] === pname && !F.srcOff.has(p[0] === 'PP Belawan' ? 'PP BELAWAN SBA' : p[0].toUpperCase())).map(p => {
    const x = (p[1] - 95) * 200, y = (6.2 - p[2]) * 200;
    return '<g class="map-pin" data-tip="' + esc(p[0] + '|Plant / packing plant asal pengiriman') + '"><circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + (fs * 0.55).toFixed(1) + '" fill="#fff" stroke="#000" stroke-width="1.6" vector-effect="non-scaling-stroke"/><circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + (fs * 0.22).toFixed(1) + '" fill="var(--red)"/>' +
      '<text x="' + (x + fs * 0.9).toFixed(1) + '" y="' + (y + fs * 0.35).toFixed(1) + '" class="pin-lbl" style="font-size:' + (fs * 0.95).toFixed(1) + 'px">' + esc(p[0]) + '</text></g>';
  }).join('');
  const svg = '<svg class="map-svg" viewBox="' + G.vb.join(' ') + '" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta capaian per distrik di ' + esc(tc(pname)) + '">' +
    defs + '<rect x="' + vx + '" y="' + vy + '" width="' + vw + '" height="' + vh + '" fill="#EEF4F7"/>' + shapes + selShape + labels + pins + '</svg>';
  const list = [...agg.entries()].map(([i, a]) => Object.assign({n: G.d[i].n}, a))
    .sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  const legend = '<div class="map-row map-head"><span class="dot"></span><span class="nm"></span><span class="tn">Realisasi</span><span class="pc">vs target</span><span class="so">Real/SO</span></div>' + list.map((a, r) => '<a class="map-row' + (opt.sel != null && a.keys.includes(opt.sel) ? ' is-sel' : '') + '" href="#/distrik/' + enc(B.dims.dist[a.keys[0]]) + '"><span class="dot" style="background:' + tierColor(a.pct) + '"></span><span class="nm">#' + (r + 1) + ' ' + esc(a.n) + '</span><span class="tn">' + fmt(a.vol) + ' t</span><span class="pc" style="color:' + tierColor(a.pct) + '">' + pctTxt(a.pct) + '</span><span class="so">' + (a.of == null ? '–' : pctTxt(a.of)) + '</span></a>').join('') +
    (missing.length ? '<div class="map-row muted"><span class="dot" style="background:transparent"></span><span class="nm">Tidak ada di peta: ' + missing.map(o => esc(tc(B.dims.dist[o.key]))).join(', ') + '</span></div>' : '') +
    '<div class="map-key"><span><i style="background:#1A7A42"></i>&ge;100%</span><span><i style="background:#8A5A00"></i>85–99%</span><span><i style="background:#E02A36"></i>&lt;85%</span><span><i style="background:#D9DEDF"></i>tanpa kiriman</span></div>';
  return '<div class="map-wrap dmap">' + svg + '<div class="map-legend">' + legend + '</div></div>';
}

/* ===================== Shipping destinations: distributor & ship-to (toko/gudang) ===================== */
/* B.ship rows: [day, inc, src, prov, dist, distributor, eksp, toko, ton, trips] */
const SHK = {prov: r => r[3], dist: r => r[4], distr: r => r[5], eksp: r => r[6], toko: r => r[7]};
function shipRows(m, sc){
  sc = sc || {};
  const per = m == null ? null : toPer(m);
  return B.ship.filter(r => (per == null || inPer(r[0], per)) && incOk(r[1]) && srcOk(r[2]) &&
    (sc.prov == null || r[3] === sc.prov) && (sc.dist == null || r[4] === sc.dist) &&
    (sc.distr == null || r[5] === sc.distr) && (sc.eksp == null || r[6] === sc.eksp));
}
function buildTree(rows, levels, depth){
  depth = depth || 0;
  const lv = levels[depth], groups = new Map();
  rows.forEach(r => { const k = SHK[lv](r); let g = groups.get(k); if(!g){ g = {lv, key: k, vol: 0, trips: 0, last: -1, rows: []}; groups.set(k, g); } g.vol += r[8]; g.trips += r[9]; g.last = Math.max(g.last, r[0]); g.rows.push(r); });
  const nodes = [...groups.values()].sort((a, b) => b.vol - a.vol);
  nodes.forEach(n => { if(depth + 1 < levels.length) n.children = buildTree(n.rows, levels, depth + 1); n.rows = null; });
  return nodes;
}
const LV_LBL = {prov: 'Provinsi', dist: 'Distrik', distr: 'Distributor', eksp: 'Ekspeditur', toko: 'Toko / gudang tujuan'};
function nodeLabel(n){
  if(n.lv === 'prov') return esc(tc(B.dims.prov[n.key]));
  if(n.lv === 'dist') return esc(tc(B.dims.dist[n.key]));
  if(n.lv === 'distr') return esc(tc(B.dims.distr[n.key]) || '(tanpa distributor)');
  if(n.lv === 'eksp') return esc(B.dims.eksp[n.key] || 'FOT (angkut sendiri)');
  const t = B.dims.toko[n.key];
  if(!t || !t[0]) return '<span class="t-unk">Toko belum tercatat</span><small>data tujuan tidak tersedia untuk periode ini</small>';
  return esc(t[1] || t[0]) + '<small>' + esc([t[2], t[0]].filter(Boolean).join(' · ')) + '</small>';
}
function nodeHref(n){
  if(n.lv === 'prov') return '#/provinsi/' + enc(B.dims.prov[n.key]);
  if(n.lv === 'dist') return '#/distrik/' + enc(B.dims.dist[n.key]);
  if(n.lv === 'distr') return B.dims.distr[n.key] ? '#/distributor/' + enc(B.dims.distr[n.key]) : null;
  if(n.lv === 'eksp') return B.dims.eksp[n.key] ? '#/ekspeditur/' + enc(B.dims.eksp[n.key]) : null;
  return null;
}
/* Expandable table: click a row to open the next level; "›" opens that item's own page.
   opt.tgt = {prov: Map, dist: Map} summaries (from summarize) to show target & achievement at those levels. */
function treeHTML(nodes, levels, opt){
  opt = opt || {};
  const id = 'tr' + Math.random().toString(36).slice(2, 8);
  const hasT = !!opt.tgt && frcOn();
  let html = '<div class="tree-tools"><span>' + levels.map(l => LV_LBL[l]).join(' &rsaquo; ') + '</span><span><a data-tree-all="' + id + ':1">Buka semua</a> · <a data-tree-all="' + id + ':0">Tutup semua</a></span></div>' +
    '<div class="tbl-wrap"><table class="tbl tree" id="' + id + '"><thead><tr><th>Tujuan</th><th class="num">Realisasi (t)</th>' + (hasT ? '<th class="num">Target MTD (t)</th><th class="num">Capaian</th><th class="num">SO (t)</th><th class="num">Real/SO</th>' : '') + '<th class="num">Trip</th><th class="num">Kirim terakhir</th><th></th></tr></thead><tbody>';
  const walk = (list, depth, parent) => list.forEach((n, i) => {
    const path = parent ? parent + '-' + i : String(i), kids = n.children && n.children.length;
    const t = hasT && opt.tgt[n.lv] ? opt.tgt[n.lv].get(n.key) : null;
    const href = nodeHref(n);
    html += '<tr class="lv' + depth + (depth ? ' tr-hide' : '') + (kids ? ' has-kids' : '') + '" data-path="' + path + '" data-parent="' + (parent || '') + '">' +
      '<td class="name" style="padding-left:' + (12 + depth * 22) + 'px"><span class="caret-t">' + (kids ? '▸' : '') + '</span><span class="t-lbl">' + nodeLabel(n) + '</span>' + (kids ? '<span class="t-cnt">' + n.children.length + ' ' + LV_LBL[levels[depth + 1]].toLowerCase() + '</span>' : '') + '</td>' +
      '<td class="num">' + fmt(n.vol) + '</td>' +
      (hasT ? (t && t.hasTgt ? '<td class="num">' + fmt(t.tgt) + '</td><td class="num pc" style="color:' + tierColor(t.pct) + '">' + pctTxt(t.pct) + '</td>' : '<td class="num">–</td><td class="num">–</td>') +
        (t && t.so ? '<td class="num">' + fmt(t.so) + '</td><td class="num pc">' + pctTxt(t.of) + '</td>' : '<td class="num">–</td><td class="num">–</td>') : '') +
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
    if(r.classList.contains('has-kids')){ r.classList.toggle('open', open); $('.caret-t', r).textContent = open ? '▾' : '▸'; }
  });
}
function distrSummary(rows){
  const per = new Map();
  rows.forEach(r => { let o = per.get(r[5]); if(!o){ o = {key: r[5], frc: 0, fot: 0, vol: 0, trips: 0, prov: new Set(), dist: new Set(), toko: new Set(), last: -1}; per.set(r[5], o); }
    if(r[1] === 0) o.frc += r[8]; else o.fot += r[8]; o.vol += r[8]; o.trips += r[9]; o.prov.add(r[3]); o.dist.add(r[4]); if(B.dims.toko[r[7]] && B.dims.toko[r[7]][0]) o.toko.add(r[7]); o.last = Math.max(o.last, r[0]); });
  return [...per.values()];
}

function pageDistrList(){
  const m = PER(), rows = shipRows(m), list = distrSummary(rows).sort((a, b) => b.vol - a.vol);
  list.forEach((o, i) => o.rank = i + 1);
  const tot = list.reduce((a, o) => a + o.vol, 0), unk = rows.filter(r => !(B.dims.toko[r[7]] && B.dims.toko[r[7]][0])).reduce((a, r) => a + r[8], 0);
  const cols = [{h:'#', cls:'rank', val: o => o.rank, html: o => '#' + o.rank}, {h:'Distributor', cls:'name', val: o => tc(B.dims.distr[o.key])},
    {h:'FRC (t)', num:true, val: o => o.frc, html: o => o.frc ? fmt(o.frc) : '–'}, {h:'FOT (t)', num:true, val: o => o.fot, html: o => o.fot ? fmt(o.fot) : '–'},
    {h:'Total (t)', num:true, val: o => o.vol, html: o => '<b>' + fmt(o.vol) + '</b>'}, {h:'Share', num:true, val: o => o.vol / (tot || 1), html: o => pctTxt(pctOf(o.vol, tot))},
    {h:'Trip', num:true, val: o => o.trips}, {h:'Provinsi', num:true, val: o => o.prov.size}, {h:'Distrik', num:true, val: o => o.dist.size},
    {h:'Toko / gudang', num:true, val: o => o.toko.size}, {h:'Kirim terakhir', num:true, val: o => o.last, html: o => dayLabel(B.days[o.last])}];
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], ['Distributor']]) +
    '<div class="dp-head"><div><div class="nm">Distributor &amp; Tujuan Kirim</div><div class="sub">' + fmt(list.length) + ' distributor · ' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + ' · ' + periodSub() + '</div></div></div>' +
    (unk > 0 ? '<div class="dp-review">' + pctTxt(pctOf(unk, tot)) + ' tonase periode ini belum punya data toko/gudang tujuan (' + fmt(unk) + ' t). Distributor tetap tercatat; toko tujuan akan muncul begitu kolom KODE_TOKO/NAMA_TOKO diisi di master Excel.</div>' : '') +
    table(cols, list, {go: o => '#/distributor/' + enc(B.dims.distr[o.key]), limit: 400}) +
    noteBasis('FOT diangkut sendiri oleh distributor (transportir = distributor). Klik distributor untuk melihat tujuan kirim per provinsi › distrik › toko.') + '</div>', 'wilayah');
}

function pageDistr(name){
  const di = B.dims.distr.indexOf(name);
  if(di < 0) return notFound('Distributor ' + name);
  const m = PER(), sc = {distr: di}, rows = shipRows(m, sc);
  const s = distrSummary(rows)[0] || {frc: 0, fot: 0, vol: 0, trips: 0, prov: new Set(), dist: new Set(), toko: new Set(), last: -1};
  const ekspTree = buildTree(rows, ['eksp']);
  // monthly volume, last 6 months; bar click lists that month's ship-to points
  const idx = D.months.indexOf(m.month), months = m.kind === 'month' ? D.months.slice(Math.max(0, idx - 5), idx + 1) : D.months.filter(x => x >= m.from.slice(0, 7) && x <= m.month);
  const all = shipRows(null, sc);
  const data = months.map(mm => ({label: monthShort(mm), long: monthLabel(mm), vol: all.filter(r => D.dayMonth[r[0]] === mm).reduce((a, r) => a + r[8], 0), tgt: null, pct: null, month: mm}));
  TRENDS.dm = {id: 'dm', data, click: i => {
    const mm = data[i].month, nodes = buildTree(shipRows(mm, sc), ['prov', 'dist', 'toko']);
    openModal(tc(name) + ' · ' + monthLabel(mm), 'Tujuan kirim per provinsi › distrik › toko', nodes.length ? treeHTML(nodes, ['prov', 'dist', 'toko']) : '<div class="dp-review">Tidak ada pengiriman di bulan ini.</div>');
  }};
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], ['Distributor', '#/distributor'], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">Distributor · ' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + ' · ' + periodSub() + '</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('box', 'var(--black)', 'Realisasi', fmt(s.vol) + ' <small>ton</small>', {hint: 'FRC ' + fmt(s.frc) + ' · FOT ' + fmt(s.fot)}) +
      kpi('trend', 'var(--black)', 'Trip', fmt(s.trips), {hint: s.last >= 0 ? 'terakhir ' + dayLabel(B.days[s.last]) : ''}) +
      kpi('pin', 'var(--amber)', 'Toko / gudang', fmt(s.toko.size), {hint: s.prov.size + ' provinsi · ' + s.dist.size + ' distrik'}) +
      kpi('truck', 'var(--black)', 'Ekspeditur', fmt(ekspTree.filter(n => B.dims.eksp[n.key]).length), {hint: s.fot ? 'termasuk angkut sendiri (FOT)' : ''}) +
    '</div>' +
    '<div class="dp-cols"><div class="dp-panel"><div class="pt">Volume bulanan (ton)</div>' + barsHTML('dm', data) + '<div class="trend-detail">Klik batang untuk tujuan kirim bulan itu.</div></div>' +
    '<div class="dp-panel"><div class="pt">Diangkut oleh</div>' + (ekspTree.length ? ekspTree.map(n => '<a class="rate-row" ' + (nodeHref(n) ? 'href="' + nodeHref(n) + '"' : '') + '><div class="rl">' + esc(B.dims.eksp[n.key] || 'FOT (angkut sendiri)') + '</div><div class="rtrack"><div class="rfill" style="width:' + (n.vol / ekspTree[0].vol * 100) + '%;background:var(--black)"></div></div><div class="rv">' + fmt(n.vol) + '</div></a>').join('') : '<div class="empty">Tidak ada pengiriman untuk filter ini.</div>') + '</div></div>' +
    '<div class="section-head"><div class="section-title">Tujuan pengiriman</div></div>' +
    (rows.length ? treeHTML(buildTree(rows, ['prov', 'dist', 'toko']), ['prov', 'dist', 'toko']) : '<div class="dp-review">Tidak ada pengiriman untuk filter ini.</div>') +
    noteBasis('Klik baris untuk membuka level berikutnya; tanda › membuka halaman provinsi/distrik.') + '</div>', 'wilayah');
}

/* ===================== Pages ===================== */
const view = () => $('#view');
function page(html, nav){ view().innerHTML = '<div class="page-fade">' + html + '</div>'; setActiveNav(nav); window.scrollTo(0, 0); }
function crumb(parts){ return '<div class="crumb"><a href="#/">Home</a>' + parts.map(p => '<span class="sep">/</span>' + (p[1] ? '<a href="' + p[1] + '">' + esc(p[0]) + '</a>' : esc(p[0]))).join('') + '</div>'; }
function periodSub(){ return PER().sub; }
function noteBasis(extra){
  return '<div class="note"><b>Sumber:</b> MASTER_DATA_OUTBOUND_LOGISTIC.xlsx, dibangun ' + esc(B.built) + ', data realisasi s.d. ' + dayLabel(B.asOf) + '. ' +
    '<b>Capaian</b> = realisasi FRC ÷ target SNOP MTD (jumlah target harian tgl 1 s.d. hari data terakhir). ' + (extra || '') + '</div>';
}

function pageHome(){
  const m = PER(), t = total(m);
  const pm = summarize('prov', m);
  const provRows = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0);
  const pp = prevPer(m), prev = pp.to >= B.days[0] && pp.from >= B.days[0] ? total(pp) : null;
  const chg = prev && prev.vol ? (t.vol / prev.vol - 1) * 100 : null;
  const isLatest = m.isLatest && m.kind === 'month';
  const incTxt = F.inc === 'ALL' ? 'FRC + FOT' : F.inc;
  const prevTxt = m.kind === 'month' ? 'periode sama bulan lalu' : 'periode ' + m.days + ' hari sebelumnya';
  const head = 'Volume ' + incTxt + ' ' + esc(m.label) + ' <b>' + fmt(t.vol) + ' ton</b>' +
    (t.pct != null ? (F.inc === 'ALL' ? '; FRC ' : ', ') + (isLatest ? 'sudah ' : '') + 'mencapai <b>' + pctTxt(t.pct) + '</b> dari target SNOP' + (isLatest ? ' MTD' : '') : '') +
    (chg != null ? ' — ' + (chg >= 0 ? 'naik ' : 'turun ') + Math.abs(Math.round(chg)) + '% dari ' + prevTxt + '.' : '.');
  const withPct = provRows.filter(o => o.pct != null);
  const worst = withPct.length ? withPct.reduce((a, b) => a.gap < b.gap ? a : b) : null;
  const attn = worst && worst.gap < 0 ? '<div class="attn"><div><span class="dot"></span></div><div><div class="t1">Perlu perhatian — ' + esc(tc(B.dims.prov[worst.key])) + '</div><div class="t2">Capaian ' + pctTxt(worst.pct) + ' dari target SNOP MTD — gap ' + fmt(-worst.gap) + ' ton, terbesar di antara semua provinsi. Realisasi ' + fmt(worst.frc) + ' t vs target ' + fmt(worst.tgt) + ' t.</div></div><a class="cta" href="#/provinsi/' + enc(B.dims.prov[worst.key]) + '">Lihat rincian &rsaquo;</a></div>' : '';
  const trucks = trucksFor(m);
  const prog = B.prog;
  page(
    mapHero(provRows, t, head, chg, isLatest) +
    '<div class="wrap"><div class="data-asof">Data realisasi s.d. ' + dayLabel(B.asOf) + (prog.date ? ' · snapshot Prognosa ' + dayLabel(prog.date) + ' ' + esc(prog.time) : '') + ' · ' + esc(B.build || '') + '</div>' +
    attn +
    '<div class="section-head"><div class="section-title">Capaian ' + esc(m.label) + '</div></div>' +
    '<div class="main-chart">' + trendPanel({}, 'harian', 'Realisasi vs target SNOP', {grans: ['harian', 'mingguan', 'bulanan'], noCmp: true}) + '</div>' +
    '<div class="section-head"><div class="section-title">Ringkasan operasional</div><a class="section-link" href="#/armada">Armada &rsaquo;</a></div>' +
    '<div class="kpi-row">' +
      kpi('trend', 'var(--black)', 'Trip', fmt(t.trips), {href:'#/armada', hint: incTxt + ' · ' + esc(m.label)}) +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {href:'#/armada', hint: fmt1(t.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata dwell', t.dwell == null ? '–' : fmt1(t.dwell) + ' <small>jam</small>', {href:'#/armada', hint:'masuk → keluar pabrik'}) +
      kpi('check', 'var(--green)', 'Realisasi / SO', pctTxt(t.of), {href:'#/ekspeditur', hint: t.so ? 'SO ' + fmt(t.so) + ' t' + ' (bulan penuh yang tercakup)' : 'SO tidak tersedia'}) +
    '</div>' +
    noteBasis('Peta: batas provinsi BAKOSURTANAL 1:250.000; provinsi abu-abu tidak dilayani via darat. Provinsi yang berkedip paling tertinggal.') + '</div>', 'home');
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
        chip('Target SNOP' + (isLatest ? ' MTD' : ''), t.hasTgt ? fmt(t.tgt) + ' <small>t</small>' : '–', t.hasTgt ? 'gap ' + signed(t.gap) + ' t' : '') +
        chip('Capaian', pctTxt(t.pct), 'FRC vs SNOP', t.pct == null ? null : t.pct >= 100 ? '#5FD08F' : t.pct >= 85 ? '#E8B04A' : '#FF6B74') +
        chip(m.kind === 'month' ? 'vs bulan lalu' : 'vs periode lalu', chg == null ? '–' : (chg >= 0 ? '▲ ' : '▼ ') + Math.abs(Math.round(chg)) + '%', 'periode sama', chg == null ? null : chg >= 0 ? '#5FD08F' : '#FF6B74') +
      '</div>' +
      '<div class="mh-list"><div class="mh-lh"><span>Rank provinsi</span><a href="#/provinsi">Lihat tabel &rsaquo;</a></div><div class="mh-row mh-head"><span></span><span></span><span class="mh-tn">Realisasi</span><span class="mh-pc">Target</span><span class="mh-so">Real/SO</span></div>' + list + '</div>' +
      '<div class="mh-key"><span><i style="background:#2E9E5C"></i>&ge;100%</span><span><i style="background:#C9952E"></i>85–99%</span><span><i style="background:#FF6B74"></i>&lt;85%</span><span><i style="background:#34424B"></i>tidak dilayani darat</span><span><i class="pin"></i>plant</span></div>' +
    '</div>' +
    '<div class="mh-map">' + mapHTML(provRows, {hero: true}) + '<div class="mh-hint">Arahkan kursor untuk angka · klik / ketuk provinsi untuk rincian</div></div>' +
  '</div></section>';
}

function pageProvList(){
  const rows = ranked(summarize('prov', PER()));
  const t = total(PER());
  const cols = [C.rank, {h:'Provinsi', cls:'name', val: r => tc(B.dims.prov[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.note];
  cols[2].total = fmt(t.vol); cols[3].total = fmt(t.tgt); cols[4].total = pctTxt(t.pct); cols[5].total = signed(t.gap); cols[6].total = fmt(t.so); cols[7].total = pctTxt(t.of); cols[8].total = fmt(t.trips);
  page('<div class="wrap">' + crumb([['Wilayah'], ['Semua provinsi']]) +
    '<div class="dp-head"><div><div class="nm">Semua Provinsi</div><div class="sub">Capaian tonase &amp; rank · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Total ' + pctTxt(t.pct) + ' dari target</div></div></div>' +
    table(cols, rows, {go: r => '#/provinsi/' + enc(B.dims.prov[r.key]), total: true}) +
    '<div class="dp-cols even">' + trendPanel({}, 'mingguan', 'Tren seluruh provinsi') + sourcePanel({}) + '</div>' +
    noteBasis() + '</div>', 'wilayah');
}
function sourcePanel(sc){
  const rows = [...summarize('src', PER(), sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  return '<div class="dp-panel"><div class="pt">Per source plant</div>' + (rows.length ? rows.map(o => rateRow(tc(B.dims.src[o.key]) + ' · ' + fmt(o.vol) + ' t', o.pct, null, o.hasTgt ? 'target ' + fmt(o.tgt) + ' t' : 'tanpa target')).join('') : '<div class="empty">Tidak ada data.</div>') +
    '<div class="trend-detail">Baris tanpa atribusi plant pada target dihitung di semua source. ' + (F.srcOff.size ? 'Dikecualikan: ' + [...F.srcOff].map(s => esc(tc(s))).join(', ') + '.' : '') + '</div></div>';
}
function progCells(rows, label){
  const s = k => rows.reduce((a, r) => a + (r[k] || 0), 0);
  const tgt = s('snopTarget');
  return '<div class="prog-row">' +
    '<div class="prog-cell"><div class="pl">Sudah rilis</div><div class="pv">' + fmt(s('sudahRilis')) + ' t</div><div class="pd">' + pctTxt(pctOf(s('sudahRilis'), tgt)) + ' dari target hari ini</div></div>' +
    '<div class="prog-cell"><div class="pl">Potensi</div><div class="pv">' + fmt(s('potensi')) + ' t</div><div class="pd">rilis + antri + timbang</div></div>' +
    '<div class="prog-cell"><div class="pl">Prognose</div><div class="pv" style="color:' + tierColor(pctOf(s('prognose'), tgt)) + '">' + fmt(s('prognose')) + ' t</div><div class="pd">' + pctTxt(pctOf(s('prognose'), tgt)) + ' · potensi + booking</div></div>' +
    '<div class="prog-cell"><div class="pl">Target SNOP hari ini</div><div class="pv">' + fmt(tgt) + ' t</div><div class="pd">' + esc(label) + '</div></div></div>';
}
function progRows(filter){
  const P = B.prog, fi = Object.fromEntries(P.fields.map((k, i) => [k, i + 3]));
  return P.rows.filter(r => r[2] !== 'TOTAL' && !F.srcOff.has(r[0]) && (!filter || filter(r))).map(r => {
    const o = {plant: r[0], inc: r[1], prov: r[2]};
    P.fields.forEach(k => o[k] = r[fi[k]]);
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
  const pr = progRows(r => r[2] === name && r[1] === 'FRC');
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.note];
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">Capaian tonase · ' + periodSub() + '</div></div><div class="badges">' + (me.rank ? '<div class="dp-badge">Rank ' + me.rank + ' / ' + all.rankTotal + '</div>' : '') + '<div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + (me.hasTgt ? ' Realisasi FRC ' + fmt(me.frc) + ' t vs target MTD ' + fmt(me.tgt) + ' t (gap ' + signed(me.gap) + ' t).' : '') + (me.so ? ' Realisasi/SO ' + pctTxt(me.of) + '.' : '') + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–', {hint: me.hasTgt ? 'bulan penuh ' + fmt(me.tgtFull) + ' t' : 'tidak ada target SNOP'}) +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', me.gap < 0 ? 'Sisa ke target' : 'Lebih dari target', me.hasTgt ? fmt(Math.abs(me.gap)) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>' +
    distMapSection(name, m) +
    '<div class="dp-cols">' + trendPanel(sc, 'mingguan') +
      '<div class="dp-panel"><div class="pt">Ekspeditur di provinsi ini<span class="pt-sub">capaian target · Real/SO</span></div>' + (eks.length ? eks.map(o => rateRow(B.dims.eksp[o.key] + ' · ' + fmt(o.vol) + ' t', o.pct, '#/ekspeditur/' + enc(B.dims.eksp[o.key]), o.hasTgt ? 'target ' + fmt(o.tgt) + ' t' : 'tanpa target', o.of)).join('') : '<div class="empty">' + (F.inc === 'FOT' ? 'FOT diambil langsung oleh distributor — lihat halaman FOT.' : 'Tidak ada data.') + '</div>') + '</div>' +
    '</div>' +
    (pr.length ? '<div class="section-head"><div class="section-title">Prognose hari ini (FRC)</div><a class="section-link" href="#/prognosa/hari-ini">Prognosa &rsaquo;</a></div>' + progCells(pr, 'snapshot ' + dayLabel(B.prog.date) + ' ' + B.prog.time) : '') +
    '<div class="section-head"><div class="section-title">Kinerja &amp; tujuan pengiriman per distrik</div></div>' +
    treeHTML(buildTree(shipRows(m, sc), ['dist', 'distr', 'toko']), ['dist', 'distr', 'toko'], {tgt: {dist: summarize('dist', m, sc)}}) +
    '<div class="dp-cols even">' + sourcePanel(sc) + '<div class="dp-panel"><div class="pt">Bulan penuh</div>' +
      '<div class="tgt-panel" style="margin:0;border:0;padding:0"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi FRC</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div></div></div>' +
    noteBasis('Distrik dengan target 0 dan realisasi 0 tidak ditampilkan.') + '</div>', 'wilayah');
}

function pageDist(name){
  const di = D.distIdx[name];
  if(di == null) return notFound('Distrik ' + name);
  const m = PER(), sc = {dist: di}, pname = B.dims.prov[B.dims.distProv[di]] || '';
  const me = total(m, sc);
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const ekCols = [{h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.dwell];
  const soh = [1, 2, 3].map(h => B.soh.filter(s => s[0] === h && s[2] === di && (F.inc === 'ALL' || s[3] === F.inc)).reduce((a, s) => a + s[4], 0));
  page('<div class="wrap">' + crumb([['Wilayah', '#/provinsi'], [tc(pname), '#/provinsi/' + enc(pname)], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">' + esc(tc(pname)) + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–') +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', 'Gap', me.hasTgt ? signed(me.gap) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>' +
    distMapSection(pname, m, di) +
    '<div class="dp-cols">' + trendPanel(sc, 'harian') + '<div class="dp-panel"><div class="pt">SO menunggu kirim (snapshot ' + esc(B.prog.date ? dayLabel(B.prog.date) + ' ' + B.prog.time : '–') + ')</div>' +
      ['H+1', 'H+2', 'H+3'].map((h, i) => '<div class="rate-row"><div class="rl">' + h + '</div><div class="rtrack"><div class="rfill" style="width:' + Math.min(100, soh[i] / Math.max(1, ...soh) * 100) + '%;background:var(--black)"></div></div><div class="rv">' + fmt(soh[i]) + ' t</div></div>').join('') + '</div></div>' +
    '<div class="section-head"><div class="section-title">Tujuan pengiriman: distributor &rsaquo; toko / gudang</div></div>' +
    treeHTML(buildTree(shipRows(m, sc), ['distr', 'toko']), ['distr', 'toko']) +
    '<div class="section-head"><div class="section-title">Ekspeditur di distrik ini</div></div>' +
    (eks.length ? table(ekCols, eks, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key]), sort: 1}) : '<div class="dp-review">Tidak ada ekspeditur FRC untuk filter ini.</div>') +
    noteBasis() + '</div>', 'wilayah');
}

function pageEkspList(){
  const m = PER(), t = total(m);
  const list = ranked(summarize('eksp', m)).filter(r => r.key !== D.eksEmpty);
  const tr = trucksFor(m);
  const perE = new Map(); tr.forEach(x => x.eksp.forEach((n, e) => { const o = perE.get(e) || {n: 0, one: 0}; o.n++; if(x.trips === 1) o.one++; perE.set(e, o); }));
  const multi = tr.length ? pctOf(tr.filter(x => x.trips > 1).length, tr.length) : null;
  const tk = r => (perE.get(r.key) || {n: 0, one: 0});
  const cols = [C.rank, {h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips,
    {h:'Truk', num:true, val: r => tk(r).n, html: r => fmt(tk(r).n)},
    {h:'Trip / truk', num:true, val: r => r.trips / Math.max(1, tk(r).n), html: r => fmt1(r.trips / Math.max(1, tk(r).n))},
    {h:'Truk 1 trip', num:true, val: r => tk(r).one}, C.dwell, C.note];
  page('<div class="wrap">' + crumb([['Ekspeditur']]) +
    '<div class="dp-head"><div><div class="nm">Peringkat Ekspeditur</div><div class="sub">Capaian target SNOP &amp; scorecard operasional · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Total ' + pctTxt(t.pct) + ' dari target</div></div></div>' +
    (F.inc === 'FOT' ? '<div class="dp-review">FOT diambil langsung oleh distributor (tanpa ekspeditur). Ganti filter ke FRC, atau buka <a href="#/distributor" style="color:var(--red);font-weight:700">Distributor FOT</a>.</div>' : '') +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Scorecard keseluruhan</div>' +
      rateRow('Capaian target SNOP', t.pct) + rateRow('Realisasi / SO', t.of) + rateRow('Truk aktif > 1 trip', multi) +
    '</div><div class="dp-panel"><div class="pt">Capaian per ekspeditur<span class="pt-sub">capaian target · Real/SO</span></div>' + (list.filter(r => r.pct != null).map(r => rateRow(B.dims.eksp[r.key], r.pct, '#/ekspeditur/' + enc(B.dims.eksp[r.key]), '', r.of)).join('') || '<div class="empty">Tidak ada target untuk filter ini.</div>') + '</div></div>' +
    table(cols, list, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key])}) +
    noteBasis('Real/SO = realisasi FRC ÷ SO pada periode yang sama (SO per ekspeditur = atribusi). SBA hanya memuat dari PP Belawan SBA, sehingga ikut tersembunyi saat source tersebut tidak dicentang.') + '</div>', 'ekspeditur');
}

function pageEksp(code){
  const ei = D.ekspIdx[code];
  if(ei == null) return notFound('Ekspeditur ' + code);
  const m = PER(), sc = {eksp: ei};
  const all = ranked(summarize('eksp', m));
  const me = all.find(o => o.key === ei) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
  const provs = [...summarize('prov', m, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  const dist = ranked(summarize('dist', m, sc));
  const trucks = trucksFor(m, sc).sort((a, b) => b.trips - a.trips);
  const served = new Set(B.targets.filter(t => t[0] >= m.from.slice(0, 7) && t[0] <= m.month && t[3] === ei).map(t => t[2]));
  const soh = [1, 2, 3].map(h => B.soh.filter(s => s[0] === h && served.has(s[2]) && s[3] === 'FRC').reduce((a, s) => a + s[4], 0));
  const one = trucks.filter(t => t.trips === 1).length;
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key]), html: r => esc(tc(B.dims.dist[r.key])) + '<small>' + esc(tc(B.dims.prov[B.dims.distProv[r.key]])) + '</small>'}, C.vol(), C.tgt, C.pct, C.gap, C.trips, C.note];
  const truckCols = [{h:'Nopol', cls:'name', val: r => B.dims.truck[r.truck]}, {h:'Trip', num:true, val: r => r.trips}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)},
    {h:'Hari aktif', num:true, val: r => r.days}, {h:'Dwell (jam)', num:true, val: r => r.dwN ? r.dwS / r.dwN : null, html: r => r.dwN ? fmt1(r.dwS / r.dwN) : '–'},
    {h:'Aktif', val: r => r.first, html: r => activeTxt(r)}, {h:'Provinsi utama', val: r => tc(B.dims.prov[r.prov])}];
  page('<div class="wrap">' + crumb([['Ekspeditur', '#/ekspeditur'], [code]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(code) + '</div><div class="sub">Mitra ekspeditur FRC · ' + periodSub() + '</div></div><div class="badges">' + (me.rank ? '<div class="dp-badge">Rank ' + me.rank + ' / ' + all.filter(o => o.key !== D.eksEmpty && o.rank).length + '</div>' : '') + '<div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + (me.hasTgt ? ' Realisasi ' + fmt(me.frc) + ' t vs target MTD ' + fmt(me.tgt) + ' t (gap ' + signed(me.gap) + ' t).' : '') + (one ? ' ' + one + ' truk baru 1 trip bulan ini.' : '') + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–', {hint: me.hasTgt ? 'bulan penuh ' + fmt(me.tgtFull) + ' t' : ''}) +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {hint: fmt1(me.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata dwell', me.dwell == null ? '–' : fmt1(me.dwell) + ' <small>jam</small>', {hint: me.so ? 'Real/SO ' + pctTxt(me.of) : ''}) +
    '</div>' +
    '<div class="dp-cols">' + trendPanel(sc, 'mingguan') +
      '<div class="dp-panel"><div class="pt">Capaian per provinsi<span class="pt-sub">capaian target · Real/SO</span></div>' + (provs.length ? provs.map(o => rateRow(tc(B.dims.prov[o.key]) + ' · ' + fmt(o.vol) + ' t', o.pct, '#/provinsi/' + enc(B.dims.prov[o.key]), o.hasTgt ? 'target ' + fmt(o.tgt) + ' t' : '', o.of)).join('') : '<div class="empty">Tidak ada data untuk filter ini.</div>') + '</div>' +
    '</div>' +
    '<div class="section-head"><div class="section-title">SO menunggu kirim di distrik layanan</div><a class="section-link" href="#/prognosa/hari-ini">Prognosa &rsaquo;</a></div>' +
    '<div class="prog-row">' + ['H+1', 'H+2', 'H+3'].map((h, i) => '<div class="prog-cell"><div class="pl">' + h + '</div><div class="pv">' + fmt(soh[i]) + ' t</div><div class="pd">SO FRC di ' + served.size + ' distrik ber-target ' + esc(code) + '</div></div>').join('') + '</div>' +
    '<div class="section-head"><div class="section-title">Target bulan berjalan</div></div>' +
    '<div class="tgt-panel"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Target MTD</div><div class="tv">' + (me.hasTgt ? fmt(me.tgt) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa ke target bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div>' +
    '<div class="section-head"><div class="section-title">Pengiriman per provinsi &rsaquo; distrik &rsaquo; distributor &rsaquo; toko</div></div>' +
    treeHTML(buildTree(shipRows(m, sc), ['prov', 'dist', 'distr', 'toko']), ['prov', 'dist', 'distr', 'toko'], {tgt: {prov: summarize('prov', m, sc), dist: summarize('dist', m, sc)}}) +
    '<div class="section-head"><div class="section-title">Truk ' + esc(code) + ' (' + fmt(trucks.length) + ')</div><a class="section-link" href="#/armada">Semua armada &rsaquo;</a></div>' +
    table(truckCols, trucks, {limit: 300}) +
    noteBasis('SO H+n tidak tercatat per ekspeditur di sumber; angka di atas adalah SO di distrik tempat ' + esc(code) + ' punya target SNOP.') + '</div>', 'ekspeditur');
}

function pageForecast(gran){
  gran = ['harian', 'mingguan', 'bulanan'].includes(gran) ? gran : 'harian';
  const chips = '<div class="gran-chips" style="margin-bottom:24px">' + [['harian','Hari ini'],['mingguan','Minggu ini'],['bulanan','Akhir bulan']].map(g => '<a href="#/prognosa/' + PROG_SLUG[g[0]] + '" class="' + (g[0] === gran ? 'on' : '') + '">' + g[1] + '</a>').join('') + '</div>';
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
        (hCols && hasH ? '<div class="section-head"><div class="section-title">Target SNOP vs SO, H+1 s.d. H+3</div></div>' + table(hCols, list, {sort: 1}) : '');
    }).join('');
    body += noteBasis('Rilis = sudah keluar pabrik; Potensi = rilis + antri + timbang; Prognose = potensi + booking (dari file monitoring, snapshot terakhir). Angka dalam kurung = jumlah truk.');
  } else if(gran === 'mingguan'){
    const end = parseISO(PER().to);
    const mon = new Date(end); mon.setDate(end.getDate() - ((end.getDay() + 6) % 7));
    const days = [...Array(7)].map((_, i) => { const d = new Date(mon); d.setDate(mon.getDate() + i); return d; });
    const data = days.map(d => {
      const s = iso(d), future = d > end, r = future ? rangeTotals(s, s) : rangeTotals(s, s);
      return {from: s, to: s, label: DAYS[d.getDay()].slice(0, 3) + ' ' + d.getDate(), long: DAYS[d.getDay()] + ', ' + dayLabel(s) + (future ? ' (belum terjadi)' : ''), vol: future ? 0 : r.vol, frc: future ? 0 : r.frc, tgt: r.pct == null && !future ? null : r.tgt, pct: future ? null : r.pct};
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
      '<div class="dp-panel" style="margin-bottom:34px"><div class="pt">Realisasi FRC per hari vs target SNOP, ' + days[0].getDate() + ' ' + MON3[days[0].getMonth()] + ' – ' + days[6].getDate() + ' ' + MON3[days[6].getMonth()] + '</div>' + barsHTML(id, data) +
      '<div class="trend-legend"><span><i style="background:var(--green)"></i>&ge;100%</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target SNOP</span></div><div class="trend-detail" id="tw-d">Klik batang untuk rincian hari itu per provinsi, distrik, ekspeditur &amp; source.</div></div>' +
      noteBasis('Proyeksi = realisasi minggu berjalan + rata-rata harian minggu ini × sisa hari. Bukan prognosa sistem.');
  } else {
    const m = PER().month, ld = lastDayOf(m), n = dim(m);
    const rows = [...summarize('prov', m).values()].filter(o => o.hasTgt || o.vol > 0).map(o => {
      const rr = o.frc / ld, proj = o.frc + rr * (n - ld);
      return Object.assign(o, {rr, proj, projPct: o.hasTgt ? pctOf(proj, o.tgtFull) : null});
    }).sort((a, b) => (a.projPct == null ? 999 : a.projPct) - (b.projPct == null ? 999 : b.projPct));
    const t = total(m), rr = t.frc / ld, proj = t.frc + rr * (n - ld);
    const cols = [{h:'Provinsi', cls:'name', val: r => tc(B.dims.prov[r.key])}, {h:'Realisasi FRC', num:true, val: r => r.frc, html: r => fmt(r.frc)},
      {h:'Laju / hari', num:true, val: r => r.rr, html: r => fmt(r.rr)}, {h:'Proyeksi akhir bulan', num:true, val: r => r.proj, html: r => fmt(r.proj)},
      {h:'Target bulan', num:true, val: r => r.hasTgt ? r.tgtFull : null, html: r => r.hasTgt ? fmt(r.tgtFull) : '–'},
      {h:'Proyeksi %', num:true, cls:'pc', val: r => r.projPct, html: r => pctTxt(r.projPct), style: r => 'color:' + tierColor(r.projPct)},
      {h:'Perlu / hari', num:true, val: r => r.hasTgt && n > ld ? (r.tgtFull - r.frc) / (n - ld) : null, html: r => r.hasTgt && n > ld ? fmt(Math.max(0, (r.tgtFull - r.frc) / (n - ld))) : '–'}];
    body = '<div class="prog-row">' +
      '<div class="prog-cell"><div class="pl">Realisasi FRC</div><div class="pv">' + fmt(t.frc) + ' t</div><div class="pd">tgl 1–' + ld + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Proyeksi akhir bulan</div><div class="pv" style="color:' + tierColor(pctOf(proj, t.tgtFull)) + '">' + fmt(proj) + ' t</div><div class="pd">laju ' + fmt(rr) + ' t/hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target bulan</div><div class="pv">' + fmt(t.tgtFull) + ' t</div><div class="pd">proyeksi ' + pctTxt(pctOf(proj, t.tgtFull)) + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Perlu per hari</div><div class="pv">' + (n > ld ? fmt(Math.max(0, (t.tgtFull - t.frc) / (n - ld))) : '–') + ' t</div><div class="pd">sisa ' + (n - ld) + ' hari</div></div></div>' +
      table(cols, rows, {go: r => '#/provinsi/' + enc(B.dims.prov[r.key])}) +
      noteBasis('Proyeksi = realisasi MTD ÷ hari berjalan × jumlah hari bulan ini (laju rata-rata). Untuk bulan yang sudah selesai, proyeksi = realisasi aktual.');
  }
  const pname = {harian: 'Hari ini', mingguan: 'Minggu ini', bulanan: 'Akhir bulan'}[gran];
  page('<div class="wrap">' + crumb([['Prognosa'], [pname]]) +
    '<div class="dp-head"><div><div class="nm">Prognosa — ' + pname + '</div><div class="sub">' + (gran === 'harian' ? 'Prognosa pengiriman hari ini dari file monitoring' : gran === 'mingguan' ? 'Minggu berjalan, Senin–Minggu' : 'Proyeksi akhir ' + monthLabel(PER().month)) + '</div></div></div>' +
    chips + body + '</div>', 'prognosa');
}

function pageArmada(){
  const m = PER(), t = total(m), trucks = trucksFor(m).sort((a, b) => b.trips - a.trips);
  const buckets = [['1', 1, 1], ['2–3', 2, 3], ['4–6', 4, 6], ['7–10', 7, 10], ['11–20', 11, 20], ['>20', 21, 1e9]];
  const bdata = buckets.map(b => ({label: b[0] + ' trip', long: b[0] + ' trip per truk', vol: trucks.filter(x => x.trips >= b[1] && x.trips <= b[2]).length, tgt: null, pct: null}));
  const dow = [1, 2, 3, 4, 5, 6, 0].map(d => {
    let s = 0, n = 0;
    B.facts.forEach(f => { if(inPer(f[0], m) && D.dayDow[f[0]] === d && incOk(f[1]) && srcOk(f[2])){ s += f[8]; n += f[9]; } });
    return {label: DAYS[d].slice(0, 3), long: DAYS[d], vol: n ? s / n : 0, tgt: null, pct: null, dec: true};
  });
  const dowNum = [1, 2, 3, 4, 5, 6, 0];
  dow.forEach(d => { d.unit = 'jam'; d.color = d.vol <= 3.5 ? 'var(--green)' : d.vol <= 4.5 ? 'var(--amber)' : 'var(--red)'; });
  TRENDS.ab = {id: 'ab', data: bdata, click: i => {
    const b = buckets[i], list = trucks.filter(x => x.trips >= b[1] && x.trips <= b[2]);
    openModal('Truk dengan ' + b[0] + ' trip', fmt(list.length) + ' truk · ' + esc(periodSub()), table(cols, list, {limit: 500}));
  }};
  TRENDS.ad = {id: 'ad', data: dow, click: i => {
    const wd = dowNum[i], per = new Map();
    B.facts.forEach(f => { if(inPer(f[0], m) && D.dayDow[f[0]] === wd && incOk(f[1]) && srcOk(f[2])){ const o = per.get(f[5]) || {key: f[5], s: 0, n: 0, trips: 0, vol: 0}; o.s += f[8]; o.n += f[9]; o.trips += f[7]; o.vol += f[6]; per.set(f[5], o); } });
    const rows = [...per.values()].filter(o => o.n).map(o => Object.assign(o, {dw: o.s / o.n}));
    openModal('Dwell hari ' + DAYS[wd], esc(periodSub()) + ' · per ekspeditur',
      table([{h:'Ekspeditur', cls:'name', val: o => B.dims.eksp[o.key] || 'FOT (distributor)'}, {h:'Rata-rata dwell (jam)', num:true, cls:'pc', val: o => o.dw, html: o => fmt1(o.dw), style: o => 'color:' + (o.dw <= 3.5 ? 'var(--green)' : o.dw <= 4.5 ? 'var(--amber)' : 'var(--red)')}, {h:'Trip', num:true, val: o => o.trips}, {h:'Tonase', num:true, val: o => o.vol, html: o => fmt(o.vol)}], rows, {go: o => B.dims.eksp[o.key] ? '#/ekspeditur/' + enc(B.dims.eksp[o.key]) : null, sort: 1}));
  }};
  const dwellBars = barsHTML('ad', dow);
  const one = trucks.filter(x => x.trips === 1).length;
  const cols = [{h:'#', cls:'rank', val: r => r.trips, html: (r, i) => '#' + (i + 1)}, {h:'Nopol', cls:'name', val: r => B.dims.truck[r.truck]},
    {h:'Ekspeditur', val: r => [...r.eksp.keys()].map(e => B.dims.eksp[e] || 'FOT').join(', ')}, {h:'Trip', num:true, val: r => r.trips}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)},
    {h:'Hari aktif', num:true, val: r => r.days}, {h:'Dwell (jam)', num:true, val: r => r.dwN ? r.dwS / r.dwN : null, html: r => r.dwN ? fmt1(r.dwS / r.dwN) : '–'},
    {h:'Aktif', val: r => r.first, html: r => activeTxt(r)}, {h:'Provinsi utama', val: r => tc(B.dims.prov[r.prov])}];
  page('<div class="wrap">' + crumb([['Armada']]) +
    '<div class="dp-head"><div><div class="nm">Armada</div><div class="sub">Truk aktif, trip dan dwell time · ' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge">' + fmt(trucks.length) + ' truk aktif</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {hint: one + ' truk baru 1 trip'}) +
      kpi('trend', 'var(--black)', 'Trip', fmt(t.trips), {hint: fmt1(t.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata dwell', t.dwell == null ? '–' : fmt1(t.dwell) + ' <small>jam</small>', {hint: 'masuk → keluar pabrik'}) +
      kpi('box', 'var(--black)', 'Ton / trip', fmt1(t.vol / Math.max(1, t.trips))) +
    '</div>' +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Sebaran trip per truk</div>' + barsHTML('ab', bdata) + '<div class="trend-detail" id="ab-d">Banyak truk 1 trip = armada tidak kontinu. Klik batang untuk daftar truknya.</div></div>' +
    '<div class="dp-panel"><div class="pt">Rata-rata dwell per hari (jam) — makin rendah makin baik</div>' + dwellBars + '<div class="trend-legend"><span><i style="background:var(--green)"></i>&le;3,5 jam</span><span><i style="background:var(--amber)"></i>3,5–4,5</span><span><i style="background:var(--red)"></i>&gt;4,5</span><span>klik batang = dwell per ekspeditur</span></div></div></div>' +
    '<div class="section-head"><div class="section-title">Daftar truk</div></div>' + table(cols, trucks, {limit: 400}) +
    noteBasis('Dwell = jam masuk s.d. jam keluar pabrik; baris yang ditandai CHECK (negatif / &gt;48 jam) di sumber tidak dihitung.') + '</div>', 'armada');
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
    else if(a === 'distrik' && b) pageDist(b);
    else if(a === 'ekspeditur') b ? pageEksp(b) : pageEkspList();
    else if(a === 'distributor') b ? pageDistr(b) : pageDistrList();
    else if(a === 'fot') location.replace('#/distributor');
    else if(a === 'armada') pageArmada();
    else if(a === 'prognosa') pageForecast(PROG_GRAN[b] || 'harian');
    // old addresses -> their single new home
    else if(a === 'forecast') location.replace('#/prognosa/' + PROG_SLUG[b in PROG_SLUG ? b : 'harian']);
    else if(a === 'incoterm') location.replace(b === 'FOT' ? '#/distributor' : '#/ekspeditur');
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
    const tgl = e.target.closest('[data-toggle]');
    if(tgl){
      const item = tgl.parentElement, was = item.classList.contains('open');
      closeNav();
      if(!was){
        closeFilter(); item.classList.add('open');
        const panel = item.querySelector('.dropdown,.mega,.filter-panel');   // on narrow screens panels are position:fixed under the header
        if(panel) panel.style.top = window.innerWidth <= 1100 ? ($('header').getBoundingClientRect().bottom + 4) + 'px' : '';
      }
      return;
    }
    const act = e.target.closest('[data-act]');
    if(act){
      const a = act.dataset.act;
      if(a === 'apply') applyFilter();
      if(a === 'reset'){ draft = {per: {kind: 'month', month: D.months[D.months.length - 1]}, inc: 'FRC', srcOff: new Set(DEFAULT_SRC_OFF)}; applyFilter(); }
      if(a === 'openfilter'){ e.stopPropagation(); openFilter(); }
      if(a === 'logout'){ logout(); }
      if(a === 'closemodal'){ closeModal(); }
      return;
    }
    const inc = e.target.closest('[data-inc]');
    if(inc){ draft.inc = inc.dataset.inc; $$('[data-inc]').forEach(x => x.classList.toggle('on', x === inc)); return; }
    const qm = e.target.closest('[data-qm]');
    if(qm){ const last = D.months[D.months.length - 1]; draft.per.month = qm.dataset.qm === 'cur' ? last : (D.months.includes(prevMonth(last)) ? prevMonth(last) : last); $('#fpMonth').value = draft.per.month; return; }
    const pm = e.target.closest('[data-pmode]');
    if(pm){ draft.per = {kind: pm.dataset.pmode, month: draft.per.month || draft.month}; renderFilterPanel(); return; }
    const qr = e.target.closest('[data-qr]');
    if(qr){ draft.to = B.asOf; draft.from = addDays(B.asOf, -(+qr.dataset.qr) + 1); $('#fpFrom').value = draft.from; $('#fpTo').value = draft.to; return; }
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
    const go = e.target.closest('[data-href]');
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
  document.addEventListener('keydown', e => { if(e.key === 'Escape'){ closeNav(); closeFilter(); closeModal(); $('#searchResults').classList.remove('show'); } });
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
    if(r) $$('.map-hero .map-region.live').forEach(x => { if(x.dataset.href === '#/provinsi/' + enc(r.dataset.prov)) x.classList.add('hl'); });
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
  window.addEventListener('hashchange', route);
}

/* ===================== Crypto, gate, session ===================== */
const IDLE_MS = 15 * 60 * 1000;
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
async function loadBundle(key, meta){
  const r = await fetch('data/site.enc?b=' + enc(meta.build), {cache: 'no-cache'});
  if(!r.ok) throw new Error('data/site.enc tidak ditemukan');
  const ct = await r.arrayBuffer();
  const gz = await crypto.subtle.decrypt({name: 'AES-GCM', iv: b64(meta.iv)}, key, ct);   // throws on wrong key
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
function logout(){ try { sessionStorage.removeItem('osp_k'); sessionStorage.removeItem('osp_t'); } catch(e) {} location.reload(); }
async function start(key, meta){
  B = await loadBundle(key, meta);
  B.build = meta.build;
  trimToYesterday(B);
  B.sod = B.sod || [];   // bundles built before SOCC-by-delivery-date
  B.sodMonths = new Set(B.sod.map(s => s[0].slice(0, 7)));
  prepare(); restoreFilter();
  try { sessionStorage.setItem('osp_k', toB64(await crypto.subtle.exportKey('raw', key))); } catch(e) {}
  touch();
  $('#gate').hidden = true; $('#app').hidden = false;
  buildNav(); buildSearch(); filterStrip();
  $('#foot').innerHTML = '<span>PT Semen Padang · Outbound Logistics · ' + esc(meta.build) + ' · data s.d. ' + dayLabel(B.asOf) + '</span><span>Sesi otomatis keluar setelah 15 menit tanpa aktivitas · <a data-act="logout">Keluar</a></span>';
  bindEvents(); route();
  ['click', 'keydown', 'scroll', 'mousemove', 'touchstart'].forEach(ev => window.addEventListener(ev, throttleTouch, {passive: true}));
  setInterval(() => { const t = +(sessionStorage.getItem('osp_t') || 0); if(Date.now() - t > IDLE_MS) logout(); }, 30000);
}
let lastTouch = 0;
function throttleTouch(){ const n = Date.now(); if(n - lastTouch > 5000){ lastTouch = n; touch(); } }

(async function init(){
  const msg = $('#gateMsg'), btn = $('#gateBtn');
  try {
    const k = sessionStorage.getItem('osp_k'), t = +(sessionStorage.getItem('osp_t') || 0);
    if(k && Date.now() - t < IDLE_MS){
      msg.style.color = 'var(--sub)'; msg.textContent = 'Memulihkan sesi…';
      const meta = await getMeta();
      const key = await crypto.subtle.importKey('raw', b64(k), {name: 'AES-GCM'}, true, ['decrypt']);
      await start(key, meta); return;
    }
  } catch(e) { try { sessionStorage.removeItem('osp_k'); } catch(_) {} msg.textContent = ''; }
  $('#gatePw').focus();
  $('#gateForm').addEventListener('submit', async e => {
    e.preventDefault();
    const pw = $('#gatePw').value; if(!pw) return;
    btn.disabled = true; msg.style.color = 'var(--sub)'; msg.textContent = 'Membuka data…';
    try {
      const meta = await getMeta();
      await start(await deriveKey(pw, meta), meta);
    } catch(err) {
      btn.disabled = false; msg.style.color = 'var(--red)';
      msg.textContent = err && err.name === 'OperationError' ? 'Kata sandi salah.' : 'Gagal memuat data: ' + (err.message || err);
    }
  });
})();
