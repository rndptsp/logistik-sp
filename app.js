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
function rateRow(label, p, href, sub){
  const w = p == null ? 0 : Math.min(100, p);
  const tag = href ? 'a' : 'div';
  return '<' + tag + ' class="rate-row"' + (href ? ' href="' + href + '"' : '') + ' title="' + esc(label + (sub ? ' — ' + sub : '')) + '"><div class="rl">' + esc(label) + '</div><div class="rtrack"><div class="rfill" style="width:' + w + '%;background:' + tierColor(p) + '"></div></div><div class="rv" style="color:' + tierColor(p) + '">' + pctTxt(p) + '</div></' + tag + '>';
}

/* ===================== State ===================== */
let B = null;            // decrypted bundle
let D = null;            // derived indexes
const F = { month: null, inc: 'FRC', srcOff: new Set(['PP BELAWAN SBA']) };
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
  F.month = months[months.length - 1];
}
const srcOk = i => i === D.srcEmpty || !F.srcOff.has(B.dims.src[i]);
const incOk = inc => F.inc === 'ALL' || (F.inc === 'FRC' ? inc === 0 : inc === 1);
const lastDayOf = m => Math.min(B.lastDay[m] || dim(m), dim(m));
const frcOn = () => F.inc !== 'FOT';

/* ===================== Aggregation core ===================== */
/* Group key extractors work on the shared column layout: facts/targets/so all carry prov,dist,eksp
   (facts: [day,inc,src,prov,dist,eksp,...]; targets: [month,prov,dist,eksp,src,...]; so: [month,prov,dist,eksp,ton]). */
const FK = {prov: f => f[3], dist: f => f[4], eksp: f => f[5], src: f => f[2], all: () => 0};
const TK = {prov: t => t[1], dist: t => t[2], eksp: t => t[3], src: t => t[4], all: () => 0};
const SK = {prov: s => s[1], dist: s => s[2], eksp: s => s[3], src: () => null, all: () => 0};
function scopeOkF(f, sc){ return (sc.prov == null || f[3] === sc.prov) && (sc.dist == null || f[4] === sc.dist) && (sc.eksp == null || f[5] === sc.eksp); }
function scopeOkT(t, sc){ return (sc.prov == null || t[1] === sc.prov) && (sc.dist == null || t[2] === sc.dist) && (sc.eksp == null || t[3] === sc.eksp); }

function targetMTD(t, m, toDay){
  const ld = toDay == null ? lastDayOf(m) : toDay;
  if(t[6]) { let s = 0; for(let d = 0; d < ld; d++) s += t[6][d] || 0; return s; }
  return t[5] * ld / dim(m);
}
function blank(){ return {vol:0, frc:0, trips:0, dwS:0, dwN:0, tgt:0, tgtFull:0, so:0, hasTgt:false}; }

/** Summary per group for one month. vol = selected incoterm(s); frc = FRC-only realisasi (achievement basis). */
function summarize(group, m, sc, opt){
  sc = sc || {}; opt = opt || {};
  const out = new Map();
  const get = k => { let o = out.get(k); if(!o){ o = blank(); o.key = k; out.set(k, o); } return o; };
  const fromDay = opt.fromDay || 1, toDay = opt.toDay || 31;
  B.facts.forEach(f => {
    if(D.dayMonth[f[0]] !== m || !srcOk(f[2]) || !scopeOkF(f, sc)) return;
    const dn = D.dayNum[f[0]]; if(dn < fromDay || dn > toDay) return;
    const o = get(FK[group](f));
    if(f[1] === 0) o.frc += f[6];
    if(incOk(f[1])){ o.vol += f[6]; o.trips += f[7]; o.dwS += f[8]; o.dwN += f[9]; }
  });
  if(frcOn()){
    B.targets.forEach(t => {
      if(t[0] !== m || !srcOk(t[4]) || !scopeOkT(t, sc)) return;
      const o = get(TK[group](t));
      o.tgt += targetMTD(t, m, opt.toDay); o.tgtFull += t[5]; o.hasTgt = true;
    });
    if(group !== 'src') B.so.forEach(s => {
      if(s[0] !== m || !scopeOkT(s, sc)) return;
      get(SK[group](s)).so += s[4];
    });
  }
  out.forEach(o => {
    o.pct = o.hasTgt && frcOn() ? pctOf(o.frc, o.tgt) : null;
    o.of = frcOn() && o.so > 0 ? pctOf(o.frc, o.so) : null;
    o.gap = o.frc - o.tgt;
    o.dwell = o.dwN ? o.dwS / o.dwN : null;
  });
  return out;
}
const total = (m, sc, opt) => summarize('all', m, sc, opt).get(0) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
function ranked(map){
  const arr = [...map.values()].filter(o => o.vol > 0 || o.tgt > 0);
  arr.sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  let r = 0; arr.forEach(o => { o.rank = o.pct == null ? null : ++r; });
  arr.rankTotal = r;
  return arr;
}
function trucksFor(m, sc){
  sc = sc || {};
  const per = new Map();
  B.trucks.forEach(t => {
    if(t[0] !== m || !incOk(t[2]) || !srcOk(t[4])) return;
    if(sc.eksp != null && t[3] !== sc.eksp) return;
    if(sc.prov != null && t[12] !== sc.prov) return;
    let o = per.get(t[1]);
    if(!o){ o = {truck:t[1], trips:0, ton:0, dwS:0, dwN:0, first:99, last:0, days:0, eksp:new Map(), prov:t[12], inc:t[2]}; per.set(t[1], o); }
    o.trips += t[5]; o.ton += t[6]; o.dwS += t[7]; o.dwN += t[8]; o.first = Math.min(o.first, t[9]); o.last = Math.max(o.last, t[10]);
    o.days = Math.max(o.days, t[11]); o.eksp.set(t[3], (o.eksp.get(t[3]) || 0) + t[5]);
  });
  return [...per.values()];
}
function prevMonth(m){ const y = +m.slice(0, 4), mo = +m.slice(5); return mo === 1 ? (y - 1) + '-12' : y + '-' + String(mo - 1).padStart(2, '0'); }

/* Range (date-based) realisasi + target, for weekly trends and weekly forecast. */
function rangeTotals(fromISO, toISO, sc){
  sc = sc || {};
  let vol = 0, frc = 0, tgt = 0, hasTgt = false;
  B.facts.forEach(f => {
    const ds = B.days[f[0]];
    if(ds < fromISO || ds > toISO || !srcOk(f[2]) || !scopeOkF(f, sc)) return;
    if(f[1] === 0) frc += f[6];
    if(incOk(f[1])) vol += f[6];
  });
  if(frcOn()){
    const from = parseISO(fromISO), to = parseISO(toISO);
    B.targets.forEach(t => {
      if(!srcOk(t[4]) || !scopeOkT(t, sc)) return;
      const m = t[0], y = +m.slice(0, 4), mo = +m.slice(5) - 1, n = dim(m);
      for(let d = 1; d <= n; d++){
        const dt = new Date(y, mo, d);
        if(dt < from || dt > to) continue;
        tgt += t[6] ? (t[6][d - 1] || 0) : t[5] / n; hasTgt = true;
      }
    });
  }
  return {vol, frc, tgt, pct: hasTgt && frcOn() ? pctOf(frc, tgt) : null};
}

/* ===================== Trend panel (weekly / monthly) ===================== */
function trendData(sc, gran, n){
  const asOf = B.asOf;
  if(gran === 'bulanan'){
    const idx = D.months.indexOf(F.month);
    const ms = D.months.slice(Math.max(0, idx - (n || 6) + 1), idx + 1);
    return ms.map(m => { const t = total(m, sc); return {label: monthShort(m), long: monthLabel(m) + (m === D.months[D.months.length - 1] ? ' (MTD s.d. tgl ' + lastDayOf(m) + ')' : ''), vol: t.vol, frc: t.frc, tgt: t.hasTgt ? t.tgt : null, pct: t.pct}; });
  }
  if(gran === 'harian'){
    const m = F.month, ld = lastDayOf(m), out = [];
    for(let d = Math.max(1, ld - (n || 14) + 1); d <= ld; d++){
      const ds = m + '-' + String(d).padStart(2, '0');
      const r = rangeTotals(ds, ds, sc);
      out.push({label: String(d), long: DAYS[parseISO(ds).getDay()] + ', ' + dayLabel(ds), vol: r.vol, frc: r.frc, tgt: r.pct == null ? null : r.tgt, pct: r.pct});
    }
    return out;
  }
  // mingguan: Monday–Sunday weeks ending at the selected month's last data day
  const end = parseISO(F.month + '-' + String(lastDayOf(F.month)).padStart(2, '0'));
  const out = [];
  const monday = new Date(end); monday.setDate(end.getDate() - ((end.getDay() + 6) % 7));
  for(let w = (n || 8) - 1; w >= 0; w--){
    const a = new Date(monday); a.setDate(monday.getDate() - 7 * w);
    const b = new Date(a); b.setDate(a.getDate() + 6);
    const bb = b > end ? end : b;
    const r = rangeTotals(iso(a), iso(bb), sc);
    out.push({label: a.getDate() + '/' + (a.getMonth() + 1), long: a.getDate() + ' ' + MON3[a.getMonth()] + ' – ' + bb.getDate() + ' ' + MON3[bb.getMonth()] + (bb < b ? ' (berjalan)' : ''), vol: r.vol, frc: r.frc, tgt: r.pct == null ? null : r.tgt, pct: r.pct});
  }
  return out;
}
const TRENDS = {};
function barsHTML(id, data){
  const max = Math.max(1, ...data.map(d => Math.max(d.vol, d.tgt || 0)));
  return '<div class="bars">' + data.map((d, i) => {
    const h = Math.max(2, d.vol / max * 100);
    const th = d.tgt != null ? Math.min(100, d.tgt / max * 100) : null;
    const c = d.pct != null ? tierColor(d.pct) : 'var(--black)';
    return '<div class="bcol click" data-bar="' + id + ':' + i + '">' +
      (d.pct != null ? '<div class="bpct" style="color:' + c + '">' + Math.round(d.pct) + '%</div>' : '') +
      '<div class="bval">' + fmt(d.vol) + '</div>' +
      '<div class="bwrap"><div class="b" style="height:' + h + '%;background:' + c + ';animation-delay:' + (i * 50) + 'ms"></div>' +
      (th != null ? '<div class="tick" style="bottom:' + th + '%"></div>' : '') + '</div>' +
      '<div class="blabel">' + esc(d.label) + '</div></div>';
  }).join('') + '</div>';
}
function trendPanel(sc, gran, title){
  const id = 't' + Math.random().toString(36).slice(2, 8);
  TRENDS[id] = {sc, gran: gran || 'mingguan', title: title || 'Tren realisasi'};
  return '<div class="dp-panel" id="' + id + '">' + trendInner(id) + '</div>';
}
function trendInner(id){
  const T = TRENDS[id], data = trendData(T.sc, T.gran);
  T.data = data;
  const chips = [['harian','Harian'],['mingguan','Mingguan'],['bulanan','Bulanan']].map(g => '<span class="' + (g[0] === T.gran ? 'on' : '') + '" data-gran="' + id + ':' + g[0] + '">' + g[1] + '</span>').join('');
  return '<div class="pt">' + esc(T.title) + ' (ton)<div class="gran-chips">' + chips + '</div></div>' + barsHTML(id, data) +
    '<div class="trend-legend">' + (frcOn() ? '<span><i style="background:var(--green)"></i>&ge;100% target</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target SNOP</span>' : '<span><i style="background:var(--black)"></i>volume FOT (tanpa target)</span>') + '</div>' +
    '<div class="trend-detail" id="' + id + '-d">Klik salah satu batang untuk lihat rincian periode.</div>';
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
  of: {h:'Real/SO', num:true, val: r => r.of, html: r => pctTxt(r.of), style: r => 'color:' + tierColor(r.of)},
  trips: {h:'Trip', num:true, val: r => r.trips, html: r => fmt(r.trips)},
  dwell: {h:'Dwell (jam)', num:true, val: r => r.dwell, html: r => r.dwell == null ? '–' : fmt1(r.dwell)},
  note: {h:'Catatan', cls:'note-c', val: r => reviewText(r.pct), sortVal: r => r.pct}
};

/* ===================== Shell: nav, filter, search ===================== */
function setActiveNav(key){ $$('#mainNav .navitem').forEach(n => n.classList.toggle('on', n.dataset.nav === key)); }
function closeNav(){ $$('#mainNav .navitem').forEach(n => n.classList.remove('open')); }
function buildNav(){
  const pm = summarize('prov', F.month);
  const provs = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  $('#navProvList').innerHTML = provs.map(o => '<a href="#/provinsi/' + enc(B.dims.prov[o.key]) + '">' + esc(tc(B.dims.prov[o.key])) + '<span>' + pctTxt(o.pct) + '</span></a>').join('') || '<span class="empty">—</span>';
  const em = summarize('eksp', F.month);
  $('#navEkspList').innerHTML = D.ekspList.map(code => {
    const o = em.get(D.ekspIdx[code]);
    const p = o ? o.pct : null;
    return '<a class="eksp-item" href="#/ekspeditur/' + enc(code) + '"><span class="en">' + esc(code) + '</span><span class="ep" style="color:' + tierColor(p) + '">' + (o && o.hasTgt ? pctTxt(p) : (o && o.vol ? fmt(o.vol) + ' t' : '–')) + '</span></a>';
  }).join('');
  $('#dayLabel').textContent = 'Data s.d. ' + dayLabel(B.asOf);
}
let draft = null;
function buildFilterPanel(){
  draft = {month: F.month, inc: F.inc, srcOff: new Set(F.srcOff)};
  const monthOpts = D.months.slice().reverse().map(m => '<option value="' + m + '"' + (m === draft.month ? ' selected' : '') + '>' + monthLabel(m) + (m === D.months[D.months.length - 1] ? ' (berjalan)' : '') + '</option>').join('');
  const inc = [['FRC','FRC'],['FOT','FOT'],['ALL','FRC + FOT']].map(x => '<span data-inc="' + x[0] + '" class="' + (draft.inc === x[0] ? 'on' : '') + '">' + x[1] + '</span>').join('');
  const quick = '<span data-qm="cur">Bulan ini</span><span data-qm="prev">Bulan lalu</span>';
  const src = D.sources.map(s => '<label class="fp-check"><input type="checkbox" data-src="' + esc(s) + '"' + (draft.srcOff.has(s) ? '' : ' checked') + '> ' + esc(tc(s)) + (DEFAULT_SRC_OFF.includes(s) ? ' <small>(default tidak dicentang)</small>' : '') + '</label>').join('');
  $('#filterPanel').innerHTML =
    '<div class="fp-row"><div class="fp-group"><label class="fl">Periode</label><select class="fp-select" id="fpMonth">' + monthOpts + '</select><div class="fp-chips" style="margin-top:8px">' + quick + '</div></div>' +
    '<div class="fp-group"><label class="fl">Incoterm</label><div class="fp-chips">' + inc + '</div><div style="font-size:11.5px;color:var(--sub);margin-top:8px;line-height:1.5">% pencapaian selalu dihitung dari FRC vs target SNOP. FOT tidak punya target.</div></div></div>' +
    '<div class="fp-row"><div class="fp-group"><label class="fl">Source plant</label>' + src + '</div>' +
    '<div class="fp-group"><label class="fl">Lompat ke</label><select class="fp-select" id="fpJump"><option value="">— provinsi / ekspeditur —</option>' +
      [...D.liveProv].map(i => B.dims.prov[i]).sort().map(p => '<option value="#/provinsi/' + enc(p) + '">Provinsi: ' + esc(tc(p)) + '</option>').join('') +
      D.ekspList.map(e => '<option value="#/ekspeditur/' + enc(e) + '">Ekspeditur: ' + esc(e) + '</option>').join('') + '</select></div></div>' +
    '<div class="fp-foot"><div class="fp-reset" data-act="reset">Reset filter</div><div class="fp-apply" data-act="apply">Terapkan</div></div>';
}
function filterStrip(){
  const offs = [...F.srcOff];
  const isLatest = F.month === D.months[D.months.length - 1];
  $('#filterStrip').innerHTML = '<span>Menampilkan</span><span class="tag"><b>' + monthLabel(F.month) + '</b>' + (isLatest ? ' · MTD s.d. ' + lastDayOf(F.month) + ' ' + MON3[+F.month.slice(5) - 1] : '') + '</span>' +
    '<span class="tag"><b>' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + '</b></span>' +
    (offs.length ? '<span class="tag">tanpa ' + offs.map(s => esc(tc(s))).join(', ') + '</span>' : '<span class="tag">semua source</span>') +
    '<span class="edit" data-act="openfilter">Ubah filter</span>';
}
function applyFilter(){
  F.month = draft.month; F.inc = draft.inc; F.srcOff = new Set(draft.srcOff);
  try { sessionStorage.setItem('osp_f', JSON.stringify({month: F.month, inc: F.inc, srcOff: [...F.srcOff]})); } catch(e) {}
  closeNav(); buildNav(); filterStrip(); route();
}
function restoreFilter(){
  try {
    const s = JSON.parse(sessionStorage.getItem('osp_f') || 'null');
    if(s && D.months.includes(s.month)){ F.month = s.month; F.inc = s.inc; F.srcOff = new Set(s.srcOff); }
  } catch(e) {}
}

let searchIndex = [];
function buildSearch(){
  const P = B.dims.prov, Ds = B.dims.dist;
  searchIndex = [];
  [...D.liveProv].forEach(i => searchIndex.push({t: tc(P[i]), k: 'Provinsi', h: '#/provinsi/' + enc(P[i])}));
  [...D.liveDist].forEach(i => searchIndex.push({t: tc(Ds[i]), k: tc(P[B.dims.distProv[i]] || ''), h: '#/distrik/' + enc(Ds[i])}));
  D.ekspList.forEach(e => searchIndex.push({t: e, k: 'Ekspeditur', h: '#/ekspeditur/' + enc(e)}));
  [['Forecast harian','#/forecast/harian'],['Forecast mingguan','#/forecast/mingguan'],['Forecast bulanan (proyeksi)','#/forecast/bulanan'],['Tren tonase','#/tren'],['Armada / truk','#/armada'],['Scorecard ekspeditur','#/scorecard'],['FRC','#/incoterm/FRC'],['FOT','#/incoterm/FOT'],['Semua provinsi','#/provinsi'],['Semua ekspeditur','#/ekspeditur']]
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
const MAP = [
  {id:'aceh', label:'Aceh', x:85, y:5, w:90, h:50},
  {id:'SUMATERA UTARA', label:'Sum. Utara', x:65, y:60, w:115, h:65},
  {id:'SUMATERA BARAT', label:'Sum. Barat', x:45, y:130, w:70, h:65},
  {id:'RIAU DARATAN', label:'Riau', x:120, y:130, w:95, h:55},
  {id:'kepri', label:'Kep. Riau', x:195, y:140, w:55, h:35},
  {id:'JAMBI', label:'Jambi', x:80, y:200, w:100, h:50},
  {id:'SUMATERA SELATAN', label:'Sum. Selatan', x:75, y:255, w:105, h:55},
  {id:'BENGKULU', label:'Bengkulu', x:25, y:255, w:45, h:80},
  {id:'babel', label:'Babel', x:185, y:270, w:55, h:45},
  {id:'lampung', label:'Lampung', x:85, y:315, w:90, h:50}
];
function mapHTML(provRows){
  const byName = Object.fromEntries(provRows.map(o => [B.dims.prov[o.key], o]));
  const withPct = provRows.filter(o => o.pct != null);
  const worst = withPct.length ? withPct.reduce((a, b) => a.pct < b.pct ? a : b) : null;
  const defs = '<defs>' +
    '<linearGradient id="gGreen" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2E9E5C"/><stop offset="100%" stop-color="#1A7A42"/></linearGradient>' +
    '<linearGradient id="gAmber" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#C08A2A"/><stop offset="100%" stop-color="#8A5A00"/></linearGradient>' +
    '<linearGradient id="gRed" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FF5A64"/><stop offset="100%" stop-color="#F5333F"/></linearGradient>' +
    '<linearGradient id="gInk" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#5A5A5A"/><stop offset="100%" stop-color="#2B2B2B"/></linearGradient>' +
    '<filter id="softShadow" x="-40%" y="-40%" width="180%" height="180%"><feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="#000" flood-opacity="0.22"/></filter></defs>';
  const island = '<path d="M60,8 C95,-2 135,8 155,30 C178,54 172,88 198,118 C222,144 216,176 202,212 C188,246 192,272 178,302 C168,326 152,348 132,372 C112,392 90,406 68,396 C48,386 54,360 44,334 C28,298 34,258 24,224 C14,188 26,148 20,114 C14,78 30,44 60,8 Z" fill="#E6EBEB" stroke="#D3D8D9" stroke-width="1.5"/>';
  const regions = MAP.map(p => {
    const o = byName[p.id];
    const cx = p.x + p.w / 2, cy = p.y + p.h / 2;
    if(!o){
      return '<rect class="map-region" x="' + p.x + '" y="' + p.y + '" width="' + p.w + '" height="' + p.h + '" rx="16" fill="#D7DCDD" data-tip="' + esc(p.label + '|Tidak ada pengiriman darat di data') + '"></rect>' +
        '<text x="' + cx + '" y="' + (cy + 3) + '" text-anchor="middle" font-size="10" font-weight="700" fill="#8C9293" style="pointer-events:none">' + p.label + '</text>';
    }
    const fill = o.pct != null ? 'url(#g' + tierName(o.pct) + ')' : 'url(#gInk)';
    const tip = tc(p.id) + '|Realisasi ' + fmt(o.vol) + ' t' + (o.hasTgt ? ' · Target MTD ' + fmt(o.tgt) + ' t' : '') + '|' + (o.pct != null ? 'Capaian ' + pctTxt(o.pct) : 'Tanpa target SNOP');
    return '<rect class="map-region live' + (worst && o === worst ? ' is-attn' : '') + '" x="' + p.x + '" y="' + p.y + '" width="' + p.w + '" height="' + p.h + '" rx="16" fill="' + fill + '" filter="url(#softShadow)" data-href="#/provinsi/' + enc(p.id) + '" data-tip="' + esc(tip) + '"></rect>' +
      '<text x="' + cx + '" y="' + (cy - 3) + '" text-anchor="middle" font-size="' + Math.min(10.5, p.w / 5.6).toFixed(1) + '" font-weight="700" fill="#fff" style="pointer-events:none">' + p.label + '</text>' +
      '<text x="' + cx + '" y="' + (cy + 11) + '" text-anchor="middle" font-size="11" font-weight="800" fill="#fff" style="pointer-events:none">' + (o.pct != null ? pctTxt(o.pct) : fmt(o.vol) + ' t') + '</text>';
  }).join('');
  const sorted = provRows.slice().sort((a, b) => (b.pct == null ? -1 : b.pct) - (a.pct == null ? -1 : a.pct) || b.vol - a.vol);
  const legend = sorted.map((o, i) => {
    const name = B.dims.prov[o.key];
    return '<a class="map-row" href="#/provinsi/' + enc(name) + '"><span class="dot" style="background:' + tierColor(o.pct) + '"></span><span class="nm">#' + (i + 1) + ' ' + esc(tc(name)) + '</span><span class="tn">' + fmt(o.vol) + ' t</span><span class="pc" style="color:' + tierColor(o.pct) + '">' + pctTxt(o.pct) + '</span></a>';
  }).join('') + '<div class="map-row muted"><span class="dot" style="background:#D7DCDD"></span><span class="nm">Aceh, Kep. Riau, Babel, Lampung</span><span class="pc">—</span></div>';
  return '<div class="map-wrap"><svg class="map-svg" viewBox="0 0 260 420" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta pencapaian per provinsi">' + defs + island + regions + '</svg><div class="map-legend">' + legend + '</div></div>';
}

/* ===================== Pages ===================== */
const view = () => $('#view');
function page(html, nav){ view().innerHTML = '<div class="page-fade">' + html + '</div>'; setActiveNav(nav); window.scrollTo(0, 0); }
function crumb(parts){ return '<div class="crumb"><a href="#/">Home</a>' + parts.map(p => '<span class="sep">/</span>' + (p[1] ? '<a href="' + p[1] + '">' + esc(p[0]) + '</a>' : esc(p[0]))).join('') + '</div>'; }
function periodSub(){ const latest = F.month === D.months[D.months.length - 1]; return monthLabel(F.month) + (latest ? ' · MTD s.d. ' + lastDayOf(F.month) + ' ' + MON3[+F.month.slice(5) - 1] : ''); }
function noteBasis(extra){
  return '<div class="note"><b>Sumber:</b> MASTER_DATA_OUTBOUND_LOGISTIC.xlsx, dibangun ' + esc(B.built) + ', data realisasi s.d. ' + dayLabel(B.asOf) + '. ' +
    '<b>Capaian</b> = realisasi FRC ÷ target SNOP MTD (jumlah target harian tgl 1 s.d. hari data terakhir). ' + (extra || '') + '</div>';
}

function pageHome(){
  const m = F.month, t = total(m);
  const pm = summarize('prov', m);
  const provRows = [...pm.values()].filter(o => o.vol > 0 || o.tgt > 0);
  const ld = lastDayOf(m), pmn = prevMonth(m);
  const prev = D.months.includes(pmn) ? total(pmn, {}, {toDay: ld}) : null;
  const chg = prev && prev.vol ? (t.vol / prev.vol - 1) * 100 : null;
  const isLatest = m === D.months[D.months.length - 1];
  const incTxt = F.inc === 'ALL' ? 'FRC + FOT' : F.inc;
  const head = 'Volume ' + incTxt + ' ' + monthLabel(m) + ' <b>' + fmt(t.vol) + ' ton</b>' +
    (t.pct != null ? (F.inc === 'ALL' ? '; FRC ' : ', ') + (isLatest ? 'sudah ' : '') + 'mencapai <b>' + pctTxt(t.pct) + '</b> dari target SNOP' + (isLatest ? ' MTD' : '') : '') +
    (chg != null ? ' — ' + (chg >= 0 ? 'naik ' : 'turun ') + Math.abs(Math.round(chg)) + '% dari periode sama bulan lalu.' : '.');
  const withPct = provRows.filter(o => o.pct != null);
  const worst = withPct.length ? withPct.reduce((a, b) => a.gap < b.gap ? a : b) : null;
  const attn = worst && worst.gap < 0 ? '<div class="attn"><div><span class="dot"></span></div><div><div class="t1">Perlu perhatian — ' + esc(tc(B.dims.prov[worst.key])) + '</div><div class="t2">Capaian ' + pctTxt(worst.pct) + ' dari target SNOP MTD — gap ' + fmt(-worst.gap) + ' ton, terbesar di antara semua provinsi. Realisasi ' + fmt(worst.frc) + ' t vs target ' + fmt(worst.tgt) + ' t.</div></div><a class="cta" href="#/provinsi/' + enc(B.dims.prov[worst.key]) + '">Lihat rincian &rsaquo;</a></div>' : '';
  const cards = provRows.slice().sort((a, b) => b.vol - a.vol).map(o => {
    const name = B.dims.prov[o.key], bad = o.pct != null && o.pct < 85;
    return '<a class="pcard" style="border-left-color:' + tierColor(o.pct) + '" href="#/provinsi/' + enc(name) + '"><div class="nm">' + esc(tc(name)) + (o.pct != null ? '<span>gap ' + signed(o.gap) + ' t</span>' : '') + '</div><div class="pc' + (bad ? ' bad' : '') + '">' + (o.pct != null ? pctTxt(o.pct) : fmt(o.vol) + ' t') + '</div><div class="tn">' + fmt(o.vol) + ' ton' + (o.hasTgt ? ' dari target ' + fmt(o.tgt) + ' t' : ' · tanpa target') + '</div><div class="track"><div class="fill' + (bad ? ' bad' : '') + '" style="width:' + (o.pct != null ? Math.min(100, o.pct) : 100) + '%"></div></div></a>';
  }).join('');
  const trucks = trucksFor(m);
  const prog = B.prog;
  page(
    '<div class="hero"><img src="assets/hero-logistik.jpg" alt=""><div class="hero-grad"></div><div class="hero-text"><div class="eyebrow">Ringkasan ' + monthLabel(m) + (isLatest ? ' — bulan berjalan' : '') + '</div><h1>' + head + '</h1></div><div class="hero-arrow"></div></div>' +
    '<div class="wrap"><div class="data-asof">Data realisasi s.d. ' + dayLabel(B.asOf) + (prog.date ? ' · snapshot Prognosa ' + dayLabel(prog.date) + ' ' + esc(prog.time) : '') + ' · ' + esc(B.build || '') + '</div>' +
    attn +
    '<div class="section-head"><div class="section-title">Peta Interaktif Sumatera</div><a class="section-link" href="#/provinsi">Lihat semua &rsaquo;</a></div>' +
    mapHTML(provRows) +
    '<div class="section-sub">Siluet blok skematik per provinsi (bukan batas geografis presisi), diwarnai menurut capaian FRC vs target SNOP. Provinsi yang berkedip merah paling tertinggal. Arahkan kursor untuk angka, klik untuk rincian.</div>' +
    '<div class="section-head"><div class="section-title">Pencapaian per provinsi</div><a class="section-link" href="#/provinsi">Lihat tabel &rsaquo;</a></div>' +
    '<div class="card-row">' + cards + '</div>' +
    '<div class="section-head"><div class="section-title">Jelajahi</div></div>' +
    '<div class="quick-row">' +
      '<a class="qcard" href="#/forecast/harian"><div class="photo p1"></div><div class="body"><div class="qt">Forecast</div><div class="qd">Rilis, potensi dan prognose hari ini; SO H+1 sampai H+3; proyeksi akhir bulan.</div><div class="qlink">Lihat forecast &rsaquo;</div></div></a>' +
      '<a class="qcard" href="#/ekspeditur"><div class="photo p2"></div><div class="body"><div class="qt">Ekspeditur</div><div class="qd">Rank capaian target, trip, truk dan dwell tiap mitra ekspedisi.</div><div class="qlink">Lihat ekspeditur &rsaquo;</div></div></a>' +
      '<a class="qcard" href="#/provinsi"><div class="photo p3"></div><div class="body"><div class="qt">Semua provinsi</div><div class="qd">Rincian tonase, target, SO dan distrik per wilayah.</div><div class="qlink">Lihat semua &rsaquo;</div></div></a>' +
    '</div>' +
    '<div class="section-head"><div class="section-title">Ringkasan operasional</div><a class="section-link" href="#/armada">Armada &rsaquo;</a></div>' +
    '<div class="kpi-row">' +
      kpi('trend', 'var(--black)', 'Trip', fmt(t.trips), {href:'#/armada', hint: incTxt + ' · ' + MON3[+m.slice(5) - 1]}) +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {href:'#/armada', hint: fmt1(t.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata dwell', t.dwell == null ? '–' : fmt1(t.dwell) + ' <small>jam</small>', {href:'#/armada', hint:'masuk → keluar pabrik'}) +
      kpi('check', 'var(--green)', 'Realisasi / SO', pctTxt(t.of), {href:'#/scorecard', hint: t.so ? 'SO ' + MON3[+m.slice(5) - 1] + ' ' + fmt(t.so) + ' t' + (isLatest ? ' (s.d. akhir bulan)' : '') : 'SO tidak tersedia'}) +
    '</div>' +
    noteBasis('Peta hanya mewarnai provinsi yang punya pengiriman darat di data; provinsi abu-abu tidak dilayani via darat.') + '</div>', 'home');
}

function pageProvList(){
  const rows = ranked(summarize('prov', F.month));
  const t = total(F.month);
  const cols = [C.rank, {h:'Provinsi', cls:'name', val: r => tc(B.dims.prov[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.note];
  cols[2].total = fmt(t.vol); cols[3].total = fmt(t.tgt); cols[4].total = pctTxt(t.pct); cols[5].total = signed(t.gap); cols[6].total = fmt(t.so); cols[7].total = pctTxt(t.of); cols[8].total = fmt(t.trips);
  page('<div class="wrap">' + crumb([['Semua provinsi']]) +
    '<div class="dp-head"><div><div class="nm">Semua Provinsi</div><div class="sub">Capaian tonase &amp; rank · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Total ' + pctTxt(t.pct) + ' dari target</div></div></div>' +
    table(cols, rows, {go: r => '#/provinsi/' + enc(B.dims.prov[r.key]), total: true}) +
    '<div class="dp-cols even">' + trendPanel({}, 'mingguan', 'Tren seluruh provinsi') + sourcePanel({}) + '</div>' +
    noteBasis() + '</div>', 'overview');
}
function sourcePanel(sc){
  const rows = [...summarize('src', F.month, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
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
  const m = F.month, sc = {prov: pi};
  const all = ranked(summarize('prov', m));
  const me = all.find(o => o.key === pi) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
  const dist = ranked(summarize('dist', m, sc));
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const pr = progRows(r => r[2] === name && r[1] === 'FRC');
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key])}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.note];
  page('<div class="wrap">' + crumb([['Provinsi', '#/provinsi'], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">Capaian tonase · ' + periodSub() + '</div></div><div class="badges">' + (me.rank ? '<div class="dp-badge">Rank ' + me.rank + ' / ' + all.rankTotal + '</div>' : '') + '<div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + (me.hasTgt ? ' Realisasi FRC ' + fmt(me.frc) + ' t vs target MTD ' + fmt(me.tgt) + ' t (gap ' + signed(me.gap) + ' t).' : '') + (me.so ? ' Realisasi/SO ' + pctTxt(me.of) + '.' : '') + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–', {hint: me.hasTgt ? 'bulan penuh ' + fmt(me.tgtFull) + ' t' : 'tidak ada target SNOP'}) +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', me.gap < 0 ? 'Sisa ke target' : 'Lebih dari target', me.hasTgt ? fmt(Math.abs(me.gap)) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>' +
    '<div class="dp-cols">' + trendPanel(sc, 'mingguan') +
      '<div class="dp-panel"><div class="pt">Ekspeditur di provinsi ini</div>' + (eks.length ? eks.map(o => rateRow(B.dims.eksp[o.key] + ' · ' + fmt(o.vol) + ' t', o.pct, '#/ekspeditur/' + enc(B.dims.eksp[o.key]), o.hasTgt ? 'target ' + fmt(o.tgt) + ' t' : 'tanpa target')).join('') : '<div class="empty">' + (F.inc === 'FOT' ? 'FOT diambil langsung oleh distributor — lihat halaman FOT.' : 'Tidak ada data.') + '</div>') + '</div>' +
    '</div>' +
    (pr.length ? '<div class="section-head"><div class="section-title">Prognose hari ini (FRC)</div><a class="section-link" href="#/forecast/harian">Forecast &rsaquo;</a></div>' + progCells(pr, 'snapshot ' + dayLabel(B.prog.date) + ' ' + B.prog.time) : '') +
    '<div class="section-head"><div class="section-title">Kinerja per distrik</div></div>' +
    table(distCols, dist, {go: r => '#/distrik/' + enc(B.dims.dist[r.key])}) +
    '<div class="dp-cols even">' + sourcePanel(sc) + '<div class="dp-panel"><div class="pt">Bulan penuh</div>' +
      '<div class="tgt-panel" style="margin:0;border:0;padding:0"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi FRC</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div></div></div>' +
    noteBasis('Distrik dengan target 0 dan realisasi 0 tidak ditampilkan.') + '</div>', 'overview');
}

function pageDist(name){
  const di = D.distIdx[name];
  if(di == null) return notFound('Distrik ' + name);
  const m = F.month, sc = {dist: di}, pname = B.dims.prov[B.dims.distProv[di]] || '';
  const me = total(m, sc);
  const eks = [...summarize('eksp', m, sc).values()].filter(o => (o.vol > 0 || o.tgt > 0) && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol);
  const ekCols = [{h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips, C.dwell];
  const soh = [1, 2, 3].map(h => B.soh.filter(s => s[0] === h && s[2] === di && (F.inc === 'ALL' || s[3] === F.inc)).reduce((a, s) => a + s[4], 0));
  page('<div class="wrap">' + crumb([['Provinsi', '#/provinsi'], [tc(pname), '#/provinsi/' + enc(pname)], [tc(name)]]) +
    '<div class="dp-head"><div><div class="nm">' + esc(tc(name)) + '</div><div class="sub">' + esc(tc(pname)) + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(me.pct) + '">' + pctTxt(me.pct) + ' dari target</div></div></div>' +
    '<div class="dp-review">' + reviewText(me.pct) + '</div>' +
    '<div class="dp-grid4">' +
      kpi('check', 'var(--green)', 'Realisasi', fmt(me.vol) + ' <small>ton</small>', {hint: fmt(me.trips) + ' trip'}) +
      kpi('target', 'var(--black)', 'Target MTD', me.hasTgt ? fmt(me.tgt) + ' <small>ton</small>' : '–') +
      kpi('gap', me.gap < 0 ? 'var(--red)' : 'var(--green)', 'Gap', me.hasTgt ? signed(me.gap) + ' <small>ton</small>' : '–') +
      kpi('box', 'var(--amber)', 'SO / Real÷SO', me.so ? fmt(me.so) + ' <small>t · ' + pctTxt(me.of) + '</small>' : '–') +
    '</div>' +
    '<div class="dp-cols">' + trendPanel(sc, 'harian') + '<div class="dp-panel"><div class="pt">SO menunggu kirim (snapshot ' + esc(B.prog.date ? dayLabel(B.prog.date) + ' ' + B.prog.time : '–') + ')</div>' +
      ['H+1', 'H+2', 'H+3'].map((h, i) => '<div class="rate-row"><div class="rl">' + h + '</div><div class="rtrack"><div class="rfill" style="width:' + Math.min(100, soh[i] / Math.max(1, ...soh) * 100) + '%;background:var(--black)"></div></div><div class="rv">' + fmt(soh[i]) + ' t</div></div>').join('') + '</div></div>' +
    '<div class="section-head"><div class="section-title">Ekspeditur di distrik ini</div></div>' +
    (eks.length ? table(ekCols, eks, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key]), sort: 1}) : '<div class="dp-review">Tidak ada ekspeditur FRC untuk filter ini.</div>') +
    noteBasis() + '</div>', 'overview');
}

function pageEkspList(){
  const rows = ranked(summarize('eksp', F.month));
  const list = rows.filter(r => r.key !== D.eksEmpty);
  const tr = trucksFor(F.month);
  const perE = new Map(); tr.forEach(t => t.eksp.forEach((n, e) => perE.set(e, (perE.get(e) || 0) + 1)));
  const cols = [C.rank, {h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.tgt, C.pct, C.gap, C.so, C.of, C.trips,
    {h:'Truk', num:true, val: r => perE.get(r.key) || 0, html: r => fmt(perE.get(r.key) || 0)}, C.dwell, C.note];
  page('<div class="wrap">' + crumb([['Semua ekspeditur']]) +
    '<div class="dp-head"><div><div class="nm">Semua Ekspeditur</div><div class="sub">Rank capaian target SNOP · ' + periodSub() + '</div></div></div>' +
    (F.inc === 'FOT' ? '<div class="dp-review">FOT diambil langsung oleh distributor (tanpa ekspeditur). Ganti filter ke FRC, atau buka <a href="#/incoterm/FOT" style="color:var(--red);font-weight:700">halaman FOT</a>.</div>' : '') +
    table(cols, list, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key])}) +
    noteBasis('SBA hanya memuat dari PP Belawan SBA, sehingga ikut tersembunyi saat source tersebut tidak dicentang. SO per ekspeditur adalah atribusi (estimasi) dari SO per distrik.') + '</div>', 'ekspeditur');
}

function pageEksp(code){
  const ei = D.ekspIdx[code];
  if(ei == null) return notFound('Ekspeditur ' + code);
  const m = F.month, sc = {eksp: ei};
  const all = ranked(summarize('eksp', m));
  const me = all.find(o => o.key === ei) || Object.assign(blank(), {pct:null, of:null, gap:0, dwell:null});
  const provs = [...summarize('prov', m, sc).values()].filter(o => o.vol > 0 || o.tgt > 0).sort((a, b) => b.vol - a.vol);
  const dist = ranked(summarize('dist', m, sc));
  const trucks = trucksFor(m, sc).sort((a, b) => b.trips - a.trips);
  const served = new Set(B.targets.filter(t => t[0] === m && t[3] === ei).map(t => t[2]));
  const soh = [1, 2, 3].map(h => B.soh.filter(s => s[0] === h && served.has(s[2]) && s[3] === 'FRC').reduce((a, s) => a + s[4], 0));
  const one = trucks.filter(t => t.trips === 1).length;
  const distCols = [C.rank, {h:'Distrik', cls:'name', val: r => tc(B.dims.dist[r.key]), html: r => esc(tc(B.dims.dist[r.key])) + '<small>' + esc(tc(B.dims.prov[B.dims.distProv[r.key]])) + '</small>'}, C.vol(), C.tgt, C.pct, C.gap, C.trips, C.note];
  const truckCols = [{h:'Nopol', cls:'name', val: r => B.dims.truck[r.truck]}, {h:'Trip', num:true, val: r => r.trips}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)},
    {h:'Hari aktif', num:true, val: r => r.days}, {h:'Dwell (jam)', num:true, val: r => r.dwN ? r.dwS / r.dwN : null, html: r => r.dwN ? fmt1(r.dwS / r.dwN) : '–'},
    {h:'Aktif', val: r => r.first, html: r => 'tgl ' + r.first + '–' + r.last}, {h:'Provinsi utama', val: r => tc(B.dims.prov[r.prov])}];
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
      '<div class="dp-panel"><div class="pt">Capaian per provinsi</div>' + (provs.length ? provs.map(o => rateRow(tc(B.dims.prov[o.key]) + ' · ' + fmt(o.vol) + ' t', o.pct, '#/provinsi/' + enc(B.dims.prov[o.key]), o.hasTgt ? 'target ' + fmt(o.tgt) + ' t' : '')).join('') : '<div class="empty">Tidak ada data untuk filter ini.</div>') + '</div>' +
    '</div>' +
    '<div class="section-head"><div class="section-title">SO menunggu kirim di distrik layanan</div><a class="section-link" href="#/forecast/harian">Forecast &rsaquo;</a></div>' +
    '<div class="prog-row">' + ['H+1', 'H+2', 'H+3'].map((h, i) => '<div class="prog-cell"><div class="pl">' + h + '</div><div class="pv">' + fmt(soh[i]) + ' t</div><div class="pd">SO FRC di ' + served.size + ' distrik ber-target ' + esc(code) + '</div></div>').join('') + '</div>' +
    '<div class="section-head"><div class="section-title">Target bulan berjalan</div></div>' +
    '<div class="tgt-panel"><div class="tgt-item"><div class="tl">Target bulan</div><div class="tv">' + (me.hasTgt ? fmt(me.tgtFull) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Target MTD</div><div class="tv">' + (me.hasTgt ? fmt(me.tgt) : '–') + ' t</div></div><div class="tgt-item"><div class="tl">Realisasi</div><div class="tv">' + fmt(me.frc) + ' t</div></div><div class="tgt-item"><div class="tl">Sisa ke target bulan</div><div class="tv ' + (me.tgtFull - me.frc > 0 ? 'bad' : 'good') + '">' + (me.hasTgt ? fmt(Math.max(0, me.tgtFull - me.frc)) : '–') + ' t</div></div></div>' +
    '<div class="section-head"><div class="section-title">Kinerja per distrik yang dilayani</div></div>' +
    table(distCols, dist, {go: r => '#/distrik/' + enc(B.dims.dist[r.key])}) +
    '<div class="section-head"><div class="section-title">Truk ' + esc(code) + ' (' + fmt(trucks.length) + ')</div><a class="section-link" href="#/armada">Semua armada &rsaquo;</a></div>' +
    table(truckCols, trucks, {limit: 300}) +
    noteBasis('SO H+n tidak tercatat per ekspeditur di sumber; angka di atas adalah SO di distrik tempat ' + esc(code) + ' punya target SNOP.') + '</div>', 'ekspeditur');
}

function pageForecast(gran){
  gran = ['harian', 'mingguan', 'bulanan'].includes(gran) ? gran : 'harian';
  const chips = '<div class="gran-chips" style="margin-bottom:24px">' + [['harian','Harian'],['mingguan','Mingguan'],['bulanan','Bulanan']].map(g => '<a href="#/forecast/' + g[0] + '" class="' + (g[0] === gran ? 'on' : '') + '">' + g[1] + '</a>').join('') + '</div>';
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
    const end = parseISO(F.month + '-' + String(lastDayOf(F.month)).padStart(2, '0'));
    const mon = new Date(end); mon.setDate(end.getDate() - ((end.getDay() + 6) % 7));
    const days = [...Array(7)].map((_, i) => { const d = new Date(mon); d.setDate(mon.getDate() + i); return d; });
    const data = days.map(d => {
      const s = iso(d), future = d > end, r = future ? rangeTotals(s, s) : rangeTotals(s, s);
      return {label: DAYS[d.getDay()].slice(0, 3) + ' ' + d.getDate(), long: DAYS[d.getDay()] + ', ' + dayLabel(s) + (future ? ' (belum terjadi)' : ''), vol: future ? 0 : r.vol, frc: future ? 0 : r.frc, tgt: r.pct == null && !future ? null : r.tgt, pct: future ? null : r.pct};
    });
    const sofar = data.filter((d, i) => days[i] <= end);
    const sReal = sofar.reduce((a, d) => a + d.frc, 0), sTgt = sofar.reduce((a, d) => a + (d.tgt || 0), 0), wTgt = data.reduce((a, d) => a + (d.tgt || 0), 0);
    const runrate = sofar.length ? sReal / sofar.length : 0, proj = sReal + runrate * (7 - sofar.length);
    const id = 'tw'; TRENDS[id] = {data, sc: {}, gran: 'x'};
    body = '<div class="prog-row">' +
      '<div class="prog-cell"><div class="pl">Realisasi minggu ini</div><div class="pv">' + fmt(sReal) + ' t</div><div class="pd">' + sofar.length + ' dari 7 hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target s.d. hari ini</div><div class="pv">' + fmt(sTgt) + ' t</div><div class="pd">capaian ' + pctTxt(pctOf(sReal, sTgt)) + '</div></div>' +
      '<div class="prog-cell"><div class="pl">Proyeksi minggu</div><div class="pv" style="color:' + tierColor(pctOf(proj, wTgt)) + '">' + fmt(proj) + ' t</div><div class="pd">laju ' + fmt(runrate) + ' t/hari</div></div>' +
      '<div class="prog-cell"><div class="pl">Target minggu penuh</div><div class="pv">' + fmt(wTgt) + ' t</div><div class="pd">proyeksi ' + pctTxt(pctOf(proj, wTgt)) + '</div></div></div>' +
      '<div class="dp-panel" style="margin-bottom:34px"><div class="pt">Realisasi FRC per hari vs target SNOP, ' + days[0].getDate() + ' ' + MON3[days[0].getMonth()] + ' – ' + days[6].getDate() + ' ' + MON3[days[6].getMonth()] + '</div>' + barsHTML(id, data) +
      '<div class="trend-legend"><span><i style="background:var(--green)"></i>&ge;100%</span><span><i style="background:var(--amber)"></i>85–99%</span><span><i style="background:var(--red)"></i>&lt;85%</span><span><i class="dash"></i>target SNOP</span></div><div class="trend-detail" id="tw-d">Klik batang untuk rincian hari.</div></div>' +
      noteBasis('Proyeksi = realisasi minggu berjalan + rata-rata harian minggu ini × sisa hari. Bukan prognosa sistem.');
  } else {
    const m = F.month, ld = lastDayOf(m), n = dim(m);
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
  page('<div class="wrap">' + crumb([['Forecast'], [tc(gran)]]) +
    '<div class="dp-head"><div><div class="nm">Forecast — ' + tc(gran) + '</div><div class="sub">' + (gran === 'harian' ? 'Prognosa pengiriman hari ini dari file monitoring' : gran === 'mingguan' ? 'Minggu berjalan, Senin–Minggu' : 'Proyeksi akhir ' + monthLabel(F.month)) + '</div></div></div>' +
    chips + body + '</div>', 'forecast');
}

function pageIncoterm(code){
  code = code === 'FOT' ? 'FOT' : 'FRC';
  const m = F.month, saved = F.inc;
  F.inc = code; const pm = summarize('prov', m), t = total(m); F.inc = 'ALL'; const both = total(m); F.inc = saved;
  const rows = [...pm.values()].filter(o => o.vol > 0).sort((a, b) => b.vol - a.vol);
  const lastInc = B.days[Math.max(...B.facts.filter(f => f[1] === (code === 'FOT' ? 1 : 0)).map(f => f[0]))];
  const emptyNote = rows.length ? '' : '<div class="dp-review" style="border-left:4px solid var(--amber)">Belum ada data ' + code + ' untuk ' + monthLabel(m) + '. Data ' + code + ' di master Excel baru sampai ' + dayLabel(lastInc) + ' — tambahkan baris baru di sheet ' + (code === 'FOT' ? '“Realisasi H”' : '“Realisasi FRC”') + ' lalu jalankan build ulang. Pilih bulan lain lewat Filter.</div>';
  let side = '';
  if(code === 'FOT'){
    const per = new Map();
    B.fotDist.forEach(r => { if(r[0] === m){ const o = per.get(r[2]) || {d: r[2], ton: 0, trips: 0, prov: r[1]}; o.ton += r[3]; o.trips += r[4]; per.set(r[2], o); } });
    const list = [...per.values()].sort((a, b) => b.ton - a.ton);
    side = '<div class="section-head"><div class="section-title">Distributor (transportir FOT)</div></div>' +
      table([{h:'#', cls:'rank', val: r => 0, html: (r, i) => '#' + (i + 1)}, {h:'Distributor', cls:'name', val: r => tc(B.dims.distr[r.d])}, {h:'Provinsi', val: r => tc(B.dims.prov[r.prov])}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)}, {h:'Trip', num:true, val: r => r.trips}], list, {limit: 100});
  } else {
    F.inc = 'FRC'; const em = [...summarize('eksp', m).values()].filter(o => o.vol > 0 && o.key !== D.eksEmpty).sort((a, b) => b.vol - a.vol); F.inc = saved;
    side = '<div class="section-head"><div class="section-title">Ekspeditur</div></div>' + table([{h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.tgt, C.pct, C.trips], em, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key])});
  }
  page('<div class="wrap">' + crumb([['Forecast', '#/forecast/harian'], [code]]) +
    '<div class="dp-head"><div><div class="nm">' + code + '</div><div class="sub">' + (code === 'FOT' ? 'Free on Truck — diambil langsung oleh distributor dengan armadanya sendiri' : 'Franco — dikirim oleh ekspeditur kontrak Semen Padang') + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge">' + pctTxt(pctOf(t.vol, both.vol)) + ' dari total darat</div></div></div>' +
    emptyNote + '<div class="dp-grid4">' +
      kpi('box', 'var(--black)', 'Tonase', fmt(t.vol) + ' <small>ton</small>') +
      kpi('trend', 'var(--amber)', 'Share', pctTxt(pctOf(t.vol, both.vol)), {hint: 'dari ' + fmt(both.vol) + ' t FRC+FOT'}) +
      kpi('truck', 'var(--black)', 'Trip', fmt(t.trips), {hint: t.dwell == null ? '' : 'dwell ' + fmt1(t.dwell) + ' jam'}) +
      kpi('pin', 'var(--black)', 'Provinsi', fmt(rows.length)) +
    '</div>' +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Per provinsi</div>' + (rows.length ? rows.map(o => '<a class="rate-row" href="#/provinsi/' + enc(B.dims.prov[o.key]) + '"><div class="rl">' + esc(tc(B.dims.prov[o.key])) + '</div><div class="rtrack"><div class="rfill" style="width:' + (o.vol / rows[0].vol * 100) + '%;background:var(--black)"></div></div><div class="rv">' + fmt(o.vol) + '</div></a>').join('') : '<div class="empty">Tidak ada data.</div>') + '</div>' +
    (function(){ const s = F.inc; F.inc = code; const h = trendPanel({}, 'bulanan', 'Tren ' + code); F.inc = s; TRENDS[Object.keys(TRENDS).pop()].incFix = code; return h; })() + '</div>' +
    side + noteBasis(code === 'FOT' ? 'FOT tidak memiliki target SNOP; transportir FOT = distributor.' : '') + '</div>', 'forecast');
}

function pageTren(){
  page('<div class="wrap">' + crumb([['Overview'], ['Tren']]) +
    '<div class="dp-head"><div><div class="nm">Tren Tonase</div><div class="sub">' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + ' · seluruh provinsi · s.d. ' + periodSub() + '</div></div></div>' +
    '<div class="dp-cols even" style="margin-bottom:28px">' + trendPanel({}, 'mingguan', 'Mingguan') + trendPanel({}, 'bulanan', 'Bulanan') + '</div>' +
    '<div style="margin-bottom:34px">' + trendPanel({}, 'harian', 'Harian') + '</div>' +
    noteBasis('Minggu = Senin–Minggu. Bulan berjalan dihitung MTD.') + '</div>', 'overview');
}

function pageArmada(){
  const m = F.month, t = total(m), trucks = trucksFor(m).sort((a, b) => b.trips - a.trips);
  const buckets = [['1', 1, 1], ['2–3', 2, 3], ['4–6', 4, 6], ['7–10', 7, 10], ['11–20', 11, 20], ['>20', 21, 1e9]];
  const bdata = buckets.map(b => ({label: b[0] + ' trip', long: b[0] + ' trip per truk', vol: trucks.filter(x => x.trips >= b[1] && x.trips <= b[2]).length, tgt: null, pct: null}));
  const dow = [1, 2, 3, 4, 5, 6, 0].map(d => {
    let s = 0, n = 0;
    B.facts.forEach(f => { if(D.dayMonth[f[0]] === m && D.dayDow[f[0]] === d && incOk(f[1]) && srcOk(f[2])){ s += f[8]; n += f[9]; } });
    return {label: DAYS[d].slice(0, 3), long: DAYS[d], vol: n ? s / n : 0, tgt: null, pct: null, dec: true};
  });
  TRENDS.ab = {data: bdata.map(x => Object.assign({}, x, {unit: 'truk'}))}; TRENDS.ad = {data: dow.map(x => Object.assign({}, x, {unit: 'jam'}))};
  const dwellBars = '<div class="bars">' + dow.map((d, i) => { const mx = Math.max(...dow.map(x => x.vol), 1); const c = d.vol <= 3.5 ? 'var(--green)' : d.vol <= 4.5 ? 'var(--amber)' : 'var(--red)'; return '<div class="bcol"><div class="bval">' + fmt1(d.vol) + '</div><div class="bwrap"><div class="b" style="height:' + Math.max(2, d.vol / mx * 100) + '%;background:' + c + ';animation-delay:' + (i * 50) + 'ms"></div></div><div class="blabel">' + d.label + '</div></div>'; }).join('') + '</div>';
  const one = trucks.filter(x => x.trips === 1).length;
  const cols = [{h:'#', cls:'rank', val: r => r.trips, html: (r, i) => '#' + (i + 1)}, {h:'Nopol', cls:'name', val: r => B.dims.truck[r.truck]},
    {h:'Ekspeditur', val: r => [...r.eksp.keys()].map(e => B.dims.eksp[e] || 'FOT').join(', ')}, {h:'Trip', num:true, val: r => r.trips}, {h:'Tonase', num:true, val: r => r.ton, html: r => fmt(r.ton)},
    {h:'Hari aktif', num:true, val: r => r.days}, {h:'Dwell (jam)', num:true, val: r => r.dwN ? r.dwS / r.dwN : null, html: r => r.dwN ? fmt1(r.dwS / r.dwN) : '–'},
    {h:'Aktif', val: r => r.first, html: r => 'tgl ' + r.first + '–' + r.last}, {h:'Provinsi utama', val: r => tc(B.dims.prov[r.prov])}];
  page('<div class="wrap">' + crumb([['Overview'], ['Armada']]) +
    '<div class="dp-head"><div><div class="nm">Armada</div><div class="sub">Truk aktif, trip dan dwell time · ' + (F.inc === 'ALL' ? 'FRC + FOT' : F.inc) + ' · ' + periodSub() + '</div></div><div class="badges"><div class="dp-badge">' + fmt(trucks.length) + ' truk aktif</div></div></div>' +
    '<div class="dp-grid4">' +
      kpi('truck', 'var(--black)', 'Truk aktif', fmt(trucks.length), {hint: one + ' truk baru 1 trip'}) +
      kpi('trend', 'var(--black)', 'Trip', fmt(t.trips), {hint: fmt1(t.trips / Math.max(1, trucks.length)) + ' trip / truk'}) +
      kpi('clock', 'var(--amber)', 'Rata-rata dwell', t.dwell == null ? '–' : fmt1(t.dwell) + ' <small>jam</small>', {hint: 'masuk → keluar pabrik'}) +
      kpi('box', 'var(--black)', 'Ton / trip', fmt1(t.vol / Math.max(1, t.trips))) +
    '</div>' +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Sebaran trip per truk</div>' + barsHTML('ab', bdata) + '<div class="trend-detail" id="ab-d">Banyak truk 1 trip = armada tidak kontinu.</div></div>' +
    '<div class="dp-panel"><div class="pt">Rata-rata dwell per hari (jam) — makin rendah makin baik</div>' + dwellBars + '<div class="trend-legend"><span><i style="background:var(--green)"></i>&le;3,5 jam</span><span><i style="background:var(--amber)"></i>3,5–4,5</span><span><i style="background:var(--red)"></i>&gt;4,5</span></div></div></div>' +
    '<div class="section-head"><div class="section-title">Daftar truk</div></div>' + table(cols, trucks, {limit: 400}) +
    noteBasis('Dwell = jam masuk s.d. jam keluar pabrik; baris yang ditandai CHECK (negatif / &gt;48 jam) di sumber tidak dihitung.') + '</div>', 'overview');
}

function pageScorecard(){
  const m = F.month, t = total(m);
  const rows = ranked(summarize('eksp', m)).filter(r => r.key !== D.eksEmpty);
  const tr = trucksFor(m);
  const perE = new Map(); tr.forEach(x => x.eksp.forEach((n, e) => { const o = perE.get(e) || {n: 0, one: 0}; o.n++; if(x.trips === 1) o.one++; perE.set(e, o); }));
  const multi = tr.length ? pctOf(tr.filter(x => x.trips > 1).length, tr.length) : null;
  const cols = [C.rank, {h:'Ekspeditur', cls:'name', val: r => B.dims.eksp[r.key]}, C.vol(), C.pct, C.of,
    {h:'Truk', num:true, val: r => (perE.get(r.key) || {}).n || 0}, {h:'Trip / truk', num:true, val: r => r.trips / Math.max(1, (perE.get(r.key) || {}).n || 0), html: r => fmt1(r.trips / Math.max(1, (perE.get(r.key) || {}).n || 0))},
    {h:'Truk 1 trip', num:true, val: r => (perE.get(r.key) || {}).one || 0}, C.dwell, C.note];
  page('<div class="wrap">' + crumb([['Overview'], ['Scorecard']]) +
    '<div class="dp-head"><div><div class="nm">Scorecard Operasional</div><div class="sub">' + periodSub() + '</div></div><div class="badges"><div class="dp-badge ' + tierClass(t.pct) + '">Capaian ' + pctTxt(t.pct) + '</div></div></div>' +
    '<div class="dp-cols even"><div class="dp-panel"><div class="pt">Komponen</div>' +
      rateRow('Capaian target SNOP', t.pct) + rateRow('Realisasi / SO', t.of) + rateRow('Truk aktif > 1 trip', multi) +
    '</div><div class="dp-panel"><div class="pt">Capaian per ekspeditur</div>' + rows.filter(r => r.pct != null).map(r => rateRow(B.dims.eksp[r.key], r.pct, '#/ekspeditur/' + enc(B.dims.eksp[r.key]))).join('') + '</div></div>' +
    table(cols, rows, {go: r => '#/ekspeditur/' + enc(B.dims.eksp[r.key])}) +
    noteBasis('Real/SO = realisasi FRC ÷ SO pada periode yang sama (SO per ekspeditur = atribusi).') + '</div>', 'overview');
}

function notFound(what){ page('<div class="wrap">' + crumb([['Tidak ditemukan']]) + '<div class="dp-head"><div><div class="nm">Tidak ditemukan</div><div class="sub">' + esc(what) + ' tidak ada di data.</div></div></div><a class="section-link" href="#/">Kembali ke Home &rsaquo;</a></div>', ''); }

/* ===================== Router ===================== */
function route(){
  if(!B) return;
  closeNav(); $('#searchResults').classList.remove('show');
  for(const k in TRENDS) delete TRENDS[k];
  for(const k in TABLES) delete TABLES[k];
  const parts = (location.hash.replace(/^#\/?/, '') || '').split('/').map(decodeURIComponent);
  const [a, b] = parts;
  try {
    if(!a) pageHome();
    else if(a === 'provinsi') b ? pageProv(b) : pageProvList();
    else if(a === 'distrik' && b) pageDist(b);
    else if(a === 'ekspeditur') b ? pageEksp(b) : pageEkspList();
    else if(a === 'forecast') pageForecast(b);
    else if(a === 'incoterm') pageIncoterm(b);
    else if(a === 'tren') pageTren();
    else if(a === 'armada') pageArmada();
    else if(a === 'scorecard') pageScorecard();
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
        item.classList.add('open'); if(item.dataset.nav === 'filter') buildFilterPanel();
        const panel = item.querySelector('.dropdown,.mega,.filter-panel');   // on narrow screens panels are position:fixed under the header
        if(panel) panel.style.top = window.innerWidth <= 1100 ? ($('header').getBoundingClientRect().bottom + 4) + 'px' : '';
      }
      return;
    }
    const act = e.target.closest('[data-act]');
    if(act){
      const a = act.dataset.act;
      if(a === 'apply') applyFilter();
      if(a === 'reset'){ draft = {month: D.months[D.months.length - 1], inc: 'FRC', srcOff: new Set(DEFAULT_SRC_OFF)}; applyFilter(); }
      if(a === 'openfilter'){ e.stopPropagation(); closeNav(); const it = $('[data-nav="filter"]'); it.classList.add('open'); buildFilterPanel(); }
      if(a === 'logout'){ logout(); }
      return;
    }
    const inc = e.target.closest('[data-inc]');
    if(inc){ draft.inc = inc.dataset.inc; $$('[data-inc]').forEach(x => x.classList.toggle('on', x === inc)); return; }
    const qm = e.target.closest('[data-qm]');
    if(qm){ const last = D.months[D.months.length - 1]; draft.month = qm.dataset.qm === 'cur' ? last : (D.months.includes(prevMonth(last)) ? prevMonth(last) : last); $('#fpMonth').value = draft.month; return; }
    const src = e.target.closest('[data-src]');
    if(src){ src.checked ? draft.srcOff.delete(src.dataset.src) : draft.srcOff.add(src.dataset.src); return; }
    const gran = e.target.closest('[data-gran]');
    if(gran){ const [id, g] = gran.dataset.gran.split(':'); const T = TRENDS[id]; T.gran = g; const s = F.inc; if(T.incFix) F.inc = T.incFix; $('#' + id).innerHTML = trendInner(id); F.inc = s; return; }
    const bar = e.target.closest('[data-bar]');
    if(bar){
      const [id, i] = bar.dataset.bar.split(':'); const d = TRENDS[id].data[+i];
      $$('[data-bar^="' + id + ':"]').forEach(x => x.classList.toggle('sel', x === bar));
      const strip = $('#' + id + '-d');
      if(strip) strip.innerHTML = d.unit ? '<b>' + esc(d.long) + '</b> — ' + (d.unit === 'jam' ? fmt1(d.vol) : fmt(d.vol)) + ' ' + d.unit + '.' :
        '<b>' + esc(d.long) + '</b> — Realisasi ' + fmt(d.vol) + ' t' + (d.tgt != null ? ' · FRC ' + fmt(d.frc) + ' t vs target ' + fmt(d.tgt) + ' t · capaian <b style="color:' + tierColor(d.pct) + '">' + pctTxt(d.pct) + '</b> · gap ' + signed(d.frc - d.tgt) + ' t' : '') + '.';
      return;
    }
    const sort = e.target.closest('[data-sort]');
    if(sort){ const [id, i] = sort.dataset.sort.split(':'); const T = TABLES[id]; if(T.sort === +i) T.dir = -T.dir; else { T.sort = +i; T.dir = T.cols[+i].num ? -1 : 1; } $('#' + id).innerHTML = tableInner(id); return; }
    const go = e.target.closest('[data-href]');
    if(go && go.dataset.href && go.dataset.href !== 'null'){ location.hash = go.dataset.href; return; }
    if(!e.target.closest('#mainNav')) closeNav();
    if(!e.target.closest('.search')) $('#searchResults').classList.remove('show');
  });
  document.addEventListener('change', e => { if(e.target.id === 'fpMonth') draft.month = e.target.value; if(e.target.id === 'fpJump' && e.target.value){ location.hash = e.target.value; closeNav(); } });
  document.addEventListener('keydown', e => { if(e.key === 'Escape'){ closeNav(); $('#searchResults').classList.remove('show'); } });
  const sb = $('#searchBox');
  sb.addEventListener('input', () => doSearch(sb.value));
  sb.addEventListener('focus', () => doSearch(sb.value));
  sb.addEventListener('keydown', e => {
    const links = $$('#searchResults a'); if(!links.length) return;
    let i = links.findIndex(a => a.classList.contains('sel'));
    if(e.key === 'ArrowDown' || e.key === 'ArrowUp'){ e.preventDefault(); i = (i + (e.key === 'ArrowDown' ? 1 : -1) + links.length) % links.length; links.forEach((a, j) => a.classList.toggle('sel', j === i)); }
    if(e.key === 'Enter'){ e.preventDefault(); location.hash = links[Math.max(0, i)].getAttribute('href'); sb.value = ''; sb.blur(); $('#searchResults').classList.remove('show'); }
  });
  // map tooltip
  const tip = document.createElement('div'); tip.className = 'map-tip'; document.body.appendChild(tip);
  document.addEventListener('mousemove', e => {
    const t = e.target.closest && e.target.closest('[data-tip]');
    if(!t){ tip.style.display = 'none'; return; }
    const [a, b2, c] = t.dataset.tip.split('|');
    tip.innerHTML = '<b>' + esc(a) + '</b><br>' + esc(b2) + (c ? '<br>' + esc(c) : '');
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
function touch(){ try { sessionStorage.setItem('osp_t', String(Date.now())); } catch(e) {} }
function logout(){ try { sessionStorage.removeItem('osp_k'); sessionStorage.removeItem('osp_t'); } catch(e) {} location.reload(); }
async function start(key, meta){
  B = await loadBundle(key, meta);
  B.build = meta.build;
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
