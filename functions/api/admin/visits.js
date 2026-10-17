// GET ?days=30: dashboard visits (from /api/hit), aggregated. Older than 120 days is deleted here.
import {json, listAll} from '../../../lib/admin.js';

const KEEP = 120;
const day = off => new Date(Date.now() + 7 * 3600e3 - off * 864e5).toISOString().slice(0, 10);

export async function onRequestGet({request, env}){
  const n = Math.min(90, Math.max(1, +new URL(request.url).searchParams.get('days') || 30));
  const days = Array.from({length: n}, (_, i) => day(n - 1 - i));
  const lists = await Promise.all(days.map(d => listAll(env, 'visits/' + d + '/')));
  const hits = lists.flatMap((l, i) => l.map(o => Object.assign({day: days[i], t: o.key.slice(18, 24)}, o.customMetadata)));
  const perDay = days.map((d, i) => { const l = hits.filter(h => h.day === d); return {d, views: l.length, visitors: new Set(l.map(h => h.v)).size, sessions: new Set(l.map(h => h.s)).size}; });
  const count = (key, f = h => h[key]) => { const m = new Map(); hits.forEach(h => { const k = f(h) || '–'; const o = m.get(k) || {k, views: 0, vs: new Set()}; o.views++; o.vs.add(h.v); m.set(k, o); });
    return [...m.values()].map(o => ({k: o.k, views: o.views, visitors: o.vs.size})).sort((a, b) => b.views - a.views); };
  const section = h => { const p = (h.p || '/').split('/'); return p[0] || 'ringkasan'; };
  // returning = visitor seen on 2+ different days in the window
  const vDays = new Map(); hits.forEach(h => { const s = vDays.get(h.v) || new Set(); s.add(h.day); vDays.set(h.v, s); });
  // housekeeping: drop day folders older than KEEP days
  const cut = day(KEEP), old = [];
  let cursor; do { const r = await env.MASTER.list({prefix: 'visits/', delimiter: '/', cursor}); old.push(...r.delimitedPrefixes.filter(p => p.slice(7, 17) < cut)); cursor = r.truncated ? r.cursor : undefined; } while(cursor);
  for(const p of old.slice(0, 5)){ const l = await listAll(env, p); for(let i = 0; i < l.length; i += 1000) await env.MASTER.delete(l.slice(i, i + 1000).map(o => o.key)); }
  return json({days: perDay, total: {views: hits.length, visitors: vDays.size, returning: [...vDays.values()].filter(s => s.size > 1).length, sessions: new Set(hits.map(h => h.s)).size},
    sections: count('sec', section), pages: count('p').slice(0, 25), devices: count('d'), cities: count('c').slice(0, 15),
    recent: hits.slice(-40).reverse().map(h => ({day: h.day, t: h.t, p: h.p, d: h.d, c: h.c, v: h.v.slice(0, 4)}))});
}
