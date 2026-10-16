/* Shared helpers for the admin functions (Cloudflare Pages Functions in functions/).
   Nothing secret lives here: tokens come from the Pages environment (see ADMIN.md). */

export const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {status, headers: {'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store'}});

const b64url = s => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4)), c => c.charCodeAt(0));
let certCache = {at: 0, keys: null};

/** Verifies the Cloudflare Access login token (RS256 JWT) and returns the admin's email, or null.
    Checked here as well as by Access itself, so a misconfigured Access policy or an unprotected preview URL
    still cannot reach the admin functions. env: ACCESS_TEAM_DOMAIN (https://<team>.cloudflareaccess.com),
    ACCESS_AUD (the application's audience tag), ADMIN_EMAILS (comma-separated allowlist). */
export async function adminEmail(request, env){
  const cookie = (request.headers.get('cookie') || '').match(/(?:^|;\s*)CF_Authorization=([^;]+)/);
  const token = request.headers.get('cf-access-jwt-assertion') || (cookie && cookie[1]);
  if(!token || !env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD) return null;
  const [h, p, s] = token.split('.');
  if(!s) return null;
  let head, claims;
  try { head = JSON.parse(new TextDecoder().decode(b64url(h))); claims = JSON.parse(new TextDecoder().decode(b64url(p))); } catch(e) { return null; }
  if(head.alg !== 'RS256') return null;
  if(!certCache.keys || Date.now() - certCache.at > 3600e3){
    const r = await fetch(env.ACCESS_TEAM_DOMAIN.replace(/\/$/, '') + '/cdn-cgi/access/certs');
    if(!r.ok) return null;
    certCache = {at: Date.now(), keys: (await r.json()).keys || []};
  }
  const jwk = certCache.keys.find(k => k.kid === head.kid);
  if(!jwk) return null;
  const key = await crypto.subtle.importKey('jwk', jwk, {name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256'}, false, ['verify']);
  const ok = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, b64url(s), new TextEncoder().encode(h + '.' + p));
  const aud = [].concat(claims.aud || []);
  if(!ok || !aud.includes(env.ACCESS_AUD) || !(claims.exp > Date.now() / 1000)) return null;
  const email = String(claims.email || '').toLowerCase();
  const allowed = String(env.ADMIN_EMAILS || '').toLowerCase().split(',').map(x => x.trim()).filter(Boolean);
  return allowed.includes(email) ? email : null;
}

/** Bearer check for the build server (GitHub Action). */
export function buildAuthorized(request, env){
  const got = (request.headers.get('authorization') || '').replace(/^Bearer\s+/i, '');
  if(!env.BUILD_TOKEN || got.length !== env.BUILD_TOKEN.length) return false;
  let diff = 0;
  for(let i = 0; i < got.length; i++) diff |= got.charCodeAt(i) ^ env.BUILD_TOKEN.charCodeAt(i);
  return diff === 0;
}

/* ---- state in R2: master/latest.xlsx, state/build.json (last build), state/history.json (uploads) ---- */
export async function readJSON(env, key, fallback){
  const o = await env.MASTER.get(key);
  return o ? o.json() : fallback;
}
export const writeJSON = (env, key, data) => env.MASTER.put(key, JSON.stringify(data), {httpMetadata: {contentType: 'application/json'}});

/* ---- GitHub: start the build workflow, read/write tools/lock.json ---- */
const gh = (env, path, init = {}) => fetch((env.GITHUB_API || 'https://api.github.com') + '/repos/' + env.GITHUB_REPO + path, Object.assign({}, init, {
  headers: Object.assign({'authorization': 'Bearer ' + env.GITHUB_TOKEN, 'accept': 'application/vnd.github+json', 'user-agent': 'logistik-sp-admin', 'x-github-api-version': '2022-11-28'}, init.headers || {})
}));
export async function startBuild(env, reason){
  const r = await gh(env, '/dispatches', {method: 'POST', body: JSON.stringify({event_type: 'build-site', client_payload: {reason}})});
  if(r.status !== 204) throw new Error('GitHub menolak memulai build (' + r.status + ')');
}
export async function getRepoFile(env, path){
  const r = await gh(env, '/contents/' + path + '?ref=main');
  if(!r.ok) throw new Error('Tidak bisa membaca ' + path + ' dari GitHub (' + r.status + ')');
  const f = await r.json();
  return {sha: f.sha, text: new TextDecoder().decode(Uint8Array.from(atob(f.content.replace(/\n/g, '')), c => c.charCodeAt(0)))};
}
export async function putRepoFile(env, path, text, sha, message){
  const bytes = new TextEncoder().encode(text);
  let bin = ''; bytes.forEach(b => bin += String.fromCharCode(b));
  const r = await gh(env, '/contents/' + path, {method: 'PUT', body: JSON.stringify({message, content: btoa(bin), sha, branch: 'main'})});
  if(!r.ok) throw new Error('GitHub menolak menyimpan ' + path + ' (' + r.status + ')');
}

/* ---- master store in R2 (private): store/current/<path>, store/staging/<path> (new run, not yet promoted),
        store/versions/<ts>/<path> (last 7 states, for rollback), inbox/<slot>.xlsx (uploaded raw files) ---- */
export const SLOTS = {rilis: 'Rilis (Realisasi Shipment MTD)', so: 'SO (Sales Order CSMS)', tactical: 'Tactical demand', snop: 'SNOP', monitoring: 'Monitoring Performance', gudang: 'Lokasi gudang / toko (koordinat)', kapal: 'Master kapal (info & foto)'};
export const STORE_FILES = ['data/rilis.csv.gz', 'data/so.csv.gz', 'data/target_ekspeditur.csv.gz', 'data/target_distributor.csv.gz',
  'config/settings.json', 'config/ekspeditur_map.json', 'config/so_frozen_attrib_master_lama.json', 'data/gudang.csv.gz', 'data/kapal.csv.gz'];
export const KEEP_VERSIONS = 7;
export const okPath = p => STORE_FILES.includes(p);

export async function listAll(env, prefix){
  const out = []; let cursor;
  do { const r = await env.MASTER.list({prefix, cursor, include: ['customMetadata']}); out.push(...r.objects); cursor = r.truncated ? r.cursor : undefined; } while(cursor);
  return out;
}
async function copyObj(env, from, to){
  const o = await env.MASTER.get(from); if(!o) return false;
  await env.MASTER.put(to, await o.arrayBuffer(), {httpMetadata: o.httpMetadata, customMetadata: o.customMetadata});
  return true;
}
export const stamp = () => new Date().toISOString().replace(/[-:]/g, '').replace(/\..*/, '') ;   // 20261015T103000
/** Copies store/current/* to store/versions/<ts>/*, then drops versions beyond KEEP_VERSIONS. Returns ts (null if store empty). */
export async function snapshotCurrent(env, note){
  const cur = await listAll(env, 'store/current/'); if(!cur.length) return null;
  const ts = stamp();
  for(const o of cur) await copyObj(env, o.key, 'store/versions/' + ts + '/' + o.key.slice('store/current/'.length));
  await env.MASTER.put('store/versions/' + ts + '/_note.json', JSON.stringify({ts, note, at: new Date().toISOString()}), {httpMetadata: {contentType: 'application/json'}});
  const all = await listAll(env, 'store/versions/');
  const tss = [...new Set(all.map(o => o.key.split('/')[2]))].sort().reverse();
  for(const old of tss.slice(KEEP_VERSIONS)) for(const o of all.filter(o => o.key.split('/')[2] === old)) await env.MASTER.delete(o.key);
  return ts;
}
export async function listVersions(env){
  const all = await listAll(env, 'store/versions/'), by = {};
  for(const o of all){ const t = o.key.split('/')[2]; (by[t] = by[t] || {ts: t, size: 0, files: 0}); if(!o.key.endsWith('_note.json')){ by[t].size += o.size; by[t].files++; } }
  const out = Object.values(by).sort((a, b) => b.ts.localeCompare(a.ts));
  for(const v of out){
    for(const k of ['_note.json', '_meta.json']){
      const o = await env.MASTER.get('store/versions/' + v.ts + '/' + k);
      if(o) Object.assign(v, k === '_note.json' ? {note: (await o.json()).note} : {data: await o.json()});
    }
  }
  return out;
}
