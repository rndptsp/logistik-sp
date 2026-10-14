// GET: stored versions (last 7), newest first. POST {ts}: roll the store back to that version, then rebuild.
import {json, listVersions, snapshotCurrent, listAll, readJSON, writeJSON, startBuild} from '../../../lib/admin.js';

export async function onRequestGet({env}){ return json({versions: await listVersions(env)}); }

export async function onRequestPost({request, env, data}){
  const {ts} = await request.json().catch(() => ({}));
  if(!/^\d{8}T\d{6}$/.test(ts || '')) return json({error: 'Versi tidak valid'}, 400);
  const cur = await readJSON(env, 'state/build.json', {});
  if((cur.state === 'queued' || cur.state === 'running') && Date.now() - Date.parse(cur.updated || cur.at) < 40 * 60e3) return json({error: 'Masih ada build berjalan.'}, 409);
  const src = (await listAll(env, 'store/versions/' + ts + '/')).filter(o => !o.key.endsWith('_note.json'));
  if(!src.length) return json({error: 'Versi tidak ditemukan'}, 404);
  await snapshotCurrent(env, 'sebelum dikembalikan ke ' + ts + ' oleh ' + data.email);   // rollback itself can be undone
  for(const o of src){ const x = await env.MASTER.get(o.key); await env.MASTER.put('store/current/' + o.key.split('/').slice(3).join('/'), await x.arrayBuffer()); }
  await writeJSON(env, 'state/options.json', {});
  const name = 'kembalikan master ke versi ' + ts;
  const state = {state: 'queued', by: data.email, at: new Date().toISOString(), name};
  await writeJSON(env, 'state/build.json', state);
  try { await startBuild(env, name + ' oleh ' + data.email); }
  catch(e) { state.state = 'failed'; state.error = e.message; await writeJSON(env, 'state/build.json', state); return json(state, 502); }
  return json(state);
}
