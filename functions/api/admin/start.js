// POST {month?, force?}: build from whatever is in the inbox (Rilis / SO / Target files). With an empty inbox
// and {rebuild: true} it rebuilds the dashboard from the current store (after a rollback or password change).
import {json, listAll, readJSON, writeJSON, startBuild} from '../../../lib/admin.js';

export async function onRequestPost({request, env, data}){
  const b = await request.json().catch(() => ({}));
  const cur = await readJSON(env, 'state/build.json', {});
  if(cur.state === 'queued' || cur.state === 'running'){
    if(Date.now() - Date.parse(cur.updated || cur.at) < 40 * 60e3) return json({error: 'Masih ada build berjalan. Tunggu selesai.'}, 409);
  }
  if(!await env.MASTER.head('store/current/data/rilis.csv.gz')) return json({error: 'Master awal belum dimuat (bagian 0).'}, 409);
  const inbox = (await listAll(env, 'inbox/')).map(o => o.key.slice(6).replace(/\.xlsx$/, ''));
  const tgt = ['tactical', 'snop', 'monitoring'].filter(k => inbox.includes(k));
  if(tgt.length && tgt.length < 3) return json({error: 'Target butuh tiga file: Tactical, SNOP, dan Monitoring.'}, 400);
  if(!inbox.length && !b.rebuild) return json({error: 'Belum ada file yang diupload.'}, 400);
  const month = b.month ? Math.max(1, Math.min(12, +b.month | 0)) : null;
  await writeJSON(env, 'state/options.json', {month, force: !!b.force});
  const name = inbox.length ? inbox.join(' + ') + (month && tgt.length ? ' (target bulan ' + month + ')' : '') : 'build ulang dari master tersimpan';
  const state = {state: 'queued', by: data.email, at: new Date().toISOString(), name};
  await writeJSON(env, 'state/build.json', state);
  try { await startBuild(env, name + ' oleh ' + data.email); }
  catch(e) { state.state = 'failed'; state.error = e.message; await writeJSON(env, 'state/build.json', state); return json(state, 502); }
  const hist = await readJSON(env, 'state/history.json', []);
  hist.unshift({name, by: data.email, at: state.at});
  await writeJSON(env, 'state/history.json', hist.slice(0, 30));
  return json(state);
}
