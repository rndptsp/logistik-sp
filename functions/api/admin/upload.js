// POST the master Excel (raw body, header x-filename) -> R2 master/latest.xlsx, then start the build.
import {json, readJSON, writeJSON, startBuild} from '../../../lib/admin.js';

const MAX_MB = 60;

export async function onRequestPost({request, env, data}){
  const name = decodeURIComponent(request.headers.get('x-filename') || 'master.xlsx');
  if(!/\.xlsx$/i.test(name)) return json({error: 'File harus .xlsx'}, 400);
  const buf = await request.arrayBuffer();
  if(buf.byteLength > MAX_MB * 1e6) return json({error: 'File lebih dari ' + MAX_MB + ' MB'}, 413);
  const head = new Uint8Array(buf, 0, 2);
  if(buf.byteLength < 100 || head[0] !== 0x50 || head[1] !== 0x4b) return json({error: 'Bukan file Excel (.xlsx) yang valid'}, 400);
  const at = new Date().toISOString();
  await env.MASTER.put('master/latest.xlsx', buf, {customMetadata: {name, by: data.email, at}});
  const hist = await readJSON(env, 'state/history.json', []);
  hist.unshift({name, by: data.email, at, size: buf.byteLength});
  await writeJSON(env, 'state/history.json', hist.slice(0, 30));
  const state = {state: 'queued', by: data.email, at, name};
  await writeJSON(env, 'state/build.json', state);
  try { await startBuild(env, 'upload ' + name + ' oleh ' + data.email); }
  catch(e) { state.state = 'failed'; state.error = e.message; await writeJSON(env, 'state/build.json', state); return json(state, 502); }
  return json(state);
}
