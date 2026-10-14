// POST from the build server after the site was published: current store -> a version (keep 7), staging -> current,
// and the inbox files that were processed are removed (only if not replaced by a newer upload meanwhile).
// body: {note, asOf, build, inbox: [{slot, etag}]}
import {json, listAll, snapshotCurrent, readJSON, writeJSON} from '../../../lib/admin.js';

export async function onRequestPost({request, env}){
  const b = await request.json().catch(() => ({}));
  const st = await listAll(env, 'store/staging/');
  if(!st.length) return json({error: 'staging kosong'}, 400);
  const state = await readJSON(env, 'state/build.json', {});
  const ts = await snapshotCurrent(env, 'sebelum ' + (state.name || b.note || 'update'));
  for(const o of st){
    const x = await env.MASTER.get(o.key); await env.MASTER.put('store/current/' + o.key.slice('store/staging/'.length), await x.arrayBuffer()); await env.MASTER.delete(o.key);
  }
  await env.MASTER.put('store/current/_meta.json', JSON.stringify({at: new Date().toISOString(), by: state.by, name: state.name, asOf: b.asOf, build: b.build}), {httpMetadata: {contentType: 'application/json'}});
  for(const f of [].concat(b.inbox || [])){
    const h = await env.MASTER.head('inbox/' + f.slot + '.xlsx');
    if(h && h.etag === f.etag) await env.MASTER.delete('inbox/' + f.slot + '.xlsx');
  }
  return json({ok: true, version: ts});
}
