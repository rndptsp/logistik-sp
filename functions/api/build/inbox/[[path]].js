// Build server (Bearer BUILD_TOKEN, see _middleware): GET /api/build/inbox -> listing + options;
// GET /api/build/inbox/<slot> -> the raw file (streamed).
import {json, listAll, readJSON, SLOTS} from '../../../../lib/admin.js';

export async function onRequestGet({env, params}){
  const slot = [].concat(params.path || [])[0];
  if(!slot){
    const l = await listAll(env, 'inbox/');
    return json({files: l.map(o => ({slot: o.key.slice(6).replace(/\.xlsx$/, ''), size: o.size, etag: o.etag})), options: await readJSON(env, 'state/options.json', {})});
  }
  if(!SLOTS[slot]) return json({error: 'slot'}, 400);
  const o = await env.MASTER.get('inbox/' + slot + '.xlsx');
  if(!o) return json({error: 'tidak ada'}, 404);
  return new Response(o.body, {headers: {'content-type': 'application/octet-stream', 'content-length': String(o.size), 'cache-control': 'no-store'}});
}
