// Build server: PUT /api/build/staging/<path> -> store/staging/<path> (a finished run, not yet the current store).
import {json, okPath} from '../../../../lib/admin.js';

export async function onRequestPut({request, env, params}){
  const p = [].concat(params.path || []).join('/');
  if(!okPath(p)) return json({error: 'path'}, 400);
  const buf = await request.arrayBuffer();
  if(buf.byteLength < 2 || buf.byteLength > 90e6) return json({error: 'ukuran'}, 400);
  await env.MASTER.put('store/staging/' + p, buf);
  return json({ok: true, size: buf.byteLength});
}
