// Build server: GET /api/build/store -> file list; GET /api/build/store/<path> -> the file (store/current).
import {json, okPath, STORE_FILES} from '../../../../lib/admin.js';

export async function onRequestGet({env, params}){
  const p = [].concat(params.path || []).join('/');
  if(!p){
    const files = [];
    for(const f of STORE_FILES){ const h = await env.MASTER.head('store/current/' + f); if(h) files.push({path: f, size: h.size}); }
    return json({files});
  }
  if(!okPath(p)) return json({error: 'path'}, 400);
  const o = await env.MASTER.get('store/current/' + p);
  if(!o) return json({error: 'tidak ada'}, 404);
  return new Response(o.body, {headers: {'content-type': 'application/octet-stream', 'content-length': String(o.size), 'cache-control': 'no-store'}});
}
