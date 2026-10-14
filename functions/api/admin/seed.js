// One-time: load the initial master store (data/*.csv.gz + config/*.json from the Mac's MASTER BARU folder)
// into R2. Refused once a store exists; after that the store changes only through builds and rollbacks.
import {json, okPath} from '../../../lib/admin.js';

export async function onRequestPut({request, env}){
  const p = new URL(request.url).searchParams.get('path') || '';
  if(!okPath(p)) return json({error: 'Nama file tidak dikenal: ' + p}, 400);
  if(await env.MASTER.head('store/current/_meta.json')) return json({error: 'Master awal sudah dimuat. Gunakan upload harian atau Kembalikan versi.'}, 409);
  const buf = await request.arrayBuffer();
  if(buf.byteLength < 2 || buf.byteLength > 90e6) return json({error: 'Ukuran file tidak valid'}, 400);
  if(p.endsWith('.gz') && !(new Uint8Array(buf)[0] === 0x1f && new Uint8Array(buf)[1] === 0x8b)) return json({error: 'Bukan file .gz'}, 400);
  if(p.endsWith('.json')) { try { JSON.parse(new TextDecoder().decode(buf)); } catch(e) { return json({error: 'JSON tidak valid'}, 400); } }
  await env.MASTER.put('store/seed/' + p, buf);
  return json({ok: true, size: buf.byteLength});
}
export async function onRequestPost({env, data}){   // finish: all files present -> becomes store/current
  const {STORE_FILES} = await import('../../../lib/admin.js');
  if(await env.MASTER.head('store/current/_meta.json')) return json({error: 'Master awal sudah dimuat.'}, 409);
  const miss = [];
  for(const f of STORE_FILES) if(!await env.MASTER.head('store/seed/' + f)) miss.push(f);
  if(miss.length) return json({error: 'Belum lengkap, kurang: ' + miss.join(', ')}, 400);
  for(const f of STORE_FILES){ const o = await env.MASTER.get('store/seed/' + f); await env.MASTER.put('store/current/' + f, await o.arrayBuffer()); await env.MASTER.delete('store/seed/' + f); }
  await env.MASTER.put('store/current/_meta.json', JSON.stringify({at: new Date().toISOString(), by: data.email, name: 'master awal'}), {httpMetadata: {contentType: 'application/json'}});
  return json({ok: true});
}
