// Chunked upload of one raw Excel file into R2 (multipart), because one request body is capped (100 MB)
// and the Rilis export is bigger. The browser slices the file and calls, per slot:
//   POST ?slot=rilis&action=begin&name=<file>            -> {uploadId}
//   POST ?slot=rilis&action=part&uploadId=..&n=1..N      (body = chunk)  -> {partNumber, etag}
//   POST ?slot=rilis&action=complete&uploadId=..         (json {parts, name, size})
//   POST ?slot=rilis&action=abort&uploadId=..   |   DELETE ?slot=rilis
// Files land in inbox/<slot>.xlsx. Nothing starts building until /api/admin/start.
import {json, SLOTS} from '../../../lib/admin.js';

const MAX_MB = 400, PART_MIN = 5 * 1024 * 1024;
const UPLOAD_PAUSED = false;   // set to false when the Action in .github/workflows/build.yml runs the master pipeline (see ADMIN.md)

export async function onRequestPost({request, env, data}){
  if(UPLOAD_PAUSED) return json({error: 'Upload dinonaktifkan sementara.'}, 409);
  const u = new URL(request.url), slot = u.searchParams.get('slot'), action = u.searchParams.get('action'), id = u.searchParams.get('uploadId');
  if(!SLOTS[slot]) return json({error: 'Slot tidak dikenal'}, 400);
  const key = 'inbox/' + slot + '.xlsx';
  if(action === 'begin'){
    const name = (u.searchParams.get('name') || slot + '.xlsx').slice(0, 200);
    if(!/\.xlsx$/i.test(name)) return json({error: 'File harus .xlsx'}, 400);
    const mp = await env.MASTER.createMultipartUpload(key, {customMetadata: {name, by: data.email, at: new Date().toISOString()}});
    return json({uploadId: mp.uploadId});
  }
  if(!id) return json({error: 'uploadId kosong'}, 400);
  const mp = env.MASTER.resumeMultipartUpload(key, id);
  if(action === 'part'){
    const n = +u.searchParams.get('n');
    if(!(n >= 1 && n <= 10000)) return json({error: 'Nomor bagian tidak valid'}, 400);
    const buf = await request.arrayBuffer();
    if(buf.byteLength > 90e6) return json({error: 'Bagian terlalu besar'}, 413);
    if(n === 1){ const h = new Uint8Array(buf, 0, 2); if(buf.byteLength < 100 || h[0] !== 0x50 || h[1] !== 0x4b) return json({error: 'Bukan file Excel (.xlsx) yang valid'}, 400); }
    const p = await mp.uploadPart(n, buf);
    return json({partNumber: p.partNumber, etag: p.etag});
  }
  if(action === 'complete'){
    const b = await request.json().catch(() => ({}));
    if(!Array.isArray(b.parts) || !b.parts.length) return json({error: 'Daftar bagian kosong'}, 400);
    if(b.size > MAX_MB * 1e6) { await mp.abort(); return json({error: 'File lebih dari ' + MAX_MB + ' MB'}, 413); }
    const o = await mp.complete(b.parts.map(p => ({partNumber: p.partNumber, etag: p.etag})));
    return json({ok: true, size: o.size});
  }
  if(action === 'abort'){ await mp.abort().catch(() => {}); return json({ok: true}); }
  return json({error: 'action tidak dikenal'}, 400);
}

export async function onRequestDelete({request, env}){
  const slot = new URL(request.url).searchParams.get('slot');
  if(!SLOTS[slot]) return json({error: 'Slot tidak dikenal'}, 400);
  await env.MASTER.delete('inbox/' + slot + '.xlsx');
  return json({ok: true});
}
