// Viewer password change. The admin's browser re-locks the private key with the new password and sends
// only the new lock (salt, kiv, kpriv). The password itself never reaches this server.
import {json, getRepoFile, putRepoFile, readJSON, writeJSON, startBuild} from '../../../lib/admin.js';

export async function onRequestGet({env}){
  return json(JSON.parse((await getRepoFile(env, 'tools/lock.json')).text));
}

export async function onRequestPost({request, env, data}){
  const body = await request.json().catch(() => null), next = body && body.lock;
  const cur = await getRepoFile(env, 'tools/lock.json'), old = JSON.parse(cur.text);
  const b64 = v => typeof v === 'string' && /^[A-Za-z0-9+/]+=*$/.test(v);
  if(!next || next.pub !== old.pub || next.v !== old.v || next.iterations !== old.iterations || !b64(next.salt) || !b64(next.kiv) || !b64(next.kpriv) || next.salt === old.salt)
    return json({error: 'Kunci baru tidak valid'}, 400);
  const lock = {v: old.v, salt: next.salt, iterations: old.iterations, pub: old.pub, kiv: next.kiv, kpriv: next.kpriv};
  await putRepoFile(env, 'tools/lock.json', JSON.stringify(lock, null, 1) + '\n', cur.sha, 'Kata sandi penonton diganti oleh ' + data.email);
  if(!await env.MASTER.head('master/latest.xlsx'))
    return json({ok: true, note: 'Kunci tersimpan. Upload master supaya data dibangun ulang dengan kata sandi baru.'});
  const state = {state: 'queued', by: data.email, at: new Date().toISOString(), name: 'ganti kata sandi'};
  await writeJSON(env, 'state/build.json', state);
  await startBuild(env, 'ganti kata sandi oleh ' + data.email);
  return json({ok: true, build: state});
}
