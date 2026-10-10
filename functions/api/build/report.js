// POST from the build server: {state: running|ok|failed, run_url, asOf, build, log}
import {json, readJSON, writeJSON} from '../../../lib/admin.js';

export async function onRequestPost({request, env}){
  const r = await request.json().catch(() => ({}));
  if(!['running', 'ok', 'failed'].includes(r.state)) return json({error: 'bad state'}, 400);
  const cur = await readJSON(env, 'state/build.json', {});
  const next = Object.assign({}, cur, {state: r.state, run_url: r.run_url, updated: new Date().toISOString()});
  if(r.state !== 'running') Object.assign(next, {asOf: r.asOf || null, build: r.build || null, log: String(r.log || '').slice(-8000)});
  await writeJSON(env, 'state/build.json', next);
  return json({ok: true});
}
