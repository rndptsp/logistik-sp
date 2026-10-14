// GET: last build, upload history, inbox, store, who is asking.
import {json, readJSON, listAll, listVersions, SLOTS} from '../../../lib/admin.js';

export async function onRequestGet({env, data}){
  const [build, history, inbox, meta, versions] = await Promise.all([readJSON(env, 'state/build.json', null), readJSON(env, 'state/history.json', []),
    listAll(env, 'inbox/'), readJSON(env, 'store/current/_meta.json', null), listVersions(env)]);
  return json({email: data.email, build, history, slots: SLOTS, versions,
    inbox: inbox.map(o => Object.assign({slot: o.key.slice(6).replace(/\.xlsx$/, ''), size: o.size}, o.customMetadata)),
    store: meta});
}
