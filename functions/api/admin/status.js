// GET: last build, upload history, who is asking.
import {json, readJSON} from '../../../lib/admin.js';

export async function onRequestGet({env, data}){
  const [build, history, master] = await Promise.all([readJSON(env, 'state/build.json', null), readJSON(env, 'state/history.json', []), env.MASTER.head('master/latest.xlsx')]);
  return json({email: data.email, build, history,
    master: master ? Object.assign({size: master.size}, master.customMetadata) : null});
}
