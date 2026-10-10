// Every admin API call: allowlisted admin signed in through Cloudflare Access.
import {adminEmail, json} from '../../../lib/admin.js';

export async function onRequest(context){
  const email = await adminEmail(context.request, context.env);
  if(!email) return json({error: 'Akses ditolak'}, 403);
  context.data.email = email;
  return context.next();
}
