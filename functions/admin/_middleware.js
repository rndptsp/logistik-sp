// The admin page itself: only allowlisted admins signed in through Cloudflare Access.
import {adminEmail} from '../../lib/admin.js';

export async function onRequest(context){
  if(!await adminEmail(context.request, context.env))
    return new Response('Akses ditolak. Halaman ini hanya untuk admin.', {status: 403, headers: {'content-type': 'text/plain; charset=utf-8'}});
  return context.next();
}
