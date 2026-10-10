// Calls from the build server (GitHub Action) carry BUILD_TOKEN.
import {buildAuthorized, json} from '../../../lib/admin.js';

export async function onRequest(context){
  if(!buildAuthorized(context.request, context.env)) return json({error: 'unauthorized'}, 401);
  return context.next();
}
