// GET: the latest uploaded master Excel, for the build server.
export async function onRequestGet({env}){
  const o = await env.MASTER.get('master/latest.xlsx');
  if(!o) return new Response('no master uploaded', {status: 404});
  return new Response(o.body, {headers: {'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'cache-control': 'no-store'}});
}
