// POST (sendBeacon) from the dashboard after the viewer unlocks it: one page view.
// Stored as an empty R2 object whose key + metadata carry the view (no IP, no name): visits/<day>/<time>-<rand>.
// Read back only through /api/admin/visits (Cloudflare Access). The repo is public, the bucket is not.
const ok = (v, re) => typeof v === 'string' && re.test(v);

export async function onRequestPost({request, env}){
  const origin = request.headers.get('origin');
  if(origin && new URL(origin).host !== new URL(request.url).host) return new Response(null, {status: 403});
  let b; try { b = JSON.parse((await request.text()).slice(0, 600)); } catch(e){ return new Response(null, {status: 400}); }
  if(!ok(b.v, /^[a-z0-9]{8,16}$/) || !ok(b.s, /^[a-z0-9]{6,16}$/) || !ok(b.p, /^[^\n]{0,120}$/) || !['hp', 'tablet', 'desktop'].includes(b.d))
    return new Response(null, {status: 400});
  const t = new Date(Date.now() + 7 * 3600e3).toISOString();   // WIB
  const cf = request.cf || {};
  await env.MASTER.put('visits/' + t.slice(0, 10) + '/' + t.slice(11, 19).replace(/:/g, '') + '-' + Math.random().toString(36).slice(2, 8), '', {
    customMetadata: {v: b.v, s: b.s, p: b.p || 'ringkasan', d: b.d, c: String(cf.city || cf.country || '').slice(0, 40)}});
  return new Response(null, {status: 204});
}
