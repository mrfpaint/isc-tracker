// MRF FIELD ONE — Site Photo Verification (Cloudflare Worker: mrf-photos)
//
// Stores site-visit photos in R2 and verifies each one with Claude vision against what the
// entry claims (site name, building type, stage). The app uploads on entry save and shows the
// verdict; managers see photo + verdict in the entry Detail view.
//
// ── Setup (dashboard) ────────────────────────────────────────────────────────
//  R2                    : create bucket  fieldone-photos   (R2 → Create bucket, no public access needed)
//  Settings → Bindings   : R2 bucket binding, name PHOTOS → fieldone-photos
//  Settings → Secrets    : ANTHROPIC_API_KEY = <your Anthropic API key (same console as mrf-ai)>
//  Settings → Variables  : MODEL = claude-haiku-4-5   (optional; default shown — cheap + has vision)
//
// Endpoints:
//  POST /verify   {refId, empId, empName, claim:{siteName,buildType,siteStage,address}, image:"data:image/jpeg;base64,..."}
//                 → stores photo + verdict → {ok, key, verdict}
//  GET  /photos?refId=E123   → [{key, ts, empName, verdict, ...}] newest first
//  GET  /img/<key>           → serves the image
//  POST /invoice  {refId, empId, empName, image}   (v4.67; image = a photo or a PDF data URL)
//                 → stores the invoice image + the amount Claude reads off it → {ok, key, ts, extract}
//                   extract = {is_invoice, invoice_number, invoice_date, seller, dealer_code, total_amount, summary}
//  GET  /invoices?refId=E123 → [{key, ts, empName, extract}] oldest first, removed ones left out
//  POST /invoice/remove {refId, ts} → hides one invoice from the list (the image itself is kept for audit)
// NO secrets in this file — safe for the public repo.

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const json = (o, s=200) => new Response(JSON.stringify(o), {status:s, headers:{'Content-Type':'application/json', ...CORS}});

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return new Response(null, {headers: CORS});
    const url = new URL(request.url);
    try {
      if (request.method === 'POST' && url.pathname === '/verify') return await verify(request, env);
      if (url.pathname === '/photos') return await listPhotos(url, env);
      if (request.method === 'POST' && url.pathname === '/invoice') return await uploadInvoice(request, env);
      if (request.method === 'POST' && url.pathname === '/invoice/remove') return await removeInvoice(request, env);
      if (url.pathname === '/invoices') return await listInvoices(url, env);
      if (url.pathname.startsWith('/img/'))  return await serveImg(url, env);
      return json({ok:true, worker:'mrf-photos'});
    } catch(e) { return json({error: e.message}, 500); }
  }
};

async function verify(request, env) {
  const {refId, empId, empName, claim = {}, image} = await request.json();
  if (!refId || !image) return json({error:'refId and image required'}, 400);
  const m = /^data:(image\/\w+);base64,(.+)$/s.exec(image);
  if (!m) return json({error:'image must be a base64 data URL'}, 400);
  const [, mime, b64] = m;
  if (b64.length > 6*1024*1024) return json({error:'image too large (max ~4MB)'}, 400);

  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const ts = Date.now();
  const key = `p/${refId}/${ts}.jpg`;
  await env.PHOTOS.put(key, bytes, {httpMetadata:{contentType: mime}});

  let verdict;
  try { verdict = await aiVerdict(env, mime, b64, claim); }
  catch(e) { verdict = {verdict:'unclear', is_site:null, flags:[], summary:'AI check unavailable: '+e.message}; }

  const meta = {key, ts, refId, empId:empId||'', empName:empName||'', claim, verdict};
  await env.PHOTOS.put(`m/${refId}/${ts}.json`, JSON.stringify(meta), {httpMetadata:{contentType:'application/json'}});
  return json({ok:true, key, verdict});
}

async function aiVerdict(env, mime, b64, claim) {
  const prompt = `You verify field-sales site photos for a paint company. The salesperson claims this photo is from:
Site: ${claim.siteName||'—'} · Building type: ${claim.buildType||'—'} · Stage: ${claim.siteStage||'—'} · Address: ${claim.address||'—'}

Assess the photo. Reply with ONLY this JSON, nothing else:
{"verdict":"plausible|suspicious|unclear",
 "is_site":true/false,
 "stage_estimate":"what construction/painting stage the photo suggests",
 "matches_claim":true/false/null,
 "flags":["photo-of-screen","selfie","indoors-unrelated","vehicle","document","too-blurred","stock-like"],
 "summary":"one short plain sentence for a manager"}
Only include flags that actually apply (empty array if none). "plausible" = a genuine site photo broadly consistent with the claim.`;

  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json'},
    body: JSON.stringify({
      model: env.MODEL || 'claude-haiku-4-5',
      max_tokens: 300,
      messages: [{role:'user', content: [
        {type:'image', source:{type:'base64', media_type:mime, data:b64}},
        {type:'text', text: prompt}
      ]}]
    })
  });
  if (!r.ok) throw new Error('Anthropic API '+r.status+': '+(await r.text()).slice(0,200));
  const d = await r.json();
  const text = (d.content||[]).map(c=>c.text||'').join('');
  const jm = text.match(/\{[\s\S]*\}/);
  if (!jm) throw new Error('no JSON in AI reply');
  return JSON.parse(jm[0]);
}

// ── INVOICES (v4.67) ── The conversion value of a site is what its invoices add up to, so the amount
// is read off the invoice image here rather than typed in by the person claiming the conversion.
async function uploadInvoice(request, env) {
  const {refId, empId, empName, image} = await request.json();
  if (!refId || !image) return json({error:'refId and image required'}, 400);
  // An invoice can be a photo or the PDF the dealer sent (WhatsApp / email) — Claude reads both.
  const m = /^data:(image\/\w+|application\/pdf);base64,(.+)$/s.exec(image);
  if (!m) return json({error:'invoice must be an image or a PDF (base64 data URL)'}, 400);
  const [, mime, b64] = m;
  if (b64.length > 6*1024*1024) return json({error:'file too large (max ~4MB)'}, 400);

  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const ts = Date.now();
  const key = `inv/${refId}/${ts}.${mime==='application/pdf' ? 'pdf' : 'jpg'}`;
  await env.PHOTOS.put(key, bytes, {httpMetadata:{contentType: mime}});

  let extract;
  try { extract = await aiInvoice(env, mime, b64); }
  catch(e) { extract = {is_invoice:null, total_amount:null, summary:'AI read unavailable: '+e.message}; }

  const meta = {key, ts, refId, empId:empId||'', empName:empName||'', extract};
  await env.PHOTOS.put(`mi/${refId}/${ts}.json`, JSON.stringify(meta), {httpMetadata:{contentType:'application/json'}});
  return json({ok:true, key, ts, extract});
}

async function aiInvoice(env, mime, b64) {
  const prompt = `This photo or PDF should be a sales invoice or bill for paint products (Indian GST invoice, amounts in rupees).
Read it and reply with ONLY this JSON, nothing else:
{"is_invoice":true/false,
 "invoice_number":"as printed, or null",
 "invoice_date":"YYYY-MM-DD, or null",
 "seller":"dealer / seller name as printed, or null",
 "dealer_code":"the dealer code printed on the invoice, exactly as printed, or null",
 "total_amount":number or null,
 "summary":"one short plain sentence for a manager"}
total_amount = the final payable grand total of the invoice INCLUDING taxes, as a plain number (no commas, no currency sign).
If the photo is not an invoice, or the grand total cannot be read with confidence, set total_amount to null.
dealer_code = the code that identifies the paint dealer on this invoice — usually labelled Dealer Code, Customer Code,
Party Code or Sold-to / Bill-to code (MRF dealer codes are normally 6 digits). Never return a PIN code, phone number,
GSTIN, HSN code or invoice number as the dealer code; if no clearly labelled dealer code is visible, use null.`;

  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json'},
    body: JSON.stringify({
      model: env.INVOICE_MODEL || env.MODEL || 'claude-haiku-4-5',
      max_tokens: 300,
      messages: [{role:'user', content: [
        mime==='application/pdf'
          ? {type:'document', source:{type:'base64', media_type:mime, data:b64}}
          : {type:'image', source:{type:'base64', media_type:mime, data:b64}},
        {type:'text', text: prompt}
      ]}]
    })
  });
  if (!r.ok) throw new Error('Anthropic API '+r.status+': '+(await r.text()).slice(0,200));
  const d = await r.json();
  const text = (d.content||[]).map(c=>c.text||'').join('');
  const jm = text.match(/\{[\s\S]*\}/);
  if (!jm) throw new Error('no JSON in AI reply');
  const out = JSON.parse(jm[0]);
  // A number only: "1,23,456.00" or "₹ 4500" are cleaned, anything else becomes null (= not counted)
  const amt = Number(String(out.total_amount ?? '').replace(/[^0-9.]/g, ''));
  out.total_amount = out.is_invoice !== false && amt > 0 ? Math.round(amt*100)/100 : null;
  return out;
}

async function listInvoices(url, env) {
  const refId = url.searchParams.get('refId');
  if (!refId) return json({error:'refId required'}, 400);
  const listing = await env.PHOTOS.list({prefix: `mi/${refId}/`});
  const out = [];
  for (const obj of listing.objects) {
    try { const o = await env.PHOTOS.get(obj.key); const m = JSON.parse(await o.text()); if (!m.removed) out.push(m); } catch(e) {}
  }
  out.sort((a,b) => (a.ts||0)-(b.ts||0));
  return json(out);
}

async function removeInvoice(request, env) {
  const {refId, ts} = await request.json();
  if (!refId || !ts) return json({error:'refId and ts required'}, 400);
  const k = `mi/${refId}/${Number(ts)}.json`;
  const o = await env.PHOTOS.get(k);
  if (!o) return json({error:'not found'}, 404);
  const m = JSON.parse(await o.text());
  m.removed = Date.now();
  await env.PHOTOS.put(k, JSON.stringify(m), {httpMetadata:{contentType:'application/json'}});
  return json({ok:true});
}

async function listPhotos(url, env) {
  const refId = url.searchParams.get('refId');
  if (!refId) return json({error:'refId required'}, 400);
  const listing = await env.PHOTOS.list({prefix: `m/${refId}/`});
  const out = [];
  for (const obj of listing.objects) {
    try { const o = await env.PHOTOS.get(obj.key); out.push(JSON.parse(await o.text())); } catch(e) {}
  }
  out.sort((a,b) => (b.ts||0)-(a.ts||0));
  return json(out);
}

async function serveImg(url, env) {
  const key = decodeURIComponent(url.pathname.slice(5));
  const obj = await env.PHOTOS.get(key);
  if (!obj) return new Response('not found', {status:404, headers:CORS});
  return new Response(obj.body, {headers:{
    'Content-Type': obj.httpMetadata?.contentType || 'image/jpeg',
    'Cache-Control': 'public, max-age=31536000, immutable',
    ...CORS
  }});
}
