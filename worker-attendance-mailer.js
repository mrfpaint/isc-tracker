// MRF FIELD ONE — Daily 10:00 IST Attendance Email (Cloudflare Worker: mrf-attendance-mailer)
//
// "Agent Vikram" reports today's attendance down the full hierarchy: NSH → SH/NH → RSMs → ASMs,
// each row showing the person's OWN status plus Present/Leave/Holiday/Unmarked rolled up over
// everyone below them. Sent FROM fieldone@mrfpaint.com via Microsoft Graph (Mail.Send app).
//
// Org rules (2026-08-12): heads live in the RSMs sheet with a coarse region — 'All India' = NSH,
// 'South' = SH (South I–V), 'North' = NH (North I–II + East + West). There is NO East/West head.
// IMPORTANT: bare 'East' / 'West' are REAL RSM regions (e.g. Biswajit Chakraborty = 'East') and
// must never be treated as head markers.
//
// ── Cloudflare setup (dashboard) ─────────────────────────────────────────────
//  Settings → Bindings   : Service Binding, name SHEETS → worker mrf-isc-tracker
//  Settings → Variables  : TENANT_ID   = <Azure AD tenant id>
//                          CLIENT_ID   = <app registration client id>
//                          SENDER      = fieldone@mrfpaint.com   (shared mailbox works — no license needed)
//                          RECIPIENTS  = joshy.francis@mrfpaint.com; <nsh email>
//                          TEST_TOKEN  = <any long random string, for manual test sends>
//  Settings → Secrets    : CLIENT_SECRET = <app registration client secret value>
//  Settings → Triggers   : Cron = 30 4 * * *   (04:30 UTC = 10:00 IST daily)
//
// Manual test: open  https://<this-worker-url>/send?token=<TEST_TOKEN>
// NO secrets in this file — safe for the public repo.

// Cron triggers to configure (Settings → Triggers):
//   30 4 * * *   → 10:00 IST daily report to management
//   30 3 * * 1   → 09:00 IST Monday weekly digest
export default {
  async scheduled(event, env, ctx) {
    const cron = event.cron || '';
    if (cron === '30 3 * * 1') ctx.waitUntil(sendWeeklyDigest(env));
    else ctx.waitUntil(sendAttendanceMail(env));
  },
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/send') {
      if (!env.TEST_TOKEN || url.searchParams.get('token') !== env.TEST_TOKEN)
        return new Response('forbidden', {status: 403});
      try {
        // ?mini=<KB> pads a probe body to that size (mailbox limit testing)
        const mini = url.searchParams.get('mini');
        if (mini) {
          const kb = Math.min(parseInt(mini)||1, 500);
          const pad = kb>1 ? `<div style="display:none">${'probe '.repeat(Math.ceil(kb*1024/6))}</div>` : '';
          const token = await graphToken(env);
          await graphSend(env, token, `Agent Vikram — size probe ${kb}KB`, `<p>🕵️ Probe body ≈ ${kb}KB. If you can read this, this size clears the limit.</p>${pad}`);
          return json({ok:true, mode:'mini', approxKB:kb});
        }
        const mode = url.searchParams.get('mode') || 'daily';
        const r = mode==='digest' ? await sendWeeklyDigest(env) : await sendAttendanceMail(env);
        return json(r);
      }
      catch(e) { return json({error: e.message}, 500); }
    }
    return new Response('mrf-attendance-mailer ok');
  }
};

const json = (o, s=200) => new Response(JSON.stringify(o), {status:s, headers:{'Content-Type':'application/json'}});

// ── data access (via service binding to mrf-isc-tracker) ──
async function readSheet(env, name) {
  const resp = await env.SHEETS.fetch('https://sheets/sheets/read?sheet=' + encodeURIComponent(name));
  if (!resp.ok) throw new Error(`read ${name} failed: ${resp.status}`);
  const d = await resp.json();
  const rows = d.values || [];
  if (rows.length < 2) return [];
  const h = rows[0];
  return rows.slice(1)
    .filter(r => r.some(c => c !== '' && c !== undefined && c !== null))
    .map(r => { const o = {}; h.forEach((k,i) => o[k] = r[i] ?? ''); return o; })
    .filter(o => String(o.status||'') !== 'Purged');   // tombstoned leavers stay invisible
}

const live = r => !['inactive','vacant','purged'].includes(String(r.status||'Active').toLowerCase());
// Heads: ONLY these coarse markers. Bare 'East'/'West' are real RSM regions, not heads.
const headZone = r => { const g = String(r.region||'').trim().toLowerCase();
  if (g==='all india' || g==='all') return 'ALL';
  if (g==='south' || g==='north') return g;
  return null; };
// Zone buckets: SH covers South*, NH covers North* + East* + West*
const bucketOf = r => { const g = String(r.region||'').trim().toLowerCase();
  if (g.startsWith('south')) return 'south';
  if (g.startsWith('north') || g.startsWith('east') || g.startsWith('west')) return 'north';
  return ''; };

function todayIST() { return new Date(Date.now() + 5.5*3600*1000).toISOString().slice(0,10); }
// ── HOLIDAYS ── Sundays for everyone, plus the company holidays the MRF 2026 consolidated list marks for
// EVERY region. Mirror of PUBLIC_HOLIDAYS in index.html — keep the two in sync. On a holiday a blank is
// 'Holiday' (nobody is chased) and the weekly % is over working days only.
const PUBLIC_HOLIDAYS = {
  '2026-01-01':"New Year's Day", '2026-01-26':'Republic Day', '2026-05-01':'May Day',
  '2026-08-15':'Independence Day', '2026-10-02':'Gandhi Jayanti',
};
const holidayName = ds => PUBLIC_HOLIDAYS[ds] || (new Date(ds+'T00:00:00Z').getUTCDay()===0 ? 'Sunday' : '');
function esc(s){ return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// ── the report ──
async function sendAttendanceMail(env) {
  const [rsmSheet, asms, sts, bdos, iscs, att] = await Promise.all(
    ['RSMs','ASMs','STs','BDOs','Users','Attendance'].map(s => readSheet(env, s)));

  const today = todayIST();
  const mark = {};   // personId -> today's status
  att.forEach(a => { if (String(a.date||'') === today) mark[String(a.iscId)] = a.status; });
  const holToday = holidayName(today);
  const OWN   = id => mark[String(id)] || (holToday ? 'Holiday' : 'Unmarked');   // a blank on a holiday is a holiday, not a miss
  const badge = s => s==='Present' ? '✅ Present' : s==='Leave' ? '🏖️ Leave' : s==='Holiday' ? '🎉 Holiday' : '❌ Unmarked';

  const heads    = rsmSheet.filter(r => live(r) && headZone(r));
  const realRsms = rsmSheet.filter(r => live(r) && !headZone(r));
  const liveAsms = asms.filter(live), liveStaff = [...sts.filter(live), ...bdos.filter(live)], liveIscs = iscs.filter(live);

  const underASM = a => [a,
    ...liveStaff.filter(s => String(s.asmId||'') === String(a.id)),
    ...liveIscs.filter(i => String(i.asmId||'') === String(a.id))];
  const asmsOfRSM = r => liveAsms.filter(a => String(a.rsmId||'') === String(r.id));
  const underRSM  = r => [r, ...asmsOfRSM(r).flatMap(underASM)];

  const count = people => { const c = {ppl: people.length, P:0, L:0, H:0, U:0};
    people.forEach(p => { const s = OWN(p.id); if (s==='Present') c.P++; else if (s==='Leave') c.L++; else if (s==='Holiday') c.H++; else c.U++; });
    return c; };

  // Build hierarchical rows: NSH → zone heads → their RSMs → each RSM's ASMs
  const rows = [];   // {level, name, tag, sub, own, c}
  const byName = (a,b) => String(a.name||'').localeCompare(String(b.name||''));
  const pushRSM = (r, level) => {
    rows.push({level, name:r.name, tag:'RSM', sub:r.region||'', own:OWN(r.id), c:count(underRSM(r))});
    asmsOfRSM(r).sort(byName).forEach(a =>
      rows.push({level:level+1, name:a.name, tag:'ASM', sub:'', own:OWN(a.id), c:count(underASM(a))}));
  };

  const everyone = [...heads, ...realRsms.flatMap(underRSM)];
  const grand = count(everyone);
  heads.filter(h => headZone(h)==='ALL').forEach(h =>
    rows.push({level:0, name:h.name, tag:'NSH', sub:'All India', own:OWN(h.id), c:grand}));

  for (const zone of ['south','north']) {
    const zoneRsms = realRsms.filter(r => bucketOf(r)===zone).sort(byName);
    const zHeads = heads.filter(h => headZone(h)===zone);
    zHeads.forEach(h => rows.push({level:0, name:h.name, tag: zone==='south'?'SH':'NH',
      sub: zone==='south'?'South I–V':'North + East + West', own:OWN(h.id),
      c: count([h, ...zoneRsms.flatMap(underRSM)])}));
    zoneRsms.forEach(r => pushRSM(r, zHeads.length ? 1 : 0));
  }
  // Safety net: any RSM outside both buckets still gets listed
  realRsms.filter(r => !bucketOf(r)).sort(byName).forEach(r => pushRSM(r, 0));

  const unmarkedMgrs = [
    ...heads.filter(h => OWN(h.id)==='Unmarked').map(h => `${h.name} (${headZone(h)==='ALL'?'NSH':headZone(h)==='south'?'SH':'NH'})`),
    ...realRsms.filter(r => OWN(r.id)==='Unmarked').map(r => `${r.name} (RSM · ${r.region||''})`),
    ...liveAsms.filter(a => OWN(a.id)==='Unmarked').map(a => `${a.name} (ASM · ${a.rsmName||''})`)
  ];

  const pct = grand.ppl ? Math.round((grand.P + grand.L + grand.H) / grand.ppl * 100) : 0;
  const td = 'padding:6px 10px;border:1px solid #e2e8f0;font-size:13px';
  const th = td + ';background:#0c1220;color:#fff;text-align:left';
  const levelStyle = l => l===0 ? 'font-weight:700;background:#f8fafc' : l===1 ? 'font-weight:600' : 'color:#475569';
  const tagPill = t => `<span style="font-size:10px;background:${t==='NSH'||t==='SH'||t==='NH'?'#0c1220':'#e2e8f0'};color:${t==='NSH'||t==='SH'||t==='NH'?'#fff':'#334155'};border-radius:6px;padding:1px 6px">${t}</span>`;

  const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif;max-width:800px">
    <div style="background:#0c1220;color:#fff;border-radius:10px;padding:14px 18px;margin-bottom:16px">
      <div style="font-size:16px;font-weight:700">🕵️ Agent Vikram reporting, Sir.</div>
      <div style="font-size:13px;color:#cbd5e1;margin-top:4px">${holToday ? `🎉 ${esc(holToday)} — holiday. Only those who chose to work have marked; nobody is chased today.` : 'Attendance status across the field force as of 10:00 hours. Full dashboard below.'}</div>
    </div>
    <h2 style="margin:0 0 2px">FIELD ONE — Attendance @ 10:00</h2>
    <div style="color:#64748b;font-size:13px;margin-bottom:14px">${today} · ${holToday ? `${esc(holToday)} (holiday) · ${grand.P} of ${grand.ppl} people working` : `marked so far: ${pct}% of ${grand.ppl} people`}</div>
    <table style="border-collapse:collapse;width:100%">
      <tr><th style="${th}">Name</th><th style="${th}">Region</th><th style="${th}">Own</th>
          <th style="${th}">👥</th><th style="${th}">✅</th><th style="${th}">🏖️</th><th style="${th}">🎉</th><th style="${th}">❌ Unmarked</th></tr>
      ${rows.map(r => `<tr style="${levelStyle(r.level)}">
        <td style="${td};padding-left:${10 + r.level*18}px">${esc(r.name)} ${tagPill(r.tag)}</td>
        <td style="${td}">${esc(r.sub)}</td>
        <td style="${td}">${badge(r.own)}</td><td style="${td}">${r.c.ppl}</td>
        <td style="${td};color:#059669"><b>${r.c.P}</b></td><td style="${td}">${r.c.L}</td><td style="${td}">${r.c.H}</td>
        <td style="${td};color:${r.c.U?'#dc2626':'#059669'}"><b>${r.c.U}</b></td></tr>`).join('')}
      <tr><td style="${td};background:#f1f5f9"><b>Total (whole org)</b></td><td style="${td};background:#f1f5f9"></td><td style="${td};background:#f1f5f9"></td>
        <td style="${td};background:#f1f5f9"><b>${grand.ppl}</b></td><td style="${td};background:#f1f5f9;color:#059669"><b>${grand.P}</b></td>
        <td style="${td};background:#f1f5f9">${grand.L}</td><td style="${td};background:#f1f5f9">${grand.H}</td>
        <td style="${td};background:#f1f5f9;color:#dc2626"><b>${grand.U}</b></td></tr>
    </table>
    ${unmarkedMgrs.length ? `<h3 style="margin:16px 0 6px;font-size:14px">Managers not marked by 10:00 (${unmarkedMgrs.length})</h3>
      <div style="font-size:13px;color:#334155;line-height:1.7">${unmarkedMgrs.map(esc).join('<br>')}</div>` : ''}
    <div style="color:#94a3b8;font-size:12px;margin-top:18px">Over and out.<br><b>— Agent Vikram</b> · FIELD ONE · each row rolls up everyone below that person (STs/BDOs/ISCs included)</div>
  </div>`;

  const htmlKB = Math.round(html.length/1024);
  const counts = {rows:rows.length, htmlKB, sheetRows:{rsms:rsmSheet.length, asms:asms.length, sts:sts.length, bdos:bdos.length, iscs:iscs.length, att:att.length}};
  try {
    const token = await graphToken(env);
    await graphSend(env, token, holToday ? `🕵️ Agent Vikram — Attendance ${today} (${holToday} · holiday · ${grand.P} working)`
                                         : `🕵️ Agent Vikram — Attendance ${today} (${pct}% marked by 10:00)`, html);
  } catch(e) { throw new Error(e.message + ' · diagnostics: ' + JSON.stringify(counts)); }
  return {ok:true, date:today, holiday:holToday||undefined, people:grand.ppl, markedPct:pct, ...counts};
}

// ── MONDAY WEEKLY DIGEST — zone-wise attendance %, activity counts, idle ISCs, conversions ──
function addDaysIST(dstr, n) { const d = new Date(dstr+'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+n); return d.toISOString().slice(0,10); }
async function sendWeeklyDigest(env) {
  const [rsmSheet, asms, sts, bdos, iscs, att, entries, stV, bdoV, asmV, rsmV] = await Promise.all(
    ['RSMs','ASMs','STs','BDOs','Users','Attendance','Entries','ST_Visits','BDO_Visits','ASM_Visits','RSM_Visits'].map(s => readSheet(env, s)));
  const end = todayIST();                       // exclusive
  const d0 = addDaysIST(end, -7);               // this week  [d0, end)
  const p0 = addDaysIST(end, -14);              // prev week  [p0, d0)
  const inW  = d => d >= d0 && d < end;
  const inPW = d => d >= p0 && d < d0;

  const heads    = rsmSheet.filter(r => live(r) && headZone(r));
  const realRsms = rsmSheet.filter(r => live(r) && !headZone(r));
  const liveAsms = asms.filter(live), liveStaff = [...sts.filter(live), ...bdos.filter(live)], liveIscs = iscs.filter(live);
  const underRSMIds = r => {
    const myAsms = liveAsms.filter(a => String(a.rsmId||'')===String(r.id));
    const aIds = new Set(myAsms.map(a=>String(a.id)));
    return [String(r.id), ...myAsms.map(a=>String(a.id)),
      ...liveStaff.filter(s=>aIds.has(String(s.asmId||''))).map(s=>String(s.id)),
      ...liveIscs.filter(i=>aIds.has(String(i.asmId||''))).map(i=>String(i.id))];
  };
  const zones = [
    {key:'south', label:'South (SH)', rsms: realRsms.filter(r=>bucketOf(r)==='south')},
    {key:'north', label:'North + East + West (NH)', rsms: realRsms.filter(r=>bucketOf(r)==='north')},
  ];
  heads.forEach(h => { const z = headZone(h); const zz = zones.find(x=>x.key===z); if (zz) zz.label += ` — ${h.name}`; });

  const visitsFlat = [...entries.map(e=>({pid:String(e.iscId||''), date:String(e.createdAt||'').slice(0,10)})),
    ...[...stV,...bdoV,...asmV,...rsmV].map(v=>({pid:String(v.empId||v.stId||v.bdoId||v.userId||''), date:String(v.visitDate||v.createdAt||'').slice(0,10)}))];
  const isConf = s => /confirm/i.test(String(s||''));

  // Working days in each week — Sundays and company holidays don't count against anyone, and marks made on
  // them are extra effort rather than expected, so they stay out of both numerator and denominator.
  const workDays = (from, n) => Array.from({length:n}, (_,i)=>addDaysIST(from,i)).filter(d=>!holidayName(d)).length;
  const wdW = workDays(d0, 7), wdPW = workDays(p0, 7);
  const onWorkDay = a => !holidayName(String(a.date||'').slice(0,10));
  const zoneRow = z => {
    const ids = new Set(z.rsms.flatMap(underRSMIds));
    const ppl = ids.size || 1;
    const attW  = att.filter(a=>ids.has(String(a.iscId)) && inW(String(a.date||'')) && onWorkDay(a)).length;
    const attPW = att.filter(a=>ids.has(String(a.iscId)) && inPW(String(a.date||'')) && onWorkDay(a)).length;
    const actW  = visitsFlat.filter(v=>ids.has(v.pid) && inW(v.date)).length;
    const actPW = visitsFlat.filter(v=>ids.has(v.pid) && inPW(v.date)).length;
    const zoneIscs = liveIscs.filter(i=>ids.has(String(i.id)));
    const activeIscIds = new Set(entries.filter(e=>inW(String(e.createdAt||'').slice(0,10))).map(e=>String(e.iscId)));
    const idle = zoneIscs.filter(i=>!activeIscIds.has(String(i.id)));
    const convW  = entries.filter(e=>ids.has(String(e.iscId)) && inW(String(e.createdAt||'').slice(0,10)) && isConf(e.status)).length;
    const convPW = entries.filter(e=>ids.has(String(e.iscId)) && inPW(String(e.createdAt||'').slice(0,10)) && isConf(e.status)).length;
    const newW = entries.filter(e=>ids.has(String(e.iscId)) && inW(String(e.createdAt||'').slice(0,10))).length;
    return {label:z.label, ppl:ids.size, attPct:Math.round(attW/(ppl*(wdW||1))*100), attPctPrev:Math.round(attPW/(ppl*(wdPW||1))*100),
            actW, actPW, idleCount:idle.length, idleNames:idle.slice(0,8).map(i=>i.name), newW, convW, convPW};
  };
  const rows = zones.map(zoneRow);
  const arrow = (now, prev) => now>prev ? `<span style="color:#059669">▲ +${now-prev}</span>` : now<prev ? `<span style="color:#dc2626">▼ ${now-prev}</span>` : `<span style="color:#64748b">＝</span>`;

  const td = 'padding:6px 10px;border:1px solid #e2e8f0;font-size:13px';
  const th = td + ';background:#0c1220;color:#fff;text-align:left';
  const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif;max-width:800px">
    <div style="background:#0c1220;color:#fff;border-radius:10px;padding:14px 18px;margin-bottom:16px">
      <div style="font-size:16px;font-weight:700">🕵️ Agent Vikram — weekly debrief, Sir.</div>
      <div style="font-size:13px;color:#cbd5e1;margin-top:4px">Field week ${d0} → ${addDaysIST(end,-1)}, compared with the week before.</div>
    </div>
    <table style="border-collapse:collapse;width:100%">
      <tr><th style="${th}">Zone</th><th style="${th}">👥</th><th style="${th}">Attendance %</th>
          <th style="${th}">Activities</th><th style="${th}">New sites</th><th style="${th}">Converted</th><th style="${th}">Idle ISCs</th></tr>
      ${rows.map(r=>`<tr>
        <td style="${td}"><b>${esc(r.label)}</b></td><td style="${td}">${r.ppl}</td>
        <td style="${td}">${r.attPct}% ${arrow(r.attPct, r.attPctPrev)}</td>
        <td style="${td}">${r.actW} ${arrow(r.actW, r.actPW)}</td>
        <td style="${td}">${r.newW}</td>
        <td style="${td}">${r.convW} ${arrow(r.convW, r.convPW)}</td>
        <td style="${td};color:${r.idleCount?'#dc2626':'#059669'}"><b>${r.idleCount}</b></td></tr>`).join('')}
    </table>
    ${rows.filter(r=>r.idleCount).map(r=>`<h3 style="margin:14px 0 4px;font-size:13px">${esc(r.label)} — ISCs with zero site entries this week (${r.idleCount})</h3>
      <div style="font-size:12px;color:#334155;line-height:1.6">${r.idleNames.map(esc).join(' · ')}${r.idleCount>8?' · +'+(r.idleCount-8)+' more':''}</div>`).join('')}
    <div style="color:#94a3b8;font-size:12px;margin-top:18px">Attendance % = marks on working days ÷ (people × ${wdW} working days; Sundays and company holidays excluded). Over and out.<br><b>— Agent Vikram</b> · FIELD ONE</div>
  </div>`;

  const token = await graphToken(env);
  await graphSend(env, token, `🕵️ Agent Vikram — Weekly Digest (${d0} → ${addDaysIST(end,-1)})`, html);
  return {ok:true, mode:'digest', week:[d0, addDaysIST(end,-1)], zones:rows.map(r=>({zone:r.label, ppl:r.ppl, attPct:r.attPct, activities:r.actW, idle:r.idleCount}))};
}

// ── Microsoft Graph (client credentials → sendMail as SENDER) ──
async function graphToken(env) {
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID, client_secret: env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, {method:'POST', body});
  const d = await r.json();
  if (!d.access_token) throw new Error('Graph token failed: ' + JSON.stringify(d));
  return d.access_token;
}

// Send as base64 MIME, not JSON. The JSON sendMail endpoint rejects HTML bodies as small as
// ~50KB with "ErrorMessageSizeExceeded / MapiStream.Read" (known Graph quirk); the documented
// MIME variant handles them fine.
async function graphSend(env, token, subject, html, opts={}) {
  const to = opts.to || String(env.RECIPIENTS||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  const bcc = opts.bcc || [];
  if (!to.length) throw new Error('RECIPIENTS variable is empty');
  const mime = [
    `From: ${env.SENDER}`,
    `To: ${to.join(', ')}`,
    ...(bcc.length ? [`Bcc: ${bcc.join(', ')}`] : []),
    `Subject: =?UTF-8?B?${b64(subject)}?=`,   // RFC 2047 so the emoji survives
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: base64',
    '',
    wrap76(b64(html))
  ].join('\r\n');
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(env.SENDER)}/sendMail`, {
    method: 'POST',
    headers: {Authorization: 'Bearer '+token, 'Content-Type': 'text/plain'},
    body: b64(mime)
  });
  if (r.status !== 202) throw new Error('Graph sendMail failed: ' + r.status + ' ' + await r.text());
}
function b64(str) { const bytes = new TextEncoder().encode(str); let bin=''; for (const b of bytes) bin += String.fromCharCode(b); return btoa(bin); }
function wrap76(s) { return s.replace(/(.{76})/g, '$1\r\n'); }
