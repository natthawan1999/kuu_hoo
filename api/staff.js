// KUUHOO — /api/staff
// เรียก Supabase ผ่าน REST ตรง ๆ ไม่ใช้ SDK — ฟังก์ชันจะไม่พังตอนโหลดถ้า env ไม่ครบ
// แต่จะตอบเป็น JSON บอกสาเหตุแทน
//
//   GET  /api/staff                    → รายชื่อทั้งหมด + จำนวนใบที่ส่ง
//   GET  /api/staff?feature=recorder   → เฉพาะคนที่เปิดใช้งานและมีสิทธิ์ฟีเจอร์นั้น
//   GET  /api/staff?debug=1            → เช็กว่า env ครบไหม (ไม่โชว์ค่า key)
//   POST /api/staff  { action: add | update | setActive, ... }
//
// env ที่ต้องมีบน Vercel: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY

const FEATURES = { recorder: 'allow_recorder', compare: 'allow_compare', invoice: 'allow_invoice' };

function env() {
  const url = (process.env.SUPABASE_URL || '').replace(/\/$/, '');
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || '';
  const missing = [];
  if (!url) missing.push('SUPABASE_URL');
  if (!key) missing.push('SUPABASE_SERVICE_ROLE_KEY');
  return { url, key, missing };
}

async function sb(path, init = {}) {
  const { url, key } = env();
  const res = await fetch(`${url}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      'Content-Type': 'application/json',
      Prefer: init.method === 'POST' || init.method === 'PATCH' ? 'return=representation' : '',
      ...(init.headers || {}),
    },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) {
    const msg = body?.message || body?.hint || `HTTP ${res.status}`;
    const err = new Error(msg);
    err.status = res.status;
    err.code = body?.code;
    throw err;
  }
  return body;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  try {
    const { missing, url } = env();
    if (req.query?.debug) {
      return res.status(200).json({
        ok: missing.length === 0,
        missing,
        url_set: !!url,
        url_host: url ? url.replace(/^https?:\/\//, '').split('.')[0] + '.supabase.co' : null,
        node: process.version,
      });
    }
    if (missing.length) {
      return res.status(500).json({
        error: `ยังไม่ได้ตั้งค่า env บน Vercel: ${missing.join(', ')} — ตั้งแล้วต้อง Redeploy ด้วย`,
      });
    }

    if (req.method === 'GET') {
      const { feature } = req.query;

      // หน้าเลือกชื่อ — เฉพาะคนที่ active และมีสิทธิ์ฟีเจอร์นั้น
      if (feature) {
        const col = FEATURES[feature];
        if (!col) return res.status(400).json({ error: `feature ไม่ถูกต้อง: ${feature}` });
        const rows = await sb(`staff?select=id,name,initial,dept&active=eq.true&${col}=eq.true&order=name.asc`);
        return res.status(200).json({ staff: rows || [] });
      }

      // หน้าจัดการพนักงาน — ทั้งหมด พร้อมจำนวนใบที่เคยส่ง
      const rows = await sb('staff?select=*&order=name.asc');
      let tally = {};
      try {
        const subs = await sb('count_submissions?select=counter_id');
        for (const s of subs || []) if (s.counter_id) tally[s.counter_id] = (tally[s.counter_id] || 0) + 1;
      } catch { /* ยังไม่มีตารางใบนับก็ไม่เป็นไร */ }

      return res.status(200).json({
        staff: (rows || []).map(s => ({ ...s, submission_count: tally[s.id] || 0 })),
      });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { action } = body;

      if (action === 'add') {
        const { name, dept, initial, allow_recorder, allow_compare, allow_invoice } = body;
        if (!name?.trim()) return res.status(400).json({ error: 'ต้องมีชื่อ' });

        const existing = await sb('staff?select=id&id=like.u*');
        const next = 'u' + ((existing || []).reduce((m, r) => Math.max(m, parseInt(String(r.id).slice(1)) || 0), 0) + 1);

        try {
          const out = await sb('staff', {
            method: 'POST',
            body: JSON.stringify({
              id: next,
              name: name.trim(),
              initial: (initial || name.trim()).charAt(0),
              dept: dept || null,
              allow_recorder: !!allow_recorder,
              allow_compare: !!allow_compare,
              allow_invoice: !!allow_invoice,
            }),
          });
          return res.status(200).json({ staff: Array.isArray(out) ? out[0] : out });
        } catch (e) {
          if (e.code === '23505' || /duplicate|unique/i.test(e.message)) {
            return res.status(409).json({ error: `มีชื่อ "${name.trim()}" อยู่แล้ว` });
          }
          throw e;
        }
      }

      if (action === 'update') {
        const { id } = body;
        if (!id) return res.status(400).json({ error: 'ต้องมี id' });
        const patch = {};
        for (const k of ['name', 'initial', 'dept', 'allow_recorder', 'allow_compare', 'allow_invoice'])
          if (k in body) patch[k] = body[k];
        const out = await sb(`staff?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify(patch) });
        return res.status(200).json({ staff: Array.isArray(out) ? out[0] : out });
      }

      if (action === 'setActive') {
        const { id, active } = body;
        if (!id) return res.status(400).json({ error: 'ต้องมี id' });
        // ปิดการใช้งาน ไม่ใช่ลบ — ใบเก่าต้องมีชื่อกำกับไว้
        const out = await sb(`staff?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify({ active: !!active }) });
        return res.status(200).json({ staff: Array.isArray(out) ? out[0] : out });
      }

      return res.status(400).json({ error: `action ไม่ถูกต้อง: ${action}` });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET กับ POST' });
  } catch (e) {
    const msg = e?.message || 'เกิดข้อผิดพลาด';
    const hint = /relation .*staff.* does not exist|could not find the table/i.test(msg)
      ? ' — ยังไม่ได้รัน sql/2-staff-setup.sql บน Supabase'
      : /invalid api key|jwt/i.test(msg)
      ? ' — SUPABASE_SERVICE_ROLE_KEY ไม่ถูกต้อง (ต้องเป็น service_role ไม่ใช่ anon)'
      : '';
    return res.status(500).json({ error: msg + hint });
  }
}
