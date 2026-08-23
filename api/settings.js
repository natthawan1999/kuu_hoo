// KUUHOO — /api/settings
// ค่ากลางที่ทุกเครื่องใช้ร่วม (ตั้งครั้งเดียว) — เดิมตั้งทีละเครื่องใน localStorage
//
//   GET   /api/settings                → { settings: { drive_client_id, drive_folder_invoice, ... } }
//   PATCH /api/settings { key, value, by }  → บันทึกค่าเดียว
//   PATCH /api/settings { settings: {...}, by }  → บันทึกหลายค่า
//
// env เป็นค่าตั้งต้น ถ้ามีใน app_settings จะใช้ค่าใน DB แทน

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

const ALLOWED = [
  'drive_client_id',        // Google Client ID
  'drive_folder_invoice',   // โฟลเดอร์บิลซื้อ
  'drive_folder_recorder',  // โฟลเดอร์ใบนับ
  'drive_folder_compare',   // โฟลเดอร์เทียบยอด
  'invoice_model',          // โมเดลที่ใช้อ่านบิล
];

const DEFAULTS = {
  drive_client_id:       process.env.VITE_GOOGLE_CLIENT_ID || '',
  drive_folder_invoice:  process.env.VITE_DRIVE_FOLDER_ID || '',
  drive_folder_recorder: process.env.DRIVE_FOLDER_RECORDER || '',
  drive_folder_compare:  process.env.DRIVE_FOLDER_COMPARE || '',
  invoice_model:         '',
};

async function sb(path, init = {}) {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, 'Content-Type': 'application/json', ...(init.headers || {}) },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) throw new Error(body?.message || `HTTP ${res.status}`);
  return body;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!SB_URL || !SB_KEY) return res.status(500).json({ error: 'ยังไม่ได้ตั้ง SUPABASE_URL / SERVICE_ROLE_KEY' });

  try {
    if (req.method === 'GET') {
      let rows = [];
      try { rows = await sb('app_settings?select=key,value,updated_at,updated_by'); } catch { /* ยังไม่ได้รัน sql/14 */ }
      const settings = { ...DEFAULTS };
      const meta = {};
      for (const r of rows || []) {
        if (!ALLOWED.includes(r.key)) continue;
        if (r.value) settings[r.key] = r.value;
        meta[r.key] = { updatedAt: r.updated_at, updatedBy: r.updated_by };
      }
      return res.status(200).json({ settings, meta });
    }

    if (req.method === 'PATCH' || req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const incoming = body.settings || (body.key ? { [body.key]: body.value } : null);
      if (!incoming) return res.status(400).json({ error: 'ต้องมี key/value หรือ settings' });

      const rows = Object.entries(incoming)
        .filter(([k]) => ALLOWED.includes(k))
        .map(([key, value]) => ({ key, value: value == null ? '' : String(value).trim(),
                                  updated_at: new Date().toISOString(), updated_by: body.by || null }));
      if (!rows.length) return res.status(400).json({ error: 'ไม่มี key ที่รองรับ', allowed: ALLOWED });

      await sb('app_settings?on_conflict=key', {
        method: 'POST', headers: { Prefer: 'resolution=merge-duplicates' }, body: JSON.stringify(rows),
      });
      return res.status(200).json({ ok: true, saved: rows.map(r => r.key) });
    }

    return res.status(405).json({ error: 'method ไม่รองรับ' });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
