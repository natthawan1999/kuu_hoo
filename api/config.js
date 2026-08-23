// KUUHOO — /api/config
// รวม 3 หน้าที่ไว้ไฟล์เดียว (Vercel Hobby จำกัด 12 serverless functions)
//
//   GET   /api/config              → ค่าเชื่อมต่อ Supabase + ค่ากลาง (settings)
//   GET   /api/config?next=IV      → เลขเอกสารใหม่จากตัวนับกลาง เช่น "IV-202608220001"
//   PATCH /api/config { key, value, by }        → บันทึกค่ากลาง 1 ค่า
//   PATCH /api/config { settings: {...}, by }   → บันทึกหลายค่า
//
// anon key ส่งออกได้ (สิทธิ์ตาม RLS) — service_role ใช้เฉพาะฝั่งนี้ ไม่เคยส่งออก
// ค่ากลางเก็บใน app_settings (sql/14) · เลขเอกสารใช้ next_doc_no() (sql/1)

const SB_URL  = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SRV_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

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
    headers: { apikey: SRV_KEY, Authorization: `Bearer ${SRV_KEY}`, 'Content-Type': 'application/json', ...(init.headers || {}) },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) throw new Error(body?.message || `HTTP ${res.status}`);
  return body;
}

export default async function handler(req, res) {
  const q = req.query || {};

  // ── เลขเอกสาร / ชื่อไฟล์ จากตัวนับกลาง ──
  if (req.method === 'GET' && q.next) {
    res.setHeader('Cache-Control', 'no-store');
    if (!SB_URL || !SRV_KEY) return res.status(500).json({ error: 'ยังไม่ได้ตั้ง SUPABASE_URL / SERVICE_ROLE_KEY' });
    const raw = String(q.next).toUpperCase();
    const prefix = /^[A-Z]{2,4}$/.test(raw) ? raw : 'IV';   // กันค่าแปลกไปสร้างแถวขยะใน doc_counters
    try {
      const out = await sb('rpc/next_doc_no', { method: 'POST', body: JSON.stringify({ p_prefix: prefix }) });
      const name = typeof out === 'string' ? out : (Array.isArray(out) ? out[0] : out?.next_doc_no);
      if (!name) throw new Error('ตัวนับไม่คืนค่า');
      return res.status(200).json(String(name));
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── บันทึกค่ากลาง ──
  if (req.method === 'PATCH' || req.method === 'POST') {
    res.setHeader('Cache-Control', 'no-store');
    if (!SB_URL || !SRV_KEY) return res.status(500).json({ error: 'ยังไม่ได้ตั้ง SUPABASE_URL / SERVICE_ROLE_KEY' });
    try {
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
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── ค่าเชื่อมต่อ + ค่ากลาง ──
  res.setHeader('Cache-Control', 'no-store');   // ค่ากลางต้องสด ไม่ cache

  const url = (process.env.VITE_SUPABASE_URL || process.env.SUPABASE_URL || '').replace(/\/$/, '');
  const anonKey = process.env.VITE_SUPABASE_ANON_KEY || process.env.SUPABASE_ANON_KEY || '';

  let settings = { ...DEFAULTS };
  if (SB_URL && SRV_KEY) {
    try {
      const rows = await sb('app_settings?select=key,value');
      for (const r of rows || []) if (ALLOWED.includes(r.key) && r.value) settings[r.key] = r.value;
    } catch { /* ยังไม่ได้รัน sql/14 — ใช้ค่าตั้งต้น */ }
  }

  if (!url || !anonKey) {
    return res.status(200).json({
      configured: false,
      missing: [!url && 'VITE_SUPABASE_URL', !anonKey && 'VITE_SUPABASE_ANON_KEY'].filter(Boolean),
      settings,
    });
  }

  return res.status(200).json({
    configured: true,
    url,
    anonKey,
    tableName: process.env.VITE_SUPABASE_TABLE || 'product_price',
    stockTableName: process.env.SUPABASE_STOCK_TABLE || process.env.VITE_SUPABASE_STOCK_TABLE || 'product_stock',
    settings,
  });
}
