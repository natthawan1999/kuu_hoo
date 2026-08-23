// KUUHOO — /api/filename
// ชื่อไฟล์ = เลขเอกสาร (doc_no) จากตัวนับกลางเดียวกับใบนับ
// เช่น IV-202608220001 — เครื่องไหนขอก็ไม่ชนกัน
//
//   GET /api/filename            → "IV-202608220001"
//   GET /api/filename?prefix=RC  → "RC-202608220001"
//
// ใช้ next_doc_no() จาก sql/1-supabase-setup.sql (ไม่ต้องมีตัวนับแยก)

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!SB_URL || !SB_KEY) return res.status(500).json({ error: 'ยังไม่ได้ตั้ง SUPABASE_URL / SERVICE_ROLE_KEY' });

  // prefix สั้น ๆ ตัวพิมพ์ใหญ่ กันค่าแปลกปลอมไปสร้างแถวขยะใน doc_counters
  const raw = String((req.query && req.query.prefix) || 'IV').toUpperCase();
  const prefix = /^[A-Z]{2,4}$/.test(raw) ? raw : 'IV';

  try {
    const r = await fetch(`${SB_URL}/rest/v1/rpc/next_doc_no`, {
      method: 'POST',
      headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ p_prefix: prefix }),
    });
    const text = await r.text();
    let out = null;
    if (text) { try { out = JSON.parse(text); } catch { out = text; } }
    if (!r.ok) throw new Error(out?.message || `HTTP ${r.status}`);
    const name = typeof out === 'string' ? out : (Array.isArray(out) ? out[0] : out?.next_doc_no);
    if (!name) throw new Error('ตัวนับไม่คืนค่า');
    return res.status(200).json(String(name));
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
