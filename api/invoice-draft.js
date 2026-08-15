// KUUHOO — /api/invoice-draft
// ร่างบิลที่คีย์ค้างไว้ — ถ่าย/อ่านที่เครื่องหนึ่ง ไปแก้ต่อเครื่องอื่นได้
// เก็บผลที่ AI อ่านได้ (ai_result) ไม่เก็บรูป — รูปอยู่ในเครื่องที่ถ่าย
// ตาราง invoice_drafts (สร้างจาก invoice-setup.sql)
//
//   GET    /api/invoice-draft?keyed_by_id=u1     → ร่างของคนนั้น
//   POST   /api/invoice-draft { keyed_by_id, keyed_by, device_id, invoices:[...] }
//   DELETE /api/invoice-draft?keyed_by_id=u1     → ล้างร่าง (บันทึกเป็นบิลจริงแล้ว)

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

async function sb(path, init = {}) {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: SB_KEY,
      Authorization: `Bearer ${SB_KEY}`,
      'Content-Type': 'application/json',
      Prefer: init.method === 'POST' ? 'return=representation' : '',
      ...(init.headers || {}),
    },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) throw new Error(body?.message || `HTTP ${res.status}`);
  return body;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  try {
    if (!SB_URL || !SB_KEY) {
      return res.status(500).json({ error: 'ยังไม่ได้ตั้งค่า env บน Vercel: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY' });
    }
    const q = req.query || {};

    if (req.method === 'GET') {
      if (!q.keyed_by_id) return res.status(400).json({ error: 'ต้องมี keyed_by_id' });
      const rows = await sb(`invoice_drafts?select=*&keyed_by_id=eq.${encodeURIComponent(q.keyed_by_id)}&order=group_no.asc,page_no.asc`);
      return res.status(200).json({
        invoices: (rows || []).map(r => r.ai_result).filter(Boolean),
        savedAt: rows?.length ? rows[rows.length - 1].created_at : null,
      });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { keyed_by_id, keyed_by, device_id, invoices } = body;
      if (!keyed_by_id || !keyed_by) return res.status(400).json({ error: 'ต้องมี keyed_by_id และ keyed_by' });
      if (!Array.isArray(invoices)) return res.status(400).json({ error: 'invoices ต้องเป็น array' });

      await sb(`invoice_drafts?keyed_by_id=eq.${encodeURIComponent(keyed_by_id)}`, { method: 'DELETE' });

      if (invoices.length) {
        const rows = invoices.map((inv, i) => ({
          keyed_by_id, keyed_by,
          device_id: device_id || null,
          group_no: i + 1,
          page_no: 1,
          image_path: inv?.fileName || inv?.file_name || 'local',
          ai_result: inv,
        }));
        await sb('invoice_drafts', { method: 'POST', body: JSON.stringify(rows) });
      }
      return res.status(200).json({ ok: true, saved: invoices.length, savedAt: new Date().toISOString() });
    }

    if (req.method === 'DELETE') {
      if (!q.keyed_by_id) return res.status(400).json({ error: 'ต้องมี keyed_by_id' });
      await sb(`invoice_drafts?keyed_by_id=eq.${encodeURIComponent(q.keyed_by_id)}`, { method: 'DELETE' });
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET / POST / DELETE' });
  } catch (e) {
    const msg = e?.message || 'เกิดข้อผิดพลาด';
    const hint = /invoice_drafts.*does not exist|could not find the table/i.test(msg)
      ? ' — ยังไม่ได้รัน invoice-setup.sql บน Supabase' : '';
    return res.status(500).json({ error: msg + hint });
  }
}
