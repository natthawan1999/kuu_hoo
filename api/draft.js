// KUUHOO — /api/draft
// ร่างการนับขึ้นเซิร์ฟเวอร์ — นับที่เครื่องหนึ่ง ไปต่อที่อีกเครื่องได้
// ตาราง count_drafts (สร้างจาก supabase-setup.sql) ผูกร่างกับ counter_id ไม่ใช่เครื่อง
//
//   GET    /api/draft?counter_id=u1&feature=recorder&branch=1   → ร่างของคนนั้นในสาขานั้น
//   POST   /api/draft  { counter_id, counter_name, feature, device_id, entries:[...] }
//          → ทับร่างเดิมของคนนั้นทั้งชุด (ลบเก่า ใส่ใหม่)
//   DELETE /api/draft?counter_id=u1&feature=recorder   → ล้างร่าง (ใช้ตอนส่งใบแล้ว)

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
    const feature = q.feature || 'recorder';

    if (req.method === 'GET') {
      if (!q.counter_id) return res.status(400).json({ error: 'ต้องมี counter_id' });
      const rows = await sb(
        `count_drafts?select=*&counter_id=eq.${encodeURIComponent(q.counter_id)}` +
        `&feature_type=eq.${encodeURIComponent(feature)}` +
        `&branch=eq.${encodeURIComponent(q.branch || '1')}&order=scanned_at.asc`
      );
      return res.status(200).json({
        entries: (rows || []).map(r => ({
          id: r.id,
          barcode: r.barcode,
          productCode: r.product_code || '',
          productName: r.product_name || '',
          unit: r.unit || '',
          price: Number(r.price) || 0,
          cost: Number(r.cost) || 0,
          qty: Number(r.qty) || 0,
          location: r.location || '',
          notFound: !!r.not_found,
          countDate: r.count_date,
          scannedAt: r.scanned_at,
          deviceId: r.device_id || '',
        })),
        savedAt: rows?.length ? rows[rows.length - 1].updated_at : null,
      });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { counter_id, counter_name, device_id, entries } = body;
      const branch = String(body.branch || '1');
      if (!counter_id || !counter_name) return res.status(400).json({ error: 'ต้องมี counter_id และ counter_name' });
      if (!Array.isArray(entries)) return res.status(400).json({ error: 'entries ต้องเป็น array' });

      // ทับทั้งชุด — ร่างในเครื่องคือความจริงตอนกดบันทึก
      await sb(`count_drafts?counter_id=eq.${encodeURIComponent(counter_id)}&feature_type=eq.${encodeURIComponent(feature)}&branch=eq.${encodeURIComponent(branch)}`, { method: 'DELETE' });

      if (entries.length) {
        const rows = entries.map(e => ({
          counter_id, counter_name,
          feature_type: feature,
          branch,
          device_id: device_id || null,
          barcode: String(e.barcode || ''),
          product_code: e.productCode || e.productId || null,
          product_name: e.productName || null,
          unit: e.unit || null,
          price: Number(e.price) || 0,
          cost: Number(e.cost) || 0,
          qty: Number(e.qty) || 0,
          location: e.location || null,
          not_found: !!e.notFound,
          count_date: e.countDate || new Date().toISOString().slice(0, 10),
          scanned_at: e.scannedAt || new Date().toISOString(),
        }));
        await sb('count_drafts', { method: 'POST', body: JSON.stringify(rows) });
      }

      return res.status(200).json({ ok: true, saved: entries.length, savedAt: new Date().toISOString() });
    }

    if (req.method === 'DELETE') {
      if (!q.counter_id) return res.status(400).json({ error: 'ต้องมี counter_id' });
      await sb(`count_drafts?counter_id=eq.${encodeURIComponent(q.counter_id)}&feature_type=eq.${encodeURIComponent(feature)}&branch=eq.${encodeURIComponent(q.branch || '1')}`, { method: 'DELETE' });
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET / POST / DELETE' });
  } catch (e) {
    const msg = e?.message || 'เกิดข้อผิดพลาด';
    const hint = /count_drafts.*does not exist|could not find the table/i.test(msg)
      ? ' — ยังไม่ได้รัน supabase-setup.sql บน Supabase' : '';
    return res.status(500).json({ error: msg + hint });
  }
}
