// KUUHOO — /api/submission
// ใบที่ส่งแล้วขึ้นเซิร์ฟเวอร์ — พนักงานส่งจากเครื่องไหนก็ได้ ผู้จัดการรีวิวจากเครื่องไหนก็ได้
// ตาราง count_submissions (สร้างจาก supabase-setup.sql)
//
//   GET   /api/submission?feature=recorder            → ใบทั้งหมดของฟีเจอร์นั้น (ผู้จัดการ)
//   GET   /api/submission?counter_id=u1&feature=...   → เฉพาะใบของคนนั้น (พนักงาน)
//   POST  /api/submission  { submission: {...} }      → บันทึกใบใหม่
//   PATCH /api/submission  { id, status, review_note, reviewed_by }  → ผลรีวิว
//   DELETE /api/submission?id=...                     → ลบใบ (ผู้จัดการ)

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

async function sb(path, init = {}) {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: SB_KEY,
      Authorization: `Bearer ${SB_KEY}`,
      'Content-Type': 'application/json',
      Prefer: init.method === 'POST' || init.method === 'PATCH' ? 'return=representation' : '',
      ...(init.headers || {}),
    },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) { const e = new Error(body?.message || `HTTP ${res.status}`); e.code = body?.code; throw e; }
  return body;
}

// แถวใน DB → รูปแบบที่แอปใช้อยู่ (ไม่ต้องแก้หน้าจอ)
const toApp = (r) => ({
  id: r.id,
  docNo: r.doc_no,
  counter: r.counter_name,
  counterId: r.counter_id,
  featureType: r.feature_type,
  deviceId: r.device_id || '',
  startedAt: r.started_at,
  submittedAt: r.submitted_at,
  note: r.note || '',
  itemCount: r.item_count,
  totalQty: Number(r.total_qty) || 0,
  data: Array.isArray(r.lines) ? r.lines : [],
  status: r.status,
  reviewNote: r.review_note || '',
  reviewedAt: r.reviewed_at,
  reviewedBy: r.reviewed_by,
  compareAt: r.compare_at,
  compareData: r.compare_data || null,
});

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  try {
    if (!SB_URL || !SB_KEY) {
      return res.status(500).json({ error: 'ยังไม่ได้ตั้งค่า env บน Vercel: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY' });
    }

    const q = req.query || {};

    if (req.method === 'GET') {
      let path = 'count_submissions?select=*&order=submitted_at.desc&limit=500';
      if (q.feature) path += `&feature_type=eq.${encodeURIComponent(q.feature)}`;
      if (q.counter_id) path += `&counter_id=eq.${encodeURIComponent(q.counter_id)}`;
      if (q.status) path += `&status=eq.${encodeURIComponent(q.status)}`;
      const rows = await sb(path);
      return res.status(200).json({ submissions: (rows || []).map(toApp) });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const s = body.submission || body;
      if (!s?.docNo) return res.status(400).json({ error: 'ต้องมี docNo' });
      if (!Array.isArray(s.data)) return res.status(400).json({ error: 'data ต้องเป็น array' });

      const row = {
        doc_no: s.docNo,
        counter_id: s.counterId || 'unknown',
        counter_name: s.counter || 'พนักงาน',
        feature_type: s.featureType || 'recorder',
        device_id: s.deviceId || null,
        started_at: s.startedAt || null,
        submitted_at: s.submittedAt || new Date().toISOString(),
        note: s.note || '',
        item_count: s.itemCount ?? s.data.length,
        total_qty: s.totalQty ?? s.data.reduce((a, d) => a + (Number(d.qty) || 0), 0),
        lines: s.data,
        status: s.status || 'pending',
        review_note: s.reviewNote || '',
        compare_at: s.compareAt || null,
        compare_data: s.compareData || null,
      };

      try {
        const out = await sb('count_submissions', { method: 'POST', body: JSON.stringify(row) });
        return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
      } catch (e) {
        // เลขใบซ้ำ (สองเครื่องส่งพร้อมกัน) — บอกให้แอปขอเลขใหม่
        if (e.code === '23505' || /duplicate|unique/i.test(e.message)) {
          return res.status(409).json({ error: `เลขที่เอกสาร ${s.docNo} มีอยู่แล้ว — กดส่งอีกครั้งเพื่อขอเลขใหม่`, duplicate: true });
        }
        throw e;
      }
    }

    if (req.method === 'PATCH') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { id, status, review_note, reviewed_by } = body;
      if (!id) return res.status(400).json({ error: 'ต้องมี id' });
      const patch = { reviewed_at: new Date().toISOString() };
      if (status) patch.status = status;
      if (review_note !== undefined) patch.review_note = review_note || '';
      if (reviewed_by) patch.reviewed_by = reviewed_by;
      const out = await sb(`count_submissions?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify(patch) });
      return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
    }

    if (req.method === 'DELETE') {
      if (!q.id) return res.status(400).json({ error: 'ต้องมี id' });
      await sb(`count_submissions?id=eq.${encodeURIComponent(q.id)}`, { method: 'DELETE' });
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET / POST / PATCH / DELETE' });
  } catch (e) {
    const msg = e?.message || 'เกิดข้อผิดพลาด';
    const hint = /count_submissions.*does not exist|could not find the table/i.test(msg)
      ? ' — ยังไม่ได้รัน supabase-setup.sql บน Supabase' : '';
    return res.status(500).json({ error: msg + hint });
  }
}
