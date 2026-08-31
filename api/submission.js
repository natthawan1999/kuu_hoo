// KUUHOO — /api/submission
// ใบที่ส่งแล้วขึ้นเซิร์ฟเวอร์ — พนักงานส่งจากเครื่องไหนก็ได้ ผู้จัดการรีวิวจากเครื่องไหนก็ได้
// ตาราง count_submissions (สร้างจาก supabase-setup.sql)
//
//   GET   /api/submission?feature=recorder            → ใบทั้งหมดของฟีเจอร์นั้น (ผู้จัดการ)
//   GET   /api/submission?counter_id=u1&feature=...   → เฉพาะใบของคนนั้น (พนักงาน)
//          submission.reviseOf = id ใบที่ถูกส่งกลับ → เลขเดิม + R1, R2 … (ไม่กินเลขใหม่)
//   POST  /api/submission  { submission: {...} }      → บันทึกใบใหม่
//   PATCH /api/submission  { id, status, review_note, reviewed_by }  → ผลรีวิว
//   DELETE /api/submission?id=...                     → ลบใบ (ผู้จัดการ)

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

// ออกเลขเอกสารจากตัวนับกลางใน DB — ทุกเครื่องขอจากที่นี่ที่เดียว จึงไม่ชนกัน
// ใบที่ถูกส่งกลับแล้วแก้มา → ใช้เลขฐานเดิมต่อท้าย R1, R2 …
// ไม่เรียก next_doc_no เพราะการนับรอบเดียวควรมีเลขฐานเดียว
async function reviseDocNo(reviseOf) {
  const cur = await sb(`count_submissions?id=eq.${encodeURIComponent(reviseOf)}&select=doc_no`);
  const orig = cur?.[0];
  if (!orig?.doc_no) throw new Error('ไม่พบใบเดิมที่จะแก้');
  const baseNo = String(orig.doc_no).replace(/R\d+$/i, '');   // สาขาอยู่ใน prefix แล้ว รอบแก้จึงคงสาขาเดิมเสมอ
  // นับรอบจากเลขที่มีอยู่จริง — ปลอดภัยกว่านับจาก revise_no ถ้ามีแถวค้าง
  const sibs = await sb(`count_submissions?doc_no=like.${encodeURIComponent(baseNo + 'R*')}&select=doc_no`);
  let max = 0;
  for (const r of sibs || []) {
    const n = parseInt(String(r.doc_no).match(/R(\d+)$/i)?.[1] || '0', 10);
    if (n > max) max = n;
  }
  return { docNo: `${baseNo}R${max + 1}`, reviseNo: max + 1 };
}

// ยังไม่ได้รัน sql/18 (ไม่มีคอลัมน์ branch) → insert ทั้งก้อนล้ม ใบเลยไม่ได้เลข
// ลองใหม่แบบตัด branch ออก ดีกว่าให้พนักงานส่งไม่ได้ทั้งวัน
async function insertRow(table, row) {
  try {
    return await sb(table, { method: 'POST', body: JSON.stringify(row) });
  } catch (e) {
    if (!/branch/i.test(e.message || '')) throw e;
    const { branch, ...rest } = row;
    return await sb(table, { method: 'POST', body: JSON.stringify(rest) });
  }
}

async function nextDocNo(prefix) {
  const out = await sb('rpc/next_doc_no', { method: 'POST', body: JSON.stringify({ p_prefix: prefix }) });
  return typeof out === 'string' ? out : (Array.isArray(out) ? out[0] : out?.next_doc_no);
}

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
  branch: r.branch || '1',
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
  reviseOf: r.revise_of || null,
  reviseNo: r.revise_no || 0,
  // สถานะส่งขึ้น Drive — อยู่กับใบ ทุกเครื่องเห็นตรงกัน
  drive: {
    status: r.drive_status || null,
    filename: r.drive_filename || '',
    url: r.drive_url || '',
    uploadedAt: r.drive_uploaded_at || null,
    error: r.drive_error || '',
    tries: r.drive_tries || 0,
  },
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
      if (q.branch)  path += `&branch=eq.${encodeURIComponent(q.branch)}`;
      if (q.counter_id) path += `&counter_id=eq.${encodeURIComponent(q.counter_id)}`;
      if (q.status) path += `&status=eq.${encodeURIComponent(q.status)}`;
      const rows = await sb(path);
      return res.status(200).json({ submissions: (rows || []).map(toApp) });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const s = body.submission || body;
      if (!Array.isArray(s.data)) return res.status(400).json({ error: 'data ต้องเป็น array' });

      // เลขเอกสารแยกตัวนับต่อสาขา — RC1-… / RC2-…
      const branch = String(s.branch || '1');
      const prefix = ((s.featureType || 'recorder') === 'stock_compare' ? 'ST' : 'RC') + branch;
      const row = {
        doc_no: null,   // เซิร์ฟเวอร์ออกให้ตอนบันทึก
        counter_id: s.counterId || 'unknown',
        counter_name: s.counter || 'พนักงาน',
        feature_type: s.featureType || 'recorder',
        branch,
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
        revise_of: s.reviseOf || null,
        revise_no: 0,
      };

      // แก้ใบที่ถูกส่งกลับ → เลขเดิม + R<n> ครั้งเดียว ไม่ต้องวนขอเลขใหม่
      if (s.reviseOf) {
        const { docNo, reviseNo } = await reviseDocNo(s.reviseOf);
        row.doc_no = docNo;
        row.revise_no = reviseNo;
        const out = await insertRow('count_submissions', row);
        // ใบเดิมถือว่าถูกแทนที่แล้ว — กันผู้จัดการเห็นค้างในกล่องส่งกลับ
        try {
          await sb(`count_submissions?id=eq.${encodeURIComponent(s.reviseOf)}`, {
            method: 'PATCH', headers: { Prefer: 'return=minimal' },
            body: JSON.stringify({ status: 'superseded' }),
          });
        } catch { /* ไม่สำเร็จก็ไม่บล็อก ใบใหม่บันทึกแล้ว */ }
        return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
      }

      // กันกดส่งซ้ำ — ใบเดิมของคนเดิมภายใน 90 วินาที ถือว่าเป็นใบเดียวกัน
      try {
        if (s.reviseOf) throw new Error('skip');   // รอบแก้อาจมียอดเท่าเดิม ไม่ใช่ใบซ้ำ
        const since = new Date(Date.now() - 90 * 1000).toISOString();
        const dupes = await sb(
          `count_submissions?select=*&counter_id=eq.${encodeURIComponent(row.counter_id)}` +
          `&feature_type=eq.${encodeURIComponent(row.feature_type)}` +
          (row.branch ? `&branch=eq.${encodeURIComponent(row.branch)}` : '') +
          `&item_count=eq.${row.item_count}&total_qty=eq.${row.total_qty}` +
          `&submitted_at=gte.${encodeURIComponent(since)}&limit=1`
        );
        if (dupes?.length) {
          return res.status(200).json({ submission: toApp(dupes[0]), deduped: true });
        }
      } catch { /* เช็คซ้ำไม่ได้ก็ให้บันทึกต่อ ดีกว่าบล็อกงาน */ }

      // ขอเลขจากตัวนับกลาง ถ้าชนจริง (แข่งกันเสี้ยววินาที) ขอใหม่ให้เองเงียบ ๆ
      let lastErr;
      for (let attempt = 0; attempt < 5; attempt++) {
        try {
          row.doc_no = await nextDocNo(prefix);
          const out = await insertRow('count_submissions', row);
          return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
        } catch (e) {
          lastErr = e;
          if (e.code === '23505' || /duplicate|unique/i.test(e.message || '')) continue;
          throw e;
        }
      }
      throw lastErr;
    }

    if (req.method === 'PATCH') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { id, status, review_note, reviewed_by, drive } = body;
      if (!id) return res.status(400).json({ error: 'ต้องมี id' });

      // บันทึกผลส่งขึ้น Drive — ไม่แตะฟิลด์รีวิว
      if (drive) {
        const cur = await sb(`count_submissions?id=eq.${encodeURIComponent(id)}&select=drive_tries`);
        const tries = (cur?.[0]?.drive_tries || 0) + 1;
        const dp = drive.ok
          ? { drive_status: 'ok', drive_filename: drive.filename || null, drive_url: drive.url || null,
              drive_uploaded_at: new Date().toISOString(), drive_error: null, drive_tries: tries }
          : { drive_status: 'failed', drive_filename: drive.filename || null,
              drive_error: (drive.error || 'ส่งไม่สำเร็จ').slice(0, 500), drive_tries: tries };
        const o = await sb(`count_submissions?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify(dp) });
        return res.status(200).json({ submission: toApp(Array.isArray(o) ? o[0] : o) });
      }

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
