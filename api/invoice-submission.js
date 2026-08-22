// KUUHOO — /api/invoice-submission
// บิลรออนุมัติ: พนักงานส่ง → ผู้จัดการรีวิว → อนุมัติแล้วระบบเขียนลง bill_header + imp_data
//
//   GET    /api/invoice-submission                    → ทุกใบ (ผู้จัดการ)
//   GET    /api/invoice-submission?keyed_by_id=u1     → ใบของคนนั้น
//   POST   /api/invoice-submission { submission }     → ส่งใบใหม่ (pending)
//   PATCH  /api/invoice-submission { id, status, review_note, reviewed_by }
//          { id, drive: { ok, url, filename, error, by } } → บันทึกผลส่งขึ้น Drive
//          status=approved → เขียนลง bill_header + imp_data ให้ด้วย
//   DELETE /api/invoice-submission?id=...
//
// ต้องรัน sql/10-invoice-approval.sql ก่อน

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

async function sb(path, init = {}) {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: SB_KEY,
      Authorization: `Bearer ${SB_KEY}`,
      'Content-Type': 'application/json',
      Prefer: (init.method === 'POST' || init.method === 'PATCH') ? 'return=representation' : '',
      ...(init.headers || {}),
    },
  });
  const text = await res.text();
  let body = null;
  if (text) { try { body = JSON.parse(text); } catch { body = { message: text.slice(0, 300) }; } }
  if (!res.ok) { const e = new Error(body?.message || `HTTP ${res.status}`); e.code = body?.code; throw e; }
  return body;
}

const toApp = (r) => ({
  id: r.id,
  docNo: r.doc_no,
  keyedById: r.keyed_by_id,
  keyedBy: r.keyed_by,
  deviceId: r.device_id || '',
  invoiceNo: r.invoice_no || '',
  invoiceDate: r.invoice_date || '',
  vendorName: r.vendor_name || '',
  fileName: r.file_name || '',
  header: r.header || {},
  lines: Array.isArray(r.lines) ? r.lines : [],
  itemCount: r.item_count,
  netTotal: Number(r.net_total) || 0,
  submittedAt: r.submitted_at,
  status: r.status,
  reviewNote: r.review_note || '',
  reviewedAt: r.reviewed_at,
  reviewedBy: r.reviewed_by,
  postedAt: r.posted_at,
  drive: {
    status: r.drive_status || null,
    filename: r.drive_filename || '',
    url: r.drive_url || '',
    uploadedAt: r.drive_uploaded_at || null,
    error: r.drive_error || '',
    tries: r.drive_tries || 0,
    by: r.drive_by || '',
  },
});

// วันที่ไทย/ค.ศ. หลายรูปแบบ → YYYY-MM-DD (เหมือน th_date() ฝั่ง SQL)
function toISODate(v) {
  if (!v) return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return s;
  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m) {
    let [, d, mo, y] = m;
    let yy = +y; if (yy > 2400) yy -= 543;               // พ.ศ. → ค.ศ.
    return `${yy}-${String(+mo).padStart(2,'0')}-${String(+d).padStart(2,'0')}`;
  }
  return null;
}

const num = (v) => (v === '' || v == null || isNaN(+v) ? null : +v);

// อนุมัติแล้วเขียนลงตารางจริง
async function postToLedger(row) {
  const h = row.header || {};
  const invNo = row.invoice_no || h.invoice_no;
  const fileName = row.file_name || invNo;
  if (!invNo) throw new Error('ใบนี้ไม่มีเลขที่บิล — แก้ก่อนอนุมัติ');

  const lines = Array.isArray(row.lines) ? row.lines : [];
  const sum = (k) => lines.reduce((a, p) => a + (num(p[k]) || 0), 0);

  // 1) หัวบิล
  await sb('bill_header', {
    method: 'POST',
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify({
      invoice_no: invNo,
      file_name: fileName,
      invoice_date: toISODate(row.invoice_date || h.invoice_date) || (row.invoice_date || h.invoice_date) || null,
      vendor_name: h.vendor_name ?? row.vendor_name ?? null,
      vendor_tax_id: h.vendor_tax_id ?? null,
      document_type: h.document_type ?? null,
      vendor_address: h.vendor_address ?? null,
      vendor_branch: h.vendor_branch ?? null,
      vendor_no: h.vendor_no ?? null,
      price_type: h.price_type ?? 'incl',
      total_amount: +sum('amount').toFixed(2) || null,
      total_discount: +sum('special_discount').toFixed(2) || null,
      net_total: +sum('total').toFixed(2) || null,
      excl_vat: +sum('excl_vat').toFixed(2) || null,
      vat_amount: +sum('vat_amt').toFixed(2) || null,
      updated_at: new Date().toISOString(),
    }),
  });

  // 2) บรรทัดสินค้า
  if (lines.length) {
    const rows = lines.map((p, i) => ({
      invoice_no: invNo,
      file_name: fileName,
      no: num(p.no) ?? i + 1,
      code: p.code ?? null,
      description: p.description ?? null,
      carton_size: num(p.carton_size),
      carton: num(p.carton),
      carton_unit: p.carton_unit ?? null,
      ea: num(p.ea),
      ea_unit: p.ea_unit ?? null,
      qty: num(p.qty),
      price_ea: num(p.price_ea),
      amount: num(p.amount),
      special_discount: num(p.special_discount),
      amount_sd: num(p.amount_sd),
      total: num(p.total),
      diff: num(p.diff),
      excl_vat: num(p.excl_vat),
      vat_amt: num(p.vat_amt),
      vat: p.vat ?? null,
      barcode: p.barcode ?? null,
      updated_at: new Date().toISOString(),
    }));
    await sb('imp_data', {
      method: 'POST',
      headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
      body: JSON.stringify(rows),
    });
  }
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  try {
    if (!SB_URL || !SB_KEY) {
      return res.status(500).json({ error: 'ยังไม่ได้ตั้งค่า env บน Vercel: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY' });
    }
    const q = req.query || {};

    if (req.method === 'GET') {
      let path = 'invoice_submissions?select=*&order=submitted_at.desc&limit=300';
      if (q.keyed_by_id) path += `&keyed_by_id=eq.${encodeURIComponent(q.keyed_by_id)}`;
      if (q.status) path += `&status=eq.${encodeURIComponent(q.status)}`;
      const rows = await sb(path);
      return res.status(200).json({ submissions: (rows || []).map(toApp) });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const s = body.submission || body;
      if (!s?.keyedById || !s?.keyedBy) return res.status(400).json({ error: 'ต้องมีชื่อคนคีย์' });
      if (!Array.isArray(s.lines) || s.lines.length === 0) return res.status(400).json({ error: 'ไม่มีบรรทัดสินค้าในใบนี้' });

      // เลขใบส่งงาน IV-YYMMDDNNNN
      const now = new Date();
      const dateKey = String(now.getFullYear() % 100).padStart(2,'0')
        + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0');
      let seq = 1;
      try {
        const cur = await sb(`doc_counters?select=seq&prefix=eq.IV&date_key=eq.${dateKey}&limit=1`);
        seq = (cur?.[0]?.seq || 0) + 1;
        await sb('doc_counters', {
          method: 'POST',
          headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
          body: JSON.stringify({ prefix: 'IV', date_key: dateKey, seq }),
        });
      } catch { seq = Math.floor(Math.random() * 9000) + 1000; }

      const row = {
        doc_no: `IV-${dateKey}${String(seq).padStart(4,'0')}`,
        keyed_by_id: s.keyedById,
        keyed_by: s.keyedBy,
        device_id: s.deviceId || null,
        invoice_no: s.invoiceNo || s.header?.invoice_no || null,
        invoice_date: s.invoiceDate || s.header?.invoice_date || null,
        vendor_name: s.vendorName || s.header?.vendor_name || null,
        file_name: s.fileName || null,
        header: s.header || {},
        lines: s.lines,
        item_count: s.lines.length,
        net_total: s.netTotal ?? s.lines.reduce((a,p) => a + (num(p.total) || 0), 0),
        status: 'pending',
      };

      try {
        const out = await sb('invoice_submissions', { method: 'POST', body: JSON.stringify(row) });
        return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
      } catch (e) {
        if (e.code === '23505' || /duplicate|unique/i.test(e.message)) {
          return res.status(409).json({ error: `บิลเลขที่ ${row.invoice_no} ส่งไปแล้วและยังรออนุมัติอยู่`, duplicate: true });
        }
        throw e;
      }
    }

    if (req.method === 'PATCH') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const { id, status, review_note, reviewed_by, header, lines, drive } = body;
      if (!id) return res.status(400).json({ error: 'ต้องมี id' });

      // บันทึกผลส่งขึ้น Drive — ไม่แตะฟิลด์รีวิว ไม่ยิงลง ledger ซ้ำ
      if (drive) {
        const cur = await sb(`invoice_submissions?id=eq.${encodeURIComponent(id)}&select=drive_tries`);
        const tries = (cur?.[0]?.drive_tries || 0) + 1;
        const dp = drive.ok
          ? { drive_status: 'ok', drive_filename: drive.filename || null, drive_url: drive.url || null,
              drive_uploaded_at: new Date().toISOString(), drive_error: null, drive_tries: tries, drive_by: drive.by || null }
          : { drive_status: 'failed', drive_filename: drive.filename || null,
              drive_error: (drive.error || 'ส่งไม่สำเร็จ').slice(0, 500), drive_tries: tries, drive_by: drive.by || null };
        const o = await sb(`invoice_submissions?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify(dp) });
        return res.status(200).json({ submission: toApp(Array.isArray(o) ? o[0] : o) });
      }

      const cur = await sb(`invoice_submissions?select=*&id=eq.${encodeURIComponent(id)}&limit=1`);
      const row = cur?.[0];
      if (!row) return res.status(404).json({ error: 'ไม่พบใบนี้' });

      // ผู้จัดการแก้ข้อมูลก่อนอนุมัติได้
      const merged = {
        ...row,
        header: header ?? row.header,
        lines: lines ?? row.lines,
        invoice_no: (header?.invoice_no) ?? row.invoice_no,
        invoice_date: (header?.invoice_date) ?? row.invoice_date,
        vendor_name: (header?.vendor_name) ?? row.vendor_name,
      };

      const patch = { reviewed_at: new Date().toISOString() };
      if (header) { patch.header = merged.header; patch.invoice_no = merged.invoice_no; patch.invoice_date = merged.invoice_date; patch.vendor_name = merged.vendor_name; }
      if (lines) { patch.lines = merged.lines; patch.item_count = merged.lines.length; patch.net_total = merged.lines.reduce((a,p) => a + (num(p.total) || 0), 0); }
      if (review_note !== undefined) patch.review_note = review_note || '';
      if (reviewed_by) patch.reviewed_by = reviewed_by;
      if (status) patch.status = status;

      // อนุมัติ = เขียนลงตารางจริงก่อน แล้วค่อยปิดสถานะ
      if (status === 'approved' && row.status !== 'approved') {
        await postToLedger(merged);
        patch.posted_at = new Date().toISOString();
      }

      const out = await sb(`invoice_submissions?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', body: JSON.stringify(patch) });
      return res.status(200).json({ submission: toApp(Array.isArray(out) ? out[0] : out) });
    }

    if (req.method === 'DELETE') {
      if (!q.id) return res.status(400).json({ error: 'ต้องมี id' });
      await sb(`invoice_submissions?id=eq.${encodeURIComponent(q.id)}`, { method: 'DELETE' });
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET / POST / PATCH / DELETE' });
  } catch (e) {
    const msg = e?.message || 'เกิดข้อผิดพลาด';
    const hint = /invoice_submissions.*does not exist|could not find the table/i.test(msg)
      ? ' — ยังไม่ได้รัน sql/10-invoice-approval.sql บน Supabase' : '';
    return res.status(500).json({ error: msg + hint });
  }
}
