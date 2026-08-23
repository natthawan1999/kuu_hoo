// KUUHOO — /api/sync-status
// อ่านของจริงจาก sync_log + view sync_status ที่มีอยู่แล้วใน Supabase
// (ไม่ต้องรัน 12-sync-status.sql — ไฟล์นั้นเลิกใช้แล้ว)
//
//   GET /api/sync-status        → สถานะรายรีพอร์ต + ประวัติรอบล่าสุด + สรุปการส่ง Drive
//   GET /api/sync-status?runs=60 → ขอประวัติมากกว่าเดิม (ดีฟอลต์ 30)
//
// ตารางที่ใช้:
//   sync_log(id, run_at, report, data_date, rows_csv, rows_sent, status, error_msg, duration_sec)
//   sync_status (view) — คอลัมน์ชื่อไทย: ซิงก์ล่าสุด · วันที่ข้อมูล · สถานะ · จำนวนแถว · ชั่วโมงที่ผ่านมา · สรุป · ข้อความ_error
//   upload_summary() — จาก sql/13-drive-status.sql (ถ้ายังไม่รัน จะข้ามส่วนนี้ไป)

const SB_URL = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

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
const rpc = (fn, args = {}) => sb(`rpc/${fn}`, { method: 'POST', body: JSON.stringify(args) });

const n = (v) => (v === null || v === undefined || v === '' ? null : Number(v));

// ชื่อรายงานเป็นภาษาคน
const REPORT_LABEL = {
  product_price:    'ราคาสินค้า',
  product_stock:    'สินค้าคงคลัง',
  vendor_info:      'ข้อมูลผู้ขาย',
  sale_report_bill: 'บิลขาย',
  sale_item:        'รายการสินค้าขาย',
};

// view sync_status ตัดสินสถานะมาให้แล้วในคอลัมน์ "สรุป" — ใช้ของมัน ไม่ตั้งเกณฑ์ซ้ำ
//   ปกติ / ล้มเหลว / ไม่ได้ซิงก์นานผิดปกติ (เกิน 36 ชม.)
// เพิ่มเงื่อนไขเดียวที่ view ไม่รู้: แถวหายระหว่างทาง (rows_csv ≠ rows_sent)
function gradeReport(r) {
  if (r.rowsCsv !== null && r.rowsSent !== null && r.rowsSent < r.rowsCsv) return 'late';
  if (r.summary === 'ล้มเหลว') return 'late';
  if (r.summary === 'ไม่ได้ซิงก์นานผิดปกติ') return 'stale';
  if (r.summary === 'ปกติ') return 'ok';
  // ไม่มีค่า "สรุป" (อ่าน view ไม่ได้ ใช้ sync_log แทน) — ตัดสินเอง
  if (r.status && r.status !== 'ok') return 'late';
  if (r.hoursAgo === null) return 'unknown';
  return r.hoursAgo > 36 ? 'stale' : 'ok';
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!SB_URL || !SB_KEY) return res.status(500).json({ error: 'ยังไม่ได้ตั้ง SUPABASE_URL / SERVICE_ROLE_KEY' });

  const runLimit = Math.min(Math.max(Number(req.query?.runs) || 30, 1), 200);

  try {
    // ประวัติรอบซิงก์ — เอาไว้หารอบที่ fail และเทียบจำนวนแถว
    const log = await sb(`sync_log?select=*&order=run_at.desc&limit=${runLimit}`);
    const runs = (log || []).map(r => {
      const csv = n(r.rows_csv), sent = n(r.rows_sent);
      return {
        id: r.id, report: r.report, runAt: r.run_at, dataDate: r.data_date,
        rowsCsv: csv, rowsSent: sent,
        missing: csv !== null && sent !== null ? Math.max(csv - sent, 0) : null,
        status: r.status, error: r.error_msg || '',
        durationSec: n(r.duration_sec),
      };
    });

    // สถานะปัจจุบันรายรีพอร์ต — ใช้ view ถ้าอ่านได้ ไม่ได้ก็สรุปจาก sync_log เอง
    let reports = [];
    try {
      const view = await sb('sync_status?select=*');
      reports = (view || []).map(v => {
        const latest = runs.find(x => x.report === v.report);
        const src  = n(v['แถวต้นทาง'])  ?? latest?.rowsCsv  ?? null;
        const sent = n(v['แถวสำเร็จ'])   ?? latest?.rowsSent ?? null;
        const miss = n(v['แถวที่หาย'])   ?? latest?.missing  ?? null;
        const base = {
          report: v.report,
          lastRunAt: v['ซิงก์ล่าสุด'] || null,
          dataDate: v['วันที่ข้อมูล'] || null,
          status: v['สถานะ'] || null,
          rows: sent ?? n(v['จำนวนแถว']),
          hoursAgo: n(v['ชั่วโมงที่ผ่านมา']),
          summary: v['สรุป'] || '',
          error: v['ข้อความ_error'] || '',
          mode: v['โหมด'] || '',                       // auto = รันเอง · manual = คนสั่ง
          durationSec: n(v['ใช้เวลา_วินาที']),
          rowsCsv: src, rowsSent: sent, missing: miss,
          // snapshot ทั้งร้าน (ราคา/สต็อก/ผู้ขาย) ไม่ผูกกับวันไหน — ไม่ใช่ข้อมูลขาด
          snapshot: !v['วันที่ข้อมูล'],
        };
        return { ...base, label: REPORT_LABEL[v.report] || v.report, state: gradeReport(base) };
      });
    } catch {
      const seen = new Set();
      for (const r of runs) {
        if (seen.has(r.report)) continue;
        seen.add(r.report);
        const hoursAgo = r.runAt ? Math.round((Date.now() - new Date(r.runAt).getTime()) / 3600000) : null;
        const base = { report: r.report, lastRunAt: r.runAt, dataDate: r.dataDate, status: r.status,
                       rows: r.rowsSent, hoursAgo, summary: '', error: r.error,
                       rowsCsv: r.rowsCsv, rowsSent: r.rowsSent, missing: r.missing };
        reports.push({ ...base, label: REPORT_LABEL[r.report] || r.report, state: gradeReport(base) });
      }
    }
    reports.sort((a, b) => a.report.localeCompare(b.report));

    // บิลที่ควรตรวจสอบ (ยอดผิดปกติ) — ไม่บล็อกอะไร แค่ชี้ให้ไปดู
    let anomalies = null;
    try {
      const rows = await sb('bill_anomalies?select=*&order=วันที่.desc&limit=50');
      anomalies = {
        count: (rows || []).length,
        rows: (rows || []).map(a => ({
          table: a['ตาราง'], date: a['วันที่'], ref: a['อ้างอิง'],
          value: n(a['ค่าที่พบ']), reason: a['เหตุผล'],
        })),
      };
    } catch { /* ไม่มี view นี้ */ }

    // สินค้า/ผู้ขายที่ไม่ถูกอัปเดตเกิน 30 วัน (upsert ลบไม่เป็น ของผีจึงค้าง)
    let stale = null;
    try {
      const rows = await sb('stale_records?select=*&order=หายไปกี่วัน.desc&limit=50');
      stale = {
        count: (rows || []).length,
        rows: (rows || []).map(x => ({
          table: x['ตาราง'], code: x['รหัส'], name: x['ชื่อ'],
          updatedAt: x.updated_at, days: n(x['หายไปกี่วัน']),
        })),
      };
    } catch { /* ไม่มี view นี้ */ }

    // สรุปการส่งขึ้น Drive (sql/13-drive-status.sql)
    let uploads = null;
    try {
      const u = await rpc('upload_summary');
      const pick = (k) => (u || []).find(x => x.kind === k) || {};
      const cs = pick('count'), iv = pick('invoice');
      const z = (v) => Number(v || 0);
      uploads = {
        count:   { pending: z(cs.pending), uploaded: z(cs.uploaded), failed: z(cs.failed), lastAt: cs.last_at || null },
        invoice: { pending: z(iv.pending), uploaded: z(iv.uploaded), failed: z(iv.failed), lastAt: iv.last_at || null },
      };
      uploads.pending = uploads.count.pending + uploads.invoice.pending;
      uploads.failed  = uploads.count.failed  + uploads.invoice.failed;
      uploads.uploaded = uploads.count.uploaded + uploads.invoice.uploaded;
      uploads.lastAt  = [uploads.count.lastAt, uploads.invoice.lastAt].filter(Boolean).sort().pop() || null;
      uploads.state   = uploads.failed ? 'late' : uploads.pending ? 'stale' : 'ok';
    } catch { /* ยังไม่ได้รัน 13-drive-status.sql */ }

    const worst = ['late', 'unknown', 'stale', 'ok'].find(s => reports.some(x => x.state === s)) || 'unknown';
    return res.status(200).json({
      checkedAt: new Date().toISOString(),
      overall: worst,
      reports,
      runs,
      anomalies,
      stale,
      failedRuns: runs.filter(r => r.status && r.status !== 'ok').length,
      uploads,
    });
  } catch (e) {
    const msg = e.message || 'เกิดข้อผิดพลาด';
    const hint = /sync_log/i.test(msg) ? ' — อ่านตาราง sync_log ไม่ได้ (เช็คสิทธิ์ service role)' : '';
    return res.status(500).json({ error: msg + hint });
  }
}
