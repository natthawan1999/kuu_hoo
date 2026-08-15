// KUUHOO — /api/report
// วางไฟล์นี้ที่ api/report.js บน Vercel
//
// เรียกจากหน้า Report Builder:
//   POST /api/report
//   { topic, from, to, doc, barcode, party, person, limit, estimateOnly }
//
// topic: count | invoice | stock | in | out
// ทุกตัวยกเว้น topic / from / to ไม่บังคับ — ไม่ส่ง = ไม่กรองด้วยตัวนั้น

import { createClient } from '@supabase/supabase-js';

// รับชื่อ env สำรองด้วย เผื่อตั้งไว้เป็น VITE_SUPABASE_URL บน Vercel
const SB_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || process.env.NEXT_PUBLIC_SUPABASE_URL || '';
const SB_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || '';

// สร้างตอนเรียกใช้ ไม่ใช่ตอนโหลดไฟล์ — env ไม่ครบจะได้ตอบเป็น JSON ไม่ใช่ครashทั้งฟังก์ชัน
let _sb = null;
const client = () => (_sb ||= createClient(SB_URL, SB_KEY));

const FN = {
  count:   'report_count',
  invoice: 'report_invoice',
  stock:   'report_stock',
  in:      'report_movement',
  out:     'report_movement',
};

const nz = (v) => (v === '' || v === undefined ? null : v);

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });

  const {
    topic, from, to,
    doc = null, barcode = null, party = null, person = null, location = null,
    limit = 5000, estimateOnly = false,
  } = req.body || {};

  if (!SB_URL || !SB_KEY) {
    const missing = [!SB_URL && 'SUPABASE_URL (หรือ VITE_SUPABASE_URL)', !SB_KEY && 'SUPABASE_SERVICE_ROLE_KEY'].filter(Boolean);
    return res.status(500).json({ error: `ยังไม่ได้ตั้งค่า env บน Vercel: ${missing.join(', ')} — ตั้งแล้วต้อง Redeploy` });
  }

  if (!FN[topic]) return res.status(400).json({ error: 'topic ไม่ถูกต้อง' });
  if (topic !== 'stock' && (!from || !to)) {
    return res.status(400).json({ error: 'ต้องมีช่วงวันที่' });
  }

  // นับจำนวนแถวก่อน (ปุ่ม "ดึงรายงาน (~N แถว)")
  if (estimateOnly) {
    const { data, error } = await client().rpc('report_estimate', {
      p_topic: topic, p_from: nz(from), p_to: nz(to),
      p_doc: nz(doc), p_barcode: nz(barcode),
      p_party: nz(party), p_person: nz(person),
    });
    if (error) return res.status(500).json({ error: error.message });
    return res.status(200).json({ estimate: data });
  }

  let args;
  if (topic === 'count') {
    args = { p_from: from, p_to: to, p_doc: nz(doc), p_barcode: nz(barcode),
             p_counter: nz(person), p_limit: limit };
  } else if (topic === 'invoice') {
    args = { p_from: from, p_to: to, p_doc: nz(doc), p_barcode: nz(barcode),
             p_vendor: nz(party), p_keyed_by: nz(person), p_limit: limit };
  } else if (topic === 'stock') {
    // รายงานคงเหลือ: location = ประเภทสินค้า (products_view.category)
    args = { p_barcode: nz(barcode), p_location: nz(location) || nz(party), p_limit: limit };
  } else {
    args = { p_kind: topic, p_from: from, p_to: to, p_doc: nz(doc),
             p_barcode: nz(barcode), p_party: nz(party), p_limit: limit };
  }

  const { data, error } = await client().rpc(FN[topic], args);
  if (error) return res.status(500).json({ error: error.message });

  return res.status(200).json({
    rows: data || [],
    count: (data || []).length,
    columns: data && data.length ? Object.keys(data[0]) : [],
  });
}
