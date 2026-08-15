// KUUHOO — /api/staff
// วางไฟล์นี้ที่ api/staff.js บน Vercel
//
//   GET  /api/staff                      → รายชื่อทั้งหมด (ผู้จัดการ)
//   GET  /api/staff?feature=recorder     → เฉพาะคนที่ใช้งานได้ (หน้าเลือกชื่อ)
//   POST /api/staff  { action, ... }     → add | update | setActive
//
// ต้องมี env: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY

import { createClient } from '@supabase/supabase-js';

const db = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false },
});

const FEATURES = { recorder: 'allow_recorder', compare: 'allow_compare', invoice: 'allow_invoice' };

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      const { feature } = req.query;

      // หน้าเลือกชื่อ — กรองเฉพาะคนที่เปิดใช้งานและมีสิทธิ์ฟีเจอร์นั้น
      if (feature) {
        if (!FEATURES[feature]) return res.status(400).json({ error: 'feature ไม่ถูกต้อง' });
        const { data, error } = await db.rpc('staff_for_feature', { p_feature: feature });
        if (error) throw error;
        return res.status(200).json({ staff: data || [] });
      }

      // หน้า Staff Admin — ทั้งหมด พร้อมจำนวนใบที่เคยส่ง
      const { data, error } = await db.from('staff').select('*').order('name');
      if (error) throw error;

      const { data: subs } = await db.from('count_submissions').select('counter_id');
      const tally = {};
      for (const s of subs || []) if (s.counter_id) tally[s.counter_id] = (tally[s.counter_id] || 0) + 1;

      return res.status(200).json({
        staff: (data || []).map(s => ({ ...s, submission_count: tally[s.id] || 0 })),
      });
    }

    if (req.method === 'POST') {
      const { action } = req.body || {};

      if (action === 'add') {
        const { name, dept, initial, allow_recorder, allow_compare, allow_invoice } = req.body;
        if (!name?.trim()) return res.status(400).json({ error: 'ต้องมีชื่อ' });

        // id รูปแบบ u1, u2 … ต่อจากตัวท้ายสุด
        const { data: last } = await db.from('staff').select('id').like('id', 'u%');
        const next = 'u' + ((last || []).reduce((m, r) => Math.max(m, parseInt(r.id.slice(1)) || 0), 0) + 1);

        const { data, error } = await db.from('staff').insert({
          id: next,
          name: name.trim(),
          initial: (initial || name.trim()).charAt(0),
          dept: dept || null,
          allow_recorder: !!allow_recorder,
          allow_compare: !!allow_compare,
          allow_invoice: !!allow_invoice,
        }).select().single();

        // ชื่อซ้ำกับคนที่ยังใช้งานอยู่
        if (error?.code === '23505' || error?.message?.includes('idx_staff_active_name')) {
          return res.status(409).json({ error: `มีชื่อ "${name.trim()}" อยู่แล้ว` });
        }
        if (error) throw error;
        return res.status(200).json({ staff: data });
      }

      if (action === 'update') {
        const { id, ...fields } = req.body;
        if (!id) return res.status(400).json({ error: 'ต้องมี id' });
        const allowed = ['name', 'initial', 'dept', 'allow_recorder', 'allow_compare', 'allow_invoice'];
        const patch = {};
        for (const k of allowed) if (k in fields) patch[k] = fields[k];
        const { data, error } = await db.from('staff').update(patch).eq('id', id).select().single();
        if (error) throw error;
        return res.status(200).json({ staff: data });
      }

      if (action === 'setActive') {
        const { id, active } = req.body;
        if (!id) return res.status(400).json({ error: 'ต้องมี id' });
        // ปิดการใช้งาน ไม่ใช่ลบ — ใบเก่าต้องมีชื่อกำกับไว้ (trigger กัน DELETE อยู่)
        const { data, error } = await db.from('staff').update({ active: !!active }).eq('id', id).select().single();
        if (error) throw error;
        return res.status(200).json({ staff: data });
      }

      return res.status(400).json({ error: 'action ไม่ถูกต้อง' });
    }

    return res.status(405).json({ error: 'ใช้ได้แค่ GET กับ POST' });
  } catch (e) {
    return res.status(500).json({ error: e.message || 'เกิดข้อผิดพลาด' });
  }
}
