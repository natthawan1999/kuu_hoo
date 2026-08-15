// KUUHOO — /api/config
// ส่งค่าเชื่อมต่อ Supabase ให้ทุกเครื่องใช้ร่วมกัน — พนักงานไม่ต้องตั้งค่าเองทีละเครื่อง
// ส่งแค่ anon key (สิทธิ์อ่านตามที่ RLS อนุญาต) ไม่เคยส่ง service_role ออกไป
//
// env ที่ใช้: VITE_SUPABASE_URL หรือ SUPABASE_URL · VITE_SUPABASE_ANON_KEY หรือ SUPABASE_ANON_KEY
//            VITE_SUPABASE_TABLE (ตารางสินค้า) · SUPABASE_STOCK_TABLE (ตารางสต็อก)

export default function handler(req, res) {
  res.setHeader('Cache-Control', 'public, max-age=60');

  const url = (process.env.VITE_SUPABASE_URL || process.env.SUPABASE_URL || '').replace(/\/$/, '');
  const anonKey = process.env.VITE_SUPABASE_ANON_KEY || process.env.SUPABASE_ANON_KEY || '';

  if (!url || !anonKey) {
    return res.status(200).json({
      configured: false,
      missing: [!url && 'VITE_SUPABASE_URL', !anonKey && 'VITE_SUPABASE_ANON_KEY'].filter(Boolean),
    });
  }

  return res.status(200).json({
    configured: true,
    url,
    anonKey,
    tableName: process.env.VITE_SUPABASE_TABLE || 'product_price',
    stockTableName: process.env.SUPABASE_STOCK_TABLE || process.env.VITE_SUPABASE_STOCK_TABLE || 'product_stock',
  });
}
