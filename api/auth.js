export default function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { code, pin } = req.body || {};

  // AccessGate check
  if (code !== undefined) {
    const ACCESS_CODE = process.env.ACCESS_CODE || '1234';
    if (code === ACCESS_CODE) return res.json({ ok: true });
    return res.json({ ok: false, msg: 'รหัสไม่ถูกต้อง' });
  }

  // Manager PIN check
  if (pin !== undefined) {
    const MANAGER_PIN = process.env.MANAGER_PIN || '0000';
    if (pin === MANAGER_PIN) return res.json({ ok: true });
    return res.json({ ok: false, msg: 'PIN ไม่ถูกต้อง' });
  }

  return res.status(400).json({ ok: false, msg: 'invalid request' });
}
