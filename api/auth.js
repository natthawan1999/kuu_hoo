export default function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { pin } = req.body || {};

  if (pin !== undefined) {
    const correct = process.env.MANAGER_PIN;
    if (!correct) return res.status(500).json({ ok: false, msg: 'MANAGER_PIN not configured' });
    return res.json({ ok: pin === correct });
  }

  return res.status(400).json({ ok: false, msg: 'invalid request' });
}
