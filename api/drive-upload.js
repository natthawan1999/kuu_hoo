export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const scriptUrl = process.env.APPS_SCRIPT_URL;
  if (!scriptUrl) return res.status(500).json({ ok: false, error: 'APPS_SCRIPT_URL not set' });

  const { filename, mimeType, content, isBase64, folderId } = req.body || {};
  if (!filename || content == null) return res.status(400).json({ ok: false, error: 'filename and content required' });

  try {
    const r = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filename, mimeType, content,
        isBase64: !!isBase64,
        folderId: folderId || process.env.GOOGLE_DRIVE_FOLDER_ID,
      }),
      redirect: 'follow',
    });
    const text = await r.text();
    try {
      const d = JSON.parse(text);
      return res.status(d.ok ? 200 : 500).json(d);
    } catch {
      return res.status(500).json({ ok: false, error: 'Apps Script error: ' + text.slice(0, 200) });
    }
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
