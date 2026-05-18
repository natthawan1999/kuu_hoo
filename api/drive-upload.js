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
        filename,
        mimeType,
        content,
        isBase64: !!isBase64,
        folderId: folderId || process.env.GOOGLE_DRIVE_FOLDER_ID,
      }),
      redirect: 'follow',
    });
    const d = await r.json();
    return res.status(r.ok ? 200 : 500).json(d);
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
