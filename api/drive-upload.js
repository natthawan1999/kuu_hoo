export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const scriptUrl = process.env.APPS_SCRIPT_URL;
  if (!scriptUrl) return res.status(500).json({ ok: false, error: 'APPS_SCRIPT_URL not set' });

  const { filename, mimeType, content, isBase64, folderId, _debug } = req.body || {};
  if (!filename || content == null) return res.status(400).json({ ok: false, error: 'filename and content required' });

  try {
    const r1 = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filename, mimeType, content,
        isBase64: !!isBase64,
        folderId: folderId || process.env.GOOGLE_DRIVE_FOLDER_ID,
      }),
      redirect: 'manual',
    });

    const diag = {
      step1_status: r1.status,
      step1_location: r1.headers.get('location'),
      step1_contentType: r1.headers.get('content-type'),
      scriptUrl_preview: scriptUrl.slice(0, 60) + '...',
    };

    let text, step2 = null;
    const location = r1.headers.get('location');
    if ((r1.status === 301 || r1.status === 302 || r1.status === 303) && location) {
      const r2 = await fetch(location, { method: 'GET' });
      text = await r2.text();
      step2 = { status: r2.status, contentType: r2.headers.get('content-type') };
    } else {
      text = await r1.text();
    }

    try {
      const d = JSON.parse(text);
      return res.status(d.ok ? 200 : 500).json(d);
    } catch {
      const diagStr = `status=${diag.step1_status} loc=${diag.step1_location || 'none'} ct=${diag.step1_contentType || 'none'}${step2 ? ` | step2: status=${step2.status} ct=${step2.contentType}` : ''}`;
      return res.status(500).json({
        ok: false,
        error: `Non-JSON [${diagStr}] body: ${text.slice(0, 300)}`,
      });
    }
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
