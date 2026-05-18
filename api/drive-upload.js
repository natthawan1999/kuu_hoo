import crypto from 'crypto';

// ── JWT for Google Service Account ──────────────────────────────────────────
function b64url(buf) {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function makeJWT(email, privateKey) {
  const now = Math.floor(Date.now() / 1000);
  const header  = b64url(Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })));
  const payload = b64url(Buffer.from(JSON.stringify({
    iss: email,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud:  'https://oauth2.googleapis.com/token',
    iat:  now,
    exp:  now + 3600,
  })));
  const toSign = `${header}.${payload}`;
  const sig = b64url(crypto.createSign('RSA-SHA256').update(toSign).sign(privateKey));
  return `${toSign}.${sig}`;
}

async function getToken(email, privateKey) {
  const jwt = makeJWT(email, privateKey);
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const d = await r.json();
  if (!d.access_token) throw new Error('Token error: ' + JSON.stringify(d));
  return d.access_token;
}

// ── Multipart upload to Drive ─────────────────────────────────────────────
async function uploadFile(token, filename, mimeType, buffer, folderId) {
  const meta = JSON.stringify({ name: filename, mimeType, ...(folderId ? { parents: [folderId] } : {}) });
  const bound = 'ku_boundary_314159';
  const body = Buffer.concat([
    Buffer.from(`--${bound}\r\nContent-Type: application/json\r\n\r\n${meta}\r\n--${bound}\r\nContent-Type: ${mimeType}\r\n\r\n`),
    buffer,
    Buffer.from(`\r\n--${bound}--`),
  ]);
  const r = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink,name',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary="${bound}"`,
      },
      body,
    }
  );
  const d = await r.json();
  if (!r.ok) throw new Error('Drive upload error: ' + JSON.stringify(d));
  return d;
}

// ── Handler ───────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const email  = process.env.GOOGLE_SA_EMAIL;
  const rawKey = process.env.GOOGLE_SA_PRIVATE_KEY;
  if (!email || !rawKey) {
    return res.status(500).json({ ok: false, error: 'GOOGLE_SA_EMAIL or GOOGLE_SA_PRIVATE_KEY not set' });
  }

  const privateKey = rawKey.replace(/\\n/g, '\n');
  const { filename, mimeType, content, isBase64, folderId } = req.body || {};
  if (!filename || content == null) {
    return res.status(400).json({ ok: false, error: 'filename and content required' });
  }

  try {
    const token  = await getToken(email, privateKey);
    const buffer = isBase64 ? Buffer.from(content, 'base64') : Buffer.from(content, 'utf8');
    const targetFolder = folderId || process.env.GOOGLE_DRIVE_FOLDER_ID;
    const file   = await uploadFile(token, filename, mimeType, buffer, targetFolder);
    return res.json({ ok: true, link: file.webViewLink, id: file.id, name: file.name });
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
