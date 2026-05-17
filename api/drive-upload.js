export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { filename, mimeType, content, isBase64 } = req.body || {};
  if (!filename || !content) return res.status(400).json({ ok: false, error: 'missing fields' });

  const FOLDER_ID = process.env.DRIVE_FOLDER_ID || '';
  const SA = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

  if (!SA) return res.status(500).json({ ok: false, error: 'GOOGLE_SERVICE_ACCOUNT_JSON not set' });

  try {
    // Get access token via service account JWT
    const sa = JSON.parse(SA);
    const token = await getServiceAccountToken(sa);

    // Build file content
    const fileContent = isBase64
      ? Buffer.from(content, 'base64')
      : Buffer.from(content, 'utf-8');

    // Upload to Drive (multipart)
    const boundary = '-------kuuhoo_boundary';
    const metadata = { name: filename, mimeType, ...(FOLDER_ID ? { parents: [FOLDER_ID] } : {}) };

    const body = [
      `--${boundary}`,
      'Content-Type: application/json; charset=UTF-8',
      '',
      JSON.stringify(metadata),
      `--${boundary}`,
      `Content-Type: ${mimeType}`,
      'Content-Transfer-Encoding: base64',
      '',
      fileContent.toString('base64'),
      `--${boundary}--`,
    ].join('\r\n');

    const uploadRes = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink',
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': `multipart/related; boundary="${boundary}"`,
        },
        body,
      }
    );

    if (!uploadRes.ok) {
      const err = await uploadRes.text();
      return res.json({ ok: false, error: `Drive ${uploadRes.status}: ${err.slice(0, 200)}` });
    }

    const data = await uploadRes.json();
    return res.json({ ok: true, link: data.webViewLink, id: data.id });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
}

async function getServiceAccountToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  };

  const b64 = (obj) => Buffer.from(JSON.stringify(obj)).toString('base64url');
  const unsigned = `${b64(header)}.${b64(payload)}`;

  // Sign with RSA-SHA256 using Web Crypto
  const keyData = sa.private_key
    .replace(/-----BEGIN PRIVATE KEY-----/, '')
    .replace(/-----END PRIVATE KEY-----/, '')
    .replace(/\n/g, '');

  const keyBuffer = Buffer.from(keyData, 'base64');
  const cryptoKey = await crypto.subtle.importKey(
    'pkcs8', keyBuffer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false, ['sign']
  );

  const signatureBuffer = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    cryptoKey,
    Buffer.from(unsigned)
  );

  const signature = Buffer.from(signatureBuffer).toString('base64url');
  const jwt = `${unsigned}.${signature}`;

  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
    }),
  });

  const tokenData = await tokenRes.json();
  if (!tokenData.access_token) throw new Error('Failed to get token: ' + JSON.stringify(tokenData));
  return tokenData.access_token;
}
