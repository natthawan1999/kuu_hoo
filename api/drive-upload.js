import { google } from 'googleapis';
import { Readable } from 'stream';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const email = process.env.GOOGLE_SA_EMAIL;
  const rawKey = process.env.GOOGLE_SA_PRIVATE_KEY;

  if (!email || !rawKey) {
    return res.status(500).json({ ok: false, error: 'GOOGLE_SA_EMAIL or GOOGLE_SA_PRIVATE_KEY not set' });
  }

  // Vercel stores \n as literal \\n — convert back
  const privateKey = rawKey.replace(/\\n/g, '\n');

  const { filename, mimeType, content, isBase64, folderId } = req.body || {};
  if (!filename || !content) return res.status(400).json({ ok: false, error: 'filename and content required' });

  try {
    const auth = new google.auth.JWT({
      email,
      key: privateKey,
      scopes: ['https://www.googleapis.com/auth/drive.file'],
    });

    const drive = google.drive({ version: 'v3', auth });

    const buffer = isBase64
      ? Buffer.from(content, 'base64')
      : Buffer.from(content, 'utf8');

    const targetFolder = folderId || process.env.GOOGLE_DRIVE_FOLDER_ID;

    const file = await drive.files.create({
      requestBody: {
        name: filename,
        mimeType,
        ...(targetFolder ? { parents: [targetFolder] } : {}),
      },
      media: {
        mimeType,
        body: Readable.from(buffer),
      },
      fields: 'id,webViewLink,name',
    });

    return res.json({ ok: true, link: file.data.webViewLink, id: file.data.id, name: file.data.name });
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
