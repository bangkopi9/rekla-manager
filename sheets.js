const { google } = require('googleapis');

const SHEET_ID = '1ABJnF9O8ul27cv43nw-mbtOaxuoQX_PPF-nX2JvN2xs';
const REKLA_TAB = 'Rekla 2026';
const AH_TAB = 'Aroundhome';

function getAuth() {
  return new google.auth.JWT(
    process.env.GOOGLE_CLIENT_EMAIL,
    null,
    process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    ['https://www.googleapis.com/auth/spreadsheets']
  );
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const { action, tab } = req.query;
    const sheetTab = tab === 'aroundhome' ? AH_TAB : REKLA_TAB;

    // ── READ ──────────────────────────────────────────────────────────────
    if (req.method === 'GET' && action === 'read') {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${sheetTab}!A1:L1000`,
      });
      const rows = response.data.values || [];
      return res.status(200).json({ success: true, data: rows, tab: sheetTab });
    }

    // ── WRITE single row ──────────────────────────────────────────────────
    if (req.method === 'POST' && action === 'write') {
      const { rowIndex, values } = req.body;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetTab}!A${rowIndex}:L${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [values] },
      });
      return res.status(200).json({ success: true });
    }

    return res.status(400).json({ error: 'Unknown action' });

  } catch (err) {
    console.error('Sheets API error:', err);
    return res.status(500).json({ error: err.message });
  }
}
