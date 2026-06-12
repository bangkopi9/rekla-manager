const { google } = require('googleapis');

const SHEET_ID = '1ABJnF9O8ul27cv43nw-mbtOaxuoQX_PPF-nX2JvN2xs';
const REKLA_TAB = 'Rekla 2026';
const AH_TAB = 'Aroundhome';
const LOG_TAB = 'Log';

function getAuth() {
  return new google.auth.JWT(
    process.env.GOOGLE_CLIENT_EMAIL,
    null,
    process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    ['https://www.googleapis.com/auth/spreadsheets']
  );
}

function fmtDate(d) {
  return d.toLocaleDateString('de-DE', { timeZone: 'Europe/Berlin', day: '2-digit', month: '2-digit', year: 'numeric' });
}
function fmtDateTime(d) {
  const date = d.toLocaleDateString('de-DE', { timeZone: 'Europe/Berlin', day: '2-digit', month: '2-digit', year: 'numeric' });
  const time = d.toLocaleTimeString('de-DE', { timeZone: 'Europe/Berlin', hour: '2-digit', minute: '2-digit', second: '2-digit' });
  return `${date} ${time}`;
}

async function ensureLogTab(sheets) {
  try {
    await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${LOG_TAB}!A1:E1` });
  } catch (e) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: LOG_TAB } } }] },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${LOG_TAB}!A1:E1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['Zeitstempel', 'Mitarbeiter', 'Eingangsdatum', 'Rekla Link', 'Grund']] },
    });
  }
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
      return res.status(200).json({ success: true, data: response.data.values || [], tab: sheetTab });
    }

    // ── READ LOG ─────────────────────────────────────────────────────────
    if (req.method === 'GET' && action === 'readLog') {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: `${LOG_TAB}!A1:E500`,
        });
        return res.status(200).json({ success: true, data: response.data.values || [] });
      } catch (e) {
        return res.status(200).json({ success: true, data: [] });
      }
    }

    // ── WRITE single row (Tab 1 edits) ──────────────────────────────────────
    // IMPORTANT: Column K (index 10, "Ueberfaellig") is formula/dropdown-based
    // and must NEVER be overwritten. We split the write into A:J and L only.
    if (req.method === 'POST' && action === 'write') {
      const { rowIndex, values } = req.body;
      const ajValues = values.slice(0, 10);   // A..J
      const lValue = values.length > 11 ? values[11] : '';  // L (Notizen)

      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetTab}!A${rowIndex}:J${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [ajValues] },
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetTab}!L${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[lValue]] },
      });

      return res.status(200).json({ success: true });
    }

    // ── SUBMIT NEW LEAD (Mitarbeiter form) ──────────────────────────────────
    if (req.method === 'POST' && action === 'submitLead') {
      const { reklaLink, grund, mitarbeiter } = req.body;
      if (!reklaLink || !grund || !mitarbeiter) {
        return res.status(400).json({ error: 'Fehlende Felder' });
      }
      const now = new Date();
      const eingangsdatum = fmtDate(now);
      const zeitstempel = fmtDateTime(now);

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: `${REKLA_TAB}!A:D`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [[eingangsdatum, '', reklaLink, grund]] },
      });

      await ensureLogTab(sheets);
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: `${LOG_TAB}!A:E`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [[zeitstempel, mitarbeiter, eingangsdatum, reklaLink, grund]] },
      });

      return res.status(200).json({ success: true, eingangsdatum, zeitstempel });
    }

    return res.status(400).json({ error: 'Unknown action' });

  } catch (err) {
    console.error('Sheets API error:', err);
    return res.status(500).json({ error: err.message });
  }
}
