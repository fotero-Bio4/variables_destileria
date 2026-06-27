'use strict';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const CORS  = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type':                 'application/json',
};

// ── Rate limiting ─────────────────────────────────────────────────────────────
const rateMap = new Map();
function checkRate(ip) {
  const now   = Date.now();
  const entry = rateMap.get(ip);
  if (!entry || now - entry.t > 15 * 60 * 1000) { rateMap.set(ip, { t: now, n: 1 }); return false; }
  entry.n++;
  return entry.n > 10;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────
async function getToken() {
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET } = process.env;
  const r = await fetch(
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method:  'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body:    new URLSearchParams({
        grant_type: 'client_credentials', client_id: GRAPH_CLIENT_ID,
        client_secret: GRAPH_CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default',
      }),
    }
  );
  const d = await r.json();
  if (!d.access_token) throw new Error('Auth error: ' + (d.error_description || JSON.stringify(d)));
  return d.access_token;
}

async function readSheet(token, driveId, itemId, sheet) {
  const enc = encodeURIComponent(sheet);
  const r   = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/usedRange(valuesOnly=true)`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!r.ok) return [];
  const { values = [] } = await r.json();
  return values;
}

// ── Handler ───────────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  if (event.httpMethod !== 'POST')    return { statusCode: 405, headers: CORS, body: JSON.stringify({ error: 'Método no permitido' }) };

  const ip = ((event.headers['x-forwarded-for'] || '').split(',')[0].trim()) || 'unknown';
  if (checkRate(ip)) return { statusCode: 429, headers: CORS, body: JSON.stringify({ error: 'Demasiados intentos. Esperá 15 minutos.' }) };

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'Cuerpo inválido.' }) }; }

  const { mail, pss } = body;
  if (!mail || !pss) return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'Mail y contraseña requeridos.' }) };

  const { BD_DRIVE_ID, BD_ITEM_ID, GRAPH_CLIENT_SECRET } = process.env;
  if (!GRAPH_CLIENT_SECRET || !BD_DRIVE_ID || !BD_ITEM_ID)
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: 'Servidor no configurado.' }) };

  try {
    const token = await getToken();

    // Hoja "usarios": col A = Mail, col B = Nombre, col C = Pss
    const rows = await readSheet(token, BD_DRIVE_ID, BD_ITEM_ID, 'usarios');
    const row  = rows.slice(1).find(r =>
      String(r[0] ?? '').trim().toLowerCase() === String(mail).trim().toLowerCase() &&
      String(r[2] ?? '').trim() === String(pss).trim()
    );

    if (!row) return { statusCode: 401, headers: CORS, body: JSON.stringify({ error: 'Mail o contraseña incorrectos.' }) };

    const nombre = String(row[1] ?? '').trim() || String(row[0]).trim();

    return {
      statusCode: 200,
      headers:    CORS,
      body:       JSON.stringify({ ok: true, nombre }),
    };

  } catch (err) {
    console.error('[init]', err.message);
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: 'Error del servidor: ' + err.message }) };
  }
};
