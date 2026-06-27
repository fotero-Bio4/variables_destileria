'use strict';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const CORS  = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type':                 'application/json',
};

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
  if (!d.access_token) throw new Error('Auth: ' + (d.error_description || JSON.stringify(d)));
  return d.access_token;
}

function parseFirstRow(address) {
  // address: "DB!A1:J1"  →  firstRow = 1
  const rangePart = (address || '').split('!').pop();
  const match = rangePart.match(/\$?[A-Z]+\$?(\d+)/);
  return match ? parseInt(match[1]) : 1;
}

async function readSheet(token, driveId, itemId, sheet) {
  const enc = encodeURIComponent(sheet);
  const r   = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/usedRange(valuesOnly=true)`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!r.ok) return { values: [], firstRow: 1 };
  const d = await r.json();
  return { values: d.values || [], firstRow: parseFirstRow(d.address) };
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Convierte fecha YYYY-MM-DD a serial Excel (25569 = días entre 1900-01-01 y 1970-01-01)
function dateToExcel(dateStr) {
  return Math.floor(Date.parse(dateStr + 'T00:00:00Z') / 86400000) + 25569;
}

async function patchRange(token, driveId, itemId, sheet, addr, rowValues, rowFormats) {
  const enc = encodeURIComponent(sheet);
  const r   = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/range(address='${addr}')`,
    {
      method:  'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ values: [rowValues], numberFormat: [rowFormats] }),
    }
  );
  if (!r.ok) throw new Error(`Patch ${addr}: ${r.status} ${await r.text()}`);
}

// Calcula el número Excel (1-indexed) de la siguiente fila libre, mirando la columna clave.
function nextRowFromRows(rows, firstRow, keyColIdx = 0) {
  let lastDataRow = firstRow; // encabezados en Excel row = firstRow
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][keyColIdx];
    if (val !== null && val !== undefined && val !== '' && val !== 0) {
      lastDataRow = firstRow + i;
    }
  }
  return lastDataRow + 1;
}

// ── Validar login ───────────────────────────────────────────────────────────────
async function validateUser(token, driveId, itemId, mail, pss) {
  const { values: rows } = await readSheet(token, driveId, itemId, 'usarios');
  const row  = rows.slice(1).find(r =>
    String(r[0] ?? '').trim().toLowerCase() === String(mail).trim().toLowerCase() &&
    String(r[2] ?? '').trim() === String(pss).trim()
  );
  if (!row) return null;
  return { nombre: String(row[1] ?? '').trim() || String(row[0]).trim() };
}

// ── Handler ───────────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  if (event.httpMethod !== 'POST')    return { statusCode: 405, headers: CORS, body: JSON.stringify({ error: 'Método no permitido' }) };

  const { BD_DRIVE_ID, BD_ITEM_ID, GRAPH_CLIENT_SECRET } = process.env;
  if (!GRAPH_CLIENT_SECRET || !BD_DRIVE_ID || !BD_ITEM_ID)
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: 'Servidor no configurado.' }) };

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'Cuerpo inválido.' }) }; }

  const { mail, pss, fecha, valores } = body;
  if (!mail || !pss) return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'Credenciales requeridas.' }) };
  if (!fecha)        return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'La fecha es obligatoria.' }) };
  if (!valores || typeof valores !== 'object') return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'Faltan los valores.' }) };

  try {
    const token = await getToken();

    // Re-validar credenciales (no confiar en el cliente para el nombre)
    const user = await validateUser(token, BD_DRIVE_ID, BD_ITEM_ID, mail, pss);
    if (!user) return { statusCode: 401, headers: CORS, body: JSON.stringify({ error: 'Credenciales incorrectas.' }) };

    // Leer encabezados de la hoja DB
    const { values: rows, firstRow } = await readSheet(token, BD_DRIVE_ID, BD_ITEM_ID, 'DB');
    const header = rows[0] || [];
    if (!header.length) throw new Error('La hoja DB no tiene encabezados.');

    const nCols   = header.length;
    const nextRow = nextRowFromRows(rows, firstRow, 0); // col A (Fecha) como clave

    const excelFecha = dateToExcel(fecha);
    const rowData    = Array(nCols).fill('');
    const rowFormat  = Array(nCols).fill('General');

    for (let c = 0; c < nCols; c++) {
      const name = String(header[c] ?? '').trim();
      const low  = name.toLowerCase();

      if (low === 'fecha') {
        rowData[c]   = excelFecha;
        rowFormat[c] = 'dd/mm/yyyy';
      } else if (low === 'usuario') {
        rowData[c] = user.nombre;
      } else {
        // Buscar el valor por nombre exacto de columna
        if (!Object.prototype.hasOwnProperty.call(valores, name)) {
          throw new Error(`Falta el valor para la columna "${name}".`);
        }
        const num = Number(valores[name]);
        if (!Number.isFinite(num)) throw new Error(`El valor de "${name}" no es numérico.`);
        rowData[c] = num;
      }
    }

    const addr = `A${nextRow}:${colLetter(nCols)}${nextRow}`;
    await patchRange(token, BD_DRIVE_ID, BD_ITEM_ID, 'DB', addr, rowData, rowFormat);

    return { statusCode: 200, headers: CORS, body: JSON.stringify({ ok: true }) };

  } catch (err) {
    console.error('[submit]', err.message);
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: 'Error del servidor: ' + err.message }) };
  }
};
