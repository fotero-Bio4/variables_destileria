'use strict';

// ── Definición de variables (orden y nombre EXACTO de la columna en la hoja DB) ─
// El campo `col` debe coincidir carácter por carácter con el encabezado de la hoja DB.
const VARIABLES = [
  { col: 'GL IC445',                 label: 'GL IC445' },
  { col: 'Caudal agua TL440',        label: 'Caudal agua TL440' },
  { col: 'Agua bomba de vacio 443',  label: 'Agua bomba de vacío 443' },
  { col: 'Caudal fusel TD440',       label: 'Caudal fusel TD440' },
  { col: 'Caudal agua fusel  SC440', label: 'Caudal agua fusel SC440' },
  { col: 'GL agua de sello BO 443',  label: 'GL agua de sello BO 443' },
  { col: 'Presión BO433',            label: 'Presión BO433' },
  { col: 'GL agua lavadora TL440',   label: 'GL agua lavadora TL440' },
];

// ── Estado de la sesión ─────────────────────────────────────────────────────
let session = {
  mail:   '',
  pss:    '',
  nombre: '',   // Nombre del usuario (lo que se carga en la columna "Usuario" de DB)
};

// ── Pantallas ────────────────────────────────────────────────────────────────
const screens = {
  login:        document.getElementById('screenLogin'),
  form:         document.getElementById('screenForm'),
  validacion:   document.getElementById('screenValidacion'),
  confirmacion: document.getElementById('screenConfirmacion'),
};

function showScreen(name) {
  Object.values(screens).forEach(s => s.classList.add('hidden'));
  screens[name].classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── Mensaje helper ─────────────────────────────────────────────────────────────
function showMsg(el, type, text) {
  el.className = 'msg ' + type;
  el.textContent = text;
  el.classList.remove('hidden');
  el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function hideMsg(el) { el.classList.add('hidden'); }

// ── Fecha helper ─────────────────────────────────────────────────────────────
// Fecha local de hoy en formato YYYY-MM-DD (para el <input type="date">).
function fechaHoyISO() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

// Convierte YYYY-MM-DD a formato legible DD/MM/AAAA (para la pantalla de validación).
function fechaLegible(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

// ── Render dinámico de los campos del formulario ───────────────────────────────
function renderFormFields() {
  const cont = document.getElementById('formFields');
  cont.innerHTML = '';
  VARIABLES.forEach((v, i) => {
    const field = document.createElement('div');
    field.className = 'field';
    field.innerHTML = `
      <label for="var_${i}">${v.label} <span class="req">*</span></label>
      <input id="var_${i}" type="number" step="any" inputmode="decimal" placeholder="0">
    `;
    cont.appendChild(field);
  });
}

// ── LOGIN ──────────────────────────────────────────────────────────────────────
document.getElementById('btnLogin').addEventListener('click', doLogin);
document.getElementById('inputPss').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });

async function doLogin() {
  const btn   = document.getElementById('btnLogin');
  const msgEl = document.getElementById('msgLogin');
  const mail  = document.getElementById('inputMail').value.trim();
  const pss   = document.getElementById('inputPss').value.trim();

  hideMsg(msgEl);
  if (!mail || !pss) { showMsg(msgEl, 'error', 'Ingresá tu mail y contraseña.'); return; }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Verificando...';

  try {
    const resp = await fetch('/.netlify/functions/init', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ mail, pss }),
    });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
      showMsg(msgEl, 'error', data.error || 'Error al iniciar sesión.');
      return;
    }

    // Guardar sesión
    session.mail   = mail;
    session.pss    = pss;
    session.nombre = data.nombre;

    // Saludo en el header
    document.getElementById('headerBadge').textContent = 'Bienvenido ' + data.nombre;

    // Ir al formulario
    goToForm();
  } catch (err) {
    showMsg(msgEl, 'error', 'Error de conexión: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = 'Ingresar';
  }
}

// ── FORMULARIO ──────────────────────────────────────────────────────────────────
function goToForm() {
  renderFormFields();
  document.getElementById('campo_fecha').value = fechaHoyISO();
  hideMsg(document.getElementById('msgForm'));
  showScreen('form');
}

// Lee los valores del formulario. Devuelve { fecha, valores } o null si hay error (ya mostrado).
function leerFormulario() {
  const msgEl = document.getElementById('msgForm');
  hideMsg(msgEl);

  const fecha = document.getElementById('campo_fecha').value;
  if (!fecha) { showMsg(msgEl, 'error', 'La fecha es obligatoria.'); return null; }

  const valores = {};
  for (let i = 0; i < VARIABLES.length; i++) {
    const raw = document.getElementById('var_' + i).value.trim();
    if (raw === '') {
      showMsg(msgEl, 'error', `Falta completar: ${VARIABLES[i].label}.`);
      document.getElementById('var_' + i).focus();
      return null;
    }
    const num = Number(raw);
    if (!Number.isFinite(num)) {
      showMsg(msgEl, 'error', `El valor de "${VARIABLES[i].label}" debe ser numérico.`);
      document.getElementById('var_' + i).focus();
      return null;
    }
    valores[VARIABLES[i].col] = num;
  }
  return { fecha, valores };
}

// Botón "Revisar datos" → arma la pantalla de validación
document.getElementById('btnRevisar').addEventListener('click', () => {
  const datos = leerFormulario();
  if (!datos) return;
  construirValidacion(datos);
  showScreen('validacion');
});

// Cerrar sesión desde el formulario
document.getElementById('btnSalir').addEventListener('click', cerrarSesion);

// ── VALIDACIÓN ──────────────────────────────────────────────────────────────────
function construirValidacion({ fecha, valores }) {
  const tabla = document.getElementById('validacionTabla');
  let html = `
    <tr><th>Fecha</th><td>${fechaLegible(fecha)}</td></tr>
    <tr><th>Usuario</th><td>${session.nombre}</td></tr>
  `;
  VARIABLES.forEach(v => {
    html += `<tr><th>${v.label}</th><td>${valores[v.col]}</td></tr>`;
  });
  tabla.innerHTML = html;
  hideMsg(document.getElementById('msgValidacion'));
}

// Volver a editar
document.getElementById('btnVolverEditar').addEventListener('click', () => {
  hideMsg(document.getElementById('msgValidacion'));
  showScreen('form');
});

// Confirmar y guardar
document.getElementById('btnConfirmar').addEventListener('click', submitDatos);

async function submitDatos() {
  const btn   = document.getElementById('btnConfirmar');
  const msgEl = document.getElementById('msgValidacion');
  hideMsg(msgEl);

  // Releer del formulario para garantizar consistencia con lo mostrado
  const fecha = document.getElementById('campo_fecha').value;
  const valores = {};
  for (let i = 0; i < VARIABLES.length; i++) {
    valores[VARIABLES[i].col] = Number(document.getElementById('var_' + i).value);
  }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Guardando...';

  try {
    const resp = await fetch('/.netlify/functions/submit', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({
        mail: session.mail,
        pss:  session.pss,
        fecha,
        valores,
      }),
    });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
      showMsg(msgEl, 'error', data.error || 'Error al guardar.');
      return;
    }

    showScreen('confirmacion');
  } catch (err) {
    showMsg(msgEl, 'error', 'Error de conexión: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '✓ Confirmar y guardar';
  }
}

// ── CONFIRMACIÓN ────────────────────────────────────────────────────────────────
document.getElementById('btnNuevaCarga').addEventListener('click', goToForm);

// ── CERRAR SESIÓN ────────────────────────────────────────────────────────────────
function cerrarSesion() {
  session = { mail: '', pss: '', nombre: '' };
  document.getElementById('inputPss').value = '';
  document.getElementById('headerBadge').textContent = 'Acceso';
  hideMsg(document.getElementById('msgLogin'));
  showScreen('login');
}
