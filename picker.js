'use strict';

/**
 * Selector de rango de fechas en el navegador.
 *
 * Levanta un servidor HTTP local (node:http, sin dependencias), abre el
 * navegador con un calendario visual (presets + inputs nativos de fecha/hora)
 * y espera a que el usuario envíe el rango. La validación final ocurre
 * en el servidor (autoritativa), la del navegador es solo feedback inline.
 *
 * Uso desde index.js:
 *   const { startDate, endDate } = await askDateRangeWeb();
 *
 * Fallback si no se quiere el navegador: flag --terminal en index.js
 * conserva los prompts de enquirer.
 */

const http = require('http');
const { spawn } = require('child_process');

const PICKER_TIMEOUT_MS = 600_000; // 10 min esperando al usuario

const PAGE_HTML = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="data:,">
<title>Transferencias del día — Rango de fechas</title>
<style>
  :root { --azul:#003399; --borde:#c9d4e4; --bg:#f4f7fb; }
  * { box-sizing:border-box; }
  body { font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); margin:0;
         display:flex; min-height:100vh; align-items:center; justify-content:center; }
  .card { background:#fff; border-radius:12px; box-shadow:0 8px 30px rgba(0,51,153,.12);
          padding:32px 36px; width:560px; max-width:95vw; }
  h1 { font-size:20px; color:var(--azul); margin:0 0 4px; }
  .sub { color:#667; font-size:13px; margin:0 0 20px; }
  .presets { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:22px; }
  .presets button { border:1px solid var(--borde); background:#fff; color:#234; padding:7px 14px;
                    border-radius:20px; cursor:pointer; font-size:13px; transition:.15s; }
  .presets button:hover { border-color:var(--azul); color:var(--azul); }
  .range { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
  .field label { display:block; font-size:12px; font-weight:600; color:#345; margin-bottom:6px; }
  .field input { width:100%; padding:9px 10px; border:1px solid var(--borde); border-radius:8px;
                 font-size:15px; font-family:inherit; }
  .field input:focus { outline:2px solid rgba(0,51,153,.25); border-color:var(--azul); }
  .msg { min-height:20px; font-size:13px; margin-top:14px; }
  .msg.err { color:#c0392b; }
  .msg.warn { color:#b7791f; }
  .msg.ok { color:#1e7e34; }
  .btn { margin-top:6px; width:100%; background:var(--azul); color:#fff; border:0; padding:13px;
         border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; font-family:inherit; }
  .btn:hover { background:#002a7d; }
  .btn:disabled { background:#9db4d8; cursor:default; }
  .note { font-size:12px; color:#889; text-align:center; margin-top:16px; }
</style>
</head>
<body>
<div class="card">
  <h1>📊 Transferencias del día</h1>
  <p class="sub">Seleccioná el rango de fechas a extraer (máximo 1 semana)</p>

  <div class="presets">
    <button data-start="-1" data-end="0">Ayer → Hoy</button>
    <button data-start="0" data-end="0">Solo hoy</button>
    <button data-start="-1" data-end="-1">Solo ayer</button>
    <button data-start="-2" data-end="0">Últimos 3 días</button>
    <button data-start="week" data-end="0">Esta semana</button>
  </div>

  <div class="range">
    <div class="field">
      <label>DESDE</label>
      <input type="date" id="startDate">
      <input type="time" id="startTime" value="08:30" style="margin-top:6px">
    </div>
    <div class="field">
      <label>HASTA</label>
      <input type="date" id="endDate">
      <input type="time" id="endTime" value="23:59" style="margin-top:6px">
    </div>
  </div>

  <div class="msg" id="msg"></div>
  <button class="btn" id="go">▶ EMPEZAR EXTRACCIÓN</button>
  <p class="note">El QR de WhatsApp aparecerá en la terminal. Podés cerrar esta pestaña al terminar.</p>
</div>

<script>
  const $ = id => document.getElementById(id);
  const DAY = 86400000;
  const fmt = dt => {
    const y = dt.getFullYear(), m = String(dt.getMonth()+1).padStart(2,'0'), d = String(dt.getDate()).padStart(2,'0');
    return y + '-' + m + '-' + d;
  };
  const off = n => { const x = new Date(); x.setDate(x.getDate() + n); return x; };

  const now = new Date();
  $('startDate').value = fmt(off(-1));
  $('endDate').value = fmt(now);

  function validate() {
    const s = new Date($('startDate').value + 'T' + ($('startTime').value || '00:00'));
    const e = new Date($('endDate').value + 'T' + ($('endTime').value || '23:59'));
    const msg = $('msg');
    if (isNaN(s) || isNaN(e)) {
      msg.textContent = 'Completá las fechas y horas.';
      msg.className = 'msg err';
      return null;
    }
    if (e <= s) {
      msg.textContent = '❌ La fecha/hora FIN debe ser posterior al INICIO.';
      msg.className = 'msg err';
      return null;
    }
    if (e - s > 7 * DAY) {
      msg.textContent = '⚠ Supera 1 semana. ¿Estás seguro?';
      msg.className = 'msg warn';
    } else {
      msg.textContent = '';
    }
    return { startDate: s.toISOString(), endDate: e.toISOString() };
  }

  document.querySelectorAll('.presets button').forEach(b => b.addEventListener('click', () => {
    let start;
    if (b.dataset.start === 'week') {
      // Lunes de esta semana (si hoy es domingo, el lunes es hace 6 días)
      start = now.getDay() === 0 ? off(-6) : off(1 - now.getDay());
    } else {
      start = off(Number(b.dataset.start));
    }
    $('startDate').value = fmt(start);
    $('endDate').value = fmt(off(Number(b.dataset.end)));
    validate();
  }));

  ['startDate', 'endDate', 'startTime', 'endTime'].forEach(id => $(id).addEventListener('change', validate));

  $('go').addEventListener('click', async () => {
    const range = validate();
    if (!range) return;
    $('go').disabled = true;
    $('go').textContent = 'Enviando...';
    try {
      const r = await fetch('/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(range),
      });
      const j = await r.json();
      if (j.ok) {
        $('msg').textContent = '✅ ¡Listo! Podés cerrar esta pestaña y volver a la terminal.';
        $('msg').className = 'msg ok';
        $('go').textContent = '✓ Enviado';
      } else {
        $('msg').textContent = '❌ ' + j.error;
        $('msg').className = 'msg err';
        $('go').disabled = false;
        $('go').textContent = '▶ EMPEZAR EXTRACCIÓN';
      }
    } catch {
      $('msg').textContent = '❌ No se pudo contactar la aplicación. ¿Se cerró la terminal?';
      $('msg').className = 'msg err';
      $('go').disabled = false;
      $('go').textContent = '▶ EMPEZAR EXTRACCIÓN';
    }
  });
</script>
</body>
</html>
`;

/** Abre el navegador por defecto (macOS: open, Windows: start, Linux: xdg-open). */
function openBrowser(url) {
  if (process.platform === 'win32') {
    spawn('start', [url], { shell: true, detached: true, stdio: 'ignore' }).unref();
  } else {
    const cmd = process.platform === 'darwin' ? 'open' : 'xdg-open';
    spawn(cmd, [url], { detached: true, stdio: 'ignore' }).unref();
  }
}

/**
 * Levanta el servidor del picker. No abre el navegador (eso lo hace
 * askDateRangeWeb) — así los tests pueden ejercitar los endpoints sin UI.
 *
 * @param {Function} onSubmit  (range) => void  — se llama con el rango validado
 * @returns {Promise<{server, port, url, close}>}
 */
function createPickerServer(onSubmit) {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      if (req.method === 'GET' && req.url === '/') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(PAGE_HTML);
        return;
      }

      if (req.method === 'POST' && req.url === '/submit') {
        let body = '';
        req.on('data', c => { body += c; });
        req.on('end', () => {
          let range;
          try {
            range = JSON.parse(body);
          } catch {
            send(res, 400, { ok: false, error: 'JSON inválido.' });
            return;
          }

          const startDate = new Date(range?.startDate);
          const endDate = new Date(range?.endDate);

          if (isNaN(startDate) || isNaN(endDate)) {
            send(res, 400, { ok: false, error: 'Fechas inválidas.' });
            return;
          }
          if (endDate <= startDate) {
            send(res, 400, { ok: false, error: 'La fecha/hora FIN debe ser posterior al INICIO.' });
            return;
          }

          send(res, 200, { ok: true });
          try { onSubmit({ startDate, endDate }); } catch { /* noop */ }
          server.close();
        });
        return;
      }

      send(res, 404, { ok: false, error: 'No encontrado.' });
    });

    server.on('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      resolve({
        server,
        port,
        url: `http://127.0.0.1:${port}/`,
        close: () => new Promise(r => server.close(() => r())),
      });
    });
  });
}

function send(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}

/**
 * Flujo completo para index.js: servidor + navegador + espera del rango.
 * Devuelve { startDate, endDate } como Date. Lanza error si el usuario
 * no responde en PICKER_TIMEOUT_MS.
 */
async function askDateRangeWeb(timeoutMs = PICKER_TIMEOUT_MS) {
  return new Promise((resolve, reject) => {
    let settled = false;
    let serverRef = null;
    const settle = (fn, v) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      fn(v);
    };

    createPickerServer((range) => {
      const summary =
        `   📅 INICIO : ${range.startDate.toLocaleString('es-CO')}\n` +
        `   📅 FIN    : ${range.endDate.toLocaleString('es-CO')}`;
      console.log(`\n✅ Rango recibido:\n${summary}\n`);
      settle(resolve, range);
    }).then(({ url, close }) => {
      serverRef = close;
      console.log('\n══════════════════════════════════════════════════════');
      console.log('  📅 SELECTOR DE FECHAS — se abrió el navegador');
      console.log('══════════════════════════════════════════════════════');
      console.log(`  Si no se abrió, entrá a: ${url}`);
      console.log('  Elegí el rango y tocá "EMPEZAR EXTRACCIÓN".\n');
      openBrowser(url);
    }).catch(err => settle(reject, err));

    const timer = setTimeout(() => {
      // Cerrar el servidor para que el proceso no quede colgado escuchando
      if (serverRef) serverRef();
      settle(reject, new Error(
        `Tiempo de espera agotado (${Math.round(timeoutMs / 60_000)} min) sin recibir rango en el navegador.\n` +
        '   Ejecutá de nuevo, o usá el modo terminal:  node index.js --terminal'
      ));
    }, timeoutMs);
  });
}

module.exports = { createPickerServer, askDateRangeWeb, openBrowser };
