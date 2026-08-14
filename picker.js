'use strict';

/**
 * Panel de control en el navegador para la extracción.
 *
 * Levanta un servidor HTTP local (node:http, sin dependencias nuevas) con:
 *
 *   1. SELECCIÓN DE RANGO — página con presets + inputs nativos de fecha/hora;
 *      la validación final ocurre en el servidor (autoritativa).
 *
 *   2. ESTADO EN VIVO — después de enviar el rango, la página muestra:
 *        • el QR de WhatsApp como imagen (QRCode.toDataURL, lib ya instalada)
 *        • el log de la terminal en tiempo real (Server-Sent Events, SSE)
 *        • el archivo Word generado con botón de descarga
 *
 * El servidor solo escucha en 127.0.0.1. Se unref() para no mantener vivo
 * el proceso al terminar.
 *
 * Flujo desde index.js:
 *   const picker = await startControlServer();
 *   picker.attachConsole();                 // tee console → página
 *   const range = await picker.waitForRange();
 *   ... (cuando llega QR) picker.sendQR(qr)
 *   ... (al final)         picker.notifyFile(OUTPUT_FILE)
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { config, updateConfig } = require('./config');
const QRCode = require('qrcode');

const PICKER_TIMEOUT_MS = 600_000;  // 10 min esperando el rango
const EVENT_BUFFER_MAX = 200;
const SSE_KEEPALIVE_MS = 25_000;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const PAGE_HTML = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="data:,">
<title>Transferencias del día — Panel</title>
<style>
  :root { --azul:#003399; --borde:#c9d4e4; --bg:#f4f7fb; }
  * { box-sizing:border-box; }
  body { font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); margin:0;
         display:flex; min-height:100vh; align-items:center; justify-content:center; padding:16px; }
  .card { background:#fff; border-radius:12px; box-shadow:0 8px 30px rgba(0,51,153,.12);
          padding:32px 36px; width:580px; max-width:100%; }
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
  .msg.err { color:#c0392b; } .msg.warn { color:#b7791f; } .msg.ok { color:#1e7e34; }
  .btn { margin-top:6px; width:100%; background:var(--azul); color:#fff; border:0; padding:13px;
         border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; font-family:inherit; }
  .btn:hover { background:#002a7d; } .btn:disabled { background:#9db4d8; cursor:default; }
  .note { font-size:12px; color:#889; text-align:center; margin-top:16px; }
  .hidden { display:none; }
  .qrbox { text-align:center; margin:10px 0 4px; }
  .qrbox img { width:280px; height:280px; image-rendering:pixelated; }
  .logbox { background:#10141c; color:#d6e2f0; border-radius:8px; padding:10px 12px;
            font-family:ui-monospace,Consolas,monospace; font-size:12px; line-height:1.5;
            height:260px; overflow-y:auto; margin:14px 0 6px; white-space:pre-wrap; word-break:break-word; }
  .fileready { border:1px solid #b8e0c0; background:#f0faf2; border-radius:8px; padding:14px 16px; margin-top:14px; }
  .fileready .fname { font-weight:600; color:#1e5a2e; font-size:15px; }
  .fileready .fsize { color:#557; font-size:12px; margin-top:2px; }
  .dl { display:inline-block; margin-top:10px; background:#1e7e34; color:#fff; text-decoration:none;
        padding:11px 18px; border-radius:8px; font-weight:600; font-size:14px; }
  .dl:hover { background:#175f28; }
  .badge { display:inline-block; font-size:11px; font-weight:600; padding:3px 10px; border-radius:12px; }
  .badge.wait { background:#fef3c7; color:#8a6d1a; }
  .badge.live { background:#dcfce7; color:#166534; }
  .badge.bad { background:#fee2e2; color:#991b1b; }
  .connform { display:grid; grid-template-columns:1fr 1fr; gap:8px 12px; margin:10px 0; }
  .connform label { font-size:11px; font-weight:600; color:#345; display:block; }
  .connform input { width:100%; margin-top:3px; padding:7px 9px; font-size:13px;
                    border:1px solid #c8d4e8; border-radius:6px; font-family:inherit; box-sizing:border-box; }
  .connbtns { display:flex; gap:8px; flex-wrap:wrap; margin-top:4px; }
  .connnote { font-size:11px; color:#889; margin-top:6px; min-height:14px; }
  .mapcard { margin-top:16px; border:1px solid var(--borde); border-radius:8px; padding:12px 14px; }
  .maphead { display:flex; align-items:center; justify-content:space-between; margin-bottom:2px; }
  .maphead h3 { margin:0; font-size:14px; color:var(--azul); }
  .maphead button { border:1px solid var(--borde); background:#fff; color:#234; padding:5px 12px;
                    border-radius:16px; cursor:pointer; font-size:12px; }
  .maphead button:hover { border-color:var(--azul); color:var(--azul); }
  .mapsub { font-size:11px; color:#889; margin:4px 0 10px; }
  .maplist { max-height:300px; overflow-y:auto; }
  .maprow { display:grid; grid-template-columns:1fr 1.2fr auto; gap:8px; align-items:center; margin-bottom:8px; }
  .maprow input { width:100%; padding:7px 9px; border:1px solid var(--borde); border-radius:6px;
                  font-size:13px; font-family:inherit; }
  .maprow input:focus { outline:2px solid rgba(0,51,153,.25); border-color:var(--azul); }
  .mapid b { font-size:12px; color:#345; }
  .mapid .mapidsub { display:block; font-size:10px; color:#99a; }
  .mapsave { background:var(--azul); color:#fff; border:0; padding:7px 12px; border-radius:6px;
             font-size:12px; cursor:pointer; font-family:inherit; }
  .mapsave:hover { background:#002a7d; } .mapsave:disabled { background:#9db4d8; cursor:default; }
  .mapbtn { margin-top:6px; width:auto; padding:9px 18px; font-size:13px; }
  .mapempty { font-size:12px; color:#889; padding:8px 0; }
</style>
</head>
<body>
<div class="card">
  <h1>📊 Transferencias del día</h1>
  <p class="sub" id="subtitle">Seleccioná el rango de fechas a extraer (máximo 1 semana)</p>

  <!-- FASE 1: rango -->
  <div id="rangePhase">
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
    <p class="note">Al enviar, la página pasa a modo estado y la terminal sigue mostrando todo.</p>
  </div>

  <!-- FASE 2: estado en vivo -->
  <div id="statusPhase" class="hidden">
    <p class="sub">
      <span class="badge wait" id="phaseBadge">esperando conexión…</span>
      &nbsp;El progreso también se muestra en la terminal.
    </p>
    <div class="qrbox hidden" id="qrbox">
      <img id="qrImg" alt="Código QR — escanealo con WhatsApp">
      <p class="note">Escaneá este QR con WhatsApp → Dispositivos vinculados.</p>
    </div>
    <div class="logbox" id="log"></div>
    <div class="fileready hidden" id="fileready">
      <div class="fname" id="fname"></div>
      <div class="fsize" id="fsize"></div>
      <a class="dl" id="dl" href="/download" download>⬇ Descargar Word</a>
    </div>
  </div>

  <!-- FASE 3: mapear mensajeros (IDs → nombres) — visible desde el inicio -->
  <div class="mapcard" id="mapcard">
    <div class="maphead">
      <h3>👥 Mensajeros del grupo</h3>
      <button id="mapRefresh">Actualizar</button>
    </div>
    <p class="mapsub">Poné el nombre de cada mensajero — se guarda al instante en nombres_mensajeros.json y se aplica a este reporte.</p>
    <div class="maplist" id="maplist"></div>
    <button class="btn mapbtn" id="mapSaveAll">💾 Guardar todos</button>
  </div>

  <!-- FASE 4: conexión WhatsApp (configurable desde acá, sin tocar config.json) -->
  <div class="mapcard" id="conncard">
    <div class="maphead">
      <h3>📡 Conexión WhatsApp</h3>
      <span class="badge wait" id="connBadge">cargando…</span>
    </div>
    <div class="connform">
      <label>URL del VPS
        <input type="text" id="cfgUrl" placeholder="https://transferencias.redpostal.co">
      </label>
      <label>Instancia de Evolution
        <select id="cfgInstance"><option value="">cargando…</option></select>
      </label>
      <label>Nombre del grupo
        <input type="text" id="cfgGroup" placeholder="transferencias">
      </label>
      <label>API key
        <input type="password" id="cfgKey" placeholder="••••••••  (dejar en blanco = mantener)">
      </label>
    </div>
    <div class="connbtns">
      <button class="btn mapbtn" id="cfgSave">💾 Guardar conexión</button>
      <button class="btn mapbtn" id="connReconnect">📲 Reconectar WhatsApp (QR)</button>
    </div>
    <p class="connnote" id="connMsg"></p>
  </div>
</div>
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
    if (isNaN(s) || isNaN(e)) { msg.textContent = 'Completá las fechas y horas.'; msg.className = 'msg err'; return null; }
    if (e <= s) { msg.textContent = '❌ La fecha/hora FIN debe ser posterior al INICIO.'; msg.className = 'msg err'; return null; }
    if (e - s > 7 * DAY) { msg.textContent = '⚠ Supera 1 semana. ¿Estás seguro?'; msg.className = 'msg warn'; }
    else msg.textContent = '';
    return { startDate: s.toISOString(), endDate: e.toISOString() };
  }

  document.querySelectorAll('.presets button').forEach(b => b.addEventListener('click', () => {
    let start;
    if (b.dataset.start === 'week') start = now.getDay() === 0 ? off(-6) : off(1 - now.getDay());
    else start = off(Number(b.dataset.start));
    $('startDate').value = fmt(start);
    $('endDate').value = fmt(off(Number(b.dataset.end)));
    validate();
  }));
  ['startDate','endDate','startTime','endTime'].forEach(id => $(id).addEventListener('change', validate));

  $('go').addEventListener('click', async () => {
    const range = validate();
    if (!range) return;
    $('go').disabled = true; $('go').textContent = 'Enviando...';
    try {
      const r = await fetch('/submit', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(range) });
      const j = await r.json();
      if (!j.ok) {
        $('msg').textContent = '❌ ' + j.error; $('msg').className = 'msg err';
        $('go').disabled = false; $('go').textContent = '▶ EMPEZAR EXTRACCIÓN';
        return;
      }
    } catch {
      $('msg').textContent = '❌ No se pudo contactar la aplicación. ¿Se cerró la terminal?';
      $('msg').className = 'msg err';
      $('go').disabled = false; $('go').textContent = '▶ EMPEZAR EXTRACCIÓN';
      return;
    }
    // Cambiar a fase estado
    $('rangePhase').classList.add('hidden');
    $('statusPhase').classList.remove('hidden');
    $('subtitle').textContent = 'Extrayendo comprobantes…';
  });

  // ── Estado en vivo (SSE) ──
  const log = $('log');
  function appendLog(line) {
    const div = document.createElement('div');
    div.textContent = line;
    log.appendChild(div);
    while (log.childNodes.length > 500) log.removeChild(log.firstChild);
    log.scrollTop = log.scrollHeight;
  }
  function setBadge(text, cls) {
    const b = $('phaseBadge');
    b.textContent = text;
    b.className = 'badge ' + cls;
  }

  const es = new EventSource('/events');
  es.addEventListener('log', e => {
    appendLog(JSON.parse(e.data));
    setBadge('trabajando…', 'live');
  });
  es.addEventListener('qr', e => {
    $('qrImg').src = JSON.parse(e.data);
    $('qrbox').classList.remove('hidden');
    setBadge('esperando escaneo del QR', 'wait');
    appendLog('🔳 QR recibido — escanealo con WhatsApp');
  });
  es.addEventListener('file', e => {
    const f = JSON.parse(e.data);
    $('fname').textContent = f.name;
    $('fsize').textContent = (f.size / 1024).toFixed(1) + ' KB — guardado en la carpeta de la app';
    $('fileready').classList.remove('hidden');
    setBadge('✅ proceso finalizado', 'live');
    appendLog('📄 ' + f.name + ' listo (' + (f.size/1024).toFixed(1) + ' KB)');
  });
  es.addEventListener('participants', e => renderMap(JSON.parse(e.data)));
  es.onopen = () => setBadge('conectado', 'live');
  es.onerror = () => setBadge('reconectando…', 'wait');

  // ── Mapeo de mensajeros: IDs → nombres ──
  const mapList = $('maplist');
  function renderMap(list) {
    mapList.innerHTML = '';
    if (!list || !list.length) {
      const div = document.createElement('div');
      div.className = 'mapempty';
      div.textContent = 'Aún no hay datos del grupo…';
      mapList.appendChild(div);
      return;
    }
    for (const p of list) {
      const row = document.createElement('div');
      row.className = 'maprow';

      const idText = document.createElement('div');
      idText.className = 'mapid';
      const main = p.phone || (p.lid || '').replace('@lid', '');
      const sub = (p.phone && p.lid) ? p.lid.replace('@lid', '') : '';
      const b = document.createElement('b');
      b.textContent = main;          // solo números (teléfono/LID) — sin innerHTML
      idText.appendChild(b);
      if (sub) {
        const s = document.createElement('span');
        s.className = 'mapidsub';
        s.textContent = sub;
        idText.appendChild(s);
      }

      const inp = document.createElement('input');
      inp.type = 'text';
      inp.placeholder = 'Nombre del mensajero…';
      inp.value = p.name || '';
      inp.dataset.key = p.phone || p.lid || '';

      const btn = document.createElement('button');
      btn.textContent = 'Guardar';
      btn.className = 'mapsave';
      btn.onclick = () => saveName(inp, btn);
      inp.onkeydown = e => { if (e.key === 'Enter') saveName(inp, btn); };

      row.appendChild(idText); row.appendChild(inp); row.appendChild(btn);
      mapList.appendChild(row);
    }
  }
  async function saveName(inp, btn) {
    const name = inp.value.trim();
    btn.disabled = true; btn.textContent = '…';
    try {
      const r = await fetch('/api/names', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: inp.dataset.key, name }),
      });
      const j = await r.json();
      if (j.ok) { btn.textContent = '✓'; }
      else { appendLog('⚠ No se guardó: ' + j.error); btn.textContent = 'Guardar'; }
      setTimeout(() => { btn.textContent = 'Guardar'; btn.disabled = false; }, 800);
    } catch {
      btn.textContent = 'Guardar'; btn.disabled = false;
      appendLog('⚠ No se pudo guardar el nombre');
    }
  }
  $('mapSaveAll').addEventListener('click', async () => {
    const inputs = [...document.querySelectorAll('.maprow input')];
    let okCount = 0;
    for (const inp of inputs) {
      const name = inp.value.trim();
      if (!name) continue;
      try {
        const r = await fetch('/api/names', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key: inp.dataset.key, name }),
        });
        if (r.ok) okCount++;
      } catch { /* seguir con el resto */ }
    }
    appendLog('💾 ' + okCount + ' nombres guardados');
  });
  $('mapRefresh').addEventListener('click', async () => {
    try {
      const r = await fetch('/api/participants');
      const j = await r.json();
      if (j.ok) renderMap(j.participants);
    } catch { /* la página se actualiza sola por SSE */ }
  });

  // ── Configuración de conexión (panel) ──
  const connBadge = $('connBadge');
  const connMsg = $('connMsg');
  async function refreshConnState() {
    try {
      const r = await fetch('/api/instance');
      const j = await r.json();
      if (!j.ok) { connBadge.textContent = 'sin datos'; connBadge.className = 'badge bad'; return null; }
      const s = j.status;
      if (s === 'open') { connBadge.textContent = 'conectada ✓'; connBadge.className = 'badge live'; connMsg.textContent = ''; }
      else if (s === 'connecting') { connBadge.textContent = 'conectando…'; connBadge.className = 'badge wait'; }
      else if (s === 'close') { connBadge.textContent = 'desconectada'; connBadge.className = 'badge bad'; }
      else { connBadge.textContent = 'desconocido'; connBadge.className = 'badge wait'; }
      return j;
    } catch {
      connBadge.textContent = 'sin conexión con la app'; connBadge.className = 'badge bad';
      return null;
    }
  }
  (async function loadCfg() {
    const instSel = $('cfgInstance');
    try {
      const r = await fetch('/api/config');
      const j = await r.json();
      if (j.ok) {
        $('cfgUrl').value = j.evolutionUrl || '';
        $('cfgGroup').value = j.groupName || '';
        $('cfgKey').placeholder = j.hasApiKey ? '••••••••  (dejar en blanco = mantener)' : 'API key (obligatoria)';
        // Selector: instancias existentes en Evolution + la configurada
        try {
          const ri = await fetch('/api/instances');
          const ji = await ri.json();
          const opts = ji.instances || [];
          instSel.innerHTML = '';
          if (j.instance) {
            const cur = document.createElement('option');
            cur.value = j.instance; cur.textContent = j.instance + ' (configurada)';
            instSel.appendChild(cur);
          }
          for (const i of opts) {
            if (i.name && i.name !== j.instance) {
              const o = document.createElement('option');
              o.value = i.name; o.textContent = i.name + (i.status === 'open' ? ' ✓' : (i.number ? ' — ' + i.number : ''));
              instSel.appendChild(o);
            }
          }
          if (opts.length === 0 && !j.instance) {
            const e = document.createElement('option');
            e.value = ''; e.textContent = '(sin instancias — creala en Evolution)';
            instSel.appendChild(e);
          }
        } catch {
          // sin lista → dejar el valor actual como opción única
          instSel.innerHTML = '';
          const cur = document.createElement('option');
          cur.value = j.instance || ''; cur.textContent = j.instance || '(sin instancia)';
          instSel.appendChild(cur);
        }
      }
    } catch { /* la app puede no estar lista aún */ }
  })();
  refreshConnState();

  $('cfgSave').addEventListener('click', async () => {
    const body = {};
    const url = $('cfgUrl').value.trim(); if (url) body.evolutionUrl = url;
    const inst = $('cfgInstance').value.trim(); if (inst) body.instance = inst;
    const grp = $('cfgGroup').value.trim(); if (grp) body.groupName = grp;
    const key = $('cfgKey').value.trim(); if (key) body.apiKey = key;
    if (!Object.keys(body).length) { connMsg.textContent = 'Sin cambios para guardar'; return; }
    const btn = $('cfgSave');
    btn.disabled = true; btn.textContent = '…';
    try {
      const r = await fetch('/api/config', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const j = await r.json();
      if (j.ok) {
        connMsg.textContent = '✓ Conexión guardada — aplica al próximo run';
        $('cfgKey').value = '';
        $('cfgKey').placeholder = '••••••••  (dejar en blanco = mantener)';
        refreshConnState();
      } else { connMsg.textContent = '❌ ' + (j.error || 'error'); }
    } catch { connMsg.textContent = '❌ No se pudo contactar la aplicación'; }
    btn.disabled = false; btn.textContent = '💾 Guardar conexión';
  });

  let qrPoll = null;
  $('connReconnect').addEventListener('click', async () => {
    const btn = $('connReconnect');
    btn.disabled = true; btn.textContent = '…';
    connMsg.textContent = 'Solicitando QR a Evolution…';
    try {
      const r = await fetch('/api/connect', { method: 'POST' });
      const j = await r.json();
      if (!j.ok || !j.qrcode) { connMsg.textContent = '❌ ' + (j.error || 'Evolution no devolvió QR'); return; }
      // El QR vive en el recuadro de la fase de extracción → mostrarla
      const origTitle = $('subtitle').textContent;
      $('subtitle').textContent = 'Escaneá el QR con WhatsApp en el teléfono de la SIM';
      $('rangePhase').classList.add('hidden');
      $('statusPhase').classList.remove('hidden');
      $('qrImg').src = j.qrcode;
      $('qrbox').classList.remove('hidden');
      setBadge('esperando escaneo del QR', 'wait');
      appendLog('🔳 QR de reconexión — escanealo con WhatsApp');
      // Pollear hasta que la instancia quede open (o el usuario cierre la pestaña)
      qrPoll = setInterval(async () => {
        const st = await refreshConnState();
        if (st && st.ok && st.status === 'open') {
          clearInterval(qrPoll); qrPoll = null;
          $('statusPhase').classList.add('hidden');
          $('rangePhase').classList.remove('hidden');
          $('subtitle').textContent = origTitle;
          connMsg.textContent = '✓ WhatsApp reconectado';
          appendLog('✅ WhatsApp reconectado');
        }
      }, 3000);
    } catch {
      connMsg.textContent = '❌ No se pudo contactar la aplicación';
    } finally {
      btn.disabled = false; btn.textContent = '📲 Reconectar WhatsApp (QR)';
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

function sseSend(res, event, data) {
  res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
}

/**
 * Levanta el servidor de control. NO abre el navegador (eso lo hace index.js
 * con handle.openBrowser()) para que los tests puedan ejercitar todo sin UI.
 *
 * Devuelve un handle con:
 *   server, port, url, close()
 *   waitForRange(timeoutMs) → Promise<{startDate, endDate}>  (Date)
 *   pushLog(line)        — línea de log → página
 *   sendQR(qr)           — string QR → imagen PNG en la página
 *   notifyFile(filePath) — marca el Word como listo + habilita /download
 *   attachConsole()      — tee de console.log/error hacia la página
 *   openBrowser()        — abre el navegador con la página
 */
function startControlServer() {
  return new Promise((resolve, reject) => {
    const handle = {
      server: null,
      fileInfo: null,          // { name, size, path }
      events: [],              // buffer de replay { event, data }
      clients: new Set(),      // res SSE activos
      _range: null,            // rango ya recibido
      _waiters: [],            // resolvers de waitForRange pendientes
      _dlWaiters: [],          // resolvers de waitForDownload pendientes
      _downloads: 0,           // veces que se descargó el archivo
      _participants: [],       // lista para mapear nombres { lid, phone, name }
      _nameSaver: null,        // callback (key, name) → persiste en index.js
      _instanceChecker: null,  // () → {found, status} estado de la instancia
      _connector: null,        // () → data de /instance/connect (QR) o {error}
      _configApplied: null,    // ({instance, changed}) → se llama tras guardar config
    };

    function emit(event, data) {
      handle.events.push({ event, data });
      if (handle.events.length > EVENT_BUFFER_MAX) handle.events.shift();
      for (const res of handle.clients) {
        try { sseSend(res, event, data); }
        catch { handle.clients.delete(res); }
      }
    }

    handle.pushLog = (line) => emit('log', String(line));
    handle.sendQR = (qr) => {
      QRCode.toDataURL(qr, { width: 300, margin: 1 })
        .then(dataUrl => emit('qr', dataUrl))
        .catch(() => { /* si falla la imagen, el QR sigue en terminal */ });
    };
    handle.notifyFile = (filePath) => {
      try {
        const stat = fs.statSync(filePath);
        handle.fileInfo = { name: path.basename(filePath), size: stat.size, path: filePath };
        emit('file', { name: handle.fileInfo.name, size: handle.fileInfo.size });
      } catch {
        handle.pushLog('⚠ No se pudo leer el archivo generado: ' + filePath);
      }
    };

    /**
     * Espera a que el usuario descargue el archivo (o hasta timeout).
     * Devuelve 'downloaded' o 'timeout'. Mantiene el servidor vivo mientras tanto —
     * es lo que evita el clásico "comprobá tu conexión" por matar el proceso
     * antes de que el navegador descargue.
     */
    handle.waitForDownloadOrTimeout = (timeoutMs = 300_000) => new Promise(resolve => {
      if (handle._downloads > 0) { resolve('downloaded'); return; }
      const timer = setTimeout(() => {
        handle._dlWaiters = handle._dlWaiters.filter(w => w !== done);
        resolve('timeout');
      }, timeoutMs);
      const done = (r) => { clearTimeout(timer); resolve(r); };
      handle._dlWaiters.push(done);
    });

    // Lista de participantes del grupo (para mapear IDs → nombres en la página)
    handle.setParticipants = (list) => {
      handle._participants = list || [];
      emit('participants', handle._participants);
    };
    // Callback para persistir un nombre asignado (lo implementa index.js)
    handle.setNameSaver = (cb) => { handle._nameSaver = cb; };

    // Conexión: proveedor de estado / conector QR / callback post-config
    handle.setInstanceChecker = (fn) => { handle._instanceChecker = fn; };
    handle.setConnector = (fn) => { handle._connector = fn; };
    handle.setInstanceLister = (fn) => { handle._instanceLister = fn; };
    handle.setConfigApplied = (fn) => { handle._configApplied = fn; };

    handle.waitForRange = (timeoutMs = PICKER_TIMEOUT_MS) => new Promise((resolveRange, rejectRange) => {
      if (handle._range) { resolveRange(handle._range); return; }
      const timer = setTimeout(() => {
        rejectRange(new Error(
          `Tiempo de espera agotado (${Math.round(timeoutMs / 60_000)} min) sin recibir rango en el navegador.\n` +
          '   Ejecutá de nuevo, o usá el modo terminal:  node index.js --terminal'
        ));
      }, timeoutMs);
      handle._waiters.push((range) => { clearTimeout(timer); resolveRange(range); });
    });

    // Tee de la consola → página (solo una vez por proceso)
    let consoleAttached = false;
    handle.attachConsole = () => {
      if (consoleAttached) return;
      consoleAttached = true;
      const origLog = console.log.bind(console);
      const origError = console.error.bind(console);
      console.log = (...a) => { origLog(...a); try { emit('log', a.map(String).join(' ')); } catch {} };
      console.error = (...a) => { origError(...a); try { emit('log', a.map(String).join(' ')); } catch {} };
    };

    handle.openBrowser = () => openBrowser(handle.url);

    handle.close = () => new Promise(r => {
      for (const res of handle.clients) { try { res.end(); } catch {} }
      handle.clients.clear();
      if (handle.server) handle.server.close(() => r());
      else r();
    });

    const server = http.createServer((req, res) => {
      // Página
      if (req.method === 'GET' && req.url === '/') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(PAGE_HTML);
        return;
      }

      // Estado en vivo (SSE)
      if (req.method === 'GET' && req.url === '/events') {
        res.writeHead(200, {
          'Content-Type': 'text/event-stream; charset=utf-8',
          'Cache-Control': 'no-cache',
          Connection: 'keep-alive',
        });
        res.write(': conectado\n\n');
        // Replay del buffer: clientes que se conectan tarde no pierden nada
        for (const e of handle.events) {
          try { sseSend(res, e.event, e.data); } catch { break; }
        }
        handle.clients.add(res);
        const ka = setInterval(() => { try { res.write(': ping\n\n'); } catch {} }, SSE_KEEPALIVE_MS);
        req.on('close', () => { handle.clients.delete(res); clearInterval(ka); });
        return;
      }

      // Rango (autoritativo)
      if (req.method === 'POST' && req.url === '/submit') {
        let body = '';
        req.on('data', c => { body += c; });
        req.on('end', () => {
          let range;
          try { range = JSON.parse(body); }
          catch { sendJson(res, 400, { ok: false, error: 'JSON inválido.' }); return; }

          const startDate = new Date(range?.startDate);
          const endDate = new Date(range?.endDate);
          if (isNaN(startDate) || isNaN(endDate)) {
            sendJson(res, 400, { ok: false, error: 'Fechas inválidas.' }); return;
          }
          if (endDate <= startDate) {
            sendJson(res, 400, { ok: false, error: 'La fecha/hora FIN debe ser posterior al INICIO.' }); return;
          }

          sendJson(res, 200, { ok: true });
          handle._range = { startDate, endDate };
          const waiters = handle._waiters;
          handle._waiters = [];
          waiters.forEach(w => w(handle._range));
          handle.pushLog(`📅 Rango recibido: ${startDate.toLocaleString('es-CO')} → ${endDate.toLocaleString('es-CO')}`);
        });
        return;
      }

      // Lista de participantes (para mapear IDs → nombres)
      if (req.method === 'GET' && req.url === '/api/participants') {
        sendJson(res, 200, { ok: true, participants: handle._participants || [] });
        return;
      }

      // Configuración de conexión (GET: apiKey enmascarada; POST: guardar en vivo)
      if (req.method === 'GET' && req.url === '/api/config') {
        sendJson(res, 200, {
          ok: true,
          evolutionUrl: config.evolutionUrl,
          instance: config.instance,
          groupName: config.groupName,
          hasApiKey: !!config.apiKey,
        });
        return;
      }
      if (req.method === 'POST' && req.url === '/api/config') {
        let body = '';
        req.on('data', c => { body += c; });
        req.on('end', () => {
          let data;
          try { data = JSON.parse(body); } catch { sendJson(res, 400, { ok: false, error: 'JSON inválido.' }); return; }
          const prevInstance = config.instance;
          let next;
          try { next = updateConfig(data); }
          catch (err) { sendJson(res, 500, { ok: false, error: err.message }); return; }
          try { if (handle._configApplied) handle._configApplied({ instance: next.instance, changed: prevInstance !== next.instance }); } catch { /* el host decide */ }
          handle.pushLog(`⚙ Conexión guardada (instancia: ${next.instance})`);
          sendJson(res, 200, { ok: true, instance: next.instance });
        });
        return;
      }

      // Instancias existentes en Evolution (selector del panel)
      if (req.method === 'GET' && req.url === '/api/instances') {
        if (!handle._instanceLister) { sendJson(res, 200, { ok: true, instances: [] }); return; }
        Promise.resolve(handle._instanceLister())
          .then(list => sendJson(res, 200, { ok: true, instances: list || [] }))
          .catch(() => sendJson(res, 200, { ok: true, instances: [] }));
        return;
      }

      // Estado de la instancia (proveedor = index-evolution.js)
      if (req.method === 'GET' && req.url === '/api/instance') {
        if (!handle._instanceChecker) { sendJson(res, 200, { ok: false, error: 'sin proveedor de estado' }); return; }
        Promise.resolve(handle._instanceChecker())
          .then(s => sendJson(res, 200, { ok: true, ...s, instance: config.instance }))
          .catch(e => sendJson(res, 200, { ok: false, error: e.message }));
        return;
      }

      // Reconectar WhatsApp (QR) — proveedor = evolution-client.connectInstance
      if (req.method === 'POST' && req.url === '/api/connect') {
        if (!handle._connector) { sendJson(res, 500, { ok: false, error: 'La aplicación aún no conectó proveedor.' }); return; }
        Promise.resolve(handle._connector())
          .then(d => {
            if (d && d.error) { sendJson(res, 500, { ok: false, error: d.error }); return; }
            const q = d?.qrcode;
            const b64 = typeof q === 'string' ? q : q?.base64;
            if (!b64) { sendJson(res, 500, { ok: false, error: 'Evolution no devolvió QR' }); return; }
            const imgUrl = b64.startsWith('data:') ? b64 : `data:image/png;base64,${b64}`;
            try { handle.sendQR(imgUrl); } catch { /* la página puede no estar abierta */ }
            handle.pushLog('📲 QR de reconexión enviado al panel — escanealo con WhatsApp');
            sendJson(res, 200, { ok: true, qrcode: imgUrl });
          })
          .catch(e => sendJson(res, 500, { ok: false, error: e.message }));
        return;
      }

      // Guardar el nombre de un mensajero (persiste index.js en el JSON)
      if (req.method === 'POST' && req.url === '/api/names') {
        let body = '';
        req.on('data', c => { body += c; });
        req.on('end', () => {
          let data;
          try { data = JSON.parse(body); } catch { sendJson(res, 400, { ok: false, error: 'JSON inválido.' }); return; }
          const key = String(data?.key || '').trim();
          const name = String(data?.name || '').trim();
          if (!key) { sendJson(res, 400, { ok: false, error: 'Falta el identificador.' }); return; }
          if (!handle._nameSaver) { sendJson(res, 500, { ok: false, error: 'La aplicación aún no conectó.' }); return; }
          try { handle._nameSaver(key, name); }
          catch (err) { sendJson(res, 500, { ok: false, error: err.message }); return; }
          handle.pushLog(`👤 Mensajero ${key} → ${name || '(sin nombre)'}`);
          sendJson(res, 200, { ok: true });
        });
        return;
      }

      // Descarga del Word generado
      if (req.method === 'GET' && req.url === '/download') {
        if (!handle.fileInfo) {
          sendJson(res, 404, { ok: false, error: 'Todavía no se generó el archivo.' }); return;
        }
        fs.readFile(handle.fileInfo.path, (err, buf) => {
          if (err) { sendJson(res, 404, { ok: false, error: 'Archivo no disponible.' }); return; }
          res.writeHead(200, {
            'Content-Type': DOCX_MIME,
            'Content-Disposition': `attachment; filename="${handle.fileInfo.name}"`,
            'Content-Length': buf.length,
          });
          res.end(buf);
          handle._downloads += 1;
          // Despertar a quienes esperan la descarga
          const waiters = handle._dlWaiters;
          handle._dlWaiters = [];
          waiters.forEach(w => w('downloaded'));
        });
        return;
      }

      sendJson(res, 404, { ok: false, error: 'No encontrado.' });
    });

    server.on('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      handle.server = server;
      handle.port = port;
      handle.url = `http://127.0.0.1:${port}/`;
      server.unref(); // no mantener vivo el proceso al terminar
      resolve(handle);
    });
  });
}

function sendJson(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}

module.exports = { startControlServer, openBrowser };
