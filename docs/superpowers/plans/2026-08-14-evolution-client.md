# Cliente Evolution API — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reemplazar el scraping baileys por un cliente REST de Evolution API (ya hosteado en VPS) que consulta mensajes del grupo por rango de fechas y arma el Word de comprobantes.

**Architecture:** App de Node como cliente delgado. `config.js` centraliza URL/API key/instancia/CA. `evolution-client.js` consulta REST (findMessages paginado + filtro client-side + getBase64FromMediaMessage). `index-evolution.js` orquesta: picker web (reusado) → descubrir grupo → descargar media a `media_cache/` → `createWordDocument` (reusado de index.js). index.js (baileys) se conserva intacto como rollback.

**Tech Stack:** Node 24 (fetch nativo, sin dependencias nuevas), `node:https` (CA custom), `docx` (ya instalado), módulos existentes reutilizados vía `require('./index.js')`.

**Spec:** `docs/superpowers/specs/2026-08-14-evolution-client-design.md`

## Global Constraints

- Node >= 18 (fetch nativo). NO agregar dependencias nuevas.
- Toda escritura de caches/JSON con `atomicWriteFileSync` (regla anti-corrupción del proyecto). NUNCA `fs.writeFileSync` sobre caches.
- Guardia de fecha `isMsgInRange` SIEMPRE en el path de descarga (defensa en profundidad).
- TLS: fijar CA con `https.Agent({ ca })` vía `CA_FILE`. NUNCA `rejectUnauthorized: false`.
- Filtrado client-side SIEMPRE en findMessages (issue #1632: el filtro remoteJid del server puede ignorarse).
- Mensajes de consola/UI en español (idioma de la app).
- Carpeta del proyecto tiene ESPACIO FINAL: `cd "/Users/usuario/Documents/proyectos/mailboxes/extraertransferencias del día "`.
- Tests en `/tmp/test_*.js`, ejecutados DESDE la carpeta del proyecto (convención existente).
- Los tests E2E escriben archivos reales del proyecto (convención existente: `test_e2e_offline.js` genera el docx real); respaldan/restauran estado antes/después.
- `index.js` NO se modifica salvo el export de `toUnix` (Task 1). Sus funciones de caché/Word escriben a sus constantes internas (`./group_messages_cache.json`, `./failed_downloads.json`, `Comprobantes_Descargados.docx`) — el nuevo código debe convivir con eso, no redirigirlo.

---

### Task 1: Exportar toUnix + módulo de config

**Files:**
- Modify: `index.js` (module.exports — agregar `toUnix`)
- Create: `config.js`
- Create: `config.example.json`
- Modify: `.gitignore` (agregar `config.json` y `.env`)
- Test: `/tmp/test_evolution_config.js`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `require('./index.js').toUnix(ts)` → number (segundos Unix)
  - `require('./config.js').config` → `{ evolutionUrl, apiKey, instance, groupName, caFile }` (objeto congelado)
  - `require('./config.js').getAgent()` → `https.Agent | undefined` (con `ca` del PEM si `caFile` set)
  - Archivo `config.example.json` con keys `{ "evolutionUrl", "apiKey", "instance", "groupName", "caFile" }`

- [ ] **Step 1: Agregar `toUnix` a module.exports de index.js**

En `index.js`, dentro de `module.exports` (junto a `isMsgInRange`, línea ~1999):

```js
  isMsgInRange,
  toUnix,
```

- [ ] **Step 2: Escribir el test que falla (config)**

`/tmp/test_evolution_config.js`:

```js
// Tests de config.js — ejecutar DESDE la carpeta del proyecto:
//   node /tmp/test_evolution_config.js
const path = require('path');
const os = require('os');
const fs = require('fs');

let passed = 0;
const ok = (n) => { passed++; console.log(`  ✅ ${n}`); };
const bad = (n, e) => { console.log(`  ❌ ${n}: ${e?.message || e}`); process.exitCode = 1; };

// Aislar: config.json del proyecto no debe existir durante defaults → mover temporal
const proj = process.cwd();
const cfgPath = path.join(proj, 'config.json');
let moved = false;
if (fs.existsSync(cfgPath)) {
  fs.renameSync(cfgPath, cfgPath + '.bak-test');
  moved = true;
}
try {
  // Caso 1: defaults (sin config.json, sin env)
  delete require.cache[require.resolve(path.join(proj, 'config.js'))];
  const { config: def, getAgent } = require(path.join(proj, 'config.js'));
  if (def.evolutionUrl === 'https://transferencias.redpostal.co') ok('default evolutionUrl = VPS');
  else bad('default evolutionUrl', def.evolutionUrl);
  if (def.instance === 'comprobantes') ok('default instance = comprobantes');
  else bad('default instance', def.instance);
  if (def.apiKey === '') ok('apiKey default vacío');
  else bad('apiKey default', def.apiKey);
  if (getAgent() === undefined) ok('sin CA_FILE → getAgent() = undefined');
  else bad('getAgent sin CA_FILE', 'esperado undefined');

  // Caso 2: overrides por env
  process.env.EVOLUTION_URL = 'http://127.0.0.1:9999';
  process.env.EVOLUTION_API_KEY = 'secret-key';
  process.env.EVOLUTION_INSTANCE = 'otra';
  process.env.EVOLUTION_GROUP_NAME = 'Grupo X';
  delete require.cache[require.resolve(path.join(proj, 'config.js'))];
  const { config: env } = require(path.join(proj, 'config.js'));
  if (env.evolutionUrl === 'http://127.0.0.1:9999' && env.apiKey === 'secret-key' && env.instance === 'otra' && env.groupName === 'Grupo X')
    ok('env overrides aplicados');
  else bad('env overrides', JSON.stringify(env));

  // Caso 3: CA_FILE → getAgent() con ca cargado
  const pem = path.join(os.tmpdir(), 'test-ca-' + Date.now() + '.pem');
  fs.writeFileSync(pem, '-----BEGIN CERTIFICATE-----\nMIIB\n-----END CERTIFICATE-----\n');
  process.env.EVOLUTION_CA_FILE = pem;
  delete require.cache[require.resolve(path.join(proj, 'config.js'))];
  const { getAgent: getAgentCa } = require(path.join(proj, 'config.js'));
  const agent = getAgentCa();
  if (agent && agent.options.ca && agent.options.ca.length === 1 && agent.options.ca[0].includes('BEGIN CERTIFICATE'))
    ok('CA_FILE → agent con ca cargado');
  else bad('CA_FILE agent', agent ? JSON.stringify(agent.options) : 'undefined');
  fs.unlinkSync(pem);

  // Caso 4: config.example.json tiene las 5 keys
  const example = JSON.parse(fs.readFileSync(path.join(proj, 'config.example.json'), 'utf8'));
  const keys = ['evolutionUrl', 'apiKey', 'instance', 'groupName', 'caFile'];
  if (keys.every(k => k in example)) ok('config.example.json tiene las 5 keys');
  else bad('config.example.json keys', JSON.stringify(keys.filter(k => !(k in example))));

  console.log(`\n${passed} checks OK`);
} finally {
  if (moved) fs.renameSync(cfgPath + '.bak-test', cfgPath);
}
process.exit(process.exitCode || 0);
```

- [ ] **Step 3: Correr el test y verificar que falla**

Run: `node /tmp/test_evolution_config.js`
Expected: FAIL — `Cannot find module config.js` (o errores de export).

- [ ] **Step 4: Implementar config.js**

`config.js`:

```js
'use strict';
// Config del cliente Evolution API.
// Fuentes (en orden de prioridad): env vars > config.json > defaults.
// El archivo config.json NO se commitea (tiene la API key); ver config.example.json.

const fs = require('fs');
const path = require('path');
const https = require('https');

const DEFAULT_URL = 'https://transferencias.redpostal.co';

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return {};
  }
}

const fileCfg = readJson(path.join(__dirname, 'config.json'));

const config = Object.freeze({
  evolutionUrl: (process.env.EVOLUTION_URL || fileCfg.evolutionUrl || DEFAULT_URL).replace(/\/+$/, ''),
  apiKey: process.env.EVOLUTION_API_KEY || fileCfg.apiKey || '',
  instance: process.env.EVOLUTION_INSTANCE || fileCfg.instance || 'comprobantes',
  groupName: process.env.EVOLUTION_GROUP_NAME || fileCfg.groupName || 'transferencias',
  caFile: process.env.EVOLUTION_CA_FILE || fileCfg.caFile || '',
});

let cachedAgent = undefined;
/** https.Agent con la CA del VPS (si config.caFile está seteado). undefined = CAs del sistema. */
function getAgent() {
  if (config.caFile) {
    if (!cachedAgent) {
      cachedAgent = new https.Agent({ ca: fs.readFileSync(config.caFile) });
    }
    return cachedAgent;
  }
  return undefined;
}

module.exports = { config, getAgent };
```

- [ ] **Step 5: config.example.json**

```json
{
  "evolutionUrl": "https://transferencias.redpostal.co",
  "apiKey": "PONER_API_KEY_DEL_VPS",
  "instance": "comprobantes",
  "groupName": "transferencias",
  "caFile": ""
}
```

- [ ] **Step 6: .gitignore** — agregar si no está:

```gitignore
config.json
.env
```

- [ ] **Step 7: Correr el test y verificar que pasa**

Run: `node /tmp/test_evolution_config.js`
Expected: todos los checks OK.

- [ ] **Step 8: Commit**

```bash
git add index.js config.js config.example.json .gitignore /tmp/test_evolution_config.js
git commit -m "feat: export toUnix + módulo de config para cliente Evolution"
```

---

### Task 2: evolution-client.js — REST client

**Files:**
- Create: `evolution-client.js`
- Create: `/tmp/evolution_mock_server.js` (mock HTTP de Evolution para tests)
- Create: `/tmp/test_evolution_client.js`
- Test: `/tmp/test_evolution_client.js`

**Interfaces:**
- Consumes:
  - `require('./config').config` — `{ evolutionUrl, apiKey, instance }`
  - `require('./config').getAgent()`
  - `require('./index.js').toUnix(ts)` y `.isMsgInRange(msg, startTs, endTs)`
- Produces (todas usan `config.instance` internamente, sin parámetro de instancia):
  - `listGroups()` → `[{ id, subject }]` (fetchAllGroups; fallback findChats filtrando `@g.us`)
  - `findContactsMap()` → `{ '<jid>': { name: string } }` (jid = id completo devuelto por Evolution)
  - `findMediaMessages(jid, startTs, endTs)` → `[{ key, message, messageTimestamp: number }]` (filtrado client-side, ordenado por timestamp, sin duplicados por `key.id`)
  - `downloadMedia(msg)` → `{ buffer: Buffer, mimeType: string }`
  - `checkInstance()` → `{ found: boolean, status: string }` (`status` = `open`/`close`/`connecting`/`unknown`)
- Test helper exportado: `normalizeRecord(record)` → `{ key, message, messageTimestamp }`

- [ ] **Step 1: Escribir el test que falla**

`/tmp/evolution_mock_server.js`:

```js
// Mock HTTP server de Evolution v2 para tests.
// Uso:
//   const mock = await startMockEvolution(0, {...});
//   const url = mock.url; // http://127.0.0.1:PORT
//   mock.setRecords([...]); mock.setContacts([...]); mock.setGroups([...]);
//   mock.setInstances([...]); mock.setMedia('...base64...', 'image/jpeg');
//   mock.requests  // log de { method, path, body } recibidos
//   await mock.close();
const http = require('http');

function startMockEvolution(port, opts = {}) {
  const state = {
    records: opts.records || [],
    contacts: opts.contacts || [],
    groups: opts.groups || [],
    chats: opts.chats || [],
    instances: opts.instances || [],
    mediaBase64: opts.mediaBase64 || '',
    mediaMime: opts.mediaMime || 'image/jpeg',
    failPaths: new Set(opts.failPaths || []),
    requests: [],
  };
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', (c) => { body += c; });
    req.on('end', () => {
      state.requests.push({ method: req.method, path: req.url, body: body ? JSON.parse(body) : undefined });
      const send = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
      const p = req.url.split('?')[0];
      if ([...state.failPaths].some(fp => p.includes(fp))) return send(500, { error: { message: 'mock failure' } });
      if (req.method === 'POST' && p.includes('/chat/findMessages/')) {
        const skip = req.body?.skip || 0;
        return send(200, { messages: { total: state.records.length, pages: 1, currentPage: Math.floor(skip / 500) + 1, records: state.records.slice(skip, skip + 500) } });
      }
      if (req.method === 'POST' && p.includes('/chat/getBase64FromMediaMessage/')) {
        return send(200, { base64: state.mediaBase64, mimetype: state.mediaMime });
      }
      if (req.method === 'GET' && p.includes('/group/fetchAllGroups/')) return send(200, state.groups);
      if (req.method === 'GET' && p.includes('/chat/findChats/')) return send(200, state.chats);
      if (req.method === 'GET' && p.includes('/chat/findContacts/')) return send(200, state.contacts);
      if (req.method === 'GET' && p.includes('/instance/fetchInstances')) return send(200, state.instances);
      return send(404, { error: { message: 'not mocked: ' + p } });
    });
  });
  return new Promise((resolve) => {
    server.listen(port, '127.0.0.1', () => {
      const url = `http://127.0.0.1:${server.address().port}`;
      resolve({
        url,
        requests: state.requests,
        setRecords: (r) => { state.records = r; },
        setContacts: (c) => { state.contacts = c; },
        setGroups: (g) => { state.groups = g; },
        setChats: (c) => { state.chats = c; },
        setInstances: (i) => { state.instances = i; },
        setMedia: (b64, mime) => { state.mediaBase64 = b64; state.mediaMime = mime; },
        failPaths: state.failPaths,
        close: () => new Promise((r) => server.close(r)),
      });
    });
  });
}

module.exports = { startMockEvolution, JPEG_1PX: '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVN//2Q==' };
```

`/tmp/test_evolution_client.js`:

```js
// Tests de evolution-client.js contra un mock HTTP de Evolution.
// Ejecutar DESDE la carpeta del proyecto:
//   node /tmp/test_evolution_client.js
const path = require('path');
const proj = process.cwd();
const { startMockEvolution, JPEG_1PX } = require('/tmp/evolution_mock_server.js');

let passed = 0;
const ok = (n) => { passed++; console.log(`  ✅ ${n}`); };
const bad = (n, e) => { console.log(`  ❌ ${n}: ${e?.message || e}`); process.exitCode = 1; };

const mkImage = (id, ts, remoteJid, participant, pushName) => ({
  key: { remoteJid, id, fromMe: false, participant },
  message: { imageMessage: { url: 'x', mimetype: 'image/jpeg' } },
  messageType: 'imageMessage',
  fromMe: false,
  pushName,
  timestamp: ts,
});

(async () => {
  const mock = await startMockEvolution(0, {});
  try {
    process.env.EVOLUTION_URL = mock.url;
    process.env.EVOLUTION_API_KEY = 'test-key';
    delete require.cache[require.resolve(path.join(proj, 'config.js'))];
    delete require.cache[require.resolve(path.join(proj, 'evolution-client.js'))];
    const evo = require(path.join(proj, 'evolution-client.js'));

    const DAY = 86400;
    const START = Math.floor(Date.UTC(2026, 7, 1) / 1000); // 2026-08-01 UTC
    const END = START + 10 * DAY; // hasta 11 ago
    const JID = '120363123456789@g.us';

    // ── findMediaMessages: filtros + normalización ──
    mock.setRecords([
      mkImage('A', START + 100, JID, '573001112223@s.whatsapp.net', 'VEHICULO 1'),
      mkImage('B', START + 200, JID, undefined, undefined), // sin participant (sync viejo)
      mkImage('C', END + 100, JID, '573001112223@s.whatsapp.net'), // FUERA de rango
      mkImage('D', START + 300, '999@g.us', '573001112223@s.whatsapp.net'), // OTRO grupo (el filtro remoteJid del server se ignora — bug #1632)
      { key: { remoteJid: JID, id: 'E', fromMe: false }, message: { conversation: 'hola' }, messageType: 'conversation', timestamp: START + 400 }, // no media
      mkImage('F', String(START + 500), JID, '573001112223@s.whatsapp.net'), // timestamp string
      { ...mkImage('G', { low: START + 600, high: 0 }, JID, '573001112223@s.whatsapp.net') }, // timestamp Long
      mkImage('G', START + 700, JID, '573001112223@s.whatsapp.net'), // duplicado id G
      mkImage('H', '2026-08-03T10:00:00.000Z', JID, '573001112223@s.whatsapp.net'), // ISO string
    ]);
    const msgs = await evo.findMediaMessages(JID, START, END);
    const ids = msgs.map(m => m.key.id);
    if (msgs.length === 5 && ids.join(',') === 'A,B,F,G,H') ok('findMediaMessages: 5 en rango, dedupe G, ordenado asc');
    else bad('findMediaMessages', JSON.stringify(ids));
    if (msgs.every(m => typeof m.messageTimestamp === 'number' && m.messageTimestamp > 0))
      ok('normalización: messageTimestamp numérico (string/Long/ISO → number)');
    else bad('normalización timestamps', JSON.stringify(msgs.map(m => m.messageTimestamp)));
    const isoMsg = msgs.find(m => m.key.id === 'H');
    if (isoMsg && isoMsg.messageTimestamp === Date.UTC(2026, 7, 3, 10, 0, 0) / 1000) ok('ISO string convertido a segundos');
    else bad('ISO string', isoMsg && isoMsg.messageTimestamp);

    // ── checkInstance ──
    mock.setInstances([{ instanceName: 'comprobantes', connectionStatus: 'open' }]);
    const chk = await evo.checkInstance();
    if (chk.found && chk.status === 'open') ok('checkInstance: encontrada y open');
    else bad('checkInstance', JSON.stringify(chk));
    mock.setInstances([]);
    const chk2 = await evo.checkInstance();
    if (!chk2.found) ok('checkInstance: no encontrada');
    else bad('checkInstance not found', JSON.stringify(chk2));

    // ── findContactsMap ──
    mock.setContacts([
      { id: '573001112223@s.whatsapp.net', name: 'VEHICULO 1' },
      { id: '573009998887@s.whatsapp.net', pushName: 'Mensajero 2' },
    ]);
    const cmap = await evo.findContactsMap();
    if (cmap['573001112223@s.whatsapp.net']?.name === 'VEHICULO 1' && cmap['573009998887@s.whatsapp.net']?.name === 'Mensajero 2')
      ok('findContactsMap: keyed por jid completo, name/pushName');
    else bad('findContactsMap', JSON.stringify(cmap));

    // ── listGroups (fetchAllGroups) ──
    mock.setGroups([{ id: '120363123456789@g.us', subject: 'Transferencias Diwifarma' }]);
    const groups = await evo.listGroups();
    if (groups.length === 1 && groups[0].id === '120363123456789@g.us') ok('listGroups: fetchAllGroups');
    else bad('listGroups', JSON.stringify(groups));

    // ── listGroups (fallback findChats si fetchAllGroups falla) ──
    mock.failPaths.add('/group/fetchAllGroups/');
    mock.setChats([{ id: '120363123456789@g.us', name: 'Transferencias Diwifarma' }, { id: '573001112223@s.whatsapp.net', name: 'Mensajero' }]);
    const groups2 = await evo.listGroups();
    if (groups2.length === 1 && groups2[0].id === '120363123456789@g.us') ok('listGroups: fallback findChats filtra @g.us');
    else bad('listGroups fallback', JSON.stringify(groups2));
    mock.failPaths.delete('/group/fetchAllGroups/');

    // ── downloadMedia ──
    mock.setMedia(JPEG_1PX, 'image/jpeg');
    const dl = await evo.downloadMedia({ key: { remoteJid: JID, id: 'A' }, message: { imageMessage: {} } });
    if (dl.buffer instanceof Buffer && dl.buffer.length > 0 && dl.mimeType === 'image/jpeg')
      ok('downloadMedia: Buffer desde base64');
    else bad('downloadMedia', `len=${dl.buffer?.length} mime=${dl.mimeType}`);
    const lastReq = mock.requests[mock.requests.length - 1];
    if (lastReq && lastReq.body?.key?.id === 'A' && lastReq.body?.message) ok('downloadMedia envía el mensaje COMPLETO (no solo el id)');
    else bad('downloadMedia body', JSON.stringify(lastReq?.body));

    // ── error handling: 500 → throw con mensaje ──
    mock.failPaths.add('/chat/findMessages/');
    let threw = false;
    try { await evo.findMediaMessages(JID, START, END); } catch (e) { threw = e.message.includes('mock failure') || e.message.length > 0; }
    if (threw) ok('fallo del server → throw con mensaje');
    else bad('error handling', 'no throw');

    console.log(`\n${passed} checks OK`);
  } finally {
    await mock.close();
  }
  process.exit(process.exitCode || 0);
})();
```

- [ ] **Step 2: Correr el test y verificar que falla**

Run: `node /tmp/test_evolution_client.js`
Expected: FAIL — module not found / funciones undefined.

- [ ] **Step 3: Implementar evolution-client.js**

`evolution-client.js`:

```js
'use strict';
// Cliente REST de Evolution API v2 (hosteado en VPS).
// Consulta mensajes/contactos/grupos y descarga media. Filtrado client-side
// SIEMPRE: el filtro remoteJid del server puede no aplicarse (issue #1632).

const { config, getAgent } = require('./config');
const { toUnix, isMsgInRange } = require('./index.js');

const TAKE = 500;
const MAX_PAGES = 40;

/** Convierte el timestamp de un record (number | string | Long | ISO) a segundos Unix. */
function toSeconds(ts) {
  if (ts == null) return 0;
  if (typeof ts === 'number') return ts > 1e12 ? Math.floor(ts / 1000) : ts;
  if (typeof ts === 'string') {
    if (/^\d+$/.test(ts)) return toSeconds(Number(ts));
    const d = Date.parse(ts);
    return isNaN(d) ? 0 : Math.floor(d / 1000);
  }
  if (typeof ts === 'object' && ts.low != null) return Number(ts.low);
  return 0;
}

/** Normaliza un record de Evolution a la forma que usa el pipeline (como baileys). */
function normalizeRecord(record) {
  return {
    key: record.key,
    message: record.message,
    messageTimestamp: toSeconds(record.timestamp ?? record.messageTimestamp),
  };
}

/** fetch con apikey + agent (CA custom) + timeout. Lanza Error con detalle del server. */
async function apiRequest(method, path, body) {
  const url = `${config.evolutionUrl}${path}`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30_000);
  try {
    const res = await fetch(url, {
      method,
      headers: {
        'Content-Type': 'application/json',
        apikey: config.apiKey,
      },
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: controller.signal,
      agent: getAgent(),
    });
    let data = null;
    try { data = await res.json(); } catch { /* body no JSON */ }
    if (!res.ok) {
      const msg = data?.error?.message || `HTTP ${res.status}`;
      throw new Error(`Evolution ${path}: ${msg}`);
    }
    return data;
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`Evolution ${path}: timeout (30s)`);
    if (err.message.startsWith('Evolution')) throw err;
    throw new Error(`Evolution ${path}: ${err.message}`);
  } finally {
    clearTimeout(timer);
  }
}

/** Grupos disponibles: [{id, subject}]. fetchAllGroups; fallback findChats
 *  filtrando @g.us (endpoint a confirmar contra la instalación real). */
async function listGroups() {
  let data = null;
  try {
    data = await apiRequest('GET', `/group/fetchAllGroups/${config.instance}`);
  } catch {
    data = await apiRequest('GET', `/chat/findChats/${config.instance}`);
  }
  const list = Array.isArray(data) ? data : data?.chats || [];
  return list
    .filter(g => (g.id || g.jid || '').endsWith('@g.us'))
    .map(g => ({ id: g.id || g.jid, subject: g.subject || g.name || '' }));
}

/** Contactos: { '<jid>': { name } } — jid completo como lo devuelve Evolution. */
async function findContactsMap() {
  const data = await apiRequest('GET', `/chat/findContacts/${config.instance}`);
  const map = {};
  if (Array.isArray(data)) {
    for (const c of data) {
      if (c?.id) map[c.id] = { name: c.name || c.pushName || c.notify || '' };
    }
  }
  return map;
}

/** Estado de la instancia: { found, status } (status: open/close/connecting/unknown). */
async function checkInstance() {
  const data = await apiRequest('GET', `/instance/fetchInstances`);
  const list = Array.isArray(data) ? data : [];
  const inst = list.find(i => (i.instanceName || i.name) === config.instance);
  if (!inst) return { found: false, status: 'unknown' };
  return { found: true, status: inst.connectionStatus || inst.status || 'unknown' };
}

const isMedia = (msg) => !!msg.message?.imageMessage || !!msg.message?.documentMessage;

/** Mensajes media del grupo en [startTs, endTs], paginados, filtrados client-side.
 *  Escaneo completo hasta MAX_PAGES (el server puede no respetar el orden pedido). */
async function findMediaMessages(jid, startTs, endTs) {
  const out = new Map();
  for (let page = 0; page < MAX_PAGES; page++) {
    const body = {
      where: { key: { remoteJid: jid } },
      take: TAKE,
      skip: page * TAKE,
      orderBy: { timestamp: 'asc' },
    };
    const data = await apiRequest('POST', `/chat/findMessages/${config.instance}`, body);
    const records = data?.messages?.records || [];
    if (records.length === 0) break;
    for (const r of records) {
      const msg = normalizeRecord(r);
      if (msg.key?.remoteJid !== jid) continue;      // filtro server puede no aplicar (#1632)
      if (!isMsgInRange(msg, startTs, endTs)) continue;
      if (!isMedia(msg)) continue;
      if (!out.has(msg.key.id)) out.set(msg.key.id, msg);
    }
  }
  return [...out.values()].sort((a, b) => a.messageTimestamp - b.messageTimestamp);
}

/** Descarga el media de un mensaje. Envía el mensaje COMPLETO (el server
 *  consulta la DB por el objeto; solo el id falla en media viejo). */
async function downloadMedia(msg) {
  const body = { key: msg.key, message: msg.message };
  const data = await apiRequest('POST', `/chat/getBase64FromMediaMessage/${config.instance}`, body);
  if (!data?.base64) throw new Error('Evolution no devolvió media para el mensaje');
  return { buffer: Buffer.from(data.base64, 'base64'), mimeType: data.mimetype || 'image/jpeg' };
}

module.exports = {
  listGroups,
  findContactsMap,
  findMediaMessages,
  downloadMedia,
  checkInstance,
  normalizeRecord,
  toSeconds,
  apiRequest,
};
```

- [ ] **Step 4: Correr el test y verificar que pasa**

Run: `node /tmp/test_evolution_client.js`
Expected: todos OK. Ajustar el mock si el path exacto del server difiere (los tests son la fuente de verdad del contrato HTTP).

- [ ] **Step 5: Commit**

```bash
git add evolution-client.js /tmp/test_evolution_client.js /tmp/evolution_mock_server.js
git commit -m "feat: cliente REST Evolution API con filtrado client-side y descarga de media"
```

---

### Task 3: index-evolution.js — orquestación + E2E

**Files:**
- Create: `index-evolution.js`
- Create: `/tmp/test_evolution_e2e.js`
- Test: `/tmp/test_evolution_e2e.js`

**Interfaces:**
- Consumes:
  - `require('./config').config`, `.getAgent()` (via evolution-client)
  - `require('./evolution-client')` — `listGroups, findContactsMap, findMediaMessages, downloadMedia, checkInstance`
  - `require('./index.js')` — `createWordDocument, getSenderName, isMsgInRange, saveCachedGroupMessages, loadFailedDownloads, saveFailedDownload, isGroupNameLike` (todas escriben a sus constantes internas: msg cache, failed cache y docx van a la carpeta del proyecto SIEMPRE — el E2E respalda/restaura esos archivos)
  - `require('./picker').startControlServer` (entry web)
- Produces:
  - `runPipeline(opts)` — `opts: { startTs, endTs }`. Retorna `{ total, downloaded, skipped, failed, outOfRange, wordPath }`.
  - `main()` — picker web o prompts `--terminal`; guardado con `require.main === module`.

- [ ] **Step 1: Escribir el test E2E que falla**

`/tmp/test_evolution_e2e.js`:

```js
// E2E del cliente Evolution: mock HTTP → runPipeline → Word real.
// Respaldar/restaurar archivos de estado del proyecto (convención existente:
// los E2E escriben archivos reales y los restauran).
// Ejecutar DESDE la carpeta del proyecto:
//   node /tmp/test_evolution_e2e.js
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const proj = process.cwd();
const { startMockEvolution, JPEG_1PX } = require('/tmp/evolution_mock_server.js');

let passed = 0;
const ok = (n) => { passed++; console.log(`  ✅ ${n}`); };
const bad = (n, e) => { console.log(`  ❌ ${n}: ${e?.message || e}`); process.exitCode = 1; };

const DAY = 86400;
const START = Math.floor(Date.UTC(2026, 7, 1) / 1000);
const END = Math.floor(Date.UTC(2026, 7, 11, 23, 59, 59) / 1000);
const JID = '120363123456789@g.us';
const PARTICIPANT = '573001112223@s.whatsapp.net';

const mkImage = (id, ts) => ({
  key: { remoteJid: JID, id, fromMe: false, participant: PARTICIPANT },
  message: { imageMessage: { url: 'x', mimetype: 'image/jpeg' } },
  messageType: 'imageMessage',
  fromMe: false,
  pushName: 'VEHICULO 1',
  timestamp: ts,
});

// ── Backup/restore de archivos de estado del proyecto ──
const STATE_FILES = ['group_cache.json', 'group_messages_cache.json', 'failed_downloads.json', 'Comprobantes_Descargados.docx'];
const stateBak = new Map();
function backupState() {
  for (const f of STATE_FILES) {
    const src = path.join(proj, f);
    if (fs.existsSync(src)) {
      const dst = src + '.bak-e2e';
      fs.renameSync(src, dst);
      stateBak.set(src, dst);
    }
  }
  const mediaSrc = path.join(proj, 'media_cache');
  if (fs.existsSync(mediaSrc)) {
    const dst = mediaSrc + '.bak-e2e';
    fs.renameSync(mediaSrc, dst);
    stateBak.set(mediaSrc, dst);
  }
}
function restoreState() {
  for (const [src, dst] of stateBak) fs.renameSync(dst, src);
  stateBak.clear();
}

(async () => {
  backupState();
  const mock = await startMockEvolution(0, {
    instances: [{ instanceName: 'comprobantes', connectionStatus: 'open' }],
    groups: [{ id: JID, subject: 'Transferencias Diwifarma' }],
    contacts: [{ id: PARTICIPANT, name: 'VEHICULO 1' }],
    records: [
      mkImage('E2E-1', START + 100),
      mkImage('E2E-2', START + 200),
      mkImage('E2E-OUT', END + 5000), // FUERA del rango: no debe descargarse
    ],
    mediaBase64: JPEG_1PX,
  });
  try {
    process.env.EVOLUTION_URL = mock.url;
    process.env.EVOLUTION_API_KEY = 'test-key';
    process.env.EVOLUTION_INSTANCE = 'comprobantes';
    process.env.EVOLUTION_GROUP_NAME = 'transferencias';
    delete require.cache[require.resolve(path.join(proj, 'config.js'))];
    delete require.cache[require.resolve(path.join(proj, 'index-evolution.js'))];
    const app = require(path.join(proj, 'index-evolution.js'));

    // ── Primera corrida ──
    const res = await app.runPipeline({ startTs: START, endTs: END });
    if (res.total === 2 && res.downloaded === 2 && res.failed === 0 && res.outOfRange === 0)
      ok(`resumen correcto: total=${res.total} downloaded=${res.downloaded} outOfRange=${res.outOfRange}`);
    else bad('resumen', JSON.stringify(res));

    // Guardia de fecha: la imagen fuera de rango NO se descargó
    // (la filtra findMediaMessages — el contador outOfRange es defensa en profundidad)
    if (!fs.existsSync(path.join(proj, 'media_cache', 'E2E-OUT.jpg'))) ok('guardia de fecha: imagen fuera de rango NO descargada');
    else bad('guardia de fecha', 'E2E-OUT.jpg existe en media_cache');

    // Media descargadas a media_cache con el id del mensaje
    if (fs.existsSync(path.join(proj, 'media_cache', 'E2E-1.jpg')) && fs.existsSync(path.join(proj, 'media_cache', 'E2E-2.jpg')))
      ok('media_cache: 2 imágenes guardadas');
    else bad('media_cache', fs.readdirSync(path.join(proj, 'media_cache')).join(','));

    // Grupo descubierto y cacheado
    const groupCache = JSON.parse(fs.readFileSync(path.join(proj, 'group_cache.json'), 'utf8'));
    if (groupCache.groupJid === JID) ok('grupo descubierto y guardado en group_cache.json');
    else bad('group_cache.json', JSON.stringify(groupCache));

    // Caché de mensajes mergeado en disco (formato heredado)
    const cached = JSON.parse(fs.readFileSync(path.join(proj, 'group_messages_cache.json'), 'utf8'));
    if (Array.isArray(cached) && cached.some(m => m.key?.id === 'E2E-1')) ok('msg cache: mensajes guardados en formato heredado');
    else bad('msg cache', JSON.stringify((cached || []).slice(0, 1)));

    // ── Segunda corrida: skip por media_cache, sin re-consultar grupos ──
    const res2 = await app.runPipeline({ startTs: START, endTs: END });
    if (res2.downloaded === 0 && res2.skipped === 2) ok(`segunda corrida: 0 descargas, 2 skips (media_cache)`);
    else bad('segunda corrida', JSON.stringify(res2));
    const fetchCalls = mock.requests.filter(r => r.path.includes('fetchAllGroups')).length;
    if (fetchCalls === 1) ok('grupo cacheado: fetchAllGroups llamado UNA vez');
    else bad('fetchAllGroups calls', fetchCalls);

    // Word válido con 2 imágenes embebidas
    const listing = execFileSync('unzip', ['-l', path.join(proj, 'Comprobantes_Descargados.docx')]).toString();
    const imgCount = (listing.match(/word\/media\/[^\s]+/g) || []).length;
    if (imgCount === 2) ok('Word: 2 imágenes embebidas');
    else bad('Word embebidas', imgCount);

    console.log(`\n${passed} checks OK`);
  } finally {
    await mock.close();
    restoreState();
  }
  process.exit(process.exitCode || 0);
})();
```

- [ ] **Step 2: Correr el test y verificar que falla**

Run: `node /tmp/test_evolution_e2e.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Implementar index-evolution.js**

`index-evolution.js` — flujo completo. Puntos clave: `runPipeline({startTs, endTs})` sin overrides de archivos (las funciones heredadas escriben a sus constantes); grupo descubierto una vez (cache); contactos una vez por corrida; descarga con pool de 2 workers + reintentos x3; guardia `isMsgInRange` antes de descargar (defensa en profundidad); fallos permanentes → failed cache; `saveCachedGroupMessages` al final; `createWordDocument` → output. El main() borra el Word viejo al arrancar (anti-confusión del día anterior, mismo patrón que index.js:1396).

```js
'use strict';
// ════════════════════════════════════════════════════════════════════════
//   EXTRACTOR DE COMPROBANTES (cliente Evolution API) — WhatsApp → Word
//   Reemplaza el scraping baileys: consulta REST a Evolution (VPS), que
//   mantiene la sesión 24/7. Sin QR para el usuario final, sin sync frágil.
//   USO:  node index-evolution.js   (panel web)   |   --terminal (prompts)
// ════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const { config } = require('./config');
const evo = require('./evolution-client');
const M = require('./index.js'); // reutiliza Word, nombres, guardias, caches
const { atomicWriteFileSync } = require('./fs_utils');
const { startControlServer } = require('./picker');

// Archivos de estado (mismos formatos que la app baileys; las funciones
// heredadas de index.js escriben a sus constantes internas)
const OUTPUT_FILE = 'Comprobantes_Descargados.docx';
const MEDIA_CACHE_DIR = './media_cache';
const CACHE_FILE = './group_cache.json';

const CONCURRENCY = 2;
const RETRIES = 3;

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/** Pool de descarga con concurrencia fija (mismo patrón que la fase 1 heredada). */
async function downloadQueue(items, worker, concurrency) {
  let i = 0;
  const results = new Array(items.length);
  const runner = async () => {
    while (i < items.length) {
      const idx = i++;
      try { results[idx] = await worker(items[idx]); }
      catch (err) { results[idx] = { error: err }; }
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, runner));
  return results;
}

/**
 * Pipeline principal (exportado para tests; main() solo agrega el picker).
 * opts: { startTs, endTs } — segundos Unix, inclusivo.
 * Retorna { total, downloaded, skipped, failed, outOfRange, wordPath }.
 */
async function runPipeline({ startTs, endTs } = {}) {
  // 1. Estado de la instancia
  let inst;
  try {
    inst = await evo.checkInstance();
  } catch (err) {
    console.log(`\n❌ No se pudo contactar Evolution (${config.evolutionUrl}).`);
    console.log(`   ${err.message}`);
    console.log('   → Revisá: URL y API key en config.json, y que el VPS esté arriba.');
    throw err;
  }
  if (!inst.found) {
    throw new Error(`Instancia "${config.instance}" no existe en Evolution. Revisá config.json.`);
  }
  if (inst.status !== 'open') {
    console.log(`⚠ La instancia "${config.instance}" está ${inst.status} — el teléfono de la SIM debe estar conectado a WhatsApp.`);
    console.log('   Verificá la instancia en la UI de Evolution (VPS) y seguí con el caché local.');
  }

  // 2. Descubrir el grupo (cacheado una vez)
  let groupJid = null;
  let groupName = '';
  try {
    const saved = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
    groupJid = saved.groupJid || null;
    groupName = saved.groupName || '';
  } catch { /* primera vez */ }
  if (!groupJid) {
    const groups = await evo.listGroups();
    const match = groups.find(g => M.isGroupNameLike(g.subject));
    if (!match) {
      throw new Error(`No se encontró un grupo con "${config.groupName}" en "${config.instance}".`);
    }
    groupJid = match.id;
    groupName = match.subject;
    atomicWriteFileSync(CACHE_FILE, JSON.stringify({ groupJid, groupName }));
    console.log(`👥 Grupo: ${groupName} (${groupJid})`);
  } else {
    console.log(`👥 Grupo (cache): ${groupName || groupJid}`);
  }

  // 3. Contactos para nombres
  let contacts = {};
  try { contacts = await evo.findContactsMap(); }
  catch (err) { console.log(`⚠ Sin contactos de Evolution: ${err.message}`); }

  // 4. Mensajes media en rango
  const msgs = await evo.findMediaMessages(groupJid, startTs, endTs);
  console.log(`📊 Comprobantes en el rango: ${msgs.length}`);
  if (msgs.length === 0) {
    console.log('   (Si esperabas resultados: revisá que el grupo esté enviando fotos y que la instancia esté conectada.)');
  }

  // 5. Descargar media (pool 2, reintentos x3, guardia de fecha)
  const failedDownloads = M.loadFailedDownloads();
  const stats = { downloaded: 0, skipped: 0, failed: 0, outOfRange: 0 };
  const downloadOne = async (msg) => {
    const id = msg.key?.id;
    // Defensa en profundidad: la guardia de fecha sigue en el path de descarga
    if (!M.isMsgInRange(msg, startTs, endTs)) { stats.outOfRange++; return; }
    if (failedDownloads.has(id)) { stats.skipped++; return; }
    const mediaPath = `${MEDIA_CACHE_DIR}/${id}.jpg`;
    if (fs.existsSync(mediaPath)) { stats.skipped++; return; }
    let lastErr;
    for (let attempt = 1; attempt <= RETRIES; attempt++) {
      try {
        const { buffer } = await evo.downloadMedia(msg);
        fs.mkdirSync(MEDIA_CACHE_DIR, { recursive: true });
        atomicWriteFileSync(mediaPath, buffer); // mismo patrón anti-corrupción que saveCachedMedia (index.js:836)
        stats.downloaded++;
        return;
      } catch (err) {
        lastErr = err;
        if (attempt < RETRIES) await sleep(1000 * attempt);
      }
    }
    M.saveFailedDownload(id, 'evolution');
    console.log(`   ⚠ Falló tras ${RETRIES} intentos: ${id} — ${lastErr.message}`);
    stats.failed++;
  };
  await downloadQueue(msgs, downloadOne, CONCURRENCY);

  // 6. Persistir mensajes al caché heredado (merge atómico; formato: Map [id, msg])
  const store = new Map();
  for (const m of msgs) store.set(m.key.id, m);
  M.saveCachedGroupMessages(groupJid, store);

  // 7. Armar receipts y generar el Word
  const receipts = [];
  for (const msg of msgs) {
    const id = msg.key?.id;
    if (M.loadFailedDownloads().has(id)) continue;
    const mediaPath = `${MEDIA_CACHE_DIR}/${id}.jpg`;
    if (!fs.existsSync(mediaPath)) continue;
    receipts.push({
      imageBuffer: fs.readFileSync(mediaPath),
      senderName: M.getSenderName(msg, contacts),
      date: new Date(msg.messageTimestamp * 1000).toLocaleString('es-CO'),
    });
  }
  if (receipts.length === 0) {
    console.log('\n⚠ No se generó el Word: no hay imágenes disponibles en el rango.');
    return { total: msgs.length, downloaded: stats.downloaded, skipped: stats.skipped, failed: stats.failed, outOfRange: stats.outOfRange, wordPath: null };
  }
  await M.createWordDocument(receipts);
  console.log(`\n✅ Word generado: ${OUTPUT_FILE} (${receipts.length} comprobantes)`);
  if (stats.failed > 0) console.log(`   ⚠ ${stats.failed} imágenes no se pudieron descargar (reintentará en otra corrida si siguen disponibles).`);
  if (stats.outOfRange > 0) console.log(`   🛡 ${stats.outOfRange} fuera del rango seleccionado (bloqueadas por la guardia de fecha).`);

  return { total: msgs.length, downloaded: stats.downloaded, skipped: stats.skipped, failed: stats.failed, outOfRange: stats.outOfRange, wordPath: OUTPUT_FILE };
}

/** Pregunta rango por terminal (modo --terminal, sin navegador). */
async function askDateRangeTerminal() {
  const { prompt } = require('enquirer');
  const a = await prompt({ type: 'input', name: 'start', message: 'Fecha inicio (YYYY-MM-DD):', validate: v => !isNaN(Date.parse(v)) });
  const b = await prompt({ type: 'input', name: 'end', message: 'Fecha fin (YYYY-MM-DD):', validate: v => !isNaN(Date.parse(v)) });
  return { startDate: new Date(a.start), endDate: new Date(b.end) };
}

async function main() {
  const useWebPicker = !process.argv.includes('--terminal');
  let startDate, endDate;
  if (useWebPicker) {
    const activePicker = await startControlServer();
    activePicker.attachConsole();
    activePicker.openBrowser();
    console.log(`📡 Panel abierto en el navegador (${activePicker.url})`);
    ({ startDate, endDate } = await activePicker.waitForRange());
  } else {
    ({ startDate, endDate } = await askDateRangeTerminal());
  }
  const startTs = Math.floor(startDate.getTime() / 1000);
  const endTs = Math.floor(endDate.getTime() / 1000);

  // Borrar el Word de corridas anteriores (mismo patrón que index.js: el archivo
  // viejo del día pasado NO puede confundirse con un resultado nuevo)
  try { fs.rmSync(OUTPUT_FILE, { force: true }); } catch { /* si no existe, ok */ }

  await runPipeline({ startTs, endTs });
  console.log('\nListo. Podés cerrar la ventana.');
  if (useWebPicker) process.exit(0);
}

if (require.main === module) {
  main().catch((err) => {
    console.error(`\n❌ Error fatal: ${err.message}`);
    process.exit(1);
  });
}

module.exports = { runPipeline, main };
```

- [ ] **Step 4: Correr el test y verificar que pasa**

Run: `node /tmp/test_evolution_e2e.js`
Expected: todos OK (incluida la guardia de fecha y el skip por media_cache).

- [ ] **Step 5: Verificar que las suites heredadas siguen verdes**

Run:
```bash
for t in date_guard fixes picker sender_name e2e_offline recovery; do node /tmp/test_$t.js >/dev/null 2>&1 && echo "$t OK" || echo "$t FAIL"; done
```
Expected: las 6 en OK (index.js sin cambios → sin regresión).

- [ ] **Step 6: Commit**

```bash
git add index-evolution.js /tmp/test_evolution_e2e.js
git commit -m "feat: pipeline Evolution API → Word (picker reusado, guardia de fecha, media_cache)"
```

---

### Task 4: package.json, README y cierre

**Files:**
- Modify: `package.json` (scripts)
- Create: `README.md` (si no existe; si existe, agregar sección)
- Test: corrida completa de las 9 suites

- [ ] **Step 1: package.json**

```json
{
  "scripts": {
    "start": "node index-evolution.js",
    "start:legacy": "node index.js",
    "start:terminal": "node index-evolution.js --terminal"
  }
}
```
(mantener `main` y `dependencies` intactos)

- [ ] **Step 2: README — sección "Modo Evolution API (recomendado)"**

Contenido mínimo (español):
- Qué es: cliente REST contra Evolution API hosteada en VPS (`https://transferencias.redpostal.co`), sesión 24/7, sin QR para el usuario final.
- Setup: copiar `config.example.json` → `config.json`; poner `apiKey` (VPS), `instance`, `groupName` (filtro de grupo), `caFile`.
- TLS: si el VPS sirve el cert autofirmado de Traefik, extraer el PEM:
  ```bash
  echo | openssl s_client -connect transferencias.redpostal.co:443 -servername transferencias.redpostal.co 2>/dev/null | openssl x509 -out traefik.crt
  ```
  y setear `caFile: "traefik.crt"`. Fix definitivo: Let's Encrypt en Traefik (certresolver) — luego `caFile: ""`.
- Uso: `npm start` (panel web) o `npm run start:terminal`.
- Windows: mismo `config.json` con la URL del VPS.
- Operación: el teléfono de la SIM debe estar enchufado con WhatsApp conectado; si la instancia se desconecta, escanear QR una vez en la UI de Evolution del VPS.
- Rollback: `npm run start:legacy` (aplicación baileys anterior, sin cambios).

- [ ] **Step 3: Correr las 9 suites completas**

Run:
```bash
cd "/Users/usuario/Documents/proyectos/mailboxes/extraertransferencias del día " && \
node /tmp/test_evolution_config.js 2>&1 | tail -1 && \
node /tmp/test_evolution_client.js 2>&1 | tail -1 && \
node /tmp/test_evolution_e2e.js 2>&1 | tail -1 && \
for t in date_guard fixes picker sender_name e2e_offline recovery; do node /tmp/test_$t.js 2>&1 | tail -1; done
```
Expected: 9 líneas finales OK / checks OK.

- [ ] **Step 4: Commit final**

```bash
git add package.json README.md
git commit -m "docs: README modo Evolution + scripts npm (start/legacy/terminal)"
```

---

## Self-Review

**Spec coverage:**
- §3.2 config (URL/API key/instance/groupName/CA_FILE, config.json gitignored + env overrides) → Task 1 ✓
- §3.3 listGroups/findContactsMap/findMediaMessages/downloadMedia → Task 2 ✓
- §3.3 fallback findChats documentado → Task 2 (implementado: fetchAllGroups → fallback) ✓
- §3.3 filtrado client-side (#1632) → Task 2 (test con record de otro jid) ✓
- §3.3 normalización a forma baileys (`{key, message, messageTimestamp}`) → Task 2 ✓
- §3.4 flujo picker → instancia → grupo → consulta → descarga → caché → Word → Task 3 ✓
- §3.4 defensa en profundidad (isMsgInRange en path de descarga) → Task 3 + test E2E-OUT ✓
- §3.4 index.js sin cambios (solo export toUnix), start → nuevo entry, start:legacy → Task 1/4 ✓
- §5 errores (Evolution caído / instancia inexistente o desconectada / media fallida con reintentos) → Tasks 2-3 ✓
- §6 migración (caché heredado, Windows mismo config, media_cache conservado) → Tasks 3-4 ✓
- §7 testing (unit + mock HTTP + E2E + suites reusadas) → Tasks 1-4 ✓
- §8 riesgos (TLS → CA_FILE + README; media expirada → media_cache + failed cache) ✓

**Placeholders:** ninguno — todos los pasos traen código completo.

**Type consistency:**
- `runPipeline({startTs, endTs})` idéntico en Task 3 (código) y el E2E; retorna `{total, downloaded, skipped, failed, outOfRange, wordPath}` — el E2E lee exactamente esas keys.
- `downloadMedia` retorna `{buffer, mimeType}` en Task 2 y Task 3 usa solo `buffer` ✓.
- `findMediaMessages(jid, startTs, endTs)` retorna `[{key, message, messageTimestamp}]` — Task 3 lo consume igual (key.id, messageTimestamp) ✓.
- `saveFailedDownload(msgKeyId, errorType)` — aridad verificada contra index.js:807 ✓.
- `saveCachedGroupMessages(groupJid, newMessages)` itera pares `[id, msg]` (index.js:327 `for (const [, msg] of newMessages)`) → Task 3 pasa un `Map` ✓.
- `toSeconds`/`normalizeRecord` exportados en Task 2 (test) ✓; `toUnix` exportado en Task 1 (usado por evolution-client vía require('./index.js')) ✓.

**Decisiones tomadas durante el self-review (documentadas para el ejecutor):**
- El E2E respalda/restaura `group_cache.json`, `group_messages_cache.json`, `failed_downloads.json`, `Comprobantes_Descargados.docx` y `media_cache/` (las funciones heredadas escriben a constantes del módulo — no hay override posible sin tocar index.js, y eso está prohibido).
- `findMediaMessages` escanea hasta MAX_PAGES sin early-exit: el server puede ignorar el orderBy pedido (mismo patrón #1632), y el mock devuelve records en orden de fixture.
- `stats.outOfRange` queda en 0 en el E2E (la filtra findMediaMessages antes de descargar); el contador es defensa en profundidad, no un camino esperado.
- `main()` borra el Word viejo al arrancar — copia del patrón anti-confusión de index.js:1396.
