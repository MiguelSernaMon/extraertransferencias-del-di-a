'use strict';
// Cliente REST de Evolution API v2 (hosteado en VPS).
// Consulta mensajes/contactos/grupos y descarga media. Filtrado client-side
// SIEMPRE: el filtro remoteJid del server puede no aplicarse (issue #1632).

const http = require('http');
const https = require('https');
const { URL } = require('url');

const { config, getAgent } = require('./config');
const { toUnix, isMsgInRange } = require('./index.js');

// Contrato real v2.3.7 (verificado contra el VPS): `take` es IGNORADO por el
// server (página fija de 50) y `orderBy` también (siempre desc). Se manda
// take: 50 igual (inofensivo) + page explícito. MAX_PAGES cubre ~20 días a
// ~750 msgs/día (300 × 50 = 15000 records).
const TAKE = 50;
const MAX_PAGES = 300;

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

/** Normaliza un record de Evolution a la forma que usa el pipeline (como baileys).
 *  Contrato real: messageTimestamp viene DIRECTO en el record (segundos) — el
 *  campo `timestamp` no existe en producción. Se conserva el fallback legacy
 *  por compatibilidad con fixtures viejas.
 *  pushName: se preserva EXCEPTO cuando es el ECO del lid — Evolution manda
 *  pushName = "77150682091545" para participant "77150682091545@lid" cuando no
 *  conoce el nombre. Ese eco no es un nombre: si pasa, getSenderName lo devuelve
 *  como tal y el Word muestra el lid aunque exista el puente lid→teléfono. Se
 *  descarta y el fallback de getSenderName usa el teléfono real del puente. */
function normalizeRecord(record) {
  const participant = record.key?.participant;
  const pushName = record.pushName && participant?.endsWith('@lid') && record.pushName === participant.split('@')[0]
    ? null
    : record.pushName;
  return {
    key: record.key,
    message: record.message,
    messageTimestamp: toSeconds(record.messageTimestamp ?? record.timestamp),
    pushName,
  };
}

/** Lee el body completo de una respuesta HTTP y lo parsea como JSON.
 *  Rechaza si el body no es JSON (el caller lo ignora) o si se corta la lectura. */
function readJson(res) {
  return new Promise((resolve, reject) => {
    let raw = '';
    res.setEncoding('utf8');
    res.on('data', c => { raw += c; });
    res.on('end', () => {
      try { resolve(JSON.parse(raw)); } catch (e) { reject(e); }
    });
    res.on('aborted', () => reject(new Error('respuesta abortada por el server')));
    res.on('error', reject);
  });
}

/** Request HTTP con apikey + agent (CA custom) + timeout. Lanza Error con detalle del server.
 *  Usa http/https.request (NO fetch): undici ignora la opción `agent` de RequestInit,
 *  por lo que el pinning CA del plan (https.Agent({ ca })) solo funciona con el
 *  transporte nativo. Mismo contrato de retorno: JSON parseado. */
async function apiRequest(method, path, body) {
  const url = new URL(`${config.evolutionUrl}${path}`);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30_000);
  try {
    const res = await new Promise((resolve, reject) => {
      const mod = url.protocol === 'https:' ? https : http;
      const req = mod.request(url, {
        method,
        headers: {
          'Content-Type': 'application/json',
          apikey: config.apiKey,
        },
        agent: getAgent(),
        signal: controller.signal,
      }, resolve);
      req.on('error', reject);
      if (body === undefined) req.end(); else req.end(JSON.stringify(body));
    });
    let data = null;
    try { data = await readJson(res); } catch { /* body no JSON */ }
    if (!(res.statusCode >= 200 && res.statusCode < 300)) {
      const msg = data?.error?.message || `HTTP ${res.statusCode}`;
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

/** Grupos disponibles: [{id, subject}]. fetchAllGroups (requiere
 *  ?getParticipants=true — sin el param el server devuelve 400); fallback
 *  findChats filtrando @g.us. */
async function listGroups() {
  let data = null;
  try {
    data = await apiRequest('GET', `/group/fetchAllGroups/${config.instance}?getParticipants=true`);
  } catch {
    data = await apiRequest('GET', `/chat/findChats/${config.instance}`);
  }
  const list = Array.isArray(data) ? data : data?.chats || [];
  return list
    .filter(g => (g.id || g.jid || '').endsWith('@g.us'))
    .map(g => ({ id: g.id || g.jid, subject: g.subject || g.name || '' }));
}

/** Miembros de un grupo: [{id (lid), phoneNumber, admin}]. fetchAllGroups exige
 *  ?getParticipants=true (sin el param → 400). El phoneNumber es el puente
 *  lid→teléfono que usa el pipeline para resolver nombres. Fail-soft: [] si
 *  falla la llamada o el grupo no aparece. */
async function fetchGroupMembers(groupJid) {
  try {
    const data = await apiRequest('GET', `/group/fetchAllGroups/${config.instance}?getParticipants=true`);
    const list = Array.isArray(data) ? data : [];
    const g = list.find(x => x.id === groupJid);
    return g?.participants || [];
  } catch {
    return [];
  }
}

/** Contactos: { '<jid>': { name } } — jid completo como lo devuelve Evolution.
 *  Contrato real: es POST /chat/findContacts/{instance} (GET → 404) y el body
 *  SIEMPRE lleva {take: 500} — body vacío crashea el contenedor (connection
 *  reset instantáneo + API caída ~2 min). Respuesta: array directo de
 *  contacts [{remoteJid, pushName, ...}]. */
async function findContactsMap() {
  const data = await apiRequest('POST', `/chat/findContacts/${config.instance}`, { take: 500 });
  const map = {};
  if (Array.isArray(data)) {
    for (const c of data) {
      const jid = c?.remoteJid || c?.id;
      if (jid) map[jid] = { name: c.pushName || c.name || c.notify || '' };
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

/** Inicia sesión/QR de la instancia (desde el panel). Devuelve la data cruda
 *  ({qrcode: {code, base64}, pairingCode, ...}) o { error } si falla. El panel
 *  muestra el QR y la terminal/panel reflejan el estado al conectar. */
async function connectInstance() {
  try {
    return await apiRequest('POST', `/instance/connect/${config.instance}`);
  } catch (err) {
    return { error: err.message };
  }
}

/** Crea la instancia en Evolution y devuelve su QR — enlaza el número nuevo
 *  desde el panel cuando la instancia no existe (ej. tras borrar el número
 *  viejo del VPS). Contrato v2: POST /instance/create/{instance} con body
 *  {createWithQr: true}; devuelve {instance, qrcode: {code, base64}, ...}.
 *  Si el server ignora el body y no trae qrcode, se cierra con connect. */
async function createInstance() {
  try {
    const data = await apiRequest('POST', `/instance/create/${config.instance}`, { createWithQr: true });
    if (data?.qrcode) return data;
    return await connectInstance();
  } catch (err) {
    return { error: err.message };
  }
}

const isMedia = (msg) => !!msg.message?.imageMessage || !!msg.message?.documentMessage;

/** Mensajes media del grupo en [startTs, endTs], filtrados client-side.
 *  Contrato real: página fija de 50 records, paginación por `page` (1-based),
 *  respuesta {messages:{total, pages, currentPage, records}}, SIEMPRE desc
 *  (nuevo primero) — orderBy se ignora. `where` solo {key:{remoteJid}}.
 *  Dedupe por key.id obligatorio (el server devuelve records duplicados).
 *  Early-exit: si TODA la página es más vieja que startTs, las siguientes son
 *  aún más viejas (orden desc) → cortar sin escanear el resto del grupo. */
async function findMediaMessages(jid, startTs, endTs) {
  const out = new Map();
  for (let page = 1; page <= MAX_PAGES; page++) {
    const body = {
      where: { key: { remoteJid: jid } },
      take: TAKE, // el server lo ignora (página fija de 50) — se manda igual
      page,
    };
    const data = await apiRequest('POST', `/chat/findMessages/${config.instance}`, body);
    const records = data?.messages?.records || [];
    if (records.length === 0) break;
    if (records.every(r => toSeconds(r.messageTimestamp ?? r.timestamp) < startTs)) break;
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

/** Descarga el media de un mensaje. Contrato real: el body DEBE ser
 *  {message: <record completo normalizado>} — el server consulta
 *  msg.message.ephemeralMessage y sin wrapper devuelve 400. Respuesta 201:
 *  {mediaType, fileName, caption, size, mimetype, base64, buffer}; se lee
 *  base64 (+ mimetype) y se guarda como media_cache/<key.id>.jpg. */
async function downloadMedia(msg) {
  const body = { message: msg };
  const data = await apiRequest('POST', `/chat/getBase64FromMediaMessage/${config.instance}`, body);
  if (!data?.base64) throw new Error('Evolution no devolvió media para el mensaje');
  return { buffer: Buffer.from(data.base64, 'base64'), mimeType: data.mimetype || 'image/jpeg' };
}

module.exports = {
  listGroups,
  fetchGroupMembers,
  findContactsMap,
  findMediaMessages,
  downloadMedia,
  checkInstance,
  connectInstance,
  createInstance,
  normalizeRecord,
  toSeconds,
  apiRequest,
};
