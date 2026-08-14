'use strict';
// Cliente REST de Evolution API v2 (hosteado en VPS).
// Consulta mensajes/contactos/grupos y descarga media. Filtrado client-side
// SIEMPRE: el filtro remoteJid del server puede no aplicarse (issue #1632).

const http = require('http');
const https = require('https');
const { URL } = require('url');

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
