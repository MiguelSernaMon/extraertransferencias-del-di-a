'use strict';

/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║   EXTRACTOR DE COMPROBANTES  —  WhatsApp → Word                ║
 * ║                                                                 ║
 * ║   INSTALACIÓN:                                                  ║
 * ║     npm install @whiskeysockets/baileys enquirer docx pino      ║
 * ║                                                                 ║
 * ║   USO:  node index.js                                           ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

const fs     = require('fs');
const QRCode = require('qrcode');
const { prompt } = require('enquirer');
const pino = require('pino');

const {
  default: makeWASocket,
  useMultiFileAuthState,
  downloadMediaMessage,
  fetchLatestBaileysVersion,
  DisconnectReason,
  Browsers,
} = require('@whiskeysockets/baileys');

const {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Table, TableRow, TableCell,
  WidthType, BorderStyle, VerticalAlign, AlignmentType, Header,
} = require('docx');

// ─── Config ─────────────────────────────────────────────────────────────────
const GROUP_NAME    = 'TRANSFERENCIAS RED POSTAL POBLADO';
const AUTH_FOLDER   = './baileys_auth';
const CACHE_FILE    = './group_cache.json';
const MSG_CACHE_FILE  = './group_messages_cache.json';
const NAME_CACHE_FILE = './name_cache.json';
const NOMBRES_FILE    = './nombres_mensajeros.json';
const OUTPUT_FILE     = 'Comprobantes_Descargados.docx';
const BAD_MAC_THRESHOLD = 5;
const MAX_AUTO_HEAL_RETRIES = 1;
const AUTO_HEAL_CONNECT_RETRIES = 5;
const TRANSIENT_CONN_THRESHOLD = 2;
const SYNC_IDLE_MS_NORMAL = 30_000;
const SYNC_IDLE_MS_FAST = 10_000;
const SYNC_GLOBAL_MS_NORMAL = 180_000;
const SYNC_GLOBAL_MS_FAST = 60_000;

// Mapeo manual de nombres: teléfono → nombre
const manualNames = {};
// Mapeo LID → teléfono (se llena desde groupMetadata)
const lidToPhone = {};

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Tracker global para detectar sesión dañada por errores de descifrado
const badMacTracker = { count: 0, installed: false };
const transientConnTracker = { count: 0, installed: false };

function isTransientConnectionClosedError(err) {
  if (!err) return false;

  const msg = String(err?.message || err || '');
  const statusCode = err?.output?.statusCode || err?.data?.statusCode;
  const payloadMsg = String(err?.output?.payload?.message || '');

  return (
    msg.includes('Connection Closed') ||
    payloadMsg.includes('Connection Closed') ||
    statusCode === 428
  );
}

function installRuntimeErrorGuards() {
  if (transientConnTracker.installed) return;
  transientConnTracker.installed = true;

  process.on('unhandledRejection', (reason) => {
    if (isTransientConnectionClosedError(reason)) {
      transientConnTracker.count += 1;
      console.log('⚠ Evento transitorio de conexión (428 Connection Closed). Se aplicará reconexión automática.');
      return;
    }
    console.error('❌ Rechazo no manejado:', reason?.message || reason);
  });

  process.on('uncaughtException', (err) => {
    if (isTransientConnectionClosedError(err)) {
      transientConnTracker.count += 1;
      console.log('⚠ Excepción transitoria de conexión capturada. Continuando con recuperación automática...');
      return;
    }

    console.error('\n❌ Error inesperado:', err?.message || err);
    process.exit(1);
  });
}

function installBadMacFilter() {
  if (badMacTracker.installed) return;
  badMacTracker.installed = true;

  const originalError = console.error.bind(console);
  console.error = (...args) => {
    const line = args.map(a => (typeof a === 'string' ? a : String(a))).join(' ');
    const isBadMac =
      line.includes('Failed to decrypt message with any known session') ||
      line.includes('Bad MAC');

    if (isBadMac) {
      badMacTracker.count += 1;
      // Mostrar solo un aviso resumido para no saturar la consola
      if (badMacTracker.count === 1) {
        originalError('⚠ Detectados errores de descifrado (Bad MAC). Intentaremos autorreparar la sesión automáticamente.');
      }
      return;
    }

    originalError(...args);
  };
}

function shouldAutoHealSession(stats) {
  const noHistorySync = stats.historyChunks === 0 && stats.totalMsgsReceived === 0;
  const manyDecryptErrors = badMacTracker.count >= BAD_MAC_THRESHOLD;
  const manyTransientConnErrors = transientConnTracker.count >= TRANSIENT_CONN_THRESHOLD;
  return noHistorySync || manyDecryptErrors || manyTransientConnErrors;
}

// Reemplaza makeInMemoryStore: guardamos mensajes en un Map simple
const msgStore = new Map(); // `${jid}:${id}` → msg

// Caché de nombres: jid/lid → nombre de WhatsApp
const nameCache = new Map();

// Mapa completo de contactos: jid → contact object (incluye lid)
const contactsMap = {};

/** Registra un contacto en todos los maps disponibles */
function registerContact(c) {
  if (!c || !c.id) return;
  contactsMap[c.id] = c;

  // Extraer nombre del contacto
  const name = c.name || c.verifiedName || c.notify;
  if (name) {
    nameCache.set(c.id, name);
    // Si tiene LID, también mapear el LID
    if (c.lid) nameCache.set(c.lid, name);
  }

  // Si el ID es un LID y tiene nombre, registrar
  if (c.id.endsWith('@lid') && name) {
    nameCache.set(c.id, name);
  }
}

/** Carga el caché de nombres desde disco */
function loadNameCache() {
  if (!fs.existsSync(NAME_CACHE_FILE)) return;
  try {
    const data = JSON.parse(fs.readFileSync(NAME_CACHE_FILE, 'utf8'));
    for (const [jid, name] of Object.entries(data)) {
      nameCache.set(jid, name);
    }
  } catch { /* ignore */ }
}

/** Guarda el caché de nombres a disco */
function saveNameCache() {
  const obj = {};
  for (const [jid, name] of nameCache) {
    obj[jid] = name;
  }
  fs.writeFileSync(NAME_CACHE_FILE, JSON.stringify(obj, null, 2));
}

// ─── Caché persistente de mensajes ──────────────────────────────────────────

/** Carga los mensajes cacheados del grupo desde disco */
function loadCachedGroupMessages() {
  if (!fs.existsSync(MSG_CACHE_FILE)) return new Map();
  try {
    const data = JSON.parse(fs.readFileSync(MSG_CACHE_FILE, 'utf8'));
    const map = new Map();
    for (const msg of data) {
      if (msg.key?.id) map.set(msg.key.id, msg);
      // Extraer pushNames de mensajes cacheados
      const participant = msg.key?.participant;
      if (participant && msg.pushName) {
        nameCache.set(participant, msg.pushName);
      }
    }
    return map;
  } catch {
    return new Map();
  }
}

/** Guarda mensajes del grupo a disco (merge con existentes) */
function saveCachedGroupMessages(groupJid, newMessages) {
  const cached = loadCachedGroupMessages();
  let added = 0;

  // Agregar mensajes nuevos del store en memoria
  for (const [, msg] of newMessages) {
    if (msg.key?.remoteJid !== groupJid) continue;
    if (!msg.message?.imageMessage && !msg.message?.documentMessage) continue;
    if (!cached.has(msg.key.id)) added++;
    cached.set(msg.key.id, msg);
  }

  fs.writeFileSync(MSG_CACHE_FILE, JSON.stringify([...cached.values()]));
  console.log(`💾 Caché actualizado: ${cached.size} comprobantes guardados (+${added} nuevos)`);
  return cached;
}

function getCacheRangeStats(cacheMap, startTs, endTs) {
  let inRange = 0;
  let minTs = Number.MAX_SAFE_INTEGER;
  let maxTs = 0;

  for (const [, msg] of cacheMap) {
    const ts = toUnix(msg.messageTimestamp);
    if (!ts) continue;
    if (ts < minTs) minTs = ts;
    if (ts > maxTs) maxTs = ts;
    if (ts >= startTs && ts <= endTs) inRange++;
  }

  return {
    inRange,
    minTs: minTs === Number.MAX_SAFE_INTEGER ? 0 : minTs,
    maxTs,
  };
}

// ════════════════════════════════════════════════════════════════════════════
// 1. PREGUNTAR FECHAS
// ════════════════════════════════════════════════════════════════════════════
async function askDateRange() {
  console.log('\n══════════════════════════════════════════════════════');
  console.log('  CONFIGURACIÓN DE FECHAS  (máximo 1 semana)');
  console.log('══════════════════════════════════════════════════════');
  console.log('  Formato fecha: YYYY-MM-DD   Hora: HH:MM\n');

  const answers = await prompt([
    { type: 'input', name: 'startDate', message: 'Fecha INICIO (YYYY-MM-DD) [Enter = AYER]:' },
    { type: 'input', name: 'startTime', message: 'Hora  INICIO (HH:MM)      [Enter = 08:30]:' },
    { type: 'input', name: 'endDate',   message: 'Fecha FIN    (YYYY-MM-DD) [Enter = HOY ]:' },
    { type: 'input', name: 'endTime',   message: 'Hora  FIN    (HH:MM)      [Enter = 23:59]:' },
  ]);

  const parseTime = (str, defH, defM) => {
    if (!str?.trim()) return [defH, defM];
    const [h, m] = str.trim().split(':').map(Number);
    return [
      !isNaN(h) && h >= 0 && h <= 23 ? h : defH,
      !isNaN(m) && m >= 0 && m <= 59 ? m : defM,
    ];
  };

  const buildDate = (dateStr, timeStr, offsetDays, defH, defM, sec, ms) => {
    const [h, m] = parseTime(timeStr, defH, defM);
    if (dateStr?.trim()) {
      const [y, mo, d] = dateStr.trim().split('-').map(Number);
      return new Date(y, mo - 1, d, h, m, sec, ms);
    }
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate() + offsetDays, h, m, sec, ms);
  };

  const startDate = buildDate(answers.startDate, answers.startTime, -1,  8, 30,  0,   0);
  const endDate   = buildDate(answers.endDate,   answers.endTime,    0, 23, 59, 59, 999);

  console.log(`\n  📅 INICIO : ${startDate.toLocaleString('es-CO')}`);
  console.log(`  📅 FIN    : ${endDate.toLocaleString('es-CO')}\n`);

  if (endDate <= startDate) {
    console.error('❌ La fecha de fin debe ser posterior a la de inicio.');
    process.exit(1);
  }

  return { startDate, endDate };
}

// ════════════════════════════════════════════════════════════════════════════
// 2. CREAR SOCKET
// ════════════════════════════════════════════════════════════════════════════
async function createSocket() {
  const { state, saveCreds } = await useMultiFileAuthState(AUTH_FOLDER);
  const { version, isLatest } = await fetchLatestBaileysVersion();
  console.log(`  WA v${version.join('.')} ${isLatest ? '(última)' : '(desactualizada)'}\n`);

  const logger = pino({ level: 'silent' });

  const sock = makeWASocket({
    version,
    logger,
    auth: state,
    browser: Browsers.ubuntu('Chrome'),
    syncFullHistory: true,
    markOnlineOnConnect: false,
    // getMessage usa nuestro store manual
    getMessage: async key => {
      const stored = msgStore.get(`${key.remoteJid}:${key.id}`);
      return stored?.message ?? undefined;
    },
  });

  sock.ev.on('creds.update', saveCreds);

  // Guardar mensajes en el store manual + extraer pushNames
  sock.ev.on('messages.upsert', ({ messages }) => {
    for (const msg of messages) {
      if (msg.key?.remoteJid && msg.key?.id) {
        msgStore.set(`${msg.key.remoteJid}:${msg.key.id}`, msg);
      }
      // Extraer pushName para el caché de nombres
      const participant = msg.key?.participant;
      if (participant && msg.pushName) {
        nameCache.set(participant, msg.pushName);
      }
    }
  });

  sock.ev.on('messaging-history.set', (data) => {
    const { messages = [], contacts: syncContacts = [], isLatest } = data;

    // Guardar mensajes
    for (const msg of messages) {
      if (msg.key?.remoteJid && msg.key?.id) {
        msgStore.set(`${msg.key.remoteJid}:${msg.key.id}`, msg);
      }
      const participant = msg.key?.participant;
      if (participant && msg.pushName) {
        nameCache.set(participant, msg.pushName);
      }
    }

    // Registrar contactos del history sync (clave para mapear LIDs)
    if (syncContacts && syncContacts.length > 0) {
      for (const c of syncContacts) {
        registerContact(c);
      }
    }
  });

  // Capturar también los contactos de los eventos dedicados
  sock.ev.on('contacts.set',    ({ contacts: list }) => {
    if (list) list.forEach(c => registerContact(c));
  });
  sock.ev.on('contacts.upsert', list => {
    if (list) list.forEach(c => registerContact(c));
  });
  sock.ev.on('contacts.update', list => {
    if (list) list.forEach(c => {
      if (c.id && contactsMap[c.id]) {
        Object.assign(contactsMap[c.id], c);
        registerContact(contactsMap[c.id]);
      } else {
        registerContact(c);
      }
    });
  });

  return sock;
}

// ════════════════════════════════════════════════════════════════════════════
// 3. ESPERAR CONEXIÓN — devuelve 'open' | 'restart' | 'loggedOut'
// ════════════════════════════════════════════════════════════════════════════
function waitForOpen(sock) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error('Timeout: no se pudo conectar en 3 minutos.')),
      180_000
    );

    let settled = false;
    const cleanup = () => {
      clearTimeout(timer);
      sock.ev.off('connection.update', onConnectionUpdate);
    };

    const safeResolve = (value) => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(value);
    };

    const onConnectionUpdate = ({ connection, lastDisconnect, qr }) => {
      if (qr) {
        console.log('\n══════════════════════════════════════════════════════');
        console.log('  ESCANEA ESTE CÓDIGO QR CON WHATSAPP');
        console.log('══════════════════════════════════════════════════════\n');
        QRCode.toString(qr, { type: 'terminal', small: true }, (err, url) => {
          if (!err) console.log(url);
        });
      }
      if (connection === 'open') {
        console.log('✅ Conectado a WhatsApp.\n');
        safeResolve('open');
      }
      if (connection === 'close') {
        const code = lastDisconnect?.error?.output?.statusCode;
        if (code === DisconnectReason.loggedOut) {
          console.log('\n⚠  Sesión cerrada. Elimina "baileys_auth" y reescanea el QR.\n');
          fs.rmSync(AUTH_FOLDER, { recursive: true, force: true });
          safeResolve('loggedOut');
          return;
        }
        // 515 = restartRequired: WhatsApp pide reconectar
        if (code === 515 || code === DisconnectReason.restartRequired) {
          console.log('⚠  WhatsApp pidió reconexión (515). Reintentando...\n');
          safeResolve('restart');
          return;
        }
        console.log(`⚠  Conexión cerrada (código ${code}). Reintentando...\n`);
        safeResolve('restart');
      }

    };

    sock.ev.on('connection.update', onConnectionUpdate);
  });
}

// ════════════════════════════════════════════════════════════════════════════
// 4. ENCONTRAR JID DEL GRUPO
// ════════════════════════════════════════════════════════════════════════════
async function findGroupJid(sock) {
  if (fs.existsSync(CACHE_FILE)) {
    try {
      const { groupJid, groupName } = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
      console.log(`📦 Grupo en caché: ${groupName}`);
      return groupJid;
    } catch {
      fs.unlinkSync(CACHE_FILE);
    }
  }

  console.log('🔍 Buscando grupo...');
  let groups;
  try {
    groups = await sock.groupFetchAllParticipating();
  } catch (err) {
    throw new Error(`No se pudo obtener grupos: ${err.message}`);
  }

  const entries = Object.values(groups);
  const match   = entries.find(g =>
    g.subject?.toUpperCase().includes(GROUP_NAME.toUpperCase())
  );

  if (!match) {
    console.log('\n❌ Grupo no encontrado. Grupos disponibles:');
    entries.forEach(g => console.log(`   • ${g.subject}`));
    console.log('\n💡 Ajusta GROUP_NAME en el código.');
    throw new Error('Grupo no encontrado.');
  }

  fs.writeFileSync(CACHE_FILE, JSON.stringify({ groupJid: match.id, groupName: match.subject }));
  console.log(`✅ Grupo: ${match.subject}\n`);
  return match.id;
}

// ════════════════════════════════════════════════════════════════════════════
// 5. RECOLECTAR MENSAJES DEL GRUPO EN EL RANGO DE FECHAS
// ════════════════════════════════════════════════════════════════════════════

/** Convierte messageTimestamp (puede ser Long, string o number) a Unix seconds */
function toUnix(ts) {
  if (ts == null) return 0;
  // Baileys usa Long objects ({ low, high, unsigned })
  if (typeof ts === 'object' && ts.low !== undefined) {
    return Number(ts.toNumber ? ts.toNumber() : ts.low);
  }
  return Number(ts);
}

function collectMessages(sock, groupJid, startTs, endTs, options = {}) {
  return new Promise(resolve => {
    const idleMs = options.idleMs ?? SYNC_IDLE_MS_NORMAL;
    const globalMs = options.globalMs ?? SYNC_GLOBAL_MS_NORMAL;

    const collected = new Map();
    let idleTimer;
    let globalTimer;
    let finished = false;
    let historyChunks = 0;
    let totalMsgsReceived = 0;
    let totalGroupMsgs = 0;
    let oldestGroupTsSeen = Number.MAX_SAFE_INTEGER;

    const finish = () => {
      if (finished) return;
      finished = true;
      clearTimeout(idleTimer);
      clearTimeout(globalTimer);
      sock.ev.off('messaging-history.set', onHistory);
      sock.ev.off('messages.upsert',       onUpsert);

      // ── Fallback: también revisar el msgStore completo ──
      let fromStore = 0;
      for (const [key, msg] of msgStore.entries()) {
        if (!msg.key?.remoteJid || msg.key.remoteJid !== groupJid) continue;
        const ts = toUnix(msg.messageTimestamp);
        if (ts < startTs || ts > endTs) continue;
        if (!isMedia(msg)) continue;
        if (!collected.has(msg.key.id)) {
          collected.set(msg.key.id, msg);
          fromStore++;
        }
      }
      if (fromStore > 0) {
        console.log(`  📦 ${fromStore} comprobantes adicionales encontrados en caché local.`);
      }

      console.log(`\n  📈 Debug: ${historyChunks} chunks de historial, ${totalMsgsReceived} mensajes totales recibidos, ${totalGroupMsgs} del grupo`);
      resolve({
        messages: [...collected.values()],
        stats: {
          historyChunks,
          totalMsgsReceived,
          totalGroupMsgs,
          totalCollected: collected.size,
        }
      });
    };

    const resetIdle = () => {
      clearTimeout(idleTimer);
      // Espera adaptativa: normal o rápida si ya hay buen caché local
      idleTimer = setTimeout(() => {
        process.stdout.write('\n  ⏱ Sin más mensajes entrantes. Continuando...\n');
        finish();
      }, idleMs);
    };

    const isMedia = msg =>
      !!msg.message?.imageMessage || !!msg.message?.documentMessage;

    const addMsg = (msg, source) => {
      totalMsgsReceived++;
      if (!msg.key?.remoteJid) return;
      if (msg.key.remoteJid !== groupJid) return;
      totalGroupMsgs++;

      const ts = toUnix(msg.messageTimestamp);
  if (ts > 0 && ts < oldestGroupTsSeen) oldestGroupTsSeen = ts;

      // Debug: mostrar primer mensaje del grupo para verificar timestamps
      if (totalGroupMsgs <= 3) {
        const date = new Date(ts * 1000);
        console.log(`  🔎 [${source}] Mensaje del grupo: ts=${ts} (${date.toLocaleString('es-CO')}), tipo=${
          msg.message?.imageMessage ? 'imagen' :
          msg.message?.documentMessage ? 'documento' :
          msg.message?.conversation ? 'texto' :
          msg.message?.extendedTextMessage ? 'texto ext.' :
          Object.keys(msg.message || {}).join(',') || 'vacío'
        }`);
      }

      if (ts < startTs || ts > endTs) return;
      if (!isMedia(msg)) return;
      if (!collected.has(msg.key.id)) {
        collected.set(msg.key.id, msg);
        process.stdout.write(`\r  📨 Comprobantes encontrados: ${collected.size}   `);
      }
    };

    const onHistory = ({ messages: msgs, isLatest }) => {
      historyChunks++;
      console.log(`  📥 Chunk de historial #${historyChunks}: ${msgs.length} mensajes (isLatest=${isLatest})`);
      msgs.forEach(m => addMsg(m, 'history'));

      // Corte temprano: si ya llegamos a mensajes anteriores al inicio del rango,
      // no hace falta seguir sincronizando todo el historial.
      if (oldestGroupTsSeen < startTs) {
        process.stdout.write('\n  ⚡ Ya alcanzamos mensajes más viejos que el inicio del rango. Cerrando sincronización anticipadamente...\n');
        clearTimeout(idleTimer);
        idleTimer = setTimeout(finish, 2_500);
        return;
      }

      resetIdle();
      if (isLatest) {
        process.stdout.write('\n  ✅ Historial sincronizado.\n');
        clearTimeout(idleTimer);
        // Esperar 8s después de "isLatest" por si llegan más chunks
        idleTimer = setTimeout(finish, 8_000);
      }
    };

    const onUpsert = ({ messages: msgs, type }) => {
      // Procesar TODOS los tipos: 'notify' (tiempo real) y 'append' (históricos)
      console.log(`  📬 Upsert: ${msgs.length} mensajes, type=${type}`);
      msgs.forEach(m => addMsg(m, `upsert-${type}`));
      if (type === 'append') resetIdle();
    };

    sock.ev.on('messaging-history.set', onHistory);
    sock.ev.on('messages.upsert',       onUpsert);

    // Tiempo máximo: 3 minutos (antes eran 2)
    globalTimer = setTimeout(() => {
      process.stdout.write('\n  ⏱ Tiempo máximo alcanzado.\n');
      finish();
    }, globalMs);

    resetIdle();
  });
}

// ════════════════════════════════════════════════════════════════════════════
// 6. DESCARGAR MEDIA
// ════════════════════════════════════════════════════════════════════════════
async function downloadImage(sock, msg) {
  try {
    const buffer = await downloadMediaMessage(
      msg,
      'buffer',
      {},
      { logger: pino({ level: 'silent' }), reuploadRequest: sock.updateMediaMessage }
    );
    return buffer;
  } catch (err) {
    process.stdout.write(`\n  ⚠ No se pudo descargar: ${err.message}\n`);
    return null;
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 7. NOMBRE DEL REMITENTE
// ════════════════════════════════════════════════════════════════════════════
function getSenderName(msg, contacts) {
  const jid = msg.key.participant ?? msg.key.remoteJid;

  // 1. Buscar en nombres manuales (archivo nombres_mensajeros.json)
  //    El jid puede ser un LID → convertir a teléfono → buscar nombre manual
  const phone = lidToPhone[jid] || jid.split('@')[0];
  if (manualNames[phone]) return manualNames[phone];

  // 2. Buscar en contactos sincronizados y caché
  const c = contacts[jid] || contactsMap[jid];
  const name = c?.name || c?.verifiedName || msg.pushName || nameCache.get(jid) || c?.notify;
  if (name) return name;

  // 3. Mostrar el número de teléfono (más útil que el LID)
  return phone;
}

async function reconnectWithRetries(maxRetries, phaseLabel) {
  let sock;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    console.log(`🔄 ${phaseLabel} (intento ${attempt}/${maxRetries})...\n`);
    sock = await createSocket();

    const result = await waitForOpen(sock);
    if (result === 'open') return sock;

    try { await sock.end(); } catch { /* ignore */ }

    if (attempt < maxRetries) {
      await sleep(3_000);
    }
  }

  throw new Error(`No fue posible completar "${phaseLabel}" después de ${maxRetries} intentos.`);
}

// ════════════════════════════════════════════════════════════════════════════
// 8. MAIN
// ════════════════════════════════════════════════════════════════════════════
async function main() {
  installRuntimeErrorGuards();
  installBadMacFilter();
  badMacTracker.count = 0;
  transientConnTracker.count = 0;

  const { startDate, endDate } = await askDateRange();
  const startTs = Math.floor(startDate.getTime() / 1000);
  const endTs   = Math.floor(endDate.getTime()   / 1000);

  // ── Cargar cachés de ejecuciones previas ─────────────────────────────────
  loadNameCache();
  const previousCache = loadCachedGroupMessages();
  const previousCacheStats = getCacheRangeStats(previousCache, startTs, endTs);
  if (previousCache.size > 0) {
    console.log(`📦 Caché local: ${previousCache.size} comprobantes, ${nameCache.size} nombres guardados`);
    if (previousCacheStats.inRange > 0) {
      console.log(`⚡ Cache-hit en rango: ${previousCacheStats.inRange} comprobantes ya estaban guardados`);
    }
  }

  // ── Conexión con reintentos automáticos ───────────────────────────────────
  let sock = await reconnectWithRetries(5, 'Conectando con WhatsApp');
  let contacts = {};

  let groupJid = await findGroupJid(sock);

  // ── Obtener participantes y construir mapeo LID → teléfono ────────────────
  try {
    const metadata = await sock.groupMetadata(groupJid);

    // Construir mapeo LID → teléfono desde metadata del grupo
    for (const p of metadata.participants) {
      const phoneLid = p.lid || p.id;
      const phoneJid = p.jid || p.id;
      const phoneNum = phoneJid.split('@')[0];

      // Mapear LID → número de teléfono
      if (phoneLid) lidToPhone[phoneLid] = phoneNum;
      if (p.id) lidToPhone[p.id] = phoneNum;
    }

    // Generar/actualizar archivo de nombres manuales
    if (fs.existsSync(NOMBRES_FILE)) {
      // Cargar nombres existentes
      try {
        const existing = JSON.parse(fs.readFileSync(NOMBRES_FILE, 'utf8'));
        Object.assign(manualNames, existing);
      } catch { /* ignore */ }

      // Agregar nuevos participantes que no estén en el archivo
      let newEntries = 0;
      for (const p of metadata.participants) {
        const phoneNum = (p.jid || p.id).split('@')[0];
        if (!manualNames[phoneNum]) {
          manualNames[phoneNum] = '';
          newEntries++;
        }
      }
      if (newEntries > 0) {
        fs.writeFileSync(NOMBRES_FILE, JSON.stringify(manualNames, null, 2));
        console.log(`   📝 ${newEntries} nuevos participantes agregados al archivo de nombres`);
      }
    } else {
      // Primera vez: crear archivo con todos los participantes
      for (const p of metadata.participants) {
        const phoneNum = (p.jid || p.id).split('@')[0];
        manualNames[phoneNum] = '';
      }
      fs.writeFileSync(NOMBRES_FILE, JSON.stringify(manualNames, null, 2));
      console.log(`\n  ╔══════════════════════════════════════════════════════╗`);
      console.log(`  ║  📋 ARCHIVO DE NOMBRES CREADO                        ║`);
      console.log(`  ║                                                      ║`);
      console.log(`  ║  Abre: nombres_mensajeros.json                       ║`);
      console.log(`  ║  Pon el nombre de cada mensajero junto a su número.  ║`);
      console.log(`  ║  Ejemplo: "573001234567": "JUAN PÉREZ"                ║`);
      console.log(`  ╚══════════════════════════════════════════════════════╝\n`);
    }

    // Contar nombres asignados
    const assigned = Object.values(manualNames).filter(n => n).length;
    console.log(`👥 Participantes: ${metadata.participants.length}, nombres asignados: ${assigned}/${Object.keys(manualNames).length}`);
  } catch (err) {
    console.log(`  ⚠ No se pudo obtener metadata del grupo: ${err.message}`);
  }

  console.log('⏳ Esperando sincronización de historial...');
  console.log('   (30-90 s la primera vez; más rápido en ejecuciones siguientes)');
  console.log(`   🔎 Rango: ${new Date(startTs * 1000).toLocaleString('es-CO')} → ${new Date(endTs * 1000).toLocaleString('es-CO')}\n`);

  // Manejar cierre de conexión durante la recolección
  const onClose = ({ connection }) => {
    if (connection === 'close') {
      console.log('\n  ⚠ Conexión cerrada durante sincronización, procesando lo recolectado...');
    }
  };
  sock.ev.on('connection.update', onClose);

  const nowTs = Math.floor(Date.now() / 1000);
  const rangeAlreadyCoveredByCache =
    previousCacheStats.inRange > 0 &&
    previousCacheStats.maxTs >= endTs &&
    endTs <= nowTs - 300;

  const useFastSync = previousCacheStats.inRange > 0;
  const collectOptions = useFastSync
    ? { idleMs: SYNC_IDLE_MS_FAST, globalMs: SYNC_GLOBAL_MS_FAST }
    : { idleMs: SYNC_IDLE_MS_NORMAL, globalMs: SYNC_GLOBAL_MS_NORMAL };

  let collection;
  if (rangeAlreadyCoveredByCache) {
    console.log('⚡ Rango ya cubierto por caché local. Omitiendo sincronización completa para acelerar.\n');
    collection = {
      messages: [],
      stats: { historyChunks: 1, totalMsgsReceived: 1, totalGroupMsgs: 1, totalCollected: 0 }
    };
  } else {
    if (useFastSync) {
      console.log('⚡ Activando sincronización rápida (hay caché previa en el rango).\n');
    }
    collection = await collectMessages(sock, groupJid, startTs, endTs, collectOptions);
  }

  let healAttempt = 0;
  while (shouldAutoHealSession(collection.stats) && healAttempt < MAX_AUTO_HEAL_RETRIES) {
    healAttempt++;
    console.log('\n🛠 Modo autorreparación: detectamos sesión inestable o sin sincronización de historial.');
    console.log('   Se renovará la sesión automáticamente para que no tengas que borrar carpetas manualmente.\n');

    try { await sock.end(); } catch { /* ignore */ }
    fs.rmSync(AUTH_FOLDER, { recursive: true, force: true });
    msgStore.clear();
    badMacTracker.count = 0;
  transientConnTracker.count = 0;

    console.log(`🔄 Reconectando con sesión limpia (${healAttempt}/${MAX_AUTO_HEAL_RETRIES})...\n`);
    sock = await reconnectWithRetries(
      AUTO_HEAL_CONNECT_RETRIES,
      'Reabriendo WhatsApp durante autorreparación'
    );

    groupJid = await findGroupJid(sock);

    console.log('⏳ Reintentando sincronización de historial con sesión renovada...\n');
    collection = await collectMessages(sock, groupJid, startTs, endTs);
  }

  sock.ev.off('connection.update', onClose);

  // ── Guardar TODOS los mensajes nuevos al caché persistente ────────────────
  const allCached = saveCachedGroupMessages(groupJid, msgStore);

  // ── Guardar caché de nombres ────────────────────────────────────────────
  saveNameCache();
  console.log(`👤 Nombres cacheados: ${nameCache.size}`);

  // ── Filtrar por rango de fechas desde el caché completo ───────────────────
  const rawMsgs = [];
  for (const [, msg] of allCached) {
    const ts = toUnix(msg.messageTimestamp);
    if (ts >= startTs && ts <= endTs) {
      rawMsgs.push(msg);
    }
  }

  // Ordenar por fecha
  rawMsgs.sort((a, b) => toUnix(a.messageTimestamp) - toUnix(b.messageTimestamp));

  console.log(`📊 Total comprobantes en el rango: ${rawMsgs.length}`);

  if (rawMsgs.length === 0) {
    console.log('\n⚠  No se encontraron imágenes en el rango seleccionado.');
    console.log('   Si es la primera vez que usas la app, los mensajes');
    console.log('   se guardaron en caché y estarán disponibles en');
    console.log('   las próximas ejecuciones.');
  console.log('   El sistema intentó autorrepararse; si persiste, vuelve a ejecutar y espera 1-2 minutos extra de sincronización.\n');
    try { await sock.end(); } catch { /* ignore */ }
    process.exit(0);
  }

  console.log('\n⬇  Descargando imágenes...\n');
  const receipts = [];

  for (let i = 0; i < rawMsgs.length; i++) {
    const msg = rawMsgs[i];
    process.stdout.write(`\r  [${i + 1}/${rawMsgs.length}] Descargando...`);

    const buffer = await downloadImage(sock, msg);
    if (!buffer) continue;

    const mimetype =
      msg.message?.imageMessage?.mimetype ??
      msg.message?.documentMessage?.mimetype ?? '';

    if (!mimetype.startsWith('image/')) {
      process.stdout.write(`\n  ⚠ Omitido — no es imagen (${mimetype || 'sin tipo'})\n`);
      continue;
    }

    receipts.push({
      imageBuffer: buffer,
      senderName:  getSenderName(msg, contacts),
      date:        new Date(toUnix(msg.messageTimestamp) * 1000).toLocaleString('es-CO'),
    });
  }

  console.log(`\n\n✅ Imágenes válidas: ${receipts.length}`);

  try { await sock.end(); } catch { /* ignore */ }

  if (receipts.length === 0) {
    console.log('⚠  Nada que exportar.\n');
    process.exit(0);
  }

  await createWordDocument(receipts);
  process.exit(0);
}

// ════════════════════════════════════════════════════════════════════════════
// 9. GENERAR DOCUMENTO WORD
// ════════════════════════════════════════════════════════════════════════════
async function createWordDocument(receipts) {
  console.log('\n📄 Generando documento Word...');

  const order   = [];
  const grouped = {};
  for (const r of receipts) {
    if (!grouped[r.senderName]) { grouped[r.senderName] = []; order.push(r.senderName); }
    grouped[r.senderName].push(r);
  }

  const sections = [];

  for (const sender of order) {
    const list  = grouped[sender];
    const pages = Math.ceil(list.length / 6);

    for (let p = 0; p < pages; p++) {
      const chunk   = list.slice(p * 6, (p + 1) * 6);
      const isExtra = p > 0;

      const header = new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'MENSAJERO: ', bold: true, size: 36, color: '1a1a1a' }),
              new TextRun({
                text: `${sender.toUpperCase()}${isExtra ? '  (Continuación)' : ''}`,
                bold: true, size: 36, color: '003399',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 80, after: 120 },
            border: { bottom: { style: BorderStyle.THICK, size: 8, color: '003399' } },
          }),
          new Paragraph({
            children: [new TextRun({
              text: `Hoja ${p + 1} / ${pages}   |   Total comprobantes: ${list.length}   |   Valor a verificar: $________________________`,
              size: 22, color: '555555',
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 0 },
          }),
        ],
      });

      const rows = [];
      for (let ri = 0; ri < chunk.length; ri += 3) {
        const cells = [];
        for (let col = 0; col < 3; col++) {
          const idx = ri + col;
          cells.push(idx < chunk.length ? receiptCell(chunk[idx]) : emptyCell());
        }
        rows.push(new TableRow({ children: cells, height: { value: 6000, rule: 'atLeast' } }));
      }

      const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
      sections.push({
        properties: {
          page: {
            size: { width: 12240, height: 15840, orientation: 'portrait' },
            margin: { top: 1000, right: 600, bottom: 600, left: 600, header: 300, footer: 200 },
          },
        },
        headers: { default: header },
        children: [new Table({
          rows,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: { top: NO, bottom: NO, left: NO, right: NO, insideHorizontal: NO, insideVertical: NO },
        })],
      });
    }
  }

  const buffer = await Packer.toBuffer(new Document({ sections }));
  fs.writeFileSync(OUTPUT_FILE, buffer);

  console.log(`\n✅ Documento listo → ${OUTPUT_FILE}`);
  console.log(`   Mensajeros   : ${order.length}`);
  console.log(`   Hojas        : ${sections.length}`);
  console.log(`   Comprobantes : ${receipts.length}\n`);
}

// ── Celda con imagen ──────────────────────────────────────────────────────────
function receiptCell(receipt) {
  const G = { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' };
  return new TableCell({
    width: { size: 3333, type: WidthType.DXA },
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    verticalAlign: VerticalAlign.CENTER,
    borders: { top: G, bottom: G, left: G, right: G },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 5 },
        children: [new TextRun({ text: receipt.date, bold: true, size: 12, color: '333333' })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: receipt.imageBuffer, transformation: { width: 230, height: 420 } })],
      }),
    ],
  });
}

// ── Celda vacía ───────────────────────────────────────────────────────────────
function emptyCell() {
  const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new TableCell({
    children: [new Paragraph({ text: '' })],
    width: { size: 3333, type: WidthType.DXA },
    borders: { top: NO, bottom: NO, left: NO, right: NO },
  });
}

main().catch(err => {
  console.error('\n❌ Error fatal:', err.message);
  process.exit(1);
});