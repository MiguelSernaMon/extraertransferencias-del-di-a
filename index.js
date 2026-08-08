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
  atomicWriteFileSync,
  loadJSONFile,
  isValidJpeg,
} = require('./fs_utils');
const { startControlServer } = require('./picker');

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
const BAD_MAC_THRESHOLD = 15;
const MAX_AUTO_HEAL_RETRIES = 1;
const AUTO_HEAL_CONNECT_RETRIES = 5;
const SESSION_SNAPSHOT_DIR = './baileys_auth_backup_snapshot';
const SYNC_IDLE_MS_NORMAL = 30_000;
const SYNC_IDLE_MS_FAST = 10_000;
const SYNC_GLOBAL_MS_NORMAL = 180_000;
const SYNC_GLOBAL_MS_FAST = 60_000;
const FAILED_CACHE_FILE = './failed_downloads.json';
const MEDIA_CACHE_DIR   = './media_cache';
const MAX_CONCURRENT_DOWNLOADS = 2;
const MAX_DOWNLOAD_RETRIES = 4;

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

/**
 * Decisión conservadora de autorreparación: borrar/renovar la sesión es caro
 * (obliga a reescanear el QR), así que solo se hace con evidencia fuerte.
 *
 *  - Los errores Bad MAC aislados NO son daño de sesión (un remitente con
 *    clave vieja genera varios; el resto del grupo sigue descifrando bien).
 *  - Una sesión realmente rota falla al descifrar TODO → muchos errores Y
 *    casi cero mensajes válidos recibidos.
 *  - "No llegaron mensajes en el sync" tampoco es daño de sesión (WhatsApp
 *    puede no tener nada nuevo, o la red frenó) → NO dispara heal.
 *  - Errores transitorios de conexión NO disparan heal: baileys se reconecta
 *    solo conservando la sesión.
 */
function shouldAutoHealSession(stats) {
  const total = stats?.totalMsgsReceived ?? 0;
  const manyDecryptErrors = badMacTracker.count >= BAD_MAC_THRESHOLD;
  const almostNothingDecrypted = total < BAD_MAC_THRESHOLD;
  return manyDecryptErrors && almostNothingDecrypted;
}

// ── Snapshot de sesión: respaldo para recuperar sin reescanear QR ────────────
let sessionRestoreAttempted = false;

/** Copia la sesión actual a un snapshot (después de conectar OK). */
function backupSessionSnapshot() {
  if (!fs.existsSync(AUTH_FOLDER)) return;
  try {
    fs.rmSync(SESSION_SNAPSHOT_DIR, { recursive: true, force: true });
    fs.cpSync(AUTH_FOLDER, SESSION_SNAPSHOT_DIR, { recursive: true });
  } catch { /* si no se puede respaldar, seguimos igual */ }
}

/** Restaura el snapshot sobre la sesión actual. Devuelve true si hubo snapshot. */
function tryRestoreSessionSnapshot() {
  if (!fs.existsSync(SESSION_SNAPSHOT_DIR)) return false;
  try {
    fs.rmSync(AUTH_FOLDER, { recursive: true, force: true });
    fs.cpSync(SESSION_SNAPSHOT_DIR, AUTH_FOLDER, { recursive: true });
    return true;
  } catch {
    return false;
  }
}

// Reemplaza makeInMemoryStore: guardamos mensajes en un Map simple
const msgStore = new Map(); // `${jid}:${id}` → msg

// Caché de nombres: jid/lid → nombre de WhatsApp
const nameCache = new Map();

// Mapa completo de contactos: jid → contact object (incluye lid)
const contactsMap = {};

// Nombre del grupo: quirk de baileys — en el sync de historial, pushName de
// mensajes de grupo llega como el NOMBRE DEL GRUPO, no el del remitente.
// Todo nombre sospechoso de ser el grupo se descarta (daría filas con el
// nombre del grupo para todos los mensajeros en Windows/sesiones nuevas).
let knownGroupName = '';
// Nombre del dueño de la sesión: los mensajes propios (fromMe) no traen
// participant en history sync → antes caían a remoteJid = jid del GRUPO
// ("MENSAJERO: 120363192583651767"). Con esto salen con tu nombre.
let ownName = 'YO';

// Contadores de diagnóstico por remitente: algunos mensajes de history sync
// en Windows llegan SIN key.participant → el reporte los une todos en "sdc".
let senderDiag = { sin: 0, con: 0, propios: 0 };

// Lista de participantes para el panel web (mapeo de IDs → nombres)
let participantList = [];
const participantLids = new Set();

function senderToEntry(lid, phone) {
  return {
    lid,
    phone: phone || null,
    name: manualNames[phone] || manualNames[lid] || manualNames[lid?.split('@')[0]] || '',
  };
}

/** Rellena el panel SIN conectar: remitentes del caché local + archivo de nombres. */
function pushEarlyParticipants() {
  loadManualNames(); // nombres + mapeo persistido del archivo
  const senders = new Map();
  try {
    const msgs = JSON.parse(fs.readFileSync(MSG_CACHE_FILE, 'utf8'));
    if (Array.isArray(msgs)) {
      for (const m of msgs) {
        const lid = m.key?.participant;
        if (lid && lid.endsWith('@lid')) senders.set(lid, lidToPhone[lid] || null);
      }
    }
  } catch { /* todavía no hay caché: el panel muestra vacío hasta conectar */ }
  participantList = [];
  participantLids.clear();
  for (const [lid, phone] of senders) {
    participantLids.add(lid);
    participantList.push(senderToEntry(lid, phone));
  }
  if (activePicker) activePicker.setParticipants(participantList);
}

/** true si el nombre parece ser el del grupo (pushName envenenado). */
function isGroupNameLike(name) {
  if (!knownGroupName || !name) return false;
  const g = knownGroupName.toLowerCase().trim();
  const n = String(name).toLowerCase().trim();
  return n === g || n.includes(g);
}

/** pushName limpio: devuelve el nombre solo si no parece el del grupo. */
function cleanPushName(pushName) {
  return pushName && !isGroupNameLike(pushName) ? pushName : null;
}

/** Registra un contacto en todos los maps disponibles */
function registerContact(c) {
  if (!c || !c.id) return;
  contactsMap[c.id] = c;

  // Extraer nombre del contacto (notify puede ser pushName envenenado → filtrar)
  const name = c.name || c.verifiedName || cleanPushName(c.notify);
  if (name && !isGroupNameLike(name)) {
    nameCache.set(c.id, name);
    // Si tiene LID, también mapear el LID
    if (c.lid) nameCache.set(c.lid, name);
  }

  // Reforzar mapeo LID → teléfono (el archivo manual se busca por teléfono)
  if (c.lid && c.jid) {
    lidToPhone[c.lid] = c.jid.split('@')[0];
  }

  // Si el ID es un LID y tiene nombre, registrar
  if (c.id.endsWith('@lid') && name) {
    nameCache.set(c.id, name);
  }
}

/** Carga el caché de nombres desde disco */
function loadNameCache() {
  const { data, corrupt } = loadJSONFile(NAME_CACHE_FILE);
  if (corrupt) {
    console.log('   Se regenerará el caché de nombres desde cero.');
  }
  if (!data) return;
  for (const [jid, name] of Object.entries(data)) {
    nameCache.set(jid, name);
  }
}

/** Guarda el caché de nombres a disco (escritura atómica) */
function saveNameCache() {
  const obj = {};
  for (const [jid, name] of nameCache) {
    obj[jid] = name;
  }
  atomicWriteFileSync(NAME_CACHE_FILE, JSON.stringify(obj, null, 2));
}

// ─── Caché persistente de mensajes ──────────────────────────────────────────

/** Carga los mensajes cacheados del grupo desde disco */
function loadCachedGroupMessages() {
  const { data, corrupt } = loadJSONFile(MSG_CACHE_FILE);
  if (corrupt) {
    console.log('⚠  El caché de comprobantes estaba dañado. Se respaldó el archivo original');
    console.log('   (podés recuperar datos manualmente desde ese respaldo) y se continúa con caché vacío.\n');
  }
  if (!data) return new Map();
  const map = new Map();
  for (const msg of data) {
    if (msg.key?.id) map.set(msg.key.id, msg);
    // Extraer pushNames de mensajes cacheados
    const participant = msg.key?.participant;
    if (participant && cleanPushName(msg.pushName)) {
      nameCache.set(participant, msg.pushName);
    }
  }
  return map;
}

/** Guarda mensajes del grupo a disco (merge con existentes, escritura atómica) */
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

  // No reescribir el archivo entero si no hubo cambios (evita ventana de corrupción innecesaria)
  if (added > 0) {
    atomicWriteFileSync(MSG_CACHE_FILE, JSON.stringify([...cached.values()]));
    console.log(`💾 Caché actualizado: ${cached.size} comprobantes guardados (+${added} nuevos)`);
  } else {
    console.log(`💾 Caché sin cambios: ${cached.size} comprobantes (no se reescribió el archivo)`);
  }
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
      if (participant && cleanPushName(msg.pushName)) {
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
      if (participant && cleanPushName(msg.pushName)) {
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
        // Enviar el QR también al navegador (como imagen)
        if (activePicker) activePicker.sendQR(qr);
      }
      if (connection === 'open') {
        console.log('✅ Conectado a WhatsApp.\n');
        safeResolve('open');
      }
      if (connection === 'close') {
        const code = lastDisconnect?.error?.output?.statusCode;
        if (code === DisconnectReason.loggedOut) {
          // Antes de reescanear QR, intentar una sola vez con el snapshot de la
          // última sesión buena: los loggedOut aislados suelen ser caídas de
          // WhatsApp, no revocación real. Si el snapshot falla, se borra y QR.
          if (!sessionRestoreAttempted && tryRestoreSessionSnapshot()) {
            sessionRestoreAttempted = true;
            console.log('♻  Sesión caída. Restaurando snapshot de la sesión anterior...\n');
            safeResolve('restart'); // el próximo intento usa la sesión restaurada
            return;
          }
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
  const { data } = loadJSONFile(CACHE_FILE);
  if (data) {
    const { groupJid, groupName } = data;
    knownGroupName = groupName || '';
    console.log(`📦 Grupo en caché: ${groupName}`);
    return groupJid;
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

  knownGroupName = match.subject || '';
  atomicWriteFileSync(CACHE_FILE, JSON.stringify({ groupJid: match.id, groupName: match.subject }));
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
      const existing = collected.get(msg.key.id);
      // Preferir la copia CON participant (los upsert en vivo la traen; las
      // copias del history sync pueden venir sin ella en sesiones nuevas)
      const hasSender = m => m?.key?.participant || m?.participant;
      if (!existing || (!hasSender(existing) && hasSender(msg))) {
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
// 6. DESCARGAR MEDIA  —  enfoque en dos fases para no desincronizar WhatsApp
// ════════════════════════════════════════════════════════════════════════════
//
//  Fase 1 — CDN directo (rápido, concurrente, SIN re-upload):
//    Descarga todo lo que aún está en los servidores de WhatsApp.
//    Sin reuploadRequest → no se envían solicitudes al remitente.
//    WhatsApp no detecta actividad sospechosa.
//
//  Fase 2 — Re-upload serial (lento, UNO POR UNO, 25-35s entre cada uno):
//    Solo para mensajes que fallaron en fase 1 con 404/410.
//    Imita a un humano dando "retry" manual en WhatsApp.
//    Un solo worker, delays largos con jitter aleatorio.
//    WhatsApp lo ve como comportamiento natural.

// ── Caché de fallos permanentes ──────────────────────────────────────────────
function loadFailedDownloads() {
  const { data } = loadJSONFile(FAILED_CACHE_FILE);
  if (!data) return new Set();
  return new Set(Object.keys(data));
}

function saveFailedDownload(msgKeyId, errorType) {
  const { data } = loadJSONFile(FAILED_CACHE_FILE);
  const existing = (data && typeof data === 'object') ? data : {};
  existing[msgKeyId] = { error: errorType, date: new Date().toISOString() };
  atomicWriteFileSync(FAILED_CACHE_FILE, JSON.stringify(existing, null, 2));
}

// ── Caché local de imágenes descargadas ──────────────────────────────────────
function getMediaCachePath(msgKeyId) {
  if (!fs.existsSync(MEDIA_CACHE_DIR)) {
    fs.mkdirSync(MEDIA_CACHE_DIR, { recursive: true });
  }
  return `${MEDIA_CACHE_DIR}/${msgKeyId}.jpg`;
}

function loadCachedMedia(msgKeyId) {
  const path = getMediaCachePath(msgKeyId);
  if (fs.existsSync(path)) {
    try {
      const buffer = fs.readFileSync(path);
      // Validar que no sea un JPEG truncado por un crash a mitad de escritura
      if (isValidJpeg(buffer)) return buffer;
      fs.rmSync(path, { force: true });
      console.log(`⚠ Imagen corrupta descartada y será re-descargada: ${msgKeyId}`);
    } catch { /* ilegible → tratar como no existente */ }
  }
  return null;
}

function saveCachedMedia(msgKeyId, buffer) {
  try {
    atomicWriteFileSync(getMediaCachePath(msgKeyId), buffer);
  } catch { /* ignorar errores de escritura */ }
}

// ── Clasificación de errores ─────────────────────────────────────────────────
function isPermanentError(errMsg) {
  const msg = String(errMsg || '');
  const permanent = [
    'DECRYPTION_ERROR',
    'no content in message',
    'is not a media message',
    'No message present',
    'Invalid media type',
  ];
  return permanent.some(p => msg.includes(p));
}

function isExpiredMediaError(errMsg) {
  const msg = String(errMsg || '');
  // Errores que indican que la media ya no está en CDN y necesita re-upload
  return (
    msg.includes('NOT_FOUND') ||
    msg.includes('GENERAL_ERROR') ||
    msg.includes('TIMEOUT') ||
    msg.includes('Connection Closed') ||
    msg.includes('ECONNRESET') ||
    msg.includes('ETIMEDOUT') ||
    msg.includes('status code 404') ||
    msg.includes('status code 410')
  );
}

// ── Descarga SIN re-upload (solo CDN) ───────────────────────────────────────
// Rápida y segura para paralelizar. Si la media expiró, lanza error.
async function downloadFromCdn(msg, timeoutMs = 20_000) {
  const msgKeyId = msg.key?.id || 'unknown';

  const cached = loadCachedMedia(msgKeyId);
  if (cached) return { buffer: cached, source: 'cache' };

  try {
    const buffer = await Promise.race([
      // ⚠ Sin ctx → sin reuploadRequest. Si 404/410, tira error enseguida.
      downloadMediaMessage(msg, 'buffer', {}),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error('TIMEOUT')), timeoutMs)
      ),
    ]);

    saveCachedMedia(msgKeyId, buffer);
    return { buffer, source: 'cdn' };
  } catch (err) {
    return { buffer: null, error: err.message || String(err), source: 'error' };
  }
}

// ── Descarga CON re-upload (serial, un mensaje a la vez) ─────────────────────
// Lenta. Solo se usa en fase 2, de a uno, con delays largos.
async function downloadWithReupload(sock, msg, timeoutMs = 40_000) {
  const msgKeyId = msg.key?.id || 'unknown';

  try {
    const buffer = await Promise.race([
      downloadMediaMessage(
        msg,
        'buffer',
        {},
        {
          logger: pino({ level: 'silent' }),
          reuploadRequest: sock.updateMediaMessage,
        }
      ),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error('TIMEOUT')), timeoutMs)
      ),
    ]);

    saveCachedMedia(msgKeyId, buffer);
    return { buffer, source: 'reupload' };
  } catch (err) {
    const errMsg = err.message || String(err);

    // Errores de autenticación / sesión dañada → abortar toda la fase 2
    if (
      errMsg.includes('Unsupported state') ||
      errMsg.includes('unable to authentic') ||
      errMsg.includes('unauthorized') ||
      errMsg.includes('Unauthorized') ||
      errMsg.includes('loggedOut') ||
      errMsg.includes('session') ||
      String(err.output?.statusCode) === '401'
    ) {
      return { buffer: null, error: errMsg, source: 'auth_error' };
    }

    // NOT_FOUND genuino tras re-upload: guardar como fallo permanente
    if (errMsg.includes('NOT_FOUND') || errMsg.includes('DECRYPTION_ERROR')) {
      saveFailedDownload(msgKeyId, errMsg);
      return { buffer: null, error: errMsg, source: 'permanent' };
    }
    return { buffer: null, error: errMsg, source: 'error' };
  }
}

// ── FASE 1: Descarga masiva desde CDN (concurrente, sin re-upload) ───────────
async function phase1_downloadFromCdn(messages, concurrency = 2) {
  const failedDownloads = loadFailedDownloads();
  const total = messages.length;
  const receipts = [];
  const stats = {
    total,
    ok: 0,
    fromCache: 0,
    expired: 0,    // necesita re-upload (404/410/timeout)
    failed: 0,     // error permanente o no imagen
  };
  // Diagnóstico de remitentes (Windows reportaba todo como un solo mensajero)
  senderDiag = { sin: 0, con: 0, propios: 0 };
  const expiredMessages = [];  // mensajes que necesitan fase 2

  let completed = 0;
  let index = 0;
  let lastProgressLine = '';

  const writeProgress = () => {
    const active = Math.min(concurrency, total - completed);
    const line = `\r  [${completed + active}/${total}] ⬇ CDN... `
      + `(${stats.ok} OK, ${stats.fromCache} caché, ${stats.expired} expirados, ${stats.failed} fallos)   `;
    if (line !== lastProgressLine) {
      process.stdout.write(line);
      lastProgressLine = line;
    }
  };

  const worker = async () => {
    while (index < total) {
      const i = index++;
      const msg = messages[i];
      const msgKeyId = msg.key?.id || `unknown-${i}`;

      // Saltar fallos permanentes previos
      if (failedDownloads.has(msgKeyId)) {
        stats.failed++;
        completed++;
        writeProgress();
        continue;
      }

      const result = await downloadFromCdn(msg);

      if (!result.buffer) {
        if (isPermanentError(result.error)) {
          stats.failed++;
          process.stdout.write(`\n  ⚠ [${i + 1}/${total}] Error perm.: ${String(result.error).slice(0, 50)}`);
        } else if (isExpiredMediaError(result.error)) {
          stats.expired++;
          expiredMessages.push(msg);
        } else {
          stats.failed++;
          process.stdout.write(`\n  ⚠ [${i + 1}/${total}] Error: ${String(result.error).slice(0, 50)}`);
        }
        completed++;
        writeProgress();
        continue;
      }

      // Validar mimetype
      const mimetype =
        msg.message?.imageMessage?.mimetype ??
        msg.message?.documentMessage?.mimetype ?? '';

      if (!mimetype.startsWith('image/')) {
        stats.failed++;
        completed++;
        writeProgress();
        continue;
      }

      if (!(msg.key?.participant || msg.participant)) senderDiag.sin++;
      else senderDiag.con++;
      if (msg.key?.fromMe) senderDiag.propios++;

      receipts.push({
        imageBuffer: result.buffer,
        senderName: getSenderName(msg, {}),
        date: new Date(toUnix(msg.messageTimestamp) * 1000).toLocaleString('es-CO'),
      });

      if (result.source === 'cache') {
        stats.fromCache++;
      } else {
        stats.ok++;
      }
      completed++;
      writeProgress();
    }
  };

  const workers = [];
  for (let w = 0; w < concurrency && w < total; w++) {
    workers.push(worker());
  }
  await Promise.all(workers);

  process.stdout.write('\n');
  return { receipts, expiredMessages, stats };
}

// ── FASE 2: Re-upload serial (UNO POR UNO, delays largos) ───────────────────
// Respeta a WhatsApp: un solo worker, 25-35s entre cada solicitud,
// imita a un humano reintentando manualmente.
async function phase2_reuploadExpired(sock, expiredMessages) {
  const total = expiredMessages.length;
  if (total === 0) return { receipts: [], stats: { ok: 0, failed: 0, permanent: 0 } };

  console.log(`\n  📡 Fase 2: Re-upload serial de ${total} mensajes expirados`);
  console.log(`     (1 cada ~30s — imitando reintento manual humano)\n`);

  const receipts = [];
  const stats = { ok: 0, failed: 0, permanent: 0 };
  const failedDownloads = loadFailedDownloads();

  for (let i = 0; i < total; i++) {
    const msg = expiredMessages[i];
    const msgKeyId = msg.key?.id || `unknown-${i}`;
    const shortId = msgKeyId.slice(0, 14);

    // Saltar fallos permanentes previos
    if (failedDownloads.has(msgKeyId)) {
      stats.permanent++;
      process.stdout.write(`  [${i + 1}/${total}] ⏭ ${shortId} — fallo permanente previo\n`);
      continue;
    }

    process.stdout.write(`  [${i + 1}/${total}] 🔄 ${shortId} — solicitando re-upload...`);

    const result = await downloadWithReupload(sock, msg);

    if (result.buffer) {
      const mimetype =
        msg.message?.imageMessage?.mimetype ??
        msg.message?.documentMessage?.mimetype ?? '';

      if (mimetype.startsWith('image/')) {
        if (!(msg.key?.participant || msg.participant)) senderDiag.sin++;
        else senderDiag.con++;
        if (msg.key?.fromMe) senderDiag.propios++;

        receipts.push({
          imageBuffer: result.buffer,
          senderName: getSenderName(msg, {}),
          date: new Date(toUnix(msg.messageTimestamp) * 1000).toLocaleString('es-CO'),
        });
        stats.ok++;
        process.stdout.write(` ✅\n`);
      } else {
        stats.failed++;
        process.stdout.write(` ⚠ no es imagen\n`);
      }
    } else if (result.source === 'auth_error') {
      // Error de sesión/autenticación: abortar toda la fase 2
      stats.failed++;
      const remaining = total - i - 1;
      process.stdout.write(` 🔒 Sesión dañada — abortando fase 2 (${remaining} pendientes)\n`);
      console.log(`\n  ⚠ La sesión de WhatsApp necesita reautenticación.`);
      console.log(`  💡 Borra la carpeta "baileys_auth/" y vuelve a ejecutar.\n`);
      break;
    } else {
      if (result.source === 'permanent') {
        stats.permanent++;
        process.stdout.write(` ❌ ${String(result.error).slice(0, 40)}\n`);
      } else {
        stats.failed++;
        process.stdout.write(` ⚠ ${String(result.error).slice(0, 40)}\n`);
      }
    }

    // Delay entre re-uploads: 25-35s con jitter aleatorio.
    // Esto es CLAVE para que WhatsApp no detecte automatización.
    if (i < total - 1) {
      const delay = 25_000 + Math.floor(Math.random() * 10_000); // 25-35s
      const countdown = Math.ceil(delay / 1000);
      process.stdout.write(`\r  ⏳ Próximo re-upload en ${countdown}s...   `);
      await sleep(delay);
      process.stdout.write('\r' + ' '.repeat(45) + '\r');
    }
  }

  process.stdout.write('\n');
  return { receipts, stats };
}

// ── Orquestador de dos fases ─────────────────────────────────────────────────
async function downloadAllTwoPhase(messages, sock) {
  const failedDownloads = loadFailedDownloads();

  // Fase 1: CDN directo, concurrente, sin re-upload
  console.log(`\n╔══════════════════════════════════════════════════════════╗`);
  console.log(`║  FASE 1/2 — Descarga directa desde CDN (${MAX_CONCURRENT_DOWNLOADS} workers)       ║`);
  console.log(`║  Sin solicitar re-upload al remitente                    ║`);
  console.log(`╚══════════════════════════════════════════════════════════╝\n`);

  const phase1 = await phase1_downloadFromCdn(messages, MAX_CONCURRENT_DOWNLOADS);

  console.log(`\n  📊 Fase 1 completada:`);
  console.log(`     ✅ CDN / caché : ${phase1.stats.ok + phase1.stats.fromCache}`);
  console.log(`     📡 Expirados   : ${phase1.stats.expired} (necesitan re-upload)`);
  console.log(`     ❌ Fallos      : ${phase1.stats.failed}`);

  // Fase 2: Re-upload serial, uno por uno
  if (phase1.expiredMessages.length > 0) {
    console.log(`\n╔══════════════════════════════════════════════════════════╗`);
    console.log(`║  FASE 2/2 — Re-upload serial (1 cada ~30s)              ║`);
    console.log(`║  Imitando comportamiento humano para evitar baneo       ║`);
    console.log(`╚══════════════════════════════════════════════════════════╝`);

    const phase2 = await phase2_reuploadExpired(sock, phase1.expiredMessages);

    console.log(`  📊 Fase 2 completada:`);
    console.log(`     ✅ Recuperados  : ${phase2.stats.ok}`);
    console.log(`     ❌ No disponible: ${phase2.stats.permanent}`);
    console.log(`     ⚠ Otros fallos  : ${phase2.stats.failed}`);

    console.log(`  📊 Remitentes: ${senderDiag.con} con participant | ${senderDiag.sin} SIN participant | ${senderDiag.propios} propios`);

    // Juntar receipts
    const allReceipts = [...phase1.receipts, ...phase2.receipts];

    return {
      receipts: allReceipts,
      stats: {
        total: messages.length,
        ok: phase1.stats.ok + phase2.stats.ok,
        fromCache: phase1.stats.fromCache,
        expired: phase1.stats.expired,
        reuploadOk: phase2.stats.ok,
        reuploadFailed: phase2.stats.failed + phase2.stats.permanent,
        failed: phase1.stats.failed,
        skippedPermanent: failedDownloads.size,
      },
    };
  }

  console.log(`  📊 Remitentes: ${senderDiag.con} con participant | ${senderDiag.sin} SIN participant | ${senderDiag.propios} propios`);

  return {
    receipts: phase1.receipts,
    stats: {
      total: messages.length,
      ok: phase1.stats.ok,
      fromCache: phase1.stats.fromCache,
      expired: 0,
      reuploadOk: 0,
      reuploadFailed: 0,
      failed: phase1.stats.failed,
      skippedPermanent: failedDownloads.size,
    },
  };
}

// ════════════════════════════════════════════════════════════════════════════
// 7. NOMBRE DEL REMITENTE
// ════════════════════════════════════════════════════════════════════════════
function getSenderName(msg, contacts) {
  // El remitente puede venir en key.participant (normal) o en el campo
  // top-level `participant` (algunos syncs en Windows lo mandan ahí).
  const participant = msg.key?.participant || msg.participant;

  // Sin participant (history sync): el pushName limpio puede identificar al
  // remitente; si no hay, es mensaje propio → nombre propio (nunca el grupo).
  if (!participant) return cleanPushName(msg.pushName) || ownName || 'YO';

  const jid = participant;

  // 1. Buscar en nombres manuales (archivo nombres_mensajeros.json)
  //    El jid puede ser un LID → convertir a teléfono → buscar nombre manual
  const phone = lidToPhone[jid] || jid.split('@')[0];
  if (manualNames[phone]) return manualNames[phone];
  //    También admitir entradas con clave LID directa (remitentes sin teléfono)
  if (manualNames[jid]) return manualNames[jid];

  // 2. Buscar en contactos sincronizados y caché (nunca el nombre del grupo)
  const c = contacts[jid] || contactsMap[jid];
  const cached = nameCache.get(jid);
  const name =
    c?.name ||
    c?.verifiedName ||
    cleanPushName(msg.pushName) ||
    (cached && !isGroupNameLike(cached) ? cached : null) ||
    cleanPushName(c?.notify);
  if (name) return name;

  // 3. Mostrar el número de teléfono (más útil que el LID)
  return phone;
}

// ── Archivo de nombres (v2: también persiste el mapeo LID → teléfono) ───────
// El mapeo vivo sale de la metadata del grupo; persistirlo en el mismo JSON
// hace que funcione también en máquinas donde la metadata no trae teléfonos
// (sesiones nuevas, como en Windows).
function loadManualNames() {
  const { data, corrupt } = loadJSONFile(NOMBRES_FILE);
  if (corrupt) {
    console.log('⚠  nombres_mensajeros.json estaba dañado. Se respaldó el archivo original');
    console.log('   — tus nombres siguen en el respaldo y se pueden restaurar manualmente.');
  }
  const obj = (data && typeof data === 'object') ? data : {};
  for (const [k, v] of Object.entries(obj)) {
    if (k === '_lid_to_phone' && v && typeof v === 'object') {
      Object.assign(lidToPhone, v);
    } else if (typeof v === 'string') {
      manualNames[k] = v;
    }
  }
}

function saveManualNames() {
  atomicWriteFileSync(NOMBRES_FILE, JSON.stringify({ ...manualNames, _lid_to_phone: lidToPhone }, null, 2));
}

async function reconnectWithRetries(maxRetries, phaseLabel) {
  let sock;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    console.log(`🔄 ${phaseLabel} (intento ${attempt}/${maxRetries})...\n`);
    sock = await createSocket();

    const result = await waitForOpen(sock);
    if (result === 'open') {
      currentSock = sock;
      ownName = sock.user?.name?.trim() || 'YO';
      backupSessionSnapshot(); // sesión OK → respaldo para recuperar sin reescanear QR
      return sock;
    }

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

  // Selector de fechas + estado en vivo: navegador por defecto;
  // --terminal conserva los prompts de texto sin servidor web.
  const useWebPicker = !process.argv.includes('--terminal');
  if (useWebPicker) {
    activePicker = await startControlServer();
    activePicker.attachConsole();   // tee de la consola → página (en vivo)
    activePicker.openBrowser();
    console.log(`📡 Panel abierto en el navegador (${activePicker.url})`);
    // Guardar nombres asignados desde el panel → archivo + mapa en vivo
    activePicker.setNameSaver((key, name) => {
      if (name) manualNames[key] = name;
      else delete manualNames[key];
      try { saveManualNames(); }
      catch (err) { activePicker.pushLog('⚠ No se pudo guardar nombres: ' + err.message); }
    });
    // Lista inicial de mensajeros desde el caché local (sin esperar conexión)
    pushEarlyParticipants();
  }
  const { startDate, endDate } = useWebPicker
    ? await activePicker.waitForRange()
    : await askDateRange();
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

  // ── Participantes: mapeo LID → teléfono + nombres para el panel web ───────
  let metadata = null;
  try {
    metadata = await sock.groupMetadata(groupJid);
    knownGroupName = metadata.subject || knownGroupName;

    // Traer contactos de la agenda YA (los nombres reales resuelven solos
    // los remitentes guardados; en sesiones nuevas puede llegar tarde o nunca)
    try { if (typeof sock.fetchContacts === 'function') await sock.fetchContacts(); } catch { /* opcional */ }
  } catch (err) {
    console.log(`  ⚠ No se pudo obtener metadata del grupo: ${err.message}`);
  }

  if (metadata) {
    // Primero el mapeo persistido del archivo, después el fresco de metadata
    loadManualNames();
    for (const p of metadata.participants) {
      const phoneLid = p.lid || p.id;
      const phoneJid = p.jid || p.id;
      const phoneNum = phoneJid.split('@')[0];

      // Mapear LID → número de teléfono
      if (phoneLid) lidToPhone[phoneLid] = phoneNum;
      if (p.id) lidToPhone[p.id] = phoneNum;
    }

    // Entradas vacías para participantes nuevos (el panel permite nombrarlos)
    let newEntries = 0;
    for (const p of metadata.participants) {
      const phoneNum = (p.jid || p.id).split('@')[0];
      if (!manualNames[phoneNum] && !manualNames[p.id] && !manualNames[p.id?.split('@')[0]]) {
        manualNames[phoneNum] = '';
        newEntries++;
      }
    }
    const firstRun = !fs.existsSync(NOMBRES_FILE);
    if (firstRun || newEntries > 0) saveManualNames();

    if (firstRun) {
      console.log(`\n  ╔══════════════════════════════════════════════════════╗`);
      console.log(`  ║  📋 ARCHIVO DE NOMBRES CREADO                        ║`);
      console.log(`  ║                                                      ║`);
      console.log(`  ║  Asigná los nombres desde el panel web (👥 Mensajeros)║`);
      console.log(`  ║  o editando: nombres_mensajeros.json                 ║`);
      console.log(`  ╚══════════════════════════════════════════════════════╝\n`);
    } else if (newEntries > 0) {
      console.log(`   📝 ${newEntries} nuevos participantes agregados al archivo de nombres`);
    }

    // Contar nombres asignados
    const assigned = Object.values(manualNames).filter(n => n).length;
    console.log(`👥 Participantes: ${metadata.participants.length}, nombres asignados: ${assigned}/${Object.keys(manualNames).length}`);

    // Panel web: fusionar metadata con la lista temprana del caché
    for (const p of metadata.participants) {
      const lid = p.lid || p.id;
      const phone = p.jid ? p.jid.split('@')[0] : null;
      participantLids.add(lid);
      const existing = participantList.find(e => e.lid === lid);
      const entry = senderToEntry(lid, phone);
      if (existing) {
        existing.phone = entry.phone || existing.phone;
        existing.name = entry.name || existing.name;
      } else {
        participantList.push(entry);
      }
    }
    if (activePicker) activePicker.setParticipants(participantList);
  } else {
    // Sin metadata: cargar igual lo persistido para que el mapeo siga funcionando
    loadManualNames();
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

  // Sumar al panel los remitentes que ya no están en la metadata del grupo
  if (activePicker && msgStore.size > 0) {
    let added = 0;
    for (const msg of msgStore.values()) {
      const lid = msg.key?.participant;
      if (lid && !participantLids.has(lid)) {
        participantLids.add(lid);
        const phone = lidToPhone[lid] || null;
        participantList.push({
          lid,
          phone,
          name: manualNames[phone] || manualNames[lid] || manualNames[lid.split('@')[0]] || '',
        });
        added++;
      }
    }
    if (added > 0) activePicker.setParticipants(participantList);
  }

  let healAttempt = 0;
  while (shouldAutoHealSession(collection.stats) && healAttempt < MAX_AUTO_HEAL_RETRIES) {
    healAttempt++;
    console.log('\n🛠 Modo autorreparación: detectamos sesión inestable o sin sincronización de historial.');
    console.log('   Se renovará la sesión automáticamente para que no tengas que borrar carpetas manualmente.\n');

    try { await sock.end(); } catch { /* ignore */ }
    // Primero intentar restaurar el snapshot de la última sesión buena
    // (evita reescanear QR); si no hay snapshot, respaldar la sesión actual
    // y arrancar limpio.
    if (!sessionRestoreAttempted && tryRestoreSessionSnapshot()) {
      sessionRestoreAttempted = true;
      console.log('♻  Restaurando snapshot de la sesión anterior para reintentar...');
    } else {
      try {
        const backup = `${AUTH_FOLDER}_backup_${Date.now()}`;
        fs.renameSync(AUTH_FOLDER, backup);
        console.log(`📦 Sesión anterior respaldada en: ${backup}`);
      } catch { /* si no existe, no hay nada que respaldar */ }
    }
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

  // ── Cargar fallos permanentes previos ──────────────────────────────────
  const failedDownloads = loadFailedDownloads();
  if (failedDownloads.size > 0) {
    console.log(`📋 ${failedDownloads.size} mensajes con fallos permanentes previos — se omitirán`);
  }

  // ── Filtrar mensajes ya fallidos del caché y limpiar ────────────────────
  if (failedDownloads.size > 0) {
    let cleaned = 0;
    for (const [key] of allCached) {
      if (failedDownloads.has(key)) {
        allCached.delete(key);
        cleaned++;
      }
    }
    if (cleaned > 0) {
      atomicWriteFileSync(MSG_CACHE_FILE, JSON.stringify([...allCached.values()]));
      console.log(`🧹 ${cleaned} mensajes fallidos eliminados del caché de grupo`);
    }
  }

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
    if (activePicker) activePicker.pushLog('⚠ No se encontraron comprobantes en el rango seleccionado.');
    try { await sock.end(); } catch { /* ignore */ }
    await flushPickerBeforeExit();
    process.exit(0);
  }

  // ── Descarga en dos fases ──────────────────────────────────────────────
  const cacheHits = rawMsgs.filter(m => loadCachedMedia(m.key?.id)).length;
  if (cacheHits > 0) {
    console.log(`📦 ${cacheHits} imágenes ya en caché local`);
  }

  const { receipts, stats: dlStats } = await downloadAllTwoPhase(rawMsgs, sock);

  // ── Resumen de descargas ────────────────────────────────────────────────
  console.log(`\n${'─'.repeat(55)}`);
  console.log(`  📊 RESUMEN FINAL`);
  console.log(`${'─'.repeat(55)}`);
  console.log(`  Total en rango        : ${dlStats.total}`);
  console.log(`  ✅ Fase 1 (CDN/caché)  : ${dlStats.ok + dlStats.fromCache}`);
  console.log(`  📡 Fase 2 (re-upload)  : ${dlStats.reuploadOk} recuperados`);
  console.log(`  ⏭ Fallos perm. previos : ${dlStats.skippedPermanent} (omitidos)`);
  console.log(`  ❌ No disponibles      : ${dlStats.reuploadFailed}`);
  console.log(`  🚫 Otros fallos        : ${dlStats.failed}`);
  console.log(`  ✅ Imágenes válidas    : ${receipts.length}`);
  console.log(`${'─'.repeat(55)}\n`);

  try { await sock.end(); } catch { /* ignore */ }

  if (receipts.length === 0) {
    console.log('⚠  Nada que exportar.\n');
    if (activePicker) activePicker.pushLog('⚠ Nada que exportar.');
    await flushPickerBeforeExit();
    process.exit(0);
  }

  await createWordDocument(receipts);

  if (activePicker) {
    // Notificar al navegador para mostrar el botón de descarga
    activePicker.notifyFile(OUTPUT_FILE);
    activePicker.pushLog('✅ Proceso completo. Descargá el archivo desde el navegador.');

    // Mantener el servidor vivo hasta que descargues (o 5 min sin hacerlo).
    // Si saliéramos enseguida, el navegador mostraría "comprobá tu conexión".
    const result = await activePicker.waitForDownloadOrTimeout(300_000);
    if (result === 'downloaded') {
      activePicker.pushLog('👋 Descarga recibida. Cerrando aplicación...');
      await sleep(2_000); // margen para descargas repetidas
    } else {
      activePicker.pushLog('👋 Sin descarga en 5 minutos. Cerrando aplicación...');
    }
  }
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
  atomicWriteFileSync(OUTPUT_FILE, buffer);

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

// ── Shutdown seguro ───────────────────────────────────────────────────────────
// Cerrar el socket antes de salir (Ctrl+C / kill) evita que baileys quede
// a mitad de una escritura de credenciales y corrompa la sesión.
let currentSock = null;

// Panel de control web (null en modo --terminal)
let activePicker = null;

/** Da unos segundos al navegador para recibir los últimos eventos SSE antes de salir. */
async function flushPickerBeforeExit() {
  if (!activePicker) return;
  await sleep(1_500);
}

if (require.main === module) {
  const handleShutdown = (signal) => {
    console.log(`\n⚠ ${signal} recibido. Cerrando sesión de forma segura...`);
    const done = () => process.exit(0);
    if (currentSock) {
      Promise.race([currentSock.end(), sleep(5_000)]).catch(() => {}).finally(done);
    } else {
      done();
    }
  };
  process.on('SIGINT', () => handleShutdown('SIGINT'));
  process.on('SIGTERM', () => handleShutdown('SIGTERM'));

  main().catch(err => {
    console.error('\n❌ Error fatal:', err.message);
    process.exit(1);
  });
}

// Exportado para pruebas unitarias de los helpers de caché
module.exports = {
  setCurrentSock: sock => { currentSock = sock; },
  loadCachedGroupMessages,
  saveCachedGroupMessages,
  loadNameCache,
  saveNameCache,
  loadFailedDownloads,
  saveFailedDownload,
  shouldAutoHealSession,
  backupSessionSnapshot,
  tryRestoreSessionSnapshot,
  _setBadMacCount: n => { badMacTracker.count = n; }, // hook solo para tests
  _setKnownGroupName: n => { knownGroupName = n; },   // hook solo para tests
  _setOwnName: n => { ownName = n; },                 // hook solo para tests
  _setManualName: (phone, name) => { manualNames[phone] = name; }, // hook solo para tests
  _setLidToPhone: (lid, phone) => { lidToPhone[lid] = phone; },    // hook solo para tests
  cleanPushName,
  isGroupNameLike,
  getSenderName,
  loadManualNames,
  saveManualNames,
  pushEarlyParticipants,
  senderToEntry,
  _resetNamesState: () => {
    Object.keys(manualNames).forEach(k => delete manualNames[k]);
    Object.keys(lidToPhone).forEach(k => delete lidToPhone[k]);
    participantList = [];
    participantLids.clear();
  }, // hook solo para tests
  _setActivePicker: p => { activePicker = p; }, // hook solo para tests
};