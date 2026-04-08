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
const GROUP_NAME  = 'TRANSFERENCIAS RED POSTAL POBLADO';
const AUTH_FOLDER = './baileys_auth';
const CACHE_FILE  = './group_cache.json';
const OUTPUT_FILE = 'Comprobantes_Descargados.docx';

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Reemplaza makeInMemoryStore: guardamos mensajes en un Map simple
const msgStore = new Map(); // `${jid}:${id}` → msg

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

  // Guardar mensajes en el store manual
  sock.ev.on('messages.upsert', ({ messages }) => {
    for (const msg of messages) {
      if (msg.key?.remoteJid && msg.key?.id) {
        msgStore.set(`${msg.key.remoteJid}:${msg.key.id}`, msg);
      }
    }
  });

  sock.ev.on('messaging-history.set', ({ messages }) => {
    for (const msg of messages) {
      if (msg.key?.remoteJid && msg.key?.id) {
        msgStore.set(`${msg.key.remoteJid}:${msg.key.id}`, msg);
      }
    }
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

    sock.ev.on('connection.update', ({ connection, lastDisconnect, qr }) => {
      if (qr) {
        console.log('\n══════════════════════════════════════════════════════');
        console.log('  ESCANEA ESTE CÓDIGO QR CON WHATSAPP');
        console.log('══════════════════════════════════════════════════════\n');
        QRCode.toString(qr, { type: 'terminal', small: true }, (err, url) => {
          if (!err) console.log(url);
        });
      }
      if (connection === 'open') {
        clearTimeout(timer);
        console.log('✅ Conectado a WhatsApp.\n');
        resolve('open');
      }
      if (connection === 'close') {
        clearTimeout(timer);
        const code = lastDisconnect?.error?.output?.statusCode;
        if (code === DisconnectReason.loggedOut) {
          console.log('\n⚠  Sesión cerrada. Elimina "baileys_auth" y reescanea el QR.\n');
          fs.rmSync(AUTH_FOLDER, { recursive: true, force: true });
          resolve('loggedOut');
          return;
        }
        // 515 = restartRequired: WhatsApp pide reconectar
        if (code === 515 || code === DisconnectReason.restartRequired) {
          console.log('⚠  WhatsApp pidió reconexión (515). Reintentando...\n');
          resolve('restart');
          return;
        }
        console.log(`⚠  Conexión cerrada (código ${code}). Reintentando...\n`);
        resolve('restart');
      }
    });
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

function collectMessages(sock, groupJid, startTs, endTs) {
  return new Promise(resolve => {
    const collected = new Map();
    let idleTimer;
    let globalTimer;
    let finished = false;
    let historyChunks = 0;
    let totalMsgsReceived = 0;
    let totalGroupMsgs = 0;

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
      resolve([...collected.values()]);
    };

    const resetIdle = () => {
      clearTimeout(idleTimer);
      // Esperar 30s en lugar de 15s para dar tiempo a la sincronización
      idleTimer = setTimeout(() => {
        process.stdout.write('\n  ⏱ Sin más mensajes entrantes. Continuando...\n');
        finish();
      }, 30_000);
    };

    const isMedia = msg =>
      !!msg.message?.imageMessage || !!msg.message?.documentMessage;

    const addMsg = (msg, source) => {
      totalMsgsReceived++;
      if (!msg.key?.remoteJid) return;
      if (msg.key.remoteJid !== groupJid) return;
      totalGroupMsgs++;

      const ts = toUnix(msg.messageTimestamp);

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
    }, 180_000);

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
  const jid    = msg.key.participant ?? msg.key.remoteJid;
  const number = jid.split('@')[0];
  const c      = contacts[jid];
  return c?.name || c?.notify || number;
}

// ════════════════════════════════════════════════════════════════════════════
// 8. MAIN
// ════════════════════════════════════════════════════════════════════════════
async function main() {
  const { startDate, endDate } = await askDateRange();
  const startTs = Math.floor(startDate.getTime() / 1000);
  const endTs   = Math.floor(endDate.getTime()   / 1000);

  // ── Conexión con reintentos automáticos ───────────────────────────────────
  let sock;
  let contacts = {};
  const MAX_RETRIES = 5;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    console.log(`🔄 Conectando con WhatsApp (intento ${attempt}/${MAX_RETRIES})...\n`);
    contacts = {};
    sock = await createSocket();

    sock.ev.on('contacts.set',    ({ contacts: list }) => list.forEach(c => { contacts[c.id] = c; }));
    sock.ev.on('contacts.upsert', list               => list.forEach(c => { contacts[c.id] = c; }));

    const result = await waitForOpen(sock);

    if (result === 'open')      break;
    if (result === 'loggedOut') process.exit(1);

    // restart / retry: destruir socket y volver a intentar
    try { await sock.end(); } catch { /* ignore */ }
    if (attempt === MAX_RETRIES) {
      console.error('❌ No se pudo conectar después de varios intentos.');
      process.exit(1);
    }
    await sleep(3_000);
  }

  const groupJid = await findGroupJid(sock);

  console.log('⏳ Esperando sincronización de historial...');
  console.log('   (30-90 s la primera vez; más rápido en ejecuciones siguientes)');
  console.log(`   🔎 Rango de búsqueda: startTs=${startTs} endTs=${endTs}`);
  console.log(`   🔎 Inicio: ${new Date(startTs * 1000).toLocaleString('es-CO')}`);
  console.log(`   🔎 Fin:    ${new Date(endTs * 1000).toLocaleString('es-CO')}`);
  console.log(`   🔎 Mensajes en store antes de recolectar: ${msgStore.size}\n`);

  // Manejar cierre de conexión durante la recolección
  let connectionLost = false;
  const onClose = ({ connection }) => {
    if (connection === 'close') {
      connectionLost = true;
      console.log('\n  ⚠ Conexión cerrada durante sincronización, procesando lo recolectado...');
    }
  };
  sock.ev.on('connection.update', onClose);

  const rawMsgs = await collectMessages(sock, groupJid, startTs, endTs);

  sock.ev.off('connection.update', onClose);

  console.log(`\n📊 Total comprobantes en el rango: ${rawMsgs.length}`);
  console.log(`📊 Mensajes totales en store: ${msgStore.size}`);

  if (rawMsgs.length === 0) {
    // Intentar buscar directamente en el store como último recurso
    console.log('\n🔄 Verificando store completo por si los mensajes se etiquetaron diferente...');
    let storeGroupMsgs = 0;
    let storeGroupMedia = 0;
    for (const [, msg] of msgStore.entries()) {
      if (msg.key?.remoteJid === groupJid) {
        storeGroupMsgs++;
        const ts = toUnix(msg.messageTimestamp);
        const hasImage = !!msg.message?.imageMessage;
        const hasDoc = !!msg.message?.documentMessage;
        if (hasImage || hasDoc) storeGroupMedia++;
        // Mostrar algunos mensajes del grupo para diagnóstico
        if (storeGroupMsgs <= 5) {
          const date = new Date(ts * 1000);
          const types = Object.keys(msg.message || {}).join(', ');
          console.log(`   Msg ${storeGroupMsgs}: ts=${ts} (${date.toLocaleString('es-CO')}) tipos=[${types}]`);
        }
      }
    }
    console.log(`   Total mensajes del grupo en store: ${storeGroupMsgs}`);
    console.log(`   Mensajes con media en store:       ${storeGroupMedia}`);

    console.log('\n⚠  No se encontraron imágenes en el rango seleccionado.');
    console.log('   Posibles causas:');
    console.log('   1. El historial no se sincronizó completamente.');
    console.log('   2. Las fechas seleccionadas no coinciden con los mensajes.');
    console.log('   → Intenta borrar la carpeta "baileys_auth" y reescanea el QR.\n');
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