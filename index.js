const fs = require('fs');
const qrcode = require('qrcode-terminal');
const { Client, LocalAuth } = require('whatsapp-web.js');
const { prompt } = require('enquirer');

// Configuración de reintentos
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 5000;
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  VerticalAlign,
  AlignmentType,
  Header
} = require('docx');

const GROUP_NAME = 'TRANSFERENCIAS RED POSTAL POBLADO';

// Función auxiliar para esperar
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Función para crear el cliente con la configuración correcta
function createClient() {
  return new Client({
    authStrategy: new LocalAuth({ dataPath: './.wwebjs_auth' }),
    webVersionCache: {
      type: 'remote',
      remotePath: 'https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.2412.54.html',
    },
    puppeteer: {
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--disable-gpu',
        '--disable-extensions',
        '--disable-background-networking',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-breakpad',
        '--disable-component-extensions-with-background-pages',
        '--disable-features=TranslateUI',
        '--disable-ipc-flooding-protection',
        '--disable-renderer-backgrounding',
        '--enable-features=NetworkService,NetworkServiceInProcess',
        '--force-color-profile=srgb',
        '--metrics-recording-only'
      ],
      protocolTimeout: 600000,
      timeout: 300000
    }
  });
}

async function initializeClientWithRetry(client, startDate, endDate, attempt = 1) {
  return new Promise((resolve, reject) => {
    let initialized = false;
    let timeoutId;
    
    const cleanup = () => {
      if (timeoutId) clearTimeout(timeoutId);
      client.removeAllListeners('qr');
      client.removeAllListeners('ready');
      client.removeAllListeners('auth_failure');
      client.removeAllListeners('disconnected');
    };

    client.on('qr', (qr) => {
      console.log('\n======================================================');
      console.log(' POR FAVOR ESCANEA EL SIGUIENTE CÓDIGO QR CON WHATSAPP');
      console.log('======================================================\n');
      qrcode.generate(qr, { small: true });
    });

    client.on('auth_failure', (msg) => {
      console.error('❌ Error de autenticación:', msg);
      cleanup();
      reject(new Error('Authentication failed'));
    });

    client.on('disconnected', (reason) => {
      console.log('⚠ Cliente desconectado:', reason);
      if (!initialized) {
        cleanup();
        reject(new Error(`Disconnected: ${reason}`));
      }
    });

    client.on('ready', async () => {
      initialized = true;
      cleanup();
      console.log('✅ Cliente de WhatsApp listo y conectado.');
      try {
        await processGroupMessages(client, startDate, endDate);
        resolve();
      } catch (error) {
        reject(error);
      }
    });

    // Timeout de seguridad para la inicialización (3 minutos)
    timeoutId = setTimeout(() => {
      if (!initialized) {
        cleanup();
        reject(new Error('Timeout durante la inicialización'));
      }
    }, 180000);

    console.log(`🔄 Iniciando cliente de WhatsApp (intento ${attempt}/${MAX_RETRIES})...`);
    
    client.initialize().catch(err => {
      cleanup();
      reject(err);
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCIÓN PRINCIPAL - Parseo de fechas CORREGIDO
// ─────────────────────────────────────────────────────────────────────────────
async function main() {
  console.log('\n======================================================');
  console.log(' CONFIGURACIÓN DE FECHAS DE DESCARGA');
  console.log('======================================================');
  console.log('Ingresa la fecha en formato YYYY-MM-DD (ej: 2026-03-12).');
  console.log('Si dejas en blanco, usará AYER como inicio y HOY como fin.\n');

  const questions = [
    {
      type: 'input',
      name: 'startDate',
      message: 'Fecha de INICIO (YYYY-MM-DD) [Enter = AYER]:'
    },
    {
      type: 'input',
      name: 'startTime',
      message: 'Hora de INICIO (HH:MM) [Enter = 08:30]:'
    },
    {
      type: 'input',
      name: 'endDate',
      message: 'Fecha de FIN   (YYYY-MM-DD) [Enter = HOY]:'
    },
    {
      type: 'input',
      name: 'endTime',
      message: 'Hora de FIN    (HH:MM)      [Enter = 23:59]:'
    }
  ];

  const answers = await prompt(questions);

  // ── Parsear hora de inicio (por defecto 08:30) ──────────────────────────
  let startHour = 8;
  let startMinute = 30;
  if (answers.startTime && answers.startTime.trim() !== '') {
    const timeParts = answers.startTime.trim().split(':');
    // Usamos exactamente lo que el usuario escribe, sin sumar nada
    startHour   = parseInt(timeParts[0], 10);
    startMinute = parseInt(timeParts[1], 10);
    // Validación básica
    if (isNaN(startHour) || startHour < 0 || startHour > 23)   startHour   = 8;
    if (isNaN(startMinute) || startMinute < 0 || startMinute > 59) startMinute = 30;
  }

  // ── Fecha de INICIO ──────────────────────────────────────────────────────
  // Importante: construimos la fecha manualmente para evitar que setDate()
  // interactúe con los meses de forma inesperada cuando mezclamos setFullYear/setMonth/setDate.
  let startDate;
  if (answers.startDate && answers.startDate.trim() !== '') {
    const parts = answers.startDate.trim().split('-');
    const year  = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // 0-indexed
    const day   = parseInt(parts[2], 10);
    startDate = new Date(year, month, day, startHour, startMinute, 0, 0);
  } else {
    // AYER a la hora indicada
    const now = new Date();
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, startHour, startMinute, 0, 0);
  }

  // ── Parsear hora de FIN (por defecto 23:59) ──────────────────────────
  let endHour = 23;
  let endMinute = 59;
  if (answers.endTime && answers.endTime.trim() !== '') {
    const timeParts = answers.endTime.trim().split(':');
    endHour   = parseInt(timeParts[0], 10);
    endMinute = parseInt(timeParts[1], 10);
    if (isNaN(endHour)   || endHour   < 0 || endHour   > 23) endHour   = 23;
    if (isNaN(endMinute) || endMinute < 0 || endMinute > 59) endMinute = 59;
  }

  // ── Fecha de FIN ───────────────────────────────────────────────
  let endDate;
  if (answers.endDate && answers.endDate.trim() !== '') {
    const parts = answers.endDate.trim().split('-');
    const year  = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day   = parseInt(parts[2], 10);
    endDate = new Date(year, month, day, endHour, endMinute, 59, 999);
  } else {
    // HOY a la hora de fin indicada
    const now = new Date();
    endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), endHour, endMinute, 59, 999);
  }

  console.log(`\n📅 Rango seleccionado:`);
  console.log(`   INICIO: ${startDate.toLocaleString('es-CO')}`);
  console.log(`   FIN:    ${endDate.toLocaleString('es-CO')}\n`);

  let lastError;
  
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    let client;
    try {
      client = createClient();
      await initializeClientWithRetry(client, startDate, endDate, attempt);
      console.log('▶ Proceso finalizado exitosamente.');
      process.exit(0);
    } catch (error) {
      lastError = error;
      const isContextDestroyed = error.message && (
        error.message.includes('Execution context was destroyed') ||
        error.message.includes('Protocol error') ||
        error.message.includes('Target closed') ||
        error.message.includes('Session closed')
      );
      
      const isFrameDetached = error.message && (
        error.message.includes('Navigating frame was detached') ||
        error.message.includes('frame was detached') ||
        error.message.includes('frame detached')
      );
      
      const isRuntimeTimeout = error.message && (
        error.message.includes('runtime.callfuncuint timed out') ||
        error.message.includes('protocolTimeout') ||
        error.message.includes('timed out')
      );
      
      console.error(`\n❌ Error en intento ${attempt}/${MAX_RETRIES}:`, error.message);
      
      if (client) {
        try {
          await client.destroy();
        } catch (destroyErr) {
          // Ignorar errores al destruir
        }
      }
      
      if (attempt < MAX_RETRIES) {
        if (isContextDestroyed) {
          console.log(`\n⚠ Se detectó error de contexto destruido.`);
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        } else if (isFrameDetached) {
          console.log(`\n⚠ Se detectó error "frame detached" (común en Windows).`);
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        } else if (isRuntimeTimeout) {
          console.log(`\n⚠ Se detectó error de timeout.`);
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        } else {
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        }
        await sleep(RETRY_DELAY_MS);
      }
    }
  }
  
  console.error('\n❌ Se agotaron todos los intentos. Error final:', lastError?.message);
  console.log('\n💡 Sugerencias para resolver el problema:');
  console.log('   1. Elimina la carpeta .wwebjs_auth y vuelve a escanear el QR');
  console.log('   2. Asegúrate de tener una conexión a internet estable');
  console.log('   3. Intenta cerrar otras instancias de WhatsApp Web');
  console.log('   4. Reinicia la aplicación');
  console.log('   5. Usa Node.js versión 18.x o 20.x (versiones LTS recomendadas)\n');
  process.exit(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// BÚSQUEDA DE GRUPO — Con caché para acelerar ejecuciones a partir de la 2ª
// ─────────────────────────────────────────────────────────────────────────────
async function processGroupMessages(client, startDate, endDate) {
  const cacheFile = './group_cache.json';
  let group = null;

  // 1) Intentar cargar desde caché (mucho más rápido que getChats())
  if (fs.existsSync(cacheFile)) {
    console.log('📦 Cargando grupo desde caché...');
    try {
      const cachedData = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
      group = await Promise.race([
        client.getChatById(cachedData.groupId),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout de caché')), 8000))
      ]);
      console.log(`✅ Grupo cargado desde caché: ${group.name}`);
    } catch (err) {
      console.log('⚠ El caché ya no sirve (¿cambiaste de cuenta?). Buscando de nuevo...');
      fs.unlinkSync(cacheFile);
      group = null;
    }
  }

  // 2) Si no hay caché, buscar en todos los chats (más lento, solo la primera vez)
  if (!group) {
    console.log(`🔍 Buscando el grupo: "${GROUP_NAME}"`);
    console.log('⏳ Cargando todos los chats (puede tardar varios minutos la primera vez)...\n');
    
    try {
      const chats = await client.getChats();
      console.log(`✅ Se cargaron ${chats.length} chats.`);
      
      const matchedGroup = chats.find(c =>
        c.isGroup && c.name && c.name.toUpperCase().includes(GROUP_NAME.toUpperCase())
      );
      
      if (matchedGroup) {
        group = matchedGroup;
        console.log(`✅ Grupo localizado: ${group.name}`);
        fs.writeFileSync(cacheFile, JSON.stringify({ groupId: group.id._serialized }));
        console.log('💾 ID del grupo guardado en caché para futuras búsquedas instantáneas.');
      } else {
        const availableGroups = chats.filter(c => c.isGroup && c.name);
        console.log(`\n❌ No se encontró el grupo "${GROUP_NAME}".`);
        console.log(`\n📋 Grupos disponibles (${availableGroups.length}):`);
        availableGroups.slice(0, 20).forEach((g, i) => {
          console.log(`   ${i + 1}. ${g.name}`);
        });
        if (availableGroups.length > 20) {
          console.log(`   ... y ${availableGroups.length - 20} grupos más.`);
        }
        console.log(`\n💡 Verifica que el nombre del grupo sea correcto y modifica GROUP_NAME en el código.`);
        return;
      }
    } catch (e) {
      console.error('❌ Error al cargar los chats:', e.message);
      return;
    }
  }

  const startTimestamp = Math.floor(startDate.getTime() / 1000);
  const endTimestamp   = Math.floor(endDate.getTime()   / 1000);

  console.log('\nBuscando mensajes en el rango de fechas...');

  const messages = await group.fetchMessages({ limit: 1000 });
  
  const validMessages = messages.filter(msg => {
    return msg.timestamp >= startTimestamp &&
           msg.timestamp <= endTimestamp &&
           msg.hasMedia &&
           (msg.type === 'image' || msg.type === 'document');
  });

  console.log(`Se encontraron ${validMessages.length} imágenes en el rango de fechas.`);

  if (validMessages.length === 0) {
    console.log('⚠ No hay imágenes para descargar en ese rango.');
    return;
  }

  const receipts = [];

  for (let i = 0; i < validMessages.length; i++) {
    const msg = validMessages[i];
    console.log(`\nProcesando imagen ${i + 1} de ${validMessages.length}...`);
    
    try {
      const media = await msg.downloadMedia();
      if (!media || !media.mimetype.includes('image')) {
        console.log(`  ⚠ El mensaje ${i+1} no es una imagen compatible.`);
        continue;
      }

      const contact = await msg.getContact();
      const senderName = contact.name || contact.pushname || contact.number || 'Desconocido';
      
      const dateObj = new Date(msg.timestamp * 1000);
      const formattedDate = dateObj.toLocaleString('es-CO');

      const imageBuffer = Buffer.from(media.data, 'base64');
      
      console.log(`  👤 Mensajero: ${senderName} | 📅 Fecha: ${formattedDate}`);

      receipts.push({
        imageBuffer,
        senderName,
        date: formattedDate,
        mimetype: media.mimetype
      });

    } catch (err) {
      console.error(`  ❌ Error al procesar el mensaje ${i+1}:`, err.message);
    }
  }

  await createWordDocument(receipts);
}

// ─────────────────────────────────────────────────────────────────────────────
// CREACIÓN DEL DOCUMENTO WORD
// Reglas:
//   • Un solo mensajero por hoja
//   • Máximo 6 comprobantes por hoja (3 columnas × 2 filas)
//   • Orientación VERTICAL (portrait)
//   • Cada nuevo mensajero comienza en una nueva hoja
//   • El nombre del mensajero se muestra claramente en el header de cada hoja
// ─────────────────────────────────────────────────────────────────────────────
async function createWordDocument(receipts) {
  // 1) Agrupar comprobantes por mensajero, manteniendo el orden de llegada
  const senderOrder = [];
  const groupedReceipts = {};
  receipts.forEach(r => {
    if (!groupedReceipts[r.senderName]) {
      groupedReceipts[r.senderName] = [];
      senderOrder.push(r.senderName);
    }
    groupedReceipts[r.senderName].push(r);
  });

  const sections = [];

  // 2) Por cada mensajero, crear tantas secciones (hojas) como sean necesarias
  for (const senderName of senderOrder) {
    const senderReceipts = groupedReceipts[senderName];
    const totalPages = Math.ceil(senderReceipts.length / 6);

    for (let pageIdx = 0; pageIdx < totalPages; pageIdx++) {
      const chunk = senderReceipts.slice(pageIdx * 6, (pageIdx + 1) * 6);
      const isExtraPage = pageIdx > 0;

      // ── Header de página: nombre del mensajero bien visible ──────────────
      const pageHeader = new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: '📋  MENSAJERO: ',
                bold: true,
                size: 36,     // 18pt
                color: '1a1a1a'
              }),
              new TextRun({
                text: `${senderName.toUpperCase()}${isExtraPage ? '  (Continuación)' : ''}`,
                bold: true,
                size: 36,
                color: '003399'
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 80, after: 120 },
            border: {
              bottom: { style: BorderStyle.THICK, size: 8, color: '003399' }
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `Hoja ${pageIdx + 1} de ${totalPages}   |   Total comprobantes: ${senderReceipts.length}   |   Total a verificar: $ ________________________`,
                size: 22,
                color: '555555'
              })
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 0 }
          })
        ]
      });

      // ── Tabla: 3 columnas × 2 filas = 6 comprobantes ────────────────────
      const rows = [];
      for (let rIndex = 0; rIndex < chunk.length; rIndex += 3) {
        const cells = [];

        for (let col = 0; col < 3; col++) {
          const idx = rIndex + col;
          if (idx < chunk.length) {
            cells.push(createReceiptCell(chunk[idx]));
          } else {
            // Celda vacía para completar la fila
            cells.push(new TableCell({
              children: [new Paragraph({ text: '' })],
              width: { size: 3333, type: WidthType.DXA },
              borders: {
                top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
              }
            }));
          }
        }

        // Altura de fila: ~6000 twips ≈ 10.6 cm → ajustado para caber exactamente 2 filas en la página
        rows.push(new TableRow({
          children: cells,
          height: { value: 6000, rule: 'atLeast' }
        }));
      }

      const table = new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top:              { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          bottom:           { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          left:             { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          right:            { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          insideVertical:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        }
      });

      // ── Sección de la hoja ────────────────────────────────────────────────
      // Orientación VERTICAL (portrait): width=12240 twips (21.59 cm), height=15840 twips (27.94 cm)
      // Equivalente a una hoja carta, que es más común que A4 en Colombia.
      // Para A4 usar: width=11906, height=16838
      sections.push({
        properties: {
          page: {
            size: {
              width:  12240,   // Carta ancho (portrait)
              height: 15840,   // Carta alto  (portrait)
              orientation: 'portrait'
            },
            margin: {
              top:    1000,   // ≈1.76 cm  (espacio para header)
              right:   600,   // ≈1.06 cm
              bottom:  600,
              left:    600,
              header:  300,
              footer:  200
            }
          }
        },
        headers: {
          default: pageHeader
        },
        children: [table]
      });
    }
  }

  const doc = new Document({ sections });

  const buffer = await Packer.toBuffer(doc);
  const outPath = 'Comprobantes_Descargados.docx';
  fs.writeFileSync(outPath, buffer);

  console.log(`\n✅ Documento guardado: ${outPath}`);
  console.log(`   Mensajeros procesados : ${senderOrder.length}`);
  console.log(`   Hojas generadas       : ${sections.length}`);
  console.log(`   Comprobantes totales  : ${receipts.length}\n`);
}

// ─────────────────────────────────────────────────────────────────────────────
// CELDA DE COMPROBANTE — Layout vertical optimizado
// ─────────────────────────────────────────────────────────────────────────────
function createReceiptCell(receipt) {
  return new TableCell({
    // 3 columnas → cada una ocupa ~33% del ancho útil
    width:  { size: 3333, type: WidthType.DXA },
    margins: { top: 0, bottom: 0, left: 0, right: 0 }, // Eliminar márgenes verticales
    verticalAlign: VerticalAlign.CENTER,
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' }, // Bordes más delgados
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
      left:   { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' }, // Bordes más delgados
      right:  { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
    },
    children: [
      // Fecha y hora del comprobante
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 5 }, // Reducir aún más el espacio después de la fecha
        children: [
          new TextRun({ text: receipt.date, bold: true, size: 12, color: '333333' }) // Texto aún más pequeño
        ]
      }),
      // Imagen del comprobante
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [
          new ImageRun({
            data: receipt.imageBuffer,
            transformation: {
              width:  230,   // Ajustado para caber mejor en el espacio disponible
              height: 420    // Alto proporcional manteniendo ratio ~1:1.83
            }
          })
        ]
      })
    ]
  });
}

// Ejecutar todo
main();
