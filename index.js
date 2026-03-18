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
  Header,
  PageNumber,
  NumberFormat
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
        // NOTA: --single-process fue REMOVIDO porque causa "frame detached" en Windows
      ],
      protocolTimeout: 600000, // 10 minutos de timeout para protocolos (aumentado significativamente)
      timeout: 300000 // 5 minutos de timeout general (aumentado significativamente)
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

async function main() {
  console.log('\n======================================================');
  console.log(' CONFIGURACIÓN DE FECHAS DE DESCARGA');
  console.log('======================================================');
  console.log('Ingresa la fecha en formato YYYY-MM-DD (ej: 2026-03-12).');
  console.log('Si dejas en blanco, usará por defecto AYER y HOY respectivamente.\n');

  const questions = [
    {
      type: 'input',
      name: 'startDate',
      message: 'Fecha de INICIO (YYYY-MM-DD) [Presiona Enter para AYER]:'
    },
    {
      type: 'input',
      name: 'startTime',
      message: 'Hora de INICIO de AYER (HH:MM) [Presiona Enter para 08:30]:'
    },
    {
      type: 'input',
      name: 'endDate',
      message: 'Fecha de FIN (YYYY-MM-DD) [Presiona Enter para HOY]:'
    }
  ];

  const answers = await prompt(questions);

  // Parsear la hora de inicio (por defecto 8:30am)
  let startHour = 8;
  let startMinute = 30;
  if (answers.startTime && answers.startTime.trim() !== '') {
    const timeParts = answers.startTime.trim().split(':');
    startHour = parseInt(timeParts[0]) || 8;
    startMinute = parseInt(timeParts[1]) || 30;
  }

  const startDate = new Date();
  if (answers.startDate && answers.startDate.trim() !== '') {
    const parts = answers.startDate.trim().split('-');
    startDate.setFullYear(parseInt(parts[0]));
    startDate.setMonth(parseInt(parts[1]) - 1);
    startDate.setDate(parseInt(parts[2]));
    startDate.setHours(startHour, startMinute, 0, 0);
  } else {
    startDate.setDate(startDate.getDate() - 1);
    startDate.setHours(startHour, startMinute, 0, 0);
  }

  const endDate = new Date();
  if (answers.endDate && answers.endDate.trim() !== '') {
    const parts = answers.endDate.trim().split('-');
    endDate.setFullYear(parseInt(parts[0]));
    endDate.setMonth(parseInt(parts[1]) - 1);
    endDate.setDate(parseInt(parts[2]));
    // La hora de fin es la hora de inicio del día siguiente menos 1 segundo
    endDate.setHours(startHour, startMinute - 1, 59, 999);
  } else {
    // Hoy hasta la hora de inicio del siguiente ciclo (mañana a las 8:30) menos 1 segundo
    endDate.setDate(endDate.getDate() + 1);
    endDate.setHours(startHour, startMinute - 1, 59, 999);
  }

  console.log(`\n📅 Rango seleccionado: ${startDate.toLocaleString()} HASTA ${endDate.toLocaleString()}\n`);

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
      
      // Intentar cerrar el cliente si existe
      if (client) {
        try {
          await client.destroy();
        } catch (destroyErr) {
          // Ignorar errores al destruir
        }
      }
      
      if (attempt < MAX_RETRIES) {
        if (isContextDestroyed) {
          console.log(`\n⚠ Se detectó error de contexto destruido. Esto puede ocurrir cuando WhatsApp Web se actualiza.`);
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        } else if (isFrameDetached) {
          console.log(`\n⚠ Se detectó error "frame detached" (común en Windows).`);
          console.log(`💡 Se ha removido el flag --single-process para evitar este error.`);
          console.log(`🔄 Reintentando en ${RETRY_DELAY_MS / 1000} segundos...\n`);
        } else if (isRuntimeTimeout) {
          console.log(`\n⚠ Se detectó error de timeout (runtime.callfuncuint).`);
          console.log(`💡 Los timeouts ya fueron aumentados a 10 minutos para protocolos y 5 minutos general.`);
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

async function processGroupMessages(client, startDate, endDate) {
  const cacheFile = './group_cache.json';
  let group;

  if (fs.existsSync(cacheFile)) {
    console.log(`Buscando el grupo desde el caché...`);
    const cachedData = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
    try {
      // Añadimos un límite de tiempo (5 segundos) por si el ID del grupo guardado ya no existe.
      group = await Promise.race([
        client.getChatById(cachedData.groupId),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout de caché')), 5000))
      ]);
      console.log(`✅ Grupo cargado desde caché: ${group.name}`);
    } catch (err) {
      console.log('⚠ El viejo caché no sirvió (tal vez cambiaste de cuenta). Buscando de nuevo...');
      fs.unlinkSync(cacheFile); // Borramos el caché corrupto/obsoleto
      group = null;
    }
  }

  if (!group) {
    console.log(`Buscando el grupo: "${GROUP_NAME}" (Puede tardar varios minutos la primera vez)...`);
    console.log(`⏳ Por favor espera mientras se cargan todos los chats...\n`);
    
    try {
      // Cargar todos los chats - puede tardar varios minutos en cuentas grandes
      const chats = await client.getChats();
      console.log(`✅ Se cargaron ${chats.length} chats.`);
      
      // Buscar el grupo por nombre (parcial, ignorando mayúsculas/minúsculas)
      const matchedGroup = chats.find(c => 
        c.isGroup && c.name && c.name.toUpperCase().includes(GROUP_NAME.toUpperCase())
      );
      
      if (matchedGroup) {
        group = matchedGroup;
        console.log(`✅ Grupo localizado: ${group.name}`);
        fs.writeFileSync(cacheFile, JSON.stringify({ groupId: group.id._serialized }));
        console.log(`💾 ID del grupo guardado en caché para futuras descargas inmediatas.`);
      } else {
        // Mostrar los grupos disponibles para ayudar al usuario
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
  const endTimestamp = Math.floor(endDate.getTime() / 1000);

  console.log(`Buscando mensajes...`);

  // Buscamos los últimos mensajes
  const messages = await group.fetchMessages({ limit: 1000 });
  
  const validMessages = messages.filter(msg => {
    return msg.timestamp >= startTimestamp &&
           msg.timestamp <= endTimestamp &&
           msg.hasMedia &&
           (msg.type === 'image' || msg.type === 'document');
  });

  console.log(`Se encontraron ${validMessages.length} imágenes en el rango de fechas.`);

  if (validMessages.length === 0) {
    console.log('⚠ No hay imágenes para descargar.');
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
      // contact.name = Nombre como lo tienes guardado tú en tu teléfono
      // contact.pushname = Nombre que la otra persona se puso en su perfil de WhatsApp
      const senderName = contact.name || contact.pushname || contact.number || 'Desconocido';
      
      const dateObj = new Date(msg.timestamp * 1000);
      const formattedDate = dateObj.toLocaleString('es-CO'); // ej: 11/3/2026, 14:30:00

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

async function createWordDocument(receipts) {
  // Agrupar por mensajero
  const groupedReceipts = {};
  receipts.forEach(r => {
    if (!groupedReceipts[r.senderName]) groupedReceipts[r.senderName] = [];
    groupedReceipts[r.senderName].push(r);
  });

  const sections = [];

  // Crear una sección por cada mensajero (y dividirla si tiene más de 6 comprobantes)
  for (const [senderName, senderReceipts] of Object.entries(groupedReceipts)) {
    // Dividir en grupos de 6 max por página
    for (let i = 0; i < senderReceipts.length; i += 6) {
      const chunk = senderReceipts.slice(i, i + 6);
      
      const elements = [];
      const isSecondOrMorePageForSender = i > 0;

      // CREAR HEADER DE PÁGINA - Esto garantiza que SIEMPRE aparezca arriba en la impresión
      const pageHeader = new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: `Mensajero: `, bold: true, size: 32 }),
              new TextRun({ text: `${senderName}${isSecondOrMorePageForSender ? ' (Continuación)' : ''}`, bold: true, size: 32 }),
              new TextRun({ text: `    |    Total: ________________________`, bold: true, size: 28 })
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 200 },
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
            }
          })
        ]
      });
      
      // LA TABLA (2 columnas, max 3 filas)
      const rows = [];
      for (let rIndex = 0; rIndex < chunk.length; rIndex += 2) {
        const rowCells = [];
        for (let col = 0; col < 2; col++) {
          const receiptIndex = rIndex + col;
          if (receiptIndex < chunk.length) {
            rowCells.push(createReceiptCell(chunk[receiptIndex]));
          } else {
            // Celda vacía
            rowCells.push(new TableCell({
              children: [new Paragraph({text: ""})],
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              }
            }));
          }
        }
        rows.push(new TableRow({ children: rowCells, height: { value: 5200, rule: "exact" } })); // Alto de fila reducido para imágenes más pequeñas
      }
      
      const table = new Table({
        rows: rows,
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        borders: {
          top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        }
      });

      elements.push(table);

      // Agregamos la hoja (Sección) al word, con MARGEN 0 para ahorrar papel
      // Usamos margen muy pequeño (400 twips = aprox 0.7cm) porque 0 absoluto puede causar problemas de impresión en algunas impresoras.
      sections.push({
        properties: {
          page: {
            margin: { top: 850, right: 400, bottom: 400, left: 400 } // margen superior para header
          }
        },
        headers: {
          default: pageHeader  // HEADER FIJO - Garantiza que el nombre del mensajero SIEMPRE aparezca arriba
        },
        children: elements
      });
    }
  }

  const doc = new Document({
    sections: sections
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = 'Comprobantes_Descargados.docx';
  fs.writeFileSync(outPath, buffer);
  console.log(`\n✅ Documento guardado exitosamente como: ${outPath} en tu carpeta.`);
  console.log(`Se generaron ${sections.length} páginas/hojas en total.\n`);
}

function createReceiptCell(receipt) {
  return new TableCell({
    width: {
      size: 50,
      type: WidthType.PERCENTAGE,
    },
    // Margen interno reducido para imágenes más pegadas
    margins: {
      top: 20,
      bottom: 20,
      left: 50,
      right: 50,
    },
    verticalAlign: VerticalAlign.CENTER,
    borders: { // Borde del campo del comprobante para recortar/separar
      top: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
      left: { style: BorderStyle.SINGLE, size: 2, color: "888888" },
      right: { style: BorderStyle.SINGLE, size: 2, color: "888888" },
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: `${receipt.date}`, bold: true, size: 20 })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
            new ImageRun({
            data: receipt.imageBuffer,
            transformation: {
              width: 250,    // Ancho más pequeño para ahorrar espacio
              height: 400    // Alto ajustado para ratio más vertical (1:1.6)
            }
          })
        ]
      })
    ]
  });
}

// Ejecutar todo
main();
