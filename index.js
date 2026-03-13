const fs = require('fs');
const qrcode = require('qrcode-terminal');
const { Client, LocalAuth } = require('whatsapp-web.js');
const { prompt } = require('enquirer');
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
  AlignmentType
} = require('docx');

const GROUP_NAME = 'TRANSFERENCIAS RED POSTAL POBLADO';

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
      name: 'endDate',
      message: 'Fecha de FIN (YYYY-MM-DD) [Presiona Enter para HOY]:'
    }
  ];

  const answers = await prompt(questions);

  const startDate = new Date();
  if (answers.startDate && answers.startDate.trim() !== '') {
    const parts = answers.startDate.trim().split('-');
    startDate.setFullYear(parseInt(parts[0]));
    startDate.setMonth(parseInt(parts[1]) - 1);
    startDate.setDate(parseInt(parts[2]));
    startDate.setHours(0, 0, 0, 0);
  } else {
    startDate.setDate(startDate.getDate() - 1);
    startDate.setHours(0, 0, 0, 0);
  }

  const endDate = new Date();
  if (answers.endDate && answers.endDate.trim() !== '') {
    const parts = answers.endDate.trim().split('-');
    endDate.setFullYear(parseInt(parts[0]));
    endDate.setMonth(parseInt(parts[1]) - 1);
    endDate.setDate(parseInt(parts[2]));
    endDate.setHours(23, 59, 59, 999);
  } else {
    endDate.setHours(23, 59, 59, 999);
  }

  console.log(`\n📅 Rango seleccionado: ${startDate.toLocaleString()} HASTA ${endDate.toLocaleString()}\n`);

  const client = new Client({
    authStrategy: new LocalAuth({ dataPath: './.wwebjs_auth' }),
    webVersionCache: {
      type: 'remote',
      remotePath: 'https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.2412.54.html',
    },
    puppeteer: {
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--disable-gpu'],
      protocolTimeout: 0,
      timeout: 0
    }
  });

  client.on('qr', (qr) => {
    console.log('\n======================================================');
    console.log(' POR FAVOR ESCANEA EL SIGUIENTE CÓDIGO QR CON WHATSAPP');
    console.log('======================================================\n');
    qrcode.generate(qr, { small: true });
  });

  client.on('ready', async () => {
    console.log('✅ Cliente de WhatsApp listo y conectado.');
    try {
      await processGroupMessages(client, startDate, endDate);
    } catch (error) {
      console.error('❌ Error procesando los mensajes:', error);
    } finally {
      console.log('▶ Proceso finalizado. Puedes cerrar la aplicación.');
      process.exit(0);
    }
  });

  client.initialize();
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
    console.log(`Buscando el grupo: "${GROUP_NAME}" (Puede tardar un par de minutos la primera vez)...`);
    
    // Para cuentas MUY grandes, getChats() puede fallar o tardar horas.
    // Usamos una estrategia mixta: Intentamos buscar en los chats más recientes,
    // y si tarda demasiado, le pedimos al usuario que interactúe.
    
    group = await new Promise(async (resolve) => {
      let found = false;
      
      // Estrategia 1: Escuchar si llega algún mensaje nuevo a ese grupo
      const tempListener = (msg) => {
        if (msg.from.includes('@g.us')) {
          msg.getChat().then(chat => {
            if (chat.name && chat.name.toUpperCase().includes(GROUP_NAME.toUpperCase())) {
              if (!found) {
                found = true;
                client.removeListener('message', tempListener);
                resolve(chat);
              }
            }
          });
        }
      };
      client.on('message', tempListener);

      // Estrategia 2: Intentar getChats normal, pero con timeout de seguridad (30 segs)
      const timeoutId = setTimeout(() => {
        if (!found) {
          console.log('\n======================================================');
          console.log('⚠ TU CUENTA TIENE DEMASIADOS CHATS Y TARDA DEMASIADO ⚠');
          console.log(`Por favor, ve a WhatsApp en tu celular y ENVÍA UN MENSAJE (cualquiera) al grupo:`);
          console.log(`"${GROUP_NAME}"`);
          console.log('Esto permitirá localizar el grupo instantáneamente.');
          console.log('======================================================\n');
        }
      }, 15000);

      try {
        const chats = await client.getChats();
        if (!found) {
          const matchedGroup = chats.find(c => c.isGroup && c.name.toUpperCase().includes(GROUP_NAME.toUpperCase()));
          if (matchedGroup) {
            found = true;
            client.removeListener('message', tempListener);
            clearTimeout(timeoutId);
            resolve(matchedGroup);
          } else {
             // Dejar que la estrategia 1 lo resuelva si no estaba en la lista local
          }
        }
      } catch (e) {
        // Ignorar error si getChats colapsa, la estrategia 1 seguirá activa
      }
    });

    if (!group) {
      console.log(`❌ No se encontró el grupo "${GROUP_NAME}". Verifica el nombre exacto.`);
      return;
    }
    console.log(`✅ Grupo localizado: ${group.name}`);
    
    fs.writeFileSync(cacheFile, JSON.stringify({ groupId: group.id._serialized }));
    console.log(`💾 ID del grupo guardado en caché para futuras descargas inmediatas.`);
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

      // EL TÍTULO SUPERIOR (Solo una vez por página)
      elements.push(new Paragraph({
        children: [
          new TextRun({ text: `Mensajero: `, bold: true, size: 28 }),
          new TextRun({ text: `${senderName}${isSecondOrMorePageForSender ? ' (Continuación)' : ''}`, bold: false, size: 28 }),
          new TextRun({ text: `    |    Total: ________________________`, bold: true, size: 28 })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 100 }
      }));
      
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
        rows.push(new TableRow({ children: rowCells, height: { value: 4800, rule: "exact" } })); // Alto de fila ajustado para aprovechar espacio
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
            margin: { top: 400, right: 400, bottom: 400, left: 400 } // márgenes muy delgados
          }
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
    // Margen interno entre celdas un poco más reducido para aprovechar espacio
    margins: {
      top: 50,
      bottom: 50,
      left: 100,
      right: 100,
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
          new TextRun({ text: `${receipt.date}`, bold: true, size: 24 })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: receipt.imageBuffer,
            transformation: {
              width: 320,    // Imágenes mucho más grandes para aprovechar la hoja sin margenes
              height: 420
            }
          })
        ]
      })
    ]
  });
}

// Ejecutar todo
main();
