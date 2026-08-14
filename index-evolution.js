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
// index.js compara con su estado interno knownGroupName (solo lo setea una
// conexión baileys). En el cliente Evolution no hay conexión: se siembra
// desde config.groupName (spec §3.1: GROUP_NAME reusa isGroupNameLike).
M._setKnownGroupName(config.groupName);
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

  // Nombres persistidos del usuario: caché de nombres + nombres_mensajeros.json.
  // Sin esto el modo Evolution ignora el mapeo del usuario y los remitentes
  // no-contactos salen como número crudo (reviewer §3.4 paso 4). Las dos
  // funciones son idempotentes (solo cargan mapas que usa getSenderName).
  M.loadNameCache();
  M.loadManualNames();

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
