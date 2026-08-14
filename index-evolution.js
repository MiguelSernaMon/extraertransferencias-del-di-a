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
const NOMBRES_FILE = './nombres_mensajeros.json';
const MSG_CACHE_FILE = './group_messages_cache.json';

/** Runs anteriores guardaron pushName = ECO del lid (ej. "77150682091545" para
 *  participant "77150682091545@lid") ANTES del fix en normalizeRecord.
 *  loadCachedGroupMessages vuelca esos ecos al nameCache interno de index.js y
 *  getSenderName los devuelve como "nombre" → el Word muestra el lid aunque el
 *  puente lid→teléfono tenga el número. Se limpian una vez al arrancar
 *  (eco → null / borrado); nameCache se reconstruye limpio. */
function sanitizeLidEchoes() {
  for (const file of [MSG_CACHE_FILE, 'name_cache.json']) {
    let data = null;
    try { data = JSON.parse(fs.readFileSync(file, 'utf8')); } catch { continue; }
    if (!data) continue;
    let changed = false;
    if (Array.isArray(data)) { // caché de mensajes: [{key, pushName, ...}]
      for (const m of data) {
        const p = m.key?.participant;
        if (p?.endsWith('@lid') && m.pushName === p.split('@')[0]) { m.pushName = null; changed = true; }
      }
    } else { // name_cache: { jid: nombre }
      for (const [jid, name] of Object.entries(data)) {
        if (jid.endsWith('@lid') && name === jid.split('@')[0]) { delete data[jid]; changed = true; }
      }
    }
    if (changed) { atomicWriteFileSync(file, JSON.stringify(data)); console.log(`🧹 Caché ${file}: ecos de lid limpiados`); }
  }
}

const CONCURRENCY = 2;
const RETRIES = 3;

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// Estado para el panel de mapeo (miembros + contactos del último run)
let currentMembers = [];
let currentContacts = {};

/** Entries para el panel de mapeo: lids vistos en el caché + miembros del grupo
 *  (lid → teléfono) + nombres manuales/contactos ya resueltos. El mapper de
 *  picker.js pide [{lid, phone, name}] — mismo shape de senderToEntry (index.js). */
function buildParticipantEntries() {
  const entries = new Map();
  // 1. Remitentes del caché local (lids, sin esperar API)
  try {
    const cache = M.loadCachedGroupMessages();
    for (const [, m] of cache) {
      const lid = m.key?.participant;
      if (lid && lid.endsWith('@lid')) entries.set(lid, { lid, phone: null, name: '' });
    }
  } catch { /* caché aún no existe */ }
  // 2. Miembros del grupo: lid → teléfono + nombre de contacto
  for (const m of currentMembers) {
    const e = entries.get(m.id) || { lid: m.id, phone: null, name: '' };
    e.phone = m.phoneNumber || e.phone;
    if (currentContacts[m.id]?.name) e.name = currentContacts[m.id].name;
    entries.set(m.id, e);
  }
  // 3. Nombres manuales del archivo (senderToEntry usa manualNames del módulo)
  return [...entries.values()].map(e => {
    const entry = M.senderToEntry(e.lid, e.phone);
    return { lid: e.lid, phone: e.phone, name: entry.name || e.name };
  });
}

/** Persiste el puente lid→teléfono en _lid_to_phone del archivo de nombres
 *  (mismo formato que saveManualNames de index.js). El mapeo vivo que en
 *  baileys llenaba la metadata del grupo acá sale de fetchGroupMembers:
 *  persistirlo hace que nombres guardados bajo NÚMERO (formato legacy y del
 *  panel) matcheen a remitentes que llegan como lid (getSenderName resuelve
 *  lid → teléfono vía lidToPhone). Devuelve true si cambió algo. */
function persistLidBridge(members) {
  let map = {};
  try {
    const d = JSON.parse(fs.readFileSync(NOMBRES_FILE, 'utf8'));
    if (d && !Array.isArray(d)) map = d;
  } catch { /* archivo nuevo */ }
  const lidPhone = map._lid_to_phone || {};
  let changed = false;
  for (const m of members || []) {
    const bare = (m.phoneNumber || '').replace('@s.whatsapp.net', '');
    if (m.id && bare && lidPhone[m.id] !== bare) { lidPhone[m.id] = bare; changed = true; }
  }
  if (changed) {
    atomicWriteFileSync(NOMBRES_FILE, JSON.stringify({ ...map, _lid_to_phone: lidPhone }, null, 2));
    M.loadManualNames(); // refresca lidToPhone del módulo para getSenderName
  }
  return changed;
}

/** Persiste un nombre manual con el MISMO formato de saveManualNames (index.js):
 *  {...manualNames, _lid_to_phone} y refresca el mapa interno del módulo para
 *  que getSenderName lo use en la misma corrida. El panel manda key = teléfono
 *  o lid: se normaliza a número crudo (getSenderName consulta manualNames con
 *  el número sin sufijo) y el puente lid→teléfono se aplica a los miembros
 *  que usan ese número. */
function saveManualName(key, name) {
  let map = {};
  try {
    const d = JSON.parse(fs.readFileSync(NOMBRES_FILE, 'utf8'));
    if (d && !Array.isArray(d)) map = d;
  } catch { /* archivo nuevo */ }
  const bare = String(key).replace('@s.whatsapp.net', '');
  if (name) map[bare] = name;
  else delete map[bare];
  const lidPhone = map._lid_to_phone || {};
  for (const m of currentMembers) {
    if (m.id && (m.phoneNumber || '').replace('@s.whatsapp.net', '') === bare) lidPhone[m.id] = bare;
  }
  atomicWriteFileSync(NOMBRES_FILE, JSON.stringify({ ...map, _lid_to_phone: lidPhone }, null, 2));
  M.loadManualNames(); // refresh del mapa interno (getSenderName)
}

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
 * opts: { startTs, endTs, picker } — segundos Unix inclusivos; picker opcional
 * (si viene, refresca el panel de mapeo con los remitentes del run).
 * Retorna { total, downloaded, skipped, failed, outOfRange, wordPath }.
 */
async function runPipeline({ startTs, endTs, picker } = {}) {
  // 0. Limpiar ecos de lid en cachés (ver sanitizeLidEchoes) ANTES de que
  // loadNameCache/loadCachedGroupMessages poblen el nameCache del módulo.
  sanitizeLidEchoes();

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

  // 3b. Miembros del grupo: el phoneNumber es el puente lid→teléfono (la
  // metadata de grupo que llenaba lidToPhone en baileys no existe acá). Un
  // contacto que matchea por teléfono resuelve el nombre de remitentes que
  // llegan como lid — contacts[lid] es el lookup que usa getSenderName. El
  // puente se persiste en _lid_to_phone para que los nombres guardados bajo
  // número (formato legacy y del panel) matcheen en corridas futuras.
  currentMembers = [];
  try { currentMembers = await evo.fetchGroupMembers(groupJid); }
  catch (err) { console.log(`⚠ Sin miembros del grupo: ${err.message}`); }
  for (const m of currentMembers) {
    if (!m.phoneNumber) continue;
    const name = contacts[m.phoneNumber]?.name;
    if (name) contacts[m.id] = { name };
  }
  currentContacts = contacts;
  persistLidBridge(currentMembers);

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

  // 6b. Refrescar el panel de mapeo: ahora el caché trae los lids de este run
  // y los miembros el puente lid→teléfono (no-op si main no pasó picker).
  if (picker) {
    try { picker.setParticipants(buildParticipantEntries()); } catch { /* panel no disponible */ }
  }

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
  let picker = null;
  if (useWebPicker) {
    picker = await startControlServer();
    picker.attachConsole();
    picker.openBrowser();
    // Panel de mapeo de remitentes: guarda vía saveManualName (mismo formato
    // que saveManualNames de index.js, preservando _lid_to_phone) y entradas
    // sembradas desde el caché local — el usuario puede mapear nombres
    // antes de elegir el rango (runPipeline refresca con miembros al correr).
    picker.setNameSaver(saveManualName);
    picker.setParticipants(buildParticipantEntries());
    // Conexión configurable desde el panel: estado de la instancia, reconexión
    // con QR y guardado en vivo (updateConfig en config.js). Si cambia la
    // instancia, el groupJid cacheado no sirve → se limpia el caché.
    picker.setInstanceChecker(async () => {
      try { return await evo.checkInstance(); }
      catch (err) { return { found: false, status: 'unknown', error: err.message }; }
    });
    picker.setConnector(() => evo.connectInstance());
    // Selector de instancias: lista las que existen en Evolution (la crea el
    // usuario en la UI del VPS) para elegir la que se acaba de hacer.
    picker.setInstanceLister(async () => {
      const list = await evo.listInstances();
      if (list.error) return [];
      return list;
    });
    picker.setConfigApplied(({ changed }) => {
      if (changed) {
        try { fs.rmSync(CACHE_FILE, { force: true }); } catch { /* no existe */ }
        console.log('🔄 Instancia cambiada: caché de grupo limpiado');
      }
    });
    console.log(`📡 Panel abierto en el navegador (${picker.url})`);
    ({ startDate, endDate } = await picker.waitForRange());
  } else {
    ({ startDate, endDate } = await askDateRangeTerminal());
  }
  // Los días se interpretan como días LOCALES (Bogotá) y el día fin es
  // INCLUSIVO: "13 a 14" toma el 13 completo + el 14 completo. new Date('YYYY-MM-DD')
  // cae en medianoche UTC → se re-parsea la fecha ISO como medianoche local.
  const dayStartLocal = (d) => {
    const [y, m, dd] = d.toISOString().slice(0, 10).split('-').map(Number);
    return Math.floor(new Date(y, m - 1, dd).getTime() / 1000);
  };
  const startTs = dayStartLocal(startDate);
  const endTs = dayStartLocal(endDate) + 86399; // fin del día local (inclusivo)

  // Borrar el Word de corridas anteriores (mismo patrón que index.js: el archivo
  // viejo del día pasado NO puede confundirse con un resultado nuevo)
  try { fs.rmSync(OUTPUT_FILE, { force: true }); } catch { /* si no existe, ok */ }

  await runPipeline({ startTs, endTs, picker: picker || undefined });
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
