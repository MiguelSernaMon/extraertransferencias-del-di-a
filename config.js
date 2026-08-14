'use strict';
// Config del cliente Evolution API.
// Fuentes (en orden de prioridad): env vars > config.json > defaults.
// El archivo config.json NO se commitea (tiene la API key); ver config.example.json.

const fs = require('fs');
const path = require('path');
const https = require('https');
const { atomicWriteFileSync } = require('./fs_utils');

const DEFAULT_URL = 'https://transferencias.redpostal.co';

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return {};
  }
}

const fileCfg = readJson(path.join(__dirname, 'config.json'));

// NO freeze: el panel web edita la conexión en vivo (updateConfig).
const config = {
  evolutionUrl: (process.env.EVOLUTION_URL || fileCfg.evolutionUrl || DEFAULT_URL).replace(/\/+$/, ''),
  apiKey: process.env.EVOLUTION_API_KEY || fileCfg.apiKey || '',
  instance: process.env.EVOLUTION_INSTANCE || fileCfg.instance || 'comprobantes',
  groupName: process.env.EVOLUTION_GROUP_NAME || fileCfg.groupName || 'transferencias',
  caFile: process.env.EVOLUTION_CA_FILE || fileCfg.caFile || '',
};

let cachedAgent = undefined;
/** https.Agent con la CA del VPS (si config.caFile está seteado). undefined = CAs del sistema. */
function getAgent() {
  if (config.caFile) {
    if (!cachedAgent) {
      cachedAgent = new https.Agent({ ca: fs.readFileSync(config.caFile) });
    }
    return cachedAgent;
  }
  return undefined;
}

/** Aplica cambios de conexión desde el panel web (en vivo) y persiste a
 *  config.json (merge atómico: preserva claves extra y el resto de campos).
 *  Las claves vacías no se tocan. Devuelve el config resultante. */
function updateConfig(partial) {
  for (const k of ['evolutionUrl', 'apiKey', 'instance', 'groupName', 'caFile']) {
    const v = partial && partial[k];
    if (typeof v === 'string' && v.trim() !== '') config[k] = v.trim();
  }
  config.evolutionUrl = config.evolutionUrl.replace(/\/+$/, '');
  const existing = readJson(path.join(__dirname, 'config.json'));
  atomicWriteFileSync(path.join(__dirname, 'config.json'), JSON.stringify({ ...existing, ...config }, null, 2));
  cachedAgent = undefined; // si cambió caFile, el agente se reconstruye
  return { ...config };
}

module.exports = { config, getAgent, updateConfig };
