'use strict';
// Config del cliente Evolution API.
// Fuentes (en orden de prioridad): env vars > config.json > defaults.
// El archivo config.json NO se commitea (tiene la API key); ver config.example.json.

const fs = require('fs');
const path = require('path');
const https = require('https');

const DEFAULT_URL = 'https://transferencias.redpostal.co';

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return {};
  }
}

const fileCfg = readJson(path.join(__dirname, 'config.json'));

const config = Object.freeze({
  evolutionUrl: (process.env.EVOLUTION_URL || fileCfg.evolutionUrl || DEFAULT_URL).replace(/\/+$/, ''),
  apiKey: process.env.EVOLUTION_API_KEY || fileCfg.apiKey || '',
  instance: process.env.EVOLUTION_INSTANCE || fileCfg.instance || 'comprobantes',
  groupName: process.env.EVOLUTION_GROUP_NAME || fileCfg.groupName || 'transferencias',
  caFile: process.env.EVOLUTION_CA_FILE || fileCfg.caFile || '',
});

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

module.exports = { config, getAgent };
