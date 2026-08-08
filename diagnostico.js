// Diagnóstico rápido del reporte de mensajeros — correr: node diagnostico.js
'use strict';
const fs = require('fs');
const path = require('path');
const here = __dirname;

const js = fs.readFileSync(path.join(here, 'index.js'), 'utf8');
console.log('1. index.js actualizado (fix nombres)?', js.includes('ownName') ? 'SI' : 'NO ← copiá el index.js nuevo');

const cache = path.join(here, 'group_messages_cache.json');
if (fs.existsSync(cache)) {
  const m = JSON.parse(fs.readFileSync(cache, 'utf8'));
  const j = {}, p = {};
  for (const x of m) {
    const pp = x.key?.participant || '<SIN-PARTICIPANT>';
    j[pp] = (j[pp] || 0) + 1;
    if (x.pushName) p[x.pushName] = (p[x.pushName] || 0) + 1;
  }
  const top = Object.entries(j).sort((a, b) => b[1] - a[1]).slice(0, 10);
  console.log('2. caché: total', m.length, '| remitentes distintos:', Object.keys(j).length);
  console.log('   top remitentes:', JSON.stringify(top));
  console.log('   pushNames distintos:', Object.keys(p).length, '| sample:', JSON.stringify(Object.entries(p).slice(0, 5)));
} else {
  console.log('2. caché: no existe (sesión nueva — los mensajes viven en memoria)');
}

const nf = path.join(here, 'nombres_mensajeros.json');
if (fs.existsSync(nf)) {
  let n = {};
  try { n = JSON.parse(fs.readFileSync(nf, 'utf8')); } catch { console.log('3. nombres_mensajeros.json: CORRUPTO'); }
  const named = Object.entries(n).filter(([, v]) => v && String(v).trim());
  console.log('3. nombres_mensajeros.json:', Object.keys(n).length, 'entradas |', named.length, 'con nombre | claves ej:', Object.keys(n).slice(0, 3).join(', '));
} else {
  console.log('3. nombres_mensajeros.json: NO EXISTE (se crea al primer arranque)');
}
