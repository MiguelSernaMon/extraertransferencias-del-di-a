'use strict';

/**
 * Utilidades de archivo resistentes a corrupción.
 *
 * Problema que resuelve: `fs.writeFileSync` directo deja el archivo truncado
 * si el proceso muere (kill, corte de luz, crash) a mitad de escritura.
 * Un JSON truncado luego se lee "vacío" silenciosamente y se sobreescribe,
 * perdiendo el caché entero de forma IRRECUPERABLE.
 *
 * Solución:
 *  - Escritura atómica: se escribe a `<archivo>.tmp` y luego se renombra
 *    (rename sobre el mismo filesystem es atómico → el destino o es el viejo
 *    completo o el nuevo completo, nunca una mezcla).
 *  - Backup de corruptos: si un JSON no parsea, NO se ignora en silencio ni
 *    se borra — se renombra a `<archivo>.corrupt-<timestamp>` para poder
 *    recuperar datos a mano.
 */

const fs = require('fs');

/** Renombra tmp → destino. En Windows el destino existente puede bloquear el rename. */
function safeRename(tmpPath, destPath) {
  try {
    fs.renameSync(tmpPath, destPath);
    return true;
  } catch {
    // Windows: destino existente → borrar y reintentar
    try {
      fs.rmSync(destPath, { force: true });
      fs.renameSync(tmpPath, destPath);
      return true;
    } catch {
      return false;
    }
  }
}

/**
 * Escritura atómica. Si rename falla (raro), cae a writeFileSync directo
 * (riesgo de corrupción como último recurso, mejor que lanzar al usuario).
 */
function atomicWriteFileSync(filePath, data) {
  const tmpPath = `${filePath}.tmp`;
  fs.writeFileSync(tmpPath, data);
  if (safeRename(tmpPath, filePath)) return;
  try { fs.rmSync(tmpPath, { force: true }); } catch { /* ignore */ }
  fs.writeFileSync(filePath, data);
}

/**
 * Backup de un archivo dañado. Devuelve la ruta del backup o null.
 */
function backupCorruptFile(filePath) {
  try {
    const backup = `${filePath}.corrupt-${Date.now()}`;
    fs.renameSync(filePath, backup);
    return backup;
  } catch {
    return null;
  }
}

/**
 * Lee un JSON de forma segura:
 *  - No existe      → { data: null, corrupt: false }
 *  - Válido         → { data: <parsed>, corrupt: false }
 *  - Corrupto       → respalda el archivo, devuelve { data: null, corrupt: true, backup }
 *
 * El llamador decide qué hacer; NUNCA es silencioso.
 */
function loadJSONFile(filePath) {
  if (!fs.existsSync(filePath)) return { data: null, corrupt: false, backup: null };
  let raw;
  try {
    raw = fs.readFileSync(filePath, 'utf8');
  } catch (err) {
    const backup = backupCorruptFile(filePath);
    console.error(`⚠ Archivo ilegible: ${filePath} (${err.message}). ${backup ? `Respaldo en: ${backup}` : 'No se pudo respaldar.'}`);
    return { data: null, corrupt: true, backup };
  }
  try {
    return { data: JSON.parse(raw), corrupt: false, backup: null };
  } catch (err) {
    const backup = backupCorruptFile(filePath);
    console.error(`⚠ Archivo corrupto (JSON inválido): ${filePath}. ${backup ? `Respaldo en: ${backup}` : ''}`);
    return { data: null, corrupt: true, backup };
  }
}

/**
 * Valida que un buffer parezca un JPEG completo (FF D8 … FF D9).
 * Detecta archivos truncados por un crash a mitad de escritura.
 */
function isValidJpeg(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 4) return false;
  return buffer[0] === 0xff && buffer[1] === 0xd8 &&
         buffer[buffer.length - 2] === 0xff && buffer[buffer.length - 1] === 0xd9;
}

module.exports = { atomicWriteFileSync, backupCorruptFile, loadJSONFile, isValidJpeg };
