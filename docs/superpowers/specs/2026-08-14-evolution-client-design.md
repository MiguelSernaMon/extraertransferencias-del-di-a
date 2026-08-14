# Cliente Evolution API — Reemplazo del scraping baileys

**Fecha:** 2026-08-14
**Estado:** propuesto

## 1. Problema

El extractor actual usa `@whiskeysockets/baileys` (protocolo reverso no oficial).
Síntomas en producción: QR falla en Windows, history sync lo empuja el servidor
cuando quiere ("0 comprobantes" con mensajes existiendo), sesiones se corrompen,
y cada fix pelea contra la misma raíz: canal no oficial que WhatsApp puede
romper o banear en cualquier momento.

## 2. Decisión

Reemplazar el origen de datos: una instancia **Evolution API v2** (Docker, en la
misma máquina) mantiene una sesión de WhatsApp viva 24/7 con una SIM dedicada.
La app deja de conectar directo a WhatsApp y consulta **REST** a Evolution,
que persiste todo en PostgreSQL.

- Los mensajeros NO cambian nada: siguen mandando la foto al grupo.
- QR se escanea UNA vez al instalar (en la UI web de Evolution), nunca por el usuario final.
- La fecha de cada mensaje es su timestamp real guardado por Evolution — desaparecen
  sync vacío, rango equivocado y recuperación de historial cortado.

## 3. Componentes

### 3.1 Infraestructura (YA existente, fuera de alcance)

Evolution API v2 está hosteado en un VPS del usuario (URL pública + API key +
instancia creada y enlazada con el teléfono de la SIM). La app SOLO consume su
REST — no despliega ni administra el daemon. No hay docker-compose local.

### 3.2 config.js (nuevo)

- `EVOLUTION_URL` (default `https://transferencias.redpostal.co`), `API_KEY`, `INSTANCE` (default `comprobantes`).
- `GROUP_NAME` — filtro para descubrir el grupo por nombre (reusa `isGroupNameLike`).
- `CA_FILE` (default vacío) — ruta a un PEM del cert del VPS (el autofirmado de
  Traefik) para usarlo como CA propia del cliente (`https.Agent` con `ca:`).
  Sin esto, falla el handshake mientras el dominio no tenga Let's Encrypt.
  NO se desactiva la verificación TLS (evitar MITM); se fija la CA exacta.
  La solución definitiva es Let's Encrypt en Traefik — documentada en README.
- Fuentes: `config.json` (gitignored) + override por env vars. `config.example.json` en repo.
- Reemplaza `knownGroupName` hardcodeado de index.js.

### 3.3 evolution-client.js (nuevo)

Cliente REST delgado, sin estado. Funciones:

- `apiGet(instance, path)` / `apiPost(instance, path, body)` — fetch con header `apikey`,
  manejo de error con detalle de body (Evolution devuelve `{success, error}`).
- `listGroups()` — `GET /group/fetchAllGroups/{instance}`; retorna `[{id, subject}]`.
  Usado UNA vez para descubrir `groupJid`; se guarda en `cache.json` (mismo formato actual).
  *(Endpoint a confirmar contra la instalación real durante implementación; fallback
  documentado: `GET /chat/findChats/{instance}` filtrando `id` que termina en `@g.us`.)*
- `findMediaMessages(jid, startTs, endTs)`:
  - `POST /chat/findMessages/{instance}` con `{ where: { key: { remoteJid: jid } },
    take: 500, skip: n, orderBy: { timestamp: 'asc' } }`, paginando.
  - **Filtrado client-side SIEMPRE** (issue #1632: el filtro remoteJid puede no
    aplicarse): `key.remoteJid === jid`, `isMsgInRange(msg, startTs, endTs)`,
    y es media (`imageMessage` o `documentMessage`).
  - Normaliza cada record a la forma que ya usa todo el pipeline:
    `{ key, message, messageTimestamp: record.timestamp }` — así
    `isMsgInRange`, `toUnix`, `getSenderName` funcionan sin cambios.
- `downloadMedia(msg)` — `POST /chat/getBase64FromMediaMessage/{instance}` con el
  **mensaje completo** (caveat: solo el ID falla en media viejo); retorna
  `{base64, mimetype}` → Buffer → guarda en `media_cache/<key.id>.jpg`
  (misma convención y helper `getMediaCachePath` actuales).
- `findContactsMap()` — `GET /chat/findContacts/{instance}` → `{numero → nombre}`
  para resolver remitentes (alimenta el mismo `nameCache`).

### 3.4 index-evolution.js (nuevo, entry point)

Reemplaza el main() de index.js. Flujo:

1. `startControlServer()` (picker.js, **sin cambios**) → `waitForRange()`.
2. Verificación de conexión: `GET /instance/fetchInstances` — si la instancia no
   existe o está desconectada, mensaje claro con pasos (QR una vez).
3. Descubrir `groupJid` (cache.json) o `listGroups()` → filtrar por `GROUP_NAME`.
4. Cargar contactos → `loadNameCache()` (reusa) → `getSenderName` como hoy.
5. `findMediaMessages(jid, startTs, endTs)` → mensajes en rango.
6. Por cada mensaje: si `media_cache/<id>.jpg` ya existe → skip (rápido en re-runs);
   si no → `downloadMedia` con concurrencia 2 y reintentos; fallos permanentes →
   `saveFailedDownload` (reusa `FAILED_CACHE_FILE`).
7. Guardar mensajes nuevos en `group_messages_cache.json` (mismo formato —
   backfill histórico para re-runs sin consultar Evolution; formato reusa
   `loadCachedGroupMessages`/`saveCachedGroupMessages`).
8. `createWordDocument(receipts)` (reutilizada tal cual) → `Comprobantes_Descargados.docx`.
9. Resumen: totales, descargadas, fallidas, fuera de rango (guard de fecha, defensa
   en profundidad: `isMsgInRange` sigue en el path de descarga).

**index.js (baileys) se conserva SIN cambios** (referencia / rollback; las suites
de test existentes lo requieren por ruta). package.json `"start"` apunta al nuevo
entry `index-evolution.js`; script `start:legacy` conserva el viejo.

## 4. Flujo de datos

```
Teléfono SIM ──QR una vez──▶ Evolution en VPS (API + Postgres + Redis)
                                 │ REST (apikey, https)
                                 ▼
        index-evolution.js (Mac/Windows)
   - picker (rango de fechas)
   - findMediaMessages → findMessages paginado + filtro client-side
   - downloadMedia → getBase64FromMediaMessage → media_cache/<id>.jpg
   - createWordDocument (reusada) → Comprobantes_Descargados.docx
```

## 5. Manejo de errores

- **Evolution caído / inalcanzable** → mensaje claro: revisar el VPS (URL, API key,
  instancia conectada en la UI de Evolution).
- **Instancia desconectada / QR expirado** → avisar que el teléfono de la SIM debe
  estar enchufado y con WhatsApp conectado; reescaneo solo desde la UI de Evolution.
- **Media no descargable** (CDN expirado + sin archivo local): se cuenta como
  fallida y se guarda en failed cache; se reporta al final con su fecha.
- **Fallos de red al consultar**: reintento x3 con backoff corto; si persiste, abortar
  con mensaje (no producir Word a medias).

## 6. Migración y compatibilidad

- `media_cache/` y `group_messages_cache.json` actuales **se conservan** — formato
  idéntico, primera corrida con Evolution suma lo nuevo al mismo caché.
- Los Word ya generados no se tocan.
- Historial anterior al alta de la instancia no existe en Evolution; queda cubierto
  por el caché local existente.
- El usuario de Windows corre la MISMA app nueva apuntando al MISMO Evolution
  en el VPS (misma `EVOLUTION_URL` + API key en su config.json).

## 7. Testing

- **Unit (nuevo test):** normalización de records → forma baileys; filtrado
  client-side (jid + rango + media) contra fixtures; contacts map.
- **Unit (reusados):** `isMsgInRange` (12 checks) — la guardia sigue vigente.
- **E2E con mock HTTP (nuevo):** servidor local simulado de Evolution
  (`findMessages` + `getBase64FromMediaMessage`) con 2-3 imágenes fixture →
  pipeline completo → docx válido (mismo chequeo `unzip -l` de media embebida).
- **E2E offline (reusado):** `test_e2e_offline.js` debe seguir verde (path de
  Word sin cambios).

## 8. Riesgos

| Riesgo | Mitigación |
|---|---|
| Baneo del número (canal no oficial) | SIM dedicada; uso pasivo (solo leer); reenlazar = reescaneo una vez |
| Teléfono de la SIM apagado/sin red | Regla operativa; Evolution reintenta; la app avisa si la instancia está desconectada |
| Media del CDN expirada | Evolution la guarda en disco al recibirla; app cachea en media_cache; re-upload como último recurso |
| remoteJid filter roto (#1632) | Filtrado client-side siempre |
| Fuga de fechas | Guardia `isMsgInRange` en el path de descarga (defensa en profundidad) |
| TLS: Traefik sirve cert default autofirmado | `CA_FILE` con el cert del VPS (sin desactivar verificación); fix real: Let's Encrypt en Traefik (README) |
