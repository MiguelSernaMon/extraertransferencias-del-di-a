# Extractor de Comprobantes (WhatsApp → Word)

Extrae los comprobantes de transferencia del grupo de WhatsApp y genera un
documento Word agrupado por mensajero.

## Modo Evolution API (único)

El extractor es un **cliente REST de Evolution API** hosteada en un VPS
(`https://transferencias.redpostal.co`). Evolution mantiene la sesión de
WhatsApp 24/7 en el VPS (SIM dedicada): el usuario final no escanea QR ni
deja el teléfono conectado.

### Setup

1. Copiar `config.example.json` → `config.json`:
   ```bash
   cp config.example.json config.json
   ```
2. Completar `config.json`:
   - `apiKey` — API key global de Evolution (VPS).
   - `instance` — nombre de la instancia (por defecto `"Personal"`).
   - `groupName` — filtro de grupo (matchea por nombre, ej. `"transferencias"`).
   - `evolutionUrl` — URL base del VPS.

`config.json` está en `.gitignore` (no se commitea).

### TLS

El VPS sirve el certificado con **Let's Encrypt activo en Traefik** → no hace
falta `caFile`; la verificación usa las CA del sistema.

Solo si el VPS usara un cert autofirmado, extraer el PEM y setearlo:

```bash
echo | openssl s_client -connect transferencias.redpostal.co:443 -servername transferencias.redpostal.co 2>/dev/null | openssl x509 -out traefik.crt
```

```json
{ "caFile": "traefik.crt" }
```

### Uso

```bash
npm start                # panel web (elige fechas + mapea mensajeros en el navegador)
npm run start:terminal   # sin navegador, fechas por terminal
```

El panel web abre en el navegador: se elige el rango de fechas y se pueden
**mapear remitentes** (los que llegan como ID sin nombre) a nombres reales —
quedan guardados en `nombres_mensajeros.json` para las próximas corridas.

### Operación

- El teléfono de la SIM debe estar enchufado y conectado a WhatsApp.
- Si la instancia se desconecta, escanear el QR una sola vez en la UI de
  Evolution del VPS (la sesión queda persistida en el VPS).
- Salida: `Comprobantes_Descargados.docx` (en esta carpeta).

### Windows

Mismo `config.json` (la URL apunta al VPS) — no depende de la máquina local.
