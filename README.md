# Comprobantes Digitales SSA

Sistema de recepción de comprobantes (respaldo de nómina) para las 41 unidades
de salud de Tamaulipas, organizado por trimestre y quincena.

## Estructura de este repositorio

```
comprobantes-digitales-ssa/
├── apps-script/
│   ├── comprobantes-digitales-v9/   ← Script principal EN PRODUCCIÓN
│   │   ├── Codigo.gs                  (login de unidades + panel admin por API)
│   │   ├── Llenado_2do_Trimestre.gs   (script que se corrió una vez para poblar Q2)
│   │   ├── Limpiar_Nomina_3T.gs       (limpieza puntual ya aplicada a Q3)
│   │   └── appsscript.json
│   └── comprobantes-movil-legacy/   ← Versión anterior (v7), ya no se usa
│       ├── Codigo.gs
│       └── appsscript.json
├── q3-trimestre/                    ← NUEVO: 3er trimestre como Sheet/Script independiente
│   ├── Codigo_Q3.gs                   (copia de Codigo.gs apuntando al Sheet de Q3)
│   ├── GenerarQ3.gs                   (llena el Sheet caminando la carpeta de Drive)
│   └── appsscript.json
├── admin-dashboard/                 ← NUEVO: panel visual para administrar trimestres
│   ├── Admin.gs                       (backend: generar/clonar estructura de un trimestre)
│   ├── AdminPanel.html                (interfaz: botones Q1–Q4 + prompt de carpetas especiales)
│   └── appsscript.json
└── docs/
    └── notas.md
```

## Frontends (HTML servidos en este repo / Netlify)

- `index.html` — frontend EN PRODUCCIÓN que usan las 41 unidades hoy (apunta al Web App principal que mezcla todos los trimestres).
- `index_2do.html` — copia archivada apuntando al deployment usado durante el 2° trimestre.
- `index_3er.html` — **nuevo**: mismo frontend, apuntando ya al Web App independiente del 3er trimestre:
  `https://script.google.com/macros/s/AKfycbzsj0wcVMnGxiL4jsdoHdLsiPTSOHDenGdmhloKIy-LY3mJHgVXTp9Q6q07OaS5ICx0Rw/exec`
  Cuando quieras que las unidades usen el sistema de Q3 en vez del mezclado, sube este archivo como `index.html` en Netlify (o cambia el dominio a apuntar aquí).
- `admin.html` — panel de administrador anterior (simple). El nuevo panel (`admin-dashboard/AdminPanel.html`) vive como Apps Script propio, ver abajo.
- `guia-comprobantes.html` — guía para las unidades, sin cambios.

## Qué cambia respecto al sistema actual

**Antes:** un solo Google Sheet ("SST_Todos_Trimestres") con TODAS las quincenas
de TODOS los trimestres mezcladas en una pestaña. Esto es lo que se iba
poniendo lento y frágil con cada trimestre nuevo.

**Ahora:**
- El 3er trimestre vive en su propio Sheet independiente (`SST_3er_Trimestre`)
  y su propio Web App (`Codigo_Q3.gs`), sin tocar el sistema de Q1/Q2.
- Cada trimestre futuro (Q4, y en adelante) se genera desde el **panel de
  administrador** en un par de clics, en vez de repetir el trabajo manual de
  crear carpetas y llenar el Sheet a mano.
- Todo el código y los Sheets quedan organizados en una sola carpeta de Drive:
  **"Comprobantes Digitales SSA"**.

## Cómo desplegar cada pieza

### 1. Q3 independiente (`q3-trimestre/`)
El Sheet `SST_3er_Trimestre` ya está creado en Drive, dentro de la carpeta
"Comprobantes Digitales SSA". Falta:
1. Abrir Google Drive → Nuevo → Más → Google Apps Script.
2. Renombrar el proyecto a `Comprobantes_Digitales_Q3`.
3. Pegar el contenido de `Codigo_Q3.gs` y `GenerarQ3.gs` (cada uno en su
   propio archivo dentro del proyecto), y copiar el `appsscript.json` en
   configuración del proyecto (⚙️ → Mostrar archivo "appsscript.json").
4. Si el 3er trimestre tiene alguna carpeta especial (bono), edítala en el
   objeto `BONOS_Q3` al inicio de `GenerarQ3.gs` antes de correrlo.
5. Ejecutar UNA VEZ la función `generarTercerTrimestre` desde el editor
   (▶ Ejecutar). Revisa el log: debe decir "✅ Filas agregadas: N".
6. Implementar → Nueva implementación → Tipo: App web → Ejecutar como: Yo →
   Acceso: Cualquiera. Copia la URL nueva — esa es la que va en el HTML que
   suben a Netlify para las unidades del 3er trimestre (o pueden mantener el
   mismo HTML y solo intercambiar la URL del backend).

### 2. Panel de administrador (`admin-dashboard/`)
1. Nuevo proyecto de Apps Script, `Admin_Comprobantes_SSA`.
2. Pegar `Admin.gs` y `AdminPanel.html` (como archivo HTML), más el
   `appsscript.json`.
3. Antes de implementar, cambia `ADMIN_PASS_DASH` en `Admin.gs` por una
   contraseña tuya.
4. Implementar → Nueva implementación → Tipo: App web → Ejecutar como: Yo →
   Acceso: **Solo tú** (este panel no es para las unidades, es solo para ti).
5. Abre la URL, entra con `ADMIN_SST` / la contraseña que pusiste, y ahí
   verás las 4 tarjetas de trimestre. Q1, Q2 y Q3 ya aparecen como
   "Generado" (apuntando a lo que ya existe); Q4 aparece "Pendiente" — al
   darle clic te pregunta qué carpetas especiales aplican y genera todo solo.

> Nota sobre rendimiento: generar un trimestre completo crea cientos de
> carpetas (41 unidades × ~13 tipos × 2 nóminas × 6 quincenas). Apps Script
> tiene un límite de ~6 minutos de ejecución; si el trimestre es muy grande
> puede que necesites correrlo dos veces (el código no duplica carpetas que
> ya existan, así que es seguro reintentar).

### 3. GitHub
No tengo una conexión activa a GitHub desde este entorno (no hay conector ni
token configurado), así que no puedo hacer el push por ti. Te dejo todo este
repositorio ya armado — solo necesitas:

```bash
cd comprobantes-digitales-ssa
git init
git add .
git commit -m "Estructura inicial: Q3 independiente + panel admin"
git remote add origin https://github.com/<tu-usuario>/<tu-repo>.git
git push -u origin main
```

Si más adelante conectas un conector de GitHub en Claude, dímelo y puedo
ayudarte a automatizar el push directamente.

## Dónde está todo en Drive

- Carpeta contenedora: **Comprobantes Digitales SSA**
  - `Comprobantes_Digitales` (script en producción, v9)
  - `Comprobantes_Movil` (script legado, ya no se usa)
  - `SST_Todos_Trimestres` (Sheet histórico: Q1 y Q2)
  - `SST_3er_Trimestre` (Sheet nuevo, independiente, solo Q3)
  - `Control_Acceso` (habilitado/deshabilitado por unidad — compartido por todos los trimestres)
- La carpeta de carpetas de comprobantes en sí ("AÑO 2026" → 1°/2°/3° TRIMESTRE
  → QNA → nómina → tipo → unidad) se quedó donde estaba, porque la usan las
  41 unidades con links ya compartidos — moverla podría romper esos accesos.
