# App de Facturación (Google Apps Script)

Mini app de facturación web para uso personal: crea facturas con **PDF bajo demanda**,
registra clientes, controla gastos administrativos y genera **reportes mensuales**.
Datos en Google Sheets, interfaz web (HTML + vanilla JS), sin dependencias externas.

> Cada usuario despliega su **propia instancia** con su cuenta de Google.
> La app es de un solo usuario: no tiene login ni comparte datos entre cuentas.

[![Deploy con Apps Script](https://img.shields.io/badge/Apps%20Script-Deploy%20flow-blue)](https://developers.google.com/apps-script/guides/clasp)

---

## Requisitos

- Una cuenta de Google.
- Node.js (para `clasp`). Si no lo tienes: [nodejs.org](https://nodejs.org).

---

## Opción A — Despliegue automático (recomendado)

Instala `clasp` (una sola vez) y ejecuta el script de despliegue:

```bash
# 1. Instala clasp
npm install -g @google/clasp

# 2. Inicia sesión con tu cuenta de Google
clasp login

# 3. Clona el proyecto (solo si lo bajaste de GitHub)
git clone https://github.com/miguetech/appscript-finance-tracker.git
cd appscript-finance-tracker

# 4. Despliega (crea Spreadsheet + proyecto Apps Script + sube archivos + tests + web app)
./scripts/deploy.sh
```

Cuando termine, el script te imprime la **URL de tu app**. Guárdala.

> Si ya tienes un proyecto Apps Script y quieres desplegarlo ahí:
> ```bash
> ./scripts/deploy.sh <scriptId>
> ```
> El `scriptId` se ve en el editor de Apps Script → Configuración del proyecto (engranaje).

---

## Opción B — Despliegue manual (sin instalar nada)

1. Crea un **Google Spreadsheet** nuevo (será el contenedor de tus datos y de tu script).
2. En el spreadsheet: **Extensiones → Apps Script**.
3. Borra el contenido por defecto y crea estos archivos (menú **+** → Archivo), pegando el
   contenido de cada uno del repo:
   - `appsscript.json`
   - `Code.gs`
   - `Data.gs`
   - `Aux.gs`
   - `index.html`
   - `pdfTemplate.html`
4. Guarda el proyecto (Ctrl+S).
5. En la barra de funciones, selecciona `ensureSheets` y pulsa **Ejecutar** para crear las
   5 hojas y la configuración por defecto. Autoriza cuando lo pida.
6. Selecciona `runTests` y pulsa **Ejecutar**. Todas deben salir con `PASS`.

---

## Publicar la app (Web App)

Con el script de despliegue se hace solo. Manualmente, en el editor de Apps Script:

1. **Implementar → Nueva implementación → Aplicación web**.
2. **Ejecutar como**: *Yo*.
3. **Tener acceso**: *Solo yo* ← importante: así nadie más ve tus datos.
4. **Implementar** y copia la URL de la app.
5. Abre la URL en tu navegador. Verás el **Dashboard**.

---

## Configurar tus datos

La primera vez, ve a la vista **Configuración** de la app (o edita la pestaña `Config`
del spreadsheet) y pon:

| Campo | Qué es |
|---|---|
| Nombre / RFC / Dirección / Tel / Email | Datos del emisor que salen en la factura |
| Logo (URL) | Logo de empresa que sale en el PDF |
| Moneda y símbolo | Ej. `USD` / `$`, o `MXN` / `$` |
| Prefijo de folio | Ej. `FAC-` (los folios quedan `FAC-001`, `FAC-002`, …) |
| Contador actual | Número del siguiente folio |
| IVA % | Ej. `16` |
| Categorías de gastos | Separadas por coma: `Renta,Internet,Papelería,Servicios` |

---

## Uso diario

| Vista | Qué puedes hacer |
|---|---|
| **Dashboard** | Totales del mes: facturado, pendiente de cobro, gastos, utilidad |
| **Facturas** | Crear (con conceptos dinámicos + IVA automático), ver, marcar pagada/pendiente, **descargar PDF**, eliminar |
| **Clientes** | Crear, editar, buscar, eliminar (bloqueado si tiene facturas) |
| **Gastos** | Registrar con categoría y método de pago, filtrar por mes/categoría |
| **Reportes** | Facturado/cobrado/pendiente/gastos/utilidad del mes, gastos por categoría, top 5 clientes |
| **Configuración** | Empresa, moneda, folio, IVA, categorías |

---

## Auto-despliegue con GitHub Actions (opcional)

Cada vez que hagas `git push` a `main`, un Action sube el código a tu proyecto
Apps Script y corre los tests. Para configurarlo:

```bash
# 1. Obtén tu scriptId (editor Apps Script → Configuración del proyecto → ID)
# 2. Obtén tu token de clasp (~/.clasprc.json se crea tras `clasp login`)
# 3. Guarda los secretos en tu repo (Settings → Secrets and variables → Actions)
gh secret set CLASP_SCRIPT_ID --body "<tu-scriptId>"
gh secret set CLASPRC_JSON < ~/.clasprc.json
```

> **Seguridad**: estos secrets están cifrados en GitHub y solo son visibles para
> los Actions de tu repo. El repositorio público **no** expone sus valores, ni
> tus datos ni tu proyecto de Apps Script. Es el flujo estándar de CI/CD.
> El Action despliega a **un** scriptId — cada usuario sigue desplegando su
> propia instancia con `./scripts/deploy.sh` o manualmente.

---

## Estructura del proyecto

```
appscript-finance-tracker/
├── appsscript.json      # manifest (acceso "Solo yo")
├── Code.gs             # API del servidor (google.script.run) + tests
├── Data.gs             # acceso a Google Sheets (5 hojas)
├── Aux.gs              # helpers puros (totales, formato, validación, escape HTML)
├── index.html          # SPA del frontend (vanilla JS, sin CDN)
├── pdfTemplate.html    # plantilla de la factura para PDF
├── scripts/deploy.sh   # despliegue automático con clasp
└── docs/               # diseño (spec) y plan de implementación
```

---

## Notas

- Los folios usan prefijo + contador autoincremental protegido con `LockService` (sin duplicados).
- El IVA se calcula sobre el subtotal con el porcentaje configurado.
- Los PDF se generan bajo demanda; no se almacenan automáticamente.
- Facturas **internas**: no hay timbrado fiscal (CFDI/SAT). Usa un servicio externo si lo necesitas.
- Sin secretos ni datos en el repo: tu información vive solo en tu Spreadsheet.
