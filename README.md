# App de Facturación (Google Apps Script)

Mini app de facturación para un solo usuario: crea facturas con PDF bajo demanda,
registra clientes, controla gastos administrativos y genera reportes mensuales.
Datos en Google Sheets, interfaz web (HTML + vanilla JS).

## Archivos

- `appscript.json` — manifest (acceso "Solo yo")
- `Code.gs` — API del servidor (google.script.run) + tests
- `Data.gs` — acceso a Google Sheets
- `Aux.gs` — helpers puros (totales, formato, validación)
- `index.html` — SPA del frontend
- `pdfTemplate.html` — plantilla de la factura para PDF
- `docs/superpowers/specs/2026-08-11-facturacion-design.md` — diseño completo
- `docs/superpowers/plans/2026-08-11-facturacion.md` — plan de implementación

## Instalación

1. Crea un Google Spreadsheet (será el contenedor de datos y de tu script).
2. Extensiones → Apps Script. Copia `appscript.json`, `Code.gs`, `Data.gs`, `Aux.gs`,
   `index.html` y `pdfTemplate.html` al proyecto (nombres idénticos).
3. En el editor, ejecuta `ensureSheets` (crea las 5 hojas y la config por defecto).
4. Ejecuta `runTests` y verifica que todo pase (PASS).
5. Configura tus datos: pestaña `Config` de la hoja (empresa, RFC, moneda, % IVA,
   prefijo de folio, categorías de gastos) o desde la web app (vista Configuración).

## Despliegue (Web App)

1. En el editor: Implementar → Nueva implementación → Aplicación web.
2. Ejecutar como: "Yo". Tener acceso: "Solo yo".
3. Implementa y abre la URL generada.

## Uso

- **Dashboard**: totales del mes.
- **Facturas**: crear, ver, marcar pagada/pendiente, PDF bajo demanda, eliminar.
- **Clientes**: crear, editar, buscar, eliminar (no se elimina si tiene facturas).
- **Gastos**: registrar con categoría, filtrar por mes/categoría.
- **Reportes**: facturado/cobrado/pendiente/gastos/utilidad, gastos por categoría, top clientes.
- **Configuración**: datos de empresa, moneda, prefijo de folio, IVA, categorías.

## Notas

- Los folios usan prefijo + contador autoincremental (protegido con LockService).
- El IVA se calcula sobre el subtotal con el porcentaje configurado.
- Los PDF se generan bajo demanda; no se almacenan automáticamente.
- Facturas internas: sin timbrado fiscal (CFDI/SAT).
