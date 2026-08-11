# Design: Mini App de Facturación en Apps Script

Fecha: 2026-08-11
Estado: Aprobado (pendiente review del usuario)

## Propósito

Aplicación web en Apps Script (Google Apps Script / HtmlService) para un solo usuario que permite:
- Facturar (crear, editar, ver, eliminar facturas)
- Registrar y gestionar clientes
- Llevar control de gastos administrativos
- Generar reportes mensuales
- Generar PDF de facturas bajo demanda

## Decisiones principales

| Decisión | Opción elegida |
|---|---|
| Interfaz | Web app SPA en vanilla JS (sin CDN, sin dependencias externas) |
| Almacenamiento | Google Sheets (una hoja, una pestaña por módulo) |
| Impuestos | IVA configurable, facturas internas (sin cumplimiento fiscal CFDI/SAT) |
| PDF | Generado bajo demanda desde plantilla HTML, sin archivo automático en Drive |
| Folio | Prefijo personalizado + contador auto-incremental |
| Moneda | Configurable (símbolo + código) |
| Usuarios | Un solo usuario |
| Seguimiento | Estado por factura: pendiente / pagada, con fecha de pago |

## Arquitectura

- **Frontend**: `index.html` (SPA). Barra lateral con vistas: Dashboard, Facturas, Clientes, Gastos, Reportes, Configuración. Comunicación con el servidor vía `google.script.run`.
- **Backend**: `Code.js` (funciones de servidor con `google.script.run`), `Aux.js` (helpers puros: cálculos, formato moneda/fecha, normalización).
- **CSS**: `estilos.css` inline en el HTML (HtmlService no permite archivos CSS separados sin plantillas; se usará una etiqueta `<style>`).
- **PDF**: plantilla HTML de la factura → `HtmlService.createHtmlOutput()` → `.asDownload()` con nombre `FAC-xxx.pdf`.
- **Despliegue**: Web App con acceso "Solo yo".

## Estructura de datos (Google Sheets)

4 pestañas. Fila 1 = cabecera congelada. IDs generados automáticamente (timestamp-based).

### Hoja `Config` (clave-valor)

| Clave | Valor (ejemplo) |
|---|---|
| `empresa_nombre` | "Mi Empresa S.A." |
| `empresa_rfc` | "XAXX010101000" |
| `empresa_direccion` | "Calle 1 #2" |
| `empresa_telefono` | "+52..." |
| `empresa_email` | "x@y.com" |
| `empresa_logo` | URL del logo (vacío = sin logo) |
| `moneda` | "USD" |
| `moneda_simbolo` | "$" |
| `prefijo_folio` | "FAC-" |
| `contador_folio` | 5 |
| `iva_porcentaje` | 16 |
| `categorias_gastos` | "Renta,Internet,Papelería,Servicios" |

### Hoja `Clientes`
`id_cliente, nombre, rfc, email, telefono, direccion, fecha_registro`

### Hoja `Facturas`
`id_factura, folio, id_cliente, nombre_cliente, fecha_emision, fecha_vencimiento, subtotal, iva, total, estado, fecha_pago, notas`

- `nombre_cliente` está denormalizado: la factura conserva el dato aunque se edite el cliente.
- `estado`: `pendiente` | `pagada`.

### Hoja `Factura_Items`
`id_factura, descripcion, cantidad, precio_unitario, importe`
Relación 1:N con `Facturas` (una fila por concepto).

### Hoja `Gastos`
`id_gasto, fecha, categoria, descripcion, monto, metodo_pago, proveedor`

## Interfaz (SPA)

- Barra lateral izquierda con 5 vistas. Cabecera con nombre de empresa y moneda.
- **Dashboard**: tarjetas Total facturado (mes), Pendiente de cobro, Total gastos (mes), Utilidad (mes). Botones rápidos: Nueva factura, Registrar gasto.
- **Facturas**: tabla con filtros por estado y mes. Acciones por fila: Ver/Editar, Generar PDF, Marcar pagada/pendiente, Eliminar.
- **Nueva factura**: seleccionar cliente (o crear rápido), items dinámicos (descripción, cantidad, precio), cálculo automático de subtotal/IVA/total, folio asignado al guardar.
- **Clientes**: tabla + formulario (crear/editar/eliminar) + búsqueda.
- **Gastos**: formulario (fecha, categoría de lista config, descripción, monto, método, proveedor) + tabla con filtros.
- **Reportes**: mensual — facturado, cobrado, pendiente, gastos por categoría, utilidad, top 5 clientes.
- **Configuración**: formulario de empresa, moneda, prefijo folio, % IVA, categorías de gastos.

## API del servidor

Todas las funciones devuelven `{ok:true, data}` o `{ok:false, error}`.

| Función | Propósito |
|---|---|
| `getConfig()` | Leer Config |
| `saveConfig(obj)` | Actualizar Config |
| `listClientes()` | Todos los clientes |
| `saveCliente(obj)` | Crear/editar cliente |
| `deleteCliente(id)` | Eliminar (bloqueado si tiene facturas) |
| `listFacturas(filtros)` | Facturas con filtros estado/mes |
| `getFactura(id)` | Factura + items |
| `createFactura(obj)` | Cabecera + items, asigna folio con `LockService` |
| `setEstadoFactura(id, estado)` | Marcar pendiente/pagada |
| `deleteFactura(id)` | Cabecera + items |
| `listGastos(filtros)` | Filtros mes/categoría |
| `saveGasto(obj)` | Crear/editar gasto |
| `deleteGasto(id)` | Eliminar gasto |
| `getReportes(mes)` | Totales, gastos por categoría, top clientes |
| `buildPDF(idFactura)` | HTML factura → PDF |

## Manejo de errores

- Try/catch en toda función de servidor → `{ok, data|error}`.
- El cliente muestra toasts de error.
- Validación server-side: campos requeridos, montos > 0, cliente existente, folio único.
- `LockService` para el contador de folio (evita duplicados en concurrencia).

## Seguridad

- Web App acceso "Solo yo" (un solo usuario).
- Sin secretos en el código. Las URLs de logo son opcionales y sanitizadas.

## Testing

- Helpers puros (`Aux.js`) separados de I/O para lógica testeable.
- Script de setup que crea y llena la hoja con datos de ejemplo.
- Función `tearDown()` que elimina datos de prueba.
- Sin framework de tests (no disponible en Apps Script nativo); verificación manual guiada.

## Archivos del proyecto

```
appscript.json            # manifest de Apps Script
Code.gs                   # funciones de servidor
Aux.gs                    # helpers puros
index.html                # SPA (HTML + CSS + JS inline)
scripts/setup_data.gs     # datos de ejemplo para pruebas
README.md                 # instrucciones de despliegue
```

## Fuera de alcance (YAGNI)

- Facturación electrónica / timbrado fiscal (CFDI/SAT u otro).
- Multi-usuario con roles.
- Pagos parciales / abonos.
- Archivo automático de PDFs en Drive (los PDF se generan bajo demanda).
- Recordatorios de cobro automáticos.
