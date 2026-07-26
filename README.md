# TOLIMAX · Aplicativo de Pedidos (web)

- **index.html** — App del despachador (solo producto y cantidad, sin precios).
- **panel.html** — Panel contable (precios, impuestos ICUI 20% + IVA 5%, informe interactivo, alta de clientes y personas).
- **data.js / data.json** — Base de datos (838 clientes + 12 productos con precio base y de venta).
- **logo.jpeg** — Logo TOLIMAX.
- **TOLIMAX_Sincronizacion_Drive.gs** — Conector Google Apps Script (Drive).

## Publicar en GitHub Pages
1. Crea un repositorio público (ej. `tolimax`) y sube estos archivos.
2. Settings → Pages → Source: rama `main`, carpeta `/root` → Save.
3. La app queda en `https://TU-USUARIO.github.io/tolimax/` (despachador) y `.../panel.html` (contable).

## Cargar los 838 clientes en Drive
En el `.gs`, pon esa URL de Pages + `/data.json` en `DATA_URL` y ejecuta la función `importarClientes` una vez.
