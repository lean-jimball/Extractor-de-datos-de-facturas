# Libro de Facturas
 
Interfaz web para extraer automáticamente los datos y conceptos de facturas de proveedores (imagen o PDF) y exportarlos a una hoja de Excel.
 
## Qué hace
 
- Acepta facturas en **JPG, PNG, WEBP o PDF**, una o varias a la vez (arrastrar y soltar o selección manual).
- Extrae automáticamente:
  - Proveedor, CIF/NIF, número de factura, fecha y moneda
  - **Todos los conceptos** de la factura: descripción, cantidad, precio unitario, % de impuesto e importe
  - Subtotal, impuestos y total
- Permite **revisar y editar** cualquier campo o línea de concepto antes de exportar (añadir o eliminar conceptos incluido).
- Exporta todo a un archivo **`.xlsx`** con dos hojas:
  - **Resumen facturas** — una fila por factura con sus totales
  - **Conceptos** — el detalle línea por línea de todas las facturas cargadas
Los datos se procesan solo en la sesión del navegador; no se almacenan en ningún servidor.
 
## Cómo usarlo
 
1. Sube una o varias facturas.
2. Espera a que cada una se procese (aparece marcada como "extraída" en la bandeja).
3. Revisa los datos y conceptos extraídos; corrige a mano lo que haga falta.
4. Pulsa **Exportar a Excel** para descargar el archivo `.xlsx`.
## Tecnología
 
- React
- [SheetJS (xlsx)](https://docs.sheetjs.com/) para generar el archivo Excel
- [lucide-react](https://lucide.dev/) para los iconos
- Extracción de datos mediante la API de Claude (visión) de Anthropic
## Notas
 
- La calidad de la extracción depende de la nitidez del documento. Con imágenes borrosas o PDFs escaneados de baja calidad, conviene revisar los importes antes de exportar.
- Cuando un dato no se puede determinar con certeza, el sistema lo deja en blanco en lugar de inventarlo.
## Licencia
 
MIT
