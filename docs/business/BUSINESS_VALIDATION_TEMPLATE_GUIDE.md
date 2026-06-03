# Plantilla de Validación de Negocio

Este archivo acompaña la plantilla de Excel `business_validation_template.xlsx`.

Su objetivo es ayudar al dueño del proceso a confirmar las reglas reales del SOP para que la automatización deje de:

- inferir valores dudosos
- dejar espacios en blanco evitables
- depender de interpretación manual

## Cómo usar la plantilla

Llena el archivo Excel hoja por hoja.

No hace falta conocer código.

La idea es responder:

1. Qué campo se llena realmente en Lucy.
2. De dónde sale ese dato.
3. Qué se hace cuando falta o viene ambiguo.
4. Qué ejemplo real sirve como referencia.

## Hojas del archivo

### `01_Instrucciones`
Resumen del propósito de la plantilla y cómo completarla.

### `02_Header_Fields`
Campos generales de la factura en Lucy.

Aquí se debe indicar para cada campo:

- si se usa o no
- si es obligatorio
- de qué archivo sale
- cuál es la regla exacta
- qué hacer si no se encuentra

### `03_Allocated_Costs`
Campos que se llenan en la grilla de `Allocated Costs`.

Sirve para confirmar:

- cuenta contable
- posting key
- tax code
- cost center
- profit center
- texto
- monto

### `04_Distribution_and_Splits`
Reglas de 33%, 50% o cualquier otro split.

Aquí se debe aclarar:

- cómo se detecta el split
- cuántos posteos salen
- cómo se reparte subtotal, IVA y retenciones
- si el PDF por sí solo define algo o si siempre manda el archivo de distribución

### `05_PDF_Formats_and_Taxes`
Casos de facturas por formato.

Aquí se documentan reglas por tipo:

- IVA 0%
- IVA claro
- incentivo/descuento
- saldo anterior
- retenciones
- múltiples conceptos

### `06_Real_Invoice_Cases`
Casos reales completos.

Cada fila idealmente representa una factura ya resuelta por el usuario.

Esto es lo más valioso para volver la automatización 100% real.

### `07_Blocking_Rules`
Define qué pasa si falta un dato.

Ejemplos:

- detener factura
- dejar flag
- detener todo el lote

### `08_Open_Questions`
Preguntas pendientes que negocio debe responder.

## Qué es más importante llenar primero

Si el tiempo es limitado, prioriza estas hojas:

1. `06_Real_Invoice_Cases`
2. `02_Header_Fields`
3. `03_Allocated_Costs`
4. `04_Distribution_and_Splits`
5. `07_Blocking_Rules`

## Resultado esperado

Cuando esta plantilla esté llena, podremos:

- eliminar reglas inventadas
- reducir campos vacíos
- decidir qué sí sale automático y qué no
- dejar una versión mucho más estable para el bot

