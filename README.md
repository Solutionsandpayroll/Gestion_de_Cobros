# Web Cuentas por Cobrar — Tesorería

Aplicación web interna de **Solutions & Payroll** para gestionar cobros por cliente a partir de la plantilla SIGO de cuentas por cobrar.

## ¿Qué hace esta aplicación?

1. El usuario sube el archivo Excel exportado desde SIGO (plantilla de cuentas por cobrar).
2. La app lee automáticamente la columna **Cliente** de la fila 7 y lista todos los clientes únicos encontrados.
3. El usuario puede seleccionar o deseleccionar los clientes que desea incluir en el reporte (todos seleccionados por defecto). Incluye buscador de clientes.
4. Al presionar **Generar Formato**, se produce un archivo Excel basado en la plantilla interna con el siguiente procesamiento:
   - Datos agrupados por cliente, con fila **"Total [cliente]"** al final de cada grupo (en negrilla).
   - Columnas con fórmulas automáticas en cada fila y subtotal.
   - Filas fijas de la plantilla (7–23) al final del archivo con fórmulas dinámicas actualizadas.
   - Filas de datos agrupadas con la función **Agrupar** de Excel (nivel 1), colapsables por cliente.

## Tecnologías

- **React 18** + **Vite**
- **SheetJS (xlsx)** — lectura del archivo SIGO
- **ExcelJS** — escritura del archivo resultante con estilos, fórmulas y agrupación
- **JSZip** — corrección de shared formulas antes de cargar en ExcelJS

## Estructura del proyecto

```
Web cuentas por cobrar/
├── public/
│   ├── Logo syp.png
│   └── Plantilla - Cuentas por cobrar detallada.xlsx   ← plantilla base del reporte
├── src/
│   ├── App.jsx          ← lógica principal
│   ├── App.css          ← estilos
│   ├── index.css        ← estilos globales
│   └── main.jsx         ← entry point
├── index.html
├── package.json
└── vite.config.js
```

## Iniciar en desarrollo

```bash
npm install
npm run dev
```

> **Nota:** `npm run dev` usa `node node_modules/vite/bin/vite.js` directamente para evitar un problema con el carácter `&` en la ruta de la carpeta `S&P` de Windows, que rompe los scripts `.cmd`.

## Campos trasladados del SIGO a la plantilla

Los siguientes encabezados se leen desde la **fila 7** del Excel SIGO:

| Campo SIGO              | Columna en plantilla    |
|-------------------------|-------------------------|
| Cliente                 | Cliente                 |
| Documento               | Documento               |
| Fecha vencimiento       | Fecha vencimiento       |
| Vencido 1 a 30          | Vencido 1 a 30          |
| Vencido 31 a 60         | Vencido 31 a 60         |
| Vencido 61 a 90         | Vencido 61 a 90         |
| Vencido más de 91       | Vencido más de 91       |
| Saldo por vencer        | Saldo por vencer        |

## Fórmulas generadas automáticamente

### En cada fila de datos
| Columna        | Fórmula                                              |
|----------------|------------------------------------------------------|
| Total cartera  | `=SUM(Vencido1a30 : SaldoPorVencer)` de esa fila     |

### En cada fila "Total [cliente]" (subtotal por cliente)
| Columna           | Fórmula                                        |
|-------------------|------------------------------------------------|
| Total cartera     | `=SUBTOTAL(9, rango del cliente)`              |
| Vencido 1 a 30    | `=SUBTOTAL(9, rango del cliente)`              |
| Vencido 31 a 60   | `=SUBTOTAL(9, rango del cliente)`              |
| Vencido 61 a 90   | `=SUBTOTAL(9, rango del cliente)`              |
| Vencido más de 91 | `=SUBTOTAL(9, rango del cliente)`              |
| Saldo por vencer  | `=SUBTOTAL(9, rango del cliente)`              |
| %                 | `=IF($TotalCarteraFila7=0, 0, TotalCartera / $TotalCarteraFila7)` en formato `0.00%` |

### En las filas fijas al final (originalmente filas 7–23 de la plantilla)

| Fila plantilla | Columna            | Fórmula generada                                                    |
|---------------|--------------------|---------------------------------------------------------------------|
| 7             | Total cartera + Vencidos + Saldo | `=SUBTOTAL(9, X2:X{últimaFila})`                  |
| 7             | %                  | `=SUM(%)` de todo el rango de datos                                 |
| 8             | Cliente            | `Cartera Grupo HL {MesActual}` (actualiza el mes automáticamente)   |
| 8             | Total cartera + Vencidos + Saldo | `=SUM(celdas de clientes que contienen "HL")`     |
| 9             | Vencido 1 a 30     | `=SUM(Vencido31, Vencido61, Vencido91)` de la fila 7               |
| 9             | %                  | `=IF($G$fila9=0, 0, Vencido1a30_fila9 / $G$fila9)`                 |
| 10            | Vencido 1 a 30     | `=SUM(Vencido31, Vencido61, Vencido91)` de la fila 8               |
| 10            | %                  | `=IF($G$fila9=0, 0, Vencido1a30_fila10 / $G$fila9)`                |
| 13            | Vencido 1 a 30     | `=fila9 - fila10`                                                   |
| 14–17         | Vencido 1 a 30     | `=SUM(Vencido31+Vencido61+Vencido91)` del subtotal del cliente específico¹ |
| 14–18         | %                  | `=IF($G$fila13=0, 0, G{filaActual} / $G$fila13)` en formato `0.00%` |
| 18            | Vencido 1 a 30     | `=fila13 - SUM(fila14:fila17)`                                      |
| 21            | Vencido 1 a 30     | `=fila10 + fila18`                                                  |
| 22            | Vencido 1 a 30     | `=SUM(fila14:fila17)`                                               |
| 23            | Vencido 1 a 30     | `=SUM(fila21:fila22)`                                               |

¹ **Clientes fijos para filas 14–17:**
- Fila 14: PUNTO MEDICAL DISTRIBUCIONES SAS
- Fila 15: REJIMETAL SAS
- Fila 16: DINTERWEB SAS
- Fila 17: SOLUTIONS & PAYROLL PERU S.A.C

## Agrupación de filas

Las filas de datos de cada cliente tienen `outlineLevel = 1` en Excel. Esto permite colapsar cada grupo con el botón `[-]` dejando visibles solo las filas "Total [cliente]". La hoja está configurada con `summaryBelow: true` (resumen debajo del grupo).

## Notas

- La plantilla debe estar en `public/Plantilla - Cuentas por cobrar detallada.xlsx`.
- Los clientes con "HL" en el nombre se detectan automáticamente para la fila 8 — p. ej. "HL INFRAESTRUCTURAS SAS", "CONSORCIO SK-HL", "CONSORCIO HL - A&D".
- El mes en "Cartera Grupo HL {mes}" se toma de la fecha del sistema al momento de generar.
- Todas las referencias de fila en las fórmulas de las filas fijas son dinámicas: se calculan según cuántos clientes y registros haya en el SIGO cargado.

## Repositorio

GitHub: [Solutionsandpayroll/Gestion_de_Cobros](https://github.com/Solutionsandpayroll/Gestion_de_Cobros)

## Licencia

© 2026 Solutions & Payroll. Uso interno.


## Tecnologías

- **React 18** + **Vite**
- **SheetJS (xlsx)** — lectura del archivo SIGO
- **ExcelJS** — escritura del archivo resultante con estilos
- **JSZip** — corrección de shared formulas antes de cargar en ExcelJS

## Estructura del proyecto

```
Web cuentas por cobrar/
├── public/
│   ├── Logo syp.png
│   └── Plantilla - Cuentas por cobrar detallada.xlsx   ← plantilla base del reporte
├── src/
│   ├── App.jsx          ← lógica principal
│   ├── App.css          ← estilos
│   ├── index.css        ← estilos globales
│   └── main.jsx         ← entry point
├── index.html
├── package.json
└── vite.config.js
```

## Iniciar en desarrollo

```bash
npm install
npm run dev
```

> **Nota:** `npm run dev` usa `node node_modules/vite/bin/vite.js` directamente para evitar un problema con el carácter `&` en la ruta de la carpeta `S&P` de Windows, que rompe los scripts `.cmd`.

## Campos que se trasladan del SIGO a la plantilla

Los siguientes encabezados se leen desde la **fila 7** del Excel SIGO y se escriben en la **fila 1** de la plantilla resultante con el mismo nombre de columna:

| Campo SIGO              | Columna en plantilla    |
|-------------------------|-------------------------|
| Cliente                 | Cliente                 |
| Documento               | Documento               |
| Fecha vencimiento       | Fecha vencimiento       |
| Vencido 1 a 30          | Vencido 1 a 30          |
| Vencido 31 a 60         | Vencido 31 a 60         |
| Vencido 61 a 90         | Vencido 61 a 90         |
| Vencido más de 91       | Vencido más de 91       |
| Saldo por vencer        | Saldo por vencer        |

## Repositorio
