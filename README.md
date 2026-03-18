# Web Cuentas por Cobrar  Tesorería

Aplicación web interna de **Solutions & Payroll** para gestionar cobros por cliente a partir de la plantilla SIGO de cuentas por cobrar.

## Cómo usar esta aplicación

1. Sube el archivo Excel exportado desde SIGO.
2. Selecciona los clientes que deseas incluir en el reporte (todos vienen seleccionados por defecto y puedes usar el buscador).
3. Haz clic en **Generar Formato** para descargar el Excel final con los datos, subtotales, fórmulas y agrupaciones.

## Tecnologías

- **React 18** + **Vite**
- **SheetJS (xlsx)**  lectura del archivo SIGO
- **ExcelJS**  escritura del archivo resultante con estilos, fórmulas y agrupación
- **JSZip**  corrección de shared formulas antes de cargar en ExcelJS

## Estructura del proyecto

```
Web cuentas por cobrar/
 public/
    Logo syp.png
    Plantilla - Cuentas por cobrar detallada.xlsx
 src/
    App.jsx
    App.css
    index.css
    main.jsx
 index.html
 package.json
 vite.config.js
```

## Iniciar en desarrollo

```bash
npm install
npm run dev
```

> **Nota:** `npm run dev` usa `node node_modules/vite/bin/vite.js` directamente para evitar un problema con el carácter `&` en la ruta `S&P` de Windows, que rompe los scripts `.cmd`.

## Campos trasladados del SIGO a la plantilla

Los siguientes encabezados se leen desde la **fila 7** del Excel SIGO:

| Campo SIGO        | Columna en plantilla |
|-------------------|----------------------|
| Cliente           | Cliente              |
| Documento         | Documento            |
| Fecha vencimiento | Fecha vencimiento    |
| Vencido 1 a 30    | Vencido 1 a 30       |
| Vencido 31 a 60   | Vencido 31 a 60      |
| Vencido 61 a 90   | Vencido 61 a 90      |
| Vencido más de 91 | Vencido más de 91    |
| Saldo por vencer  | Saldo por vencer     |

## Fórmulas generadas automáticamente

### En cada fila de datos

| Columna       | Fórmula                                           |
|---------------|---------------------------------------------------|
| Total cartera | `=SUM(Vencido1a30 : SaldoPorVencer)` de esa fila  |

### En cada fila "Total [cliente]"

| Columna                                                                               | Fórmula                                                                            |
|---------------------------------------------------------------------------------------|------------------------------------------------------------------------------------|
| Total cartera, Vencido 1a30, Vencido 31a60, Vencido 61a90, Vencido +91, Saldo vencer | `=SUBTOTAL(9, rango del cliente)`                                                  |
| %                                                                                     | `=IF($TotalCarteraFila7=0, 0, TotalCartera / $TotalCarteraFila7)` formato `0.00%` |

### En las filas fijas al final (filas 723 de la plantilla)

| Fila  | Columna                        | Fórmula generada                                                              |
|-------|--------------------------------|-------------------------------------------------------------------------------|
| 7     | Total cartera, Vencidos, Saldo | `=SUBTOTAL(9, X2:X{últimaFila})`                                              |
| 7     | %                              | `=SUM(columna%)` de todo el rango                                             |
| 8     | Cliente                        | `Cartera Grupo HL {MesActual}` (mes del sistema al momento de generar)        |
| 8     | Total cartera, Vencidos, Saldo | `=SUM(celdas de clientes con "HL" en el nombre)`                              |
| 9     | Vencido 1 a 30                 | `=SUM(Vencido31, Vencido61, Vencido91)` de la fila 7                         |
| 9     | %                              | `=IF($G$fila9=0, 0, Vencido1a30_fila9 / $G$fila9)`                           |
| 10    | Vencido 1 a 30                 | `=SUM(Vencido31, Vencido61, Vencido91)` de la fila 8                         |
| 10    | %                              | `=IF($G$fila9=0, 0, Vencido1a30_fila10 / $G$fila9)`                          |
| 13    | Vencido 1 a 30                 | `=fila9 - fila10`                                                             |
| 1417 | Vencido 1 a 30                 | `=SUM(Vencido31+Vencido61+Vencido91)` del subtotal del cliente fijo         |
| 1418 | %                              | `=IF($G$fila13=0, 0, G{filaActual} / $G$fila13)` formato `0.00%`             |
| 18    | Vencido 1 a 30                 | `=fila13 - SUM(fila14:fila17)`                                                |
| 21    | Vencido 1 a 30                 | `=fila10 + fila18`                                                            |
| 22    | Vencido 1 a 30                 | `=SUM(fila14:fila17)`                                                         |
| 23    | Vencido 1 a 30                 | `=SUM(fila21:fila22)`                                                         |

 **Clientes fijos para filas 1417:**
- Fila 14: PUNTO MEDICAL DISTRIBUCIONES SAS
- Fila 15: REJIMETAL SAS
- Fila 16: DINTERWEB SAS
- Fila 17: SOLUTIONS & PAYROLL PERU S.A.C

## Agrupación de filas

Las filas de datos de cada cliente tienen `outlineLevel = 1`. Se pueden colapsar con el botón `[-]` en Excel, dejando visibles solo las filas "Total [cliente]". La hoja usa `summaryBelow: true`.

## Notas

- La plantilla debe estar en `public/Plantilla - Cuentas por cobrar detallada.xlsx`.
- Los clientes con "HL" en el nombre se detectan automáticamente (p. ej. "HL INFRAESTRUCTURAS SAS", "CONSORCIO SK-HL", "CONSORCIO HL - A&D").
- Todas las referencias de fila en las fórmulas fijas son **dinámicas**: se recalculan según la cantidad de clientes y registros del SIGO cargado.

## Repositorio

GitHub: [Solutionsandpayroll/Gestion_de_Cobros](https://github.com/Solutionsandpayroll/Gestion_de_Cobros)

## Licencia

 2026 Solutions & Payroll. Uso interno.
