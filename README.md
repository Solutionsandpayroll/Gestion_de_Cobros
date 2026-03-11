# Web Cuentas por Cobrar — Tesorería

Aplicación web interna de **Solutions & Payroll** para gestionar cobros por cliente a partir de la plantilla SIGO de cuentas por cobrar.

## ¿Qué hace esta aplicación?

1. El usuario sube el archivo Excel exportado desde SIGO (plantilla de cuentas por cobrar).
2. La app lee automáticamente la columna **Cliente** de la fila 7 y lista todos los clientes únicos encontrados.
3. El usuario puede seleccionar o deseleccionar los clientes que desea incluir en el reporte (todos seleccionados por defecto).
4. Al presionar **Generar Formato**, se produce un archivo Excel basado en la plantilla interna, con los datos de los clientes seleccionados y conservando los estilos, colores, fuentes y anchos de columna originales de la plantilla.

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

## Notas sobre la plantilla

- Las filas **7 a 23** de la plantilla contienen información fija que se conserva siempre. Al generar el reporte, esa información se coloca **al final** del archivo, debajo de los datos traídos del SIGO.
- La plantilla debe estar en `public/Plantilla - Cuentas por cobrar detallada.xlsx`.

- `.drop-zone` - Zona drag & drop
- `.modal-overlay` - Overlay de modal
- `.help-section` - Sección colapsable

## 💡 Tips

1. **Mantén limpio el App.jsx** - Crea componentes separados si crece mucho
2. **Usa las variables CSS** - No modifiques los colores directamente
3. **Los SVG están inline** - Puedes cambiarlos fácilmente o usar íconos de librerías
4. **Las animaciones ya están configuradas** - Se activarán automáticamente

## 📚 Recursos

- [Documentación React](https://react.dev/)
- [Documentación Vite](https://vitejs.dev/)
- [Iconos SVG](https://feathericons.com/)
- [Colores](https://tailwindcss.com/docs/customizing-colors)

## 🔒 No Subir a Git

Si inicias Git en tu nuevo proyecto, asegúrate de tener `.gitignore`:
```
node_modules
dist
.env
```

## 📄 Licencia

© 2026 Solutions & Payroll. Template de uso interno.

---

**¡Listo para crear tu próximo proyecto!** 🚀
