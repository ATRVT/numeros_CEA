# 🏫 ALTUS Contabilidad & Auditoría SAT (iPhone PWA & Mac Desktop)

Sistema contable inteligente para **Centro Educativo ALTUS** (Guatemala), diseñado para operar en **iPhone** y **Mac de escritorio**, conectado a **Google Sheets** y reglado bajo criterios de auditoría de la **SAT**, **MINEDUC** e **IGSS**.

---

## 📦 Archivos Entregados en este Proyecto

1. **`Libro_Contable_ALTUS_26_Auditado.xlsx`**:
   Nuevo libro maestro de Excel/Google Sheets con los **283 asientos históricos de 2026 auditados y corregidos**:
   - **Doble columna de saldo bancario en tiempo real:** `Saldo BI-Monetaria (1101)` (`Q 86,289.53`) y `Saldo BI-Ahorro (1102)` (`Q 93,296.67`).
   - **Eliminación de doble contabilización (`Q 45,816.00`):** Separación contable entre las transferencias mensuales de reserva (`2102 - Provisiones Laborales`) y la liquidación de pagos desde la cuenta de ahorro (`BI-Ahorro` para IGSS `2103` y Bono 14 `5104`).
   - **Reclasificación de ISR Empleados (`2105`):** Los enteros de ISR retenido (`Q 375.85`) se reclasificaron de `5210` a `2105 - ISR Retenciones por Pagar`.
   - **Matriz de Familias y Colegiaturas limpia:** Sin errores `#N/A`, con alias de padres/madres (ej. *Pamela Sosa Cuéllar* $\rightarrow$ *Luís Felipe Guzmán*) y control de solvencia de Enero a Diciembre.
   - **Presupuesto Anual 2027:** Proyección completa de 12 meses para el modelo de 2 grupos (16 alumnos).

2. **`Google_Apps_Script_ALTUS.gs`**:
   Conector directo para pegar en **Extensiones $\rightarrow$ Apps Script** dentro de tu Google Sheets y sincronizar entradas desde el iPhone o Mac sin servidores intermedios.

3. **Aplicación Web Progresiva (`index.html`, `js/`, `manifest.json`, `sw.js`)**:
   Lista para alojarse en **GitHub Pages** o **Vercel** e instalarse en iPhone y Mac.

---

## 🚀 Cómo usarla en tu iPhone y Mac

### En tu Mac (Escritorio):
1. Abre `index.html` o accede a la URL de tu repositorio en GitHub Pages.
2. En Safari en Mac: ve a **Archivo $\rightarrow$ Añadir al Dock** para tenerla como una app de escritorio independiente.

### En tu iPhone:
1. Abre el enlace de tu GitHub Pages en **Safari**.
2. Toca el botón **Compartir** (⬆️) $\rightarrow$ **Añadir a pantalla de inicio**.
3. Se instalará con el ícono de **ALTUS Contable** a pantalla completa.
