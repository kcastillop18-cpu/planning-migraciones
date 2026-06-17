# Planning del Supervisor · Migraciones

Aplicación web (corre 100% en el navegador, sin servidor) para que cada supervisor
de un call center de migraciones haga seguimiento **hora a hora** de su equipo, en
lugar de esperar la reportería del día anterior (D‑1).

Se carga el volcado del CRM (Excel/CSV) y la web calcula todo localmente — **ningún
dato sale del equipo**.

## Funcionalidades
- **Perfiles**: Supervisor (ve solo a su equipo) y Jefatura (ve todos, con PIN).
- **Ventas hora a hora** + acumulado; **horas extra** (después de 18:00) resaltadas.
- **Jornada base 09–18 (8 h efectivas)** vs horas extra, con **SPH** y **proyección**.
- **SPD** del día y acumulado.
- **Cuartilización** del equipo (Q1–Q4) con nombres.
- **Objetivos del día por agente** con bloqueo (solo Jefatura edita, con PIN).
- **Migración vs Cross‑selling** (por `RV_Tipo_Ofrecimiento`); cross con desglose.
- **Fechas multi‑selección** + **cuadro dinámico Supervisor × Hora** ordenable.
- **Monitoreo de calidad** (DNI asesor, fecha, N° llamada, DNI cliente, Migra, Ofrec. cross), registros fijos con hora.
- **Tareas del supervisor** con hora de cumplimiento y bloqueo.

## Cómo usar
1. Abre `index.html` con doble clic (no requiere internet ni instalación).
2. Carga el volcado del CRM, o pulsa **"Ver demo con datos de ejemplo"** (datos
   sintéticos incluidos en el repo).
3. Filtra solo la campaña `MIGRACION REGULAR` (automático), tu fecha, tu equipo, etc.

Más detalle de uso en **`LEEME.txt`**.

## Archivos
- `index.html` — interfaz
- `app.js` — toda la lógica de cálculo
- `xlsx.full.min.js` — librería SheetJS (lee Excel en el navegador, offline)
- `sample-data.js` — **datos sintéticos** para el botón de demo
- `extract-fast.ps1`, `make-sample.ps1`, `profile-cols.ps1` — utilitarios (procesar el volcado real)

> ⚠️ Los datos reales (volcado del CRM, demos con información real) **no** se incluyen
> en el repositorio (ver `.gitignore`).

## Nota
Es un **prototipo**: perfiles, PIN y bloqueos viven en el navegador (localStorage).
Para control inviolable (usuarios/contraseñas en servidor, bitácora) haría falta una
versión backend.
