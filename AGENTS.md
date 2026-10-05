# AGENTS.md — demo-d2-informe (Demo D2)

## Qué es este proyecto

Script CLI en Python que convierte un Excel de ventas en un **informe PDF profesional**
(KPIs + gráficos + tabla). Es la **demo D2** de *Operación Freelancer* — demo del **servicio S2**:
*"elimino tu trabajo manual de pasar de Excel a informe"*.

- Negocio ficticio: **Ferretería Ortega** (100% sintético, misma familia de demos que la D1).
- **Ficha de requisitos congelada** (única fuente de verdad de alcance):
  `C:\Users\Diego\Desktop\OperacionFreelancer\plantillas\ficha-d2.md`. Si algo no está ahí = ficha aparte.
- La **ejecución programada** (cron / Task Scheduler / GitHub Actions) está FUERA de alcance — es la demo D3.

## Stack

- Python 3.11+ con dependencias fijadas **por versión exacta** en `requirements.txt`:
  pandas, openpyxl, matplotlib, fpdf2, pytest.
- Todo se ejecuta contra el **venv local** (`venv/`, ignorado por git) — nunca contra el Python global.
- Los gráficos del PDF se renderizan a PNG con matplotlib y se incrustan con fpdf2.
- El `.env` se parsea con un mini-parser stdlib (sin python-dotenv, así `requirements.txt` queda en los 5 paquetes de la ficha).

## Estructura

```
informe.py            # Pipeline completo: lee Excel → limpieza → métricas → PDF → email opcional
generar_datos.py      # Genera el Excel sintético reproducible (seed=42)
data/ventas_ejemplo.xlsx  # Datos de ejemplo (pestaña "Ventas")
informe.pdf           # Informe de ejemplo commiteado (ver el resultado sin ejecutar nada)
tests/test_limpieza.py    # Test pytest de la limpieza
data/, docs/capturas/  # datos + capturas del PDF para el README
```

## Cómo ejecutar

```bash
python -m venv venv
venv\Scripts\pip install -r requirements.txt      # Git Bash: venv/Scripts/pip install -r requirements.txt
venv\Scripts\python informe.py --input data/ventas_ejemplo.xlsx --out informe.pdf
venv\Scripts\python -m pytest
```

- El script imprime por consola **qué corrigió/descartó durante la limpieza** — ese detalle es clave (vende el S2).
- **Modo demo por defecto**: guarda el PDF y NO envía email. El envío solo ocurre con `.env` completo
  (ver `.env.example`). `.env` está en `.gitignore` — **jamás commitear secretos**.

## Reglas de este repo

- Rama por tarea, **commits/push solo con confirmación de Diego** (él testea antes de que nada entre al historial).
- Datos 100% sintéticos con seed fija — reproducibles con `generar_datos.py`.
- Alcance congelado según ficha; máx 4 h/día.

## Estado actual

- **Demo D2 terminada**: primer commit en `main` (via `feat/informe-pdf`), repo público `DiegoVillena/demo-d2-informe`.
- Pipeline verificado de punta a punta por Diego (script + PDF + tests 4 en verde + reproducibilidad).
- GIF "script corriendo" hecho (`docs/demo-corriendo.gif`, salida real animada con Pillow; el generador `.zcode/tmp/build_gif.py` no se commitea).
- Documentación completa: README (patrón portfolio D1/Oráculo), DOC-USO.md, `.env.example`, este AGENTS.md, commonerrors.md.
- El "Excel de entrada" es la única pieza que cambia entre clientes (ver DOC-USO.md §2).
- Fuera de alcance congelado: ejecución programada (cron / Task Scheduler / Actions) = demo D3; dashboards, BD, GUI y multiinforme también fuera.