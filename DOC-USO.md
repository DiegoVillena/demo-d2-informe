# DOC-USO — Informe automático de ventas (demo D2)

Guía de uso del script `informe.py` — Ferretería **Ferretería Ortega** (negocio 100% ficticio, datos sintéticos).

---

## 1. Correr por primera vez

```bash
python -m venv venv
venv/Scripts/pip install -r requirements.txt        # Windows · en Linux/macOS: venv/bin/pip
venv/Scripts/python informe.py --input data/ventas_ejemplo.xlsx --out informe.pdf
```

El script imprime por consola el detalle de la limpieza (corregido/descartado), las métricas y dónde quedó el PDF. Por defecto queda en **modo demo**: guarda `informe.pdf` y no envía email.

## 2. Abrir con tu propio Excel

```bash
python informe.py --input mi_ventas.xlsx --out informe_octubre.pdf
```

Requisitos del Excel:

- **Pestaña "Ventas"** con exactamente estas columnas (los nombres, con acentos): `fecha · producto · categoría · cantidad · importe`.
- Una fila por línea de venta. El `importe` debe ser positivo; la `cantidad`, entera y positiva.
- Lo que no se entienda se **registra por consola y se descarta con motivo**: fechas imposibles, importes nulos o negativos, cantidades nulas, duplicados exactos. Las fechas escritas como texto (`31/12/2025`) se convierten; las categorías con erratas ("herramienta", "Tornilleria ", "FONTANERIA") se normalizan a las 6 categorías válidas.

## 3. Activar el envío por email (opcional)

```bash
cp .env.example .env   # Windows: copy .env.example .env
```

Rellena en `.env` las 5 variables:

| Variable | Qué es |
| --- | --- |
| `SMTP_HOST` | Servidor de salida (p. ej. `smtp.gmail.com`) |
| `SMTP_PORT` | Normalmente `587` (STARTTLS) |
| `SMTP_USER` | Tu cuenta de correo |
| `SMTP_PASS` | Contraseña / **app-password** (Gmail: cuenta → Seguridad → 2FA → contraseñas de aplicaciones) |
| `SMTP_DEST` | Destinatario del informe |

Con `.env` completo, cada corrida envía el PDF recién generado como adjunto. Sin `.env` (o incompleto): modo demo, nada se envía. El `.env` está en `.gitignore` — **nunca lo subas al repo**.

## 4. Regenerar los datos de ejemplo

```bash
python generar_datos.py --seed 42         # mismo resultado siempre
python generar_datos.py --seed 7          # otra variante de ferretería
```

## 5. Tests

```bash
venv/Scripts/python -m pytest
```

Cubre la limpieza: normalización de categorías/fechas, eliminación de duplicados y descartes con motivo.