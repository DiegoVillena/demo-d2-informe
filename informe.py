#!/usr/bin/env python
"""Informe de ventas — Ferretería Ortega (demo D2, servicio S2 de Operación Freelancer).

Lee un Excel de ventas, lo limpia (registrando por consola cada corrección y
descarte), calcula las métricas clave y genera un informe PDF profesional (A4).
Si hay un .env con SMTP completo, además envía el PDF por email; por defecto
queda en modo demo: guarda el informe y NO envía nada.

Uso:
    python informe.py --input data/ventas_ejemplo.xlsx --out informe.pdf
"""

import argparse
import os
import smtplib
import sys
import tempfile
import unicodedata
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace
from matplotlib.font_manager import FontProperties, findfont

NEGOCIO = "Ferretería Ortega"
CATEGORIAS = ["Herramientas", "Tornillería", "Pintura", "Plomería", "Electricidad", "Jardín"]
COLUMNAS_ESPERADAS = ["fecha", "producto", "categoría", "cantidad", "importe"]
PERIODO_TXT = ""  # se completa con las fechas reales de los datos

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

# Paleta sobria (azul corporativo + acentos cálidos)
AZUL = "#1F4E79"
AZUL_MEDIO = "#4B86B4"
GUNMETAL = "#5A6472"
PALETA = [AZUL, AZUL_MEDIO, "#9CC3E5", "#E8A33D", "#C0714F", "#6B9080"]


# --------------------------------------------------------------------------
# 1) Limpieza — con registro por consola de cada corrección/descarte
# --------------------------------------------------------------------------
def normalizar_categoria(valor):
    """Devuelve la categoría canónica o None si el valor no se entiende."""
    if not isinstance(valor, str):
        return None
    texto = valor.strip()
    if not texto:
        return None
    clave = (
        unicodedata.normalize("NFKD", texto)
        .encode("ascii", "ignore")
        .decode()
        .strip()
        .casefold()
    )
    canon = {}
    for cat in CATEGORIAS:
        k = cat.casefold()
        canon[k] = cat
        canon[unicodedata.normalize("NFKD", k).encode("ascii", "ignore").decode()] = cat
    canon.update({"herramienta": "Herramientas", "fontaneria": "Plomería",
                  "jardinera": "Jardín", "electrico": "Electricidad"})
    return canon.get(clave)


def limpiar(df: pd.DataFrame, log=print):
    """Aplica la limpieza completa de la ficha §3.2. Devuelve (df_limpio, resumen)."""
    df = df.copy()
    n_inicial = len(df)

    # a) fechas: algunas llegan como texto → convertidas; las imposibles → fuera
    es_texto = df["fecha"].map(lambda v: isinstance(v, str))
    fecha = pd.to_datetime(df["fecha"].where(~es_texto), errors="coerce")  # las ya-fechas
    de_texto = pd.to_datetime(df["fecha"].where(es_texto), dayfirst=True, errors="coerce")
    df["fecha"] = fecha.fillna(de_texto)
    fechas_convertidas = int((es_texto & df["fecha"].notna()).sum())

    # b) filas con datos vacíos o imposibles → descartadas con motivo
    motivos = df["fecha"].isna().map({True: "fecha no válida", False: ""})
    motivos += df["importe"].isna().map({True: "importe nulo", False: ""})
    motivos += df["importe"].le(0).map({True: "importe no positivo", False: ""})
    motivos += df["cantidad"].isna().map({True: "cantidad nula", False: ""})
    motivos += df["cantidad"].le(0).map({True: "cantidad no positiva", False: ""})
    malas = motivos.str.len().astype(bool)
    descartes = dict(motivos[malas].value_counts())

    # c) duplicados exactos (se comprueba sobre filas válidas)
    validas = df[~malas]
    duplicados = int(validas.duplicated().sum())

    df_final = (df[~malas]
                .drop_duplicates()
                .reset_index(drop=True))
    categorias_corregidas = int((~df_final["categoría"].isin(CATEGORIAS)).sum())
    df_final["categoría"] = df_final["categoría"].map(normalizar_categoria)
    df_final = df_final.dropna(subset=["categoría"])
    df_final["cantidad"] = df_final["cantidad"].astype(int)

    # Registro por consola — el detalle que vende el servicio S2
    log("→ Limpieza aplicada:")
    log(f"   · Filas leídas: {n_inicial}")
    if fechas_convertidas:
        log(f"   · Fechas en texto convertidas a fecha: {fechas_convertidas}")
    if duplicados:
        log(f"   · Duplicados exactos eliminados: {duplicados}")
    if descartes:
        detalle = ", ".join(f"{c} {v}" for v, c in sorted(descartes.items()))
        log(f"   · Filas descartadas: {sum(descartes.values())} ({detalle})")
    if categorias_corregidas:
        log(f"   · Categorías normalizadas: {categorias_corregidas}"
            f" (p. ej. \"herramienta\" → Herramientas)")
    log(f"   · Filas válidas finales: {len(df_final)}")
    log("")
    resumen = {
        "leidas": n_inicial,
        "fechas_convertidas": fechas_convertidas,
        "duplicados": duplicados,
        "descartes": descartes,
        "categorias_corregidas": categorias_corregidas,
        "validas": len(df_final),
    }
    return df_final, resumen


# --------------------------------------------------------------------------
# 2) Métricas
# --------------------------------------------------------------------------
def calcular_metricas(df: pd.DataFrame) -> dict:
    """KPIs, top 5 productos, reparto por categoría y serie mensual."""
    top5 = (
        df.groupby("producto", as_index=False)
        .agg(cantidad=("cantidad", "sum"), importe=("importe", "sum"))
        .sort_values("importe", ascending=False)
        .head(5)
        .reset_index(drop=True)
    )
    por_categoria = (
        df.groupby("categoría", as_index=False)
        .agg(ventas=("importe", "size"), importe=("importe", "sum"), unidades=("cantidad", "sum"))
        .sort_values("importe", ascending=False)
        .reset_index(drop=True)
    )
    total = por_categoria["importe"].sum()
    por_categoria["pct"] = por_categoria["importe"] / total
    mensual = df.groupby(pd.Grouper(key="fecha", freq="MS"))["importe"].sum()

    return {
        "total": total,
        "tickets": len(df),
        "ticket_medio": total / len(df),
        "unidades": int(df["cantidad"].sum()),
        "top5": top5,
        "por_categoria": por_categoria,
        "mensual": mensual,
    }


# --------------------------------------------------------------------------
# 3) Gráficos (matplotlib → PNG temporales)
# --------------------------------------------------------------------------
def _fmt_pct(x) -> str:
    return f"{x * 100:.1f}".replace(".", ",") + " %"


def _estilo(ax):
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.25, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=8)


def _etiqueta_mes(ts: pd.Timestamp) -> str:
    return f"{MESES[ts.month - 1]} {str(ts.year)[2:]}"


def _fmt_miles(n) -> str:
    return f"{round(n):,}".replace(",", ".")


def grafico_mensual(mensual: pd.Series, ruta: Path) -> None:
    valores = mensual.values
    fig, ax = plt.subplots(figsize=(9.5, 4.0))
    colores = [AZUL if v == valores.max() else AZUL_MEDIO for v in valores]
    ax.bar(range(len(valores)), valores, color=colores, width=0.72)
    for i, v in enumerate(valores):
        ax.annotate(_fmt_miles(v), (i, v), ha="center", va="bottom", fontsize=7, color=GUNMETAL)
    ax.set_ylim(0, max(valores) * 1.1)
    ax.set_xticks(range(len(mensual)), [_etiqueta_mes(ts) for ts in mensual.index])
    ax.set_title("Ventas por mes (€)", fontsize=11, color=AZUL, fontweight="bold", loc="left")
    _estilo(ax)
    fig.tight_layout()
    fig.savefig(ruta, dpi=150)
    plt.close(fig)


def grafico_top5(top5: pd.DataFrame, ruta: Path) -> None:
    datos = top5.iloc[::-1]  # el nº 1 arriba
    nombres = [p if len(p) <= 24 else p[:23] + "…" for p in datos["producto"]]
    fig, ax = plt.subplots(figsize=(5.0, 3.6))
    ax.barh(range(len(datos)), datos["importe"], color=AZUL_MEDIO, height=0.62)
    ymax = datos["importe"].max() * 1.25
    ax.set_xlim(0, ymax)
    for i, v in enumerate(datos["importe"]):
        ax.annotate(f"{_fmt_miles(v)} €", (v, i), va="center", fontsize=7, color=GUNMETAL)
    ax.set_yticks(range(len(datos)), nombres, fontsize=8)
    ax.set_title("Top 5 productos (importe)", fontsize=11, color=AZUL, fontweight="bold", loc="left")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="x", alpha=0.25, linewidth=0.8)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(ruta, dpi=150)
    plt.close(fig)


def grafico_categorias(por_categoria: pd.DataFrame, ruta: Path) -> None:
    fig, ax = plt.subplots(figsize=(5.0, 3.6))
    wedges, textos, porcentajes = ax.pie(
        por_categoria["importe"],
        labels=por_categoria["categoría"],
        autopct=lambda p: f"{p:.1f}".replace(".", ",") + " %",
        colors=PALETA[:len(por_categoria)],
        startangle=90,
        counterclock=False,
        wedgeprops={"linewidth": 1, "edgecolor": "white"},
        textprops={"fontsize": 8, "color": GUNMETAL},
    )
    plt.setp(porcentajes, size=7.5, color="white", fontweight="bold")
    ax.set_title("Reparto por categorías (importe)", fontsize=11, color=AZUL, fontweight="bold")
    fig.tight_layout()
    fig.savefig(ruta, dpi=150)
    plt.close(fig)


# --------------------------------------------------------------------------
# 4) PDF A4
# --------------------------------------------------------------------------
def _fuente_dejavu(bold: bool = False) -> str:
    prop = FontProperties(family="DejaVu Sans", weight="bold" if bold else "normal")
    return findfont(prop)


def _fmt_eur(n) -> str:
    texto = f"{n:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"{texto} €"


class InformePDF(FPDF):
    def header(self):
        self.set_fill_color(31, 78, 121)
        self.rect(0, 0, 210, 24, style="F")
        self.set_text_color(255, 255, 255)
        self.set_xy(14, 5)
        self.set_font("DejaVu", "B", 16)
        self.cell(100, 8, "Informe de Ventas")
        self.set_xy(14, 13)
        self.set_font("DejaVu", "", 9.5)
        self.cell(140, 5, f"{NEGOCIO} · Periodo: {PERIODO_TXT}")
        self.set_y(30)

    def footer(self):
        self.set_y(-14)
        self.set_text_color(130, 138, 150)
        self.set_font("DejaVu", "", 7.5)
        ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.cell(110, 5, f"Generado automáticamente el {ahora} · datos sintéticos de demostración")
        self.cell(0, 5, f"Página {self.page_no()}/{{nb}}", align="RIGHT")


def caja_kpi(pdf: FPDF, x: float, titulo: str, valor: str) -> None:
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(214, 224, 234)
    pdf.set_text_color(90, 100, 114)
    pdf.set_font("DejaVu", "", 7)
    pdf.rect(x, 32, 42, 21, style="DF", round_corners=True, corner_radius=2.5)
    pdf.set_xy(x + 3, 36.5)
    pdf.cell(36, 4.5, titulo.upper())
    pdf.set_text_color(31, 78, 121)
    pdf.set_font("DejaVu", "B", 13)
    pdf.set_xy(x + 3, 42)
    pdf.cell(36, 7, valor)
    pdf.set_text_color(0, 0, 0)


def _bullets_resumen(m: dict) -> list[str]:
    """Frases clave calculadas de las métricas (resumen ejecutivo del informe)."""
    mejor = m["mensual"].idxmax()
    cat = m["por_categoria"].iloc[0]
    top = m["top5"].iloc[0]
    return [
        f"Mejor mes: {MESES[mejor.month - 1]} {mejor.year} — {_fmt_eur(m['mensual'].max())} ({_fmt_pct(m['mensual'].max() / m['total'])} del total).",
        f"Categoría líder: {cat['categoría']} — {_fmt_eur(cat['importe'])} ({_fmt_pct(cat['pct'])} del total).",
        f"Producto estrella: {top['producto']} — {_fmt_eur(top['importe'])}.",
        f"Ticket medio: {_fmt_eur(m['ticket_medio'])} · {_fmt_miles(m['unidades'])} unidades vendidas.",
    ]


def _caja_bullets(pdf: FPDF, y: float, titulo: str, lineas: list[str]) -> None:
    """Título azul + caja redondeada con viñetas (bloques del pie de la página 2)."""
    pdf.set_y(y)
    pdf.set_font("DejaVu", "B", 10.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 6, titulo)
    pdf.ln(9)
    y_caja = pdf.get_y()
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(214, 224, 234)
    pdf.rect(14, y_caja, 182, 42, style="DF", round_corners=True, corner_radius=2.5)
    pdf.set_xy(20, y_caja + 5)
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(60, 70, 84)
    pdf.multi_cell(170, 8, "\n".join(f"·  {linea}" for linea in lineas))


def generar_pdf(m: dict, out: Path) -> None:
    pdf = InformePDF(orientation="P", unit="mm", format="A4")
    pdf.add_font("DejaVu", "", _fuente_dejavu())
    pdf.add_font("DejaVu", "B", _fuente_dejavu(bold=True))
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.alias_nb_pages()
    pdf.set_margins(14, 30, 14)
    pdf.add_page()

    # KPIs
    cajas = [
        ("Ventas totales", _fmt_eur(m["total"])),
        ("Tickets", _fmt_miles(m["tickets"])),
        ("Ticket medio", _fmt_eur(m["ticket_medio"])),
        ("Unidades vendidas", _fmt_miles(m["unidades"])),
    ]
    for i, (titulo, valor) in enumerate(cajas):
        caja_kpi(pdf, 14 + i * 46, titulo, valor)

    # Gráfico 1: ventas por mes (ancho completo) → 182 × 76,6 mm
    pdf.image(str(m["rutas_img"]["mensual"]), x=14, y=58, w=182)

    # Tabla resumen por categoría
    pdf.set_y(142)
    pdf.set_font("DejaVu", "B", 10.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 6, "Resumen por categoría")
    pdf.ln(8)
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(0, 0, 0)
    cabecera = FontFace(family="DejaVu", emphasis="bold", color=255, fill_color=(31, 78, 121))
    with pdf.table(col_widths=(64, 28, 44, 44), headings_style=cabecera,
                   text_align=("LEFT", "CENTER", "RIGHT", "RIGHT"), line_height=6.5,
                   padding=1.5) as tabla:
        fila_cab = tabla.row()
        for celda in ("Categoría", "Ventas", "Importe", "% del total"):
            fila_cab.cell(celda)
        for _, r in m["por_categoria"].iterrows():
            fila = tabla.row()
            fila.cell(r["categoría"])
            fila.cell(_fmt_miles(int(r["ventas"])))
            fila.cell(_fmt_eur(r["importe"]))
            fila.cell(_fmt_pct(r["pct"]))
        fila_total = tabla.row()
        for celda in ("TOTAL", _fmt_miles(m["tickets"]), _fmt_eur(m["total"]), "100 %"):
            fila_total.cell(celda, style=FontFace(emphasis="bold"))

    # Página 2: top-5 + categorías lado a lado (cada imagen 88 × 63,4 mm)
    MITAD = 88
    pdf.add_page()
    pdf.image(str(m["rutas_img"]["top5"]), x=14, y=31, w=MITAD)
    pdf.image(str(m["rutas_img"]["categorias"]), x=108, y=31, w=MITAD)

    # Resumen automático + calidad de los datos (el log de limpieza, dentro del PDF)
    _caja_bullets(pdf, 105, "Resumen automático", _bullets_resumen(m))
    limpieza = m["limpieza"]
    descartes = ", ".join(f"{c} {v}" for v, c in sorted(limpieza["descartes"].items())) or "ninguno"
    correcciones = []
    if limpieza["fechas_convertidas"]:
        correcciones.append(f"{limpieza['fechas_convertidas']} fechas en texto convertidas a fecha")
    if limpieza["categorias_corregidas"]:
        correcciones.append(f"{limpieza['categorias_corregidas']} categorías normalizadas")
    _caja_bullets(pdf, 168, "Calidad de los datos", [
        f"Filas procesadas: {limpieza['leidas']} → {limpieza['validas']} válidas.",
        f"Duplicados exactos eliminados: {limpieza['duplicados']}.",
        "Correcciones: " + (("; ".join(correcciones) or "ninguna") + "."),
        f"Descartes: {sum(limpieza['descartes'].values())} ({descartes}).",
    ])

    pdf.set_y(-38)
    pdf.set_font("DejaVu", "", 7.5)
    pdf.set_text_color(130, 138, 150)
    pdf.multi_cell(0, 4.5, "Informe generado a partir del Excel de ventas del negocio (sintético). "
                           "Limpieza, métricas y gráficos producidos automáticamente por informe.py — "
                           "demo del servicio S2 de Operación Freelancer.")
    pdf.output(str(out))


# --------------------------------------------------------------------------
# 5) Email opcional (solo con .env completo)
# --------------------------------------------------------------------------
def cargar_env(ruta_env: Path) -> dict:
    """Mini-parser stdlib para .env (KEY=VALUE). Los SMTP_* de entorno real se respetan."""
    env = {k: v for k, v in os.environ.items() if k.startswith("SMTP_")}
    if ruta_env.is_file():
        for linea in ruta_env.read_text(encoding="utf-8").splitlines():
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            clave, _, valor = linea.partition("=")
            env[clave.strip()] = valor.strip().strip('"').strip("'")
    return env


def enviar_email(pdf_path: Path, env: dict, periodo_txt: str, log=print) -> bool:
    """Envía el PDF adjunto. False = no corresponde / no disponible (se queda en demo)."""
    necesarios = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS", "SMTP_DEST"]
    faltan = [k for k in necesarios if not env.get(k)]
    if faltan:
        log(f"   Modo demo: informe guardado; sin email (.env incompleto: faltan {', '.join(faltan)})")
        return False
    msg = EmailMessage()
    msg["From"] = env["SMTP_USER"]
    msg["To"] = env["SMTP_DEST"]
    msg["Subject"] = f"Informe de ventas {periodo_txt} — {NEGOCIO}"
    msg.set_content(
        "Adjunto el informe de ventas generado automáticamente.\n\n"
        "(Demo D2 — datos 100% sintéticos.)"
    )
    msg.add_attachment(pdf_path.read_bytes(), maintype="application", subtype="pdf",
                       filename=pdf_path.name)
    with smtplib.SMTP(env["SMTP_HOST"], int(env["SMTP_PORT"])) as smtp:
        smtp.starttls()
        smtp.login(env["SMTP_USER"], env["SMTP_PASS"])
        smtp.send_message(msg)
    log(f"   Email enviado a {env['SMTP_DEST']} (asunto adjunto: {pdf_path.name})")
    return True


# --------------------------------------------------------------------------
# Pipeline completo
# --------------------------------------------------------------------------
def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=f"Informe de ventas PDF — {NEGOCIO} (demo D2)")
    parser.add_argument("--input", default="data/ventas_ejemplo.xlsx", help="Excel de entrada")
    parser.add_argument("--out", default="informe.pdf", help="PDF de salida")
    parser.add_argument("--env", default=".env", help="ruta del .env para el email opcional")
    args = parser.parse_args(argv)

    print(f"Informe de ventas — {NEGOCIO} (demo D2 · servicio S2)")
    print("=" * 60)

    print(f"1) Leyendo {args.input} (pestaña 'Ventas')…")
    df = pd.read_excel(args.input, sheet_name="Ventas")
    faltan = [c for c in COLUMNAS_ESPERADAS if c not in df.columns]
    if faltan:
        print(f"   [ERROR] Faltan columnas en el Excel: {', '.join(faltan)}")
        return 1
    print(f"   {len(df)} filas · {len(df.columns)} columnas")

    print("2) Limpieza del archivo:")
    df, resumen_limpieza = limpiar(df)

    print("3) Calculando métricas…")
    m = calcular_metricas(df)
    m["limpieza"] = resumen_limpieza
    global PERIODO_TXT
    i0, i1 = m["mensual"].index[0], m["mensual"].index[-1]
    periodo_txt = f"{MESES[i0.month - 1]} {i0.year} – {MESES[i1.month - 1]} {i1.year}"
    PERIODO_TXT = periodo_txt
    print(f"   Ventas totales: {_fmt_eur(m['total'])} · {m['tickets']} tickets · ticket medio {_fmt_eur(m['ticket_medio'])}")
    print(f"   Top producto: {m['top5'].iloc[0]['producto']} ({_fmt_eur(m['top5'].iloc[0]['importe'])})")

    print("4) Generando gráficos…")
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        m["rutas_img"] = {
            "mensual": tmp_dir / "g1_mensual.png",
            "top5": tmp_dir / "g2_top5.png",
            "categorias": tmp_dir / "g3_categorias.png",
        }
        grafico_mensual(m["mensual"], m["rutas_img"]["mensual"])
        grafico_top5(m["top5"], m["rutas_img"]["top5"])
        grafico_categorias(m["por_categoria"], m["rutas_img"]["categorias"])

        print(f"5) Generando PDF → {args.out}")
        generar_pdf(m, Path(args.out))

    print("6) Email:")
    if not enviar_email(Path(args.out), cargar_env(Path(args.env)), periodo_txt):
        pass  # el motivo ya se imprimió dentro
    print("Hecho.")
    return 0


if __name__ == "__main__":
    sys.exit(main())