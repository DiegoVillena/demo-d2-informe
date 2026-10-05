#!/usr/bin/env python
"""Genera data/ventas_ejemplo.xlsx: ventas sintéticas de la Ferretería Ortega.

Los datos son 100% FICTICIOS y reproducibles (seed fija). Incluye a propósito
unas filas "sucias" (duplicados, nulos, categorías mal escritas, fechas en
texto) para que la limpieza de informe.py tenga material real que registrar.

Uso: venv/Scripts/python generar_datos.py [--seed 42] [--out data/ventas_ejemplo.xlsx]
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

# --- Negocio ficticio -------------------------------------------------------
PERIODO = (pd.Timestamp("2025-10-01"), pd.Timestamp("2026-09-30"))  # 12 meses

# producto → (categoría, precio unitario en €, cantidad típica por venta)
CATALOGO = {
    "Martillo de uñas 16 oz":      ("Herramientas", 11.90, 1),
    "Destornillador Phillips nº2": ("Herramientas", 4.50, 1),
    "Sierra de mano 550 mm":       ("Herramientas", 14.20, 1),
    "Llave inglesa 10 pulgadas":   ("Herramientas", 9.80, 1),
    "Taladro percutor 600 W":      ("Herramientas", 62.90, 1),
    "Juego de llaves Allen (9 pzs)": ("Herramientas", 12.40, 1),
    "Nivel de burbuja 60 cm":      ("Herramientas", 13.50, 1),
    "Tornillo DIN 57 6x40 (caja 200)": ("Tornillería", 8.60, 2),
    "Taco plástico 6 mm (caja 100)":   ("Tornillería", 4.30, 2),
    "Clavo de 40 mm (kg)":         ("Tornillería", 3.20, 4),
    "Tuerca M6 (caja 100)":        ("Tornillería", 3.80, 2),
    "Pintura plástica blanca 4 L": ("Pintura", 24.90, 2),
    "Esmalte sintético blanco 750 ml": ("Pintura", 9.60, 2),
    "Rodillo de lana 22 cm":       ("Pintura", 5.40, 2),
    "Cinta de carrocero 38 mm":    ("Pintura", 2.60, 3),
    "Tubo de cobre 15 mm (1 m)":   ("Plomería", 9.90, 2),
    "Codo de cobre 15 mm":         ("Plomería", 4.70, 3),
    "Silicona sanitaria transparente": ("Plomería", 6.90, 2),
    "Grifo de lavabo monomando":   ("Plomería", 57.00, 1),
    "Sifón de lavabo":             ("Plomería", 14.60, 1),
    "Bombilla LED E27 10 W":       ("Electricidad", 3.90, 3),
    "Aplique de exterior":         ("Electricidad", 21.80, 1),
    "Cable unipolar 2,5 mm (rollo 100 m)": ("Electricidad", 46.50, 1),
    "Enchufe schuko doble":        ("Electricidad", 7.80, 2),
    "Manguera de riego 20 m":      ("Jardín", 16.90, 1),
    "Tijeras de podar":            ("Jardín", 12.90, 1),
    "Fertilizante universal 5 kg": ("Jardín", 13.70, 2),
    "Tierra vegetal 20 L":         ("Jardín", 5.90, 3),
}

# Factor de estacionalidad por categoría, mes a mes (oct, nov, … sep)
ESTACIONALIDAD = {
    "Jardín":      (0.6, 0.7, 0.9, 1.2, 1.5, 1.5, 1.4, 1.2, 1.0, 0.8, 0.6, 0.5),
    "Pintura":     (0.7, 0.7, 0.9, 1.3, 1.4, 1.2, 1.0, 0.9, 0.9, 0.8, 0.7, 0.7),
    "Tornillería": (0.9, 0.9, 1.1, 1.2, 1.1, 1.0, 0.9, 0.9, 1.0, 1.0, 1.0, 1.0),
}


def generar_ventas(seed: int) -> pd.DataFrame:
    """~290 filas de ventas válidas entre PERIODO[0] y PERIODO[1]."""
    rng = np.random.default_rng(seed)
    productos = list(CATALOGO)
    filas = []
    for dia in pd.date_range(*PERIODO, freq="D"):
        n_hoy = rng.choice([0, 1, 2], p=[0.35, 0.50, 0.15])
        for _ in range(n_hoy):
            # peso de cada producto: estacionalidad de su categoría ese mes
            pesos = np.array(
                [ESTACIONALIDAD.get(CATALOGO[p][0], (1.0,) * 12)[dia.month - 1] for p in productos],
                dtype=float,
            )
            pesos = pesos / pesos.sum()
            producto = rng.choice(productos, p=pesos)
            categoria, precio, tipico = CATALOGO[producto]
            cantidad = int(rng.integers(1, tipico + 2))
            filas.append(
                {
                    "fecha": dia,
                    "producto": producto,
                    "categoría": categoria,
                    "cantidad": cantidad,
                    "importe": round(cantidad * precio, 2),
                }
            )
    df = pd.DataFrame(filas, columns=["fecha", "producto", "categoría", "cantidad", "importe"])
    return df, rng


def añadir_filas_sucias(df: pd.DataFrame, rng) -> pd.DataFrame:
    """Simula los fallos típicos de un Excel de cliente: duplicados, nulos,
    categorías mal escritas, fechas en texto y valores imposibles."""
    sucias = []

    # 3 duplicados exactos tal cual salieron
    idx_dup = rng.choice(len(df), size=3, replace=False)
    sucias.append(df.iloc[idx_dup].copy())

    # 2 fechas escritas como texto (dd/mm/aaaa) — se rescatan durante la limpieza
    sucias.append(df.iloc[100:102].copy().assign(fecha="03/05/2026"))
    sucias.append(df.iloc[103:104].copy().assign(fecha="12/12/2025"))

    # 1 fecha imposible — se descarta durante la limpieza
    sucias.append(df.iloc[104:105].copy().assign(fecha="2026-13-01"))

    # 3 categorías mal escritas — se normalizan durante la limpieza
    typos = ["herramienta", "Tornilleria ", "ELECTRICIDAD"]
    sucias.append(df.iloc[200:203].reset_index(drop=True).assign(categoría=typos))

    # nulos y un importe negativo — se descartan durante la limpieza
    sucias.append(df.iloc[250:251].copy().assign(importe=np.nan))
    sucias.append(df.iloc[251:252].copy().assign(cantidad=np.nan))
    sucias.append(df.iloc[252:253].copy().assign(importe=-15.90))

    return pd.concat([df] + sucias, ignore_index=True)[
        ["fecha", "producto", "categoría", "cantidad", "importe"]
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera el Excel sintético de la Ferretería Ortega")
    parser.add_argument("--seed", type=int, default=42, help="semilla para reproducibilidad")
    parser.add_argument("--out", default="data/ventas_ejemplo.xlsx")
    args = parser.parse_args()

    df_validas, rng = generar_ventas(args.seed)
    df = añadir_filas_sucias(df_validas, rng)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out, sheet_name="Ventas", index=False)

    print(f"OK → {out}")
    print(f"  Filas totales: {len(df)} (~{len(df) - 12} válidas + 12 sucias a propósito)")
    print(f"  Pestaña: 'Ventas' · columnas: {', '.join(df.columns)}")
    print(f"  Periodo: {PERIODO[0]:%d/%m/%Y} – {PERIODO[1]:%d/%m/%Y} · seed={args.seed}")


if __name__ == "__main__":
    main()