"""Tests del paso de limpieza de informe.py (ficha D2 §4: test mínimo)."""

import pandas as pd

import informe


def df_sucio() -> pd.DataFrame:
    """Mini-Excel con los mismos fallos que el archivo real de la demo."""
    return pd.DataFrame(
        {
            "fecha": ["01/03/2026", "05/03/2026", "05/03/2026", "2026-13-40", "20/03/2026"],
            "producto": ["Martillo de uñas 16 oz", "Martillo de uñas 16 oz",
                         "Martillo de uñas 16 oz", "Grifo de lavabo monomando", "Pintura plástica blanca 4 L"],
            "categoría": ["herramienta", "Plomería", "Plomería", "Herramientas", "pinturas "],
            "cantidad": [1.0, 2.0, 2.0, 1.0, 1.0],
            "importe": [11.90, 23.80, 23.80, 57.00, pd.NA],
        },
        columns=["fecha", "producto", "categoría", "cantidad", "importe"],
    )


def test_normaliza_categorias_y_fechas_en_texto():
    df, resumen = informe.limpiar(df_sucio(), log=lambda *_: None)
    assert set(df["categoría"]) == {"Herramientas", "Plomería"}
    assert list(df["fecha"].dt.strftime("%Y-%m-%d")) == ["2026-03-01", "2026-03-05"]
    assert "herramienta" not in set(df["categoría"])
    assert resumen["fechas_convertidas"] == 4  # las 5 fechas son texto; 4 se rescatan, 1 es imposible
    assert resumen["categorias_corregidas"] == 1


def test_elimina_duplicados_exactos():
    df, resumen = informe.limpiar(df_sucio(), log=lambda *_: None)
    assert len(df) == 2  # 5 leídas − 1 duplicado − 2 descartes
    assert resumen["duplicados"] == 1
    assert resumen["validas"] == 2


def test_descarta_nulos_y_fechas_imposibles_con_motivo():
    _, resumen = informe.limpiar(df_sucio(), log=lambda *_: None)
    assert resumen["descartes"].get("importe nulo") == 1
    assert resumen["descartes"].get("fecha no válida") == 1
    assert sum(resumen["descartes"].values()) == 2


def test_normalizar_categoria_unit():
    assert informe.normalizar_categoria("herramienta ") == "Herramientas"
    assert informe.normalizar_categoria("TORnilleria") == "Tornillería"
    assert informe.normalizar_categoria("Plomería") == "Plomería"
    assert informe.normalizar_categoria("no-existe") is None
    assert informe.normalizar_categoria(pd.NA) is None