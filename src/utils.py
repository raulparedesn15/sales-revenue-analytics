# -*- coding: utf-8 -*-

"""
Entrega Caso Práctico | Especialista de Calidad de Gestión Comercial
===================================================================

Objetivo
-------------------
Tomar el Excel del caso práctico (3 hojas: BD, Ciudad-Region, Presupuesto),
limpiar y cruzar la información, para calcular los KPIs solicitados, agregar además un análisis
libre extra en este caso decidi enfocarlo a cartera, concentración, reactivación de clientes y
exportar un Excel final con varias pestañas listas para usar en tablas dinámicas y un dashboard.

Cómo correrlo (en terminal)
---------------------------
1) Ejecuta (Primero verificando que el xlsx crudo este en la carpeta donde se va a correr):

2) Al final imprime la ruta del archivo generado.

Qué genera (pestañas principales)
-------------------------------
- Base_Cruzada: la base ya limpia y cruzada (lista para pivots en Excel).
- KPI1: ingreso mensual (total / región / ciudad / vendedor).
- KPI2: ingreso trimestral y % crecimiento (total / región / vendedor).
- KPI3: proyección anual simple con lo disponible.
- KPI4: cumplimiento vs presupuesto por vendedor.
- Ejercicio 3: análisis extra de cartera/concentración y clientes a reactivar."""

import argparse
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# ===============================================================
# 1) FUNCIONES PEQUEÑAS (LIMPIEZA DE TEXTO Y CREACIÓN DE LLAVES)
# ===============================================================


def norm_spaces(value: object) -> str:
    """Dejo el texto limpio en cuanto a espacios.
    - quito espacios al inicio/fin
    - si hay muchos espacios seguidos, los reduzco a uno
    Lo uso para que los nombres de vendedor/cliente no fallen por detalles de formato.
    """
    return " ".join(str(value or "").strip().split())


def normalize_vendor_name(raw: object, remove_leading_digits: bool = False) -> str:
    """Creo una llave estándar para el vendedor: `Vendedor_key`.
    Mi intención aquí es que variaciones como:
    - "Irving   Hernandez"
    - "IRVING HERNANDEZ"
    - "  001 Irving Hernandez"  (cuando viene con números)
    terminen siendo exactamente el mismo texto, para poder hacer merges confiables."""
    s = norm_spaces(raw)
    if remove_leading_digits:
        # En Ciudad-Region viene "001 NOMBRE APELLIDO".
        # Aquí le quito esos dígitos para que las tablas sean congruentes.
        s = re.sub(r"^\d+\s*", "", s)
    # Finalmente, estandarizo a mayúsculas.
    return norm_spaces(s).upper()


def presupuesto_to_common_key(raw: object) -> str:
    """Ajusto el nombre del vendedor en Presupuesto para que use el mismo formato.
    En Presupuesto el nombre viene como: "APELLIDO NOMBRE".
    En BD/Ciudad-Region viene como: "NOMBRE APELLIDO".
    Entonces aquí lo que hago es voltear el orden para generar una llave compatible."""
    s = normalize_vendor_name(raw)
    if not s:
        return s
    parts = s.split()
    if len(parts) >= 2:
        # Ej: "HERNANDEZ IRVING" -> "IRVING HERNANDEZ"
        return " ".join(parts[1:] + [parts[0]])
    return s


# ============================================================
# 2) CARGA Y VALIDACIONES
# ============================================================


def require_columns(df: pd.DataFrame, cols: list[str], df_name: str) -> None:
    """Valido que existan columnas necesarias antes de avanzar.
    Esto me evita llegar al final y descubrir que algo faltaba.
    Prefiero fallar rápido y con un mensaje claro."""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas en '{df_name}': {missing}. Columnas disponibles: {list(df.columns)}"
        )


def load_sheets(xlsx_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # Cargo las 3 hojas del Excel del caso: BD, Ciudad-Region y Presupuesto.
    xls = pd.ExcelFile(xlsx_path)
    required = ["BD", "Ciudad-Region", "Presupuesto"]
    missing = [s for s in required if s not in xls.sheet_names]
    if missing:
        raise ValueError(
            f"Faltan hojas en el Excel: {missing}. Hojas disponibles: {xls.sheet_names}"
        )
    bd = pd.read_excel(xlsx_path, sheet_name="BD")
    cdrg = pd.read_excel(xlsx_path, sheet_name="Ciudad-Region")
    pres = pd.read_excel(xlsx_path, sheet_name="Presupuesto")
    return bd, cdrg, pres


def prepare_base(
    bd: pd.DataFrame, cdrg: pd.DataFrame, pres: pd.DataFrame
) -> pd.DataFrame:
    # Limpio, creo llaves, cruzo tablas y devuelvo la base final (Base_Cruzada).
    require_columns(
        bd,
        ["Fecha Operación", "Vendedor", "Ingreso Operación", "No. Cliente", "Guia"],
        "BD",
    )
    require_columns(cdrg, ["NOMBRE", "CIUDAD", "REGION"], "Ciudad-Region")
    require_columns(pres, ["Vendedor", "Presupuesto"], "Presupuesto")

    bd = bd.copy()
    cdrg = cdrg.copy()
    pres = pres.copy()

    # Me aseguro de que la fecha sea fecha real, no texto.
    bd["Fecha Operación"] = pd.to_datetime(bd["Fecha Operación"], errors="coerce")
    if bd["Fecha Operación"].isna().any():
        # Si hay fechas invalidas con esto puedo saberlo para no afectar a los KPIs
        raise ValueError(
            "Hay fechas inválidas en BD (Fecha Operación). Revisa el formato de la columna."
        )

    # Creo columnas de tiempo (mes/trimestre). Uso Period para ordenar bien y luego lo convierto a texto.
    bd["Mes_Period"] = bd["Fecha Operación"].dt.to_period("M")  # type: ignore[union-attr]
    bd["Trimestre_Period"] = bd["Fecha Operación"].dt.to_period("Q")  # type: ignore[union-attr]
    bd["Mes"] = bd["Mes_Period"].astype(str)
    bd["Trimestre"] = bd["Trimestre_Period"].astype(str)

    # Creo mi llave estándar (Vendedor_key) en cada tabla, con esto puedo cruzar sin errores.
    bd["Vendedor_key"] = bd["Vendedor"].apply(normalize_vendor_name)
    cdrg["Vendedor_key"] = cdrg["NOMBRE"].apply(
        lambda x: normalize_vendor_name(x, remove_leading_digits=True)
    )
    pres["Vendedor_key"] = pres["Vendedor"].apply(presupuesto_to_common_key)

    # Me quedo con lo mínimo para el cruce y evito duplicados.
    # Así no se multiplico filas al hacer merge.
    cdrg_min = cdrg[["Vendedor_key", "CIUDAD", "REGION"]].drop_duplicates(
        subset=["Vendedor_key"]
    )
    pres_min = pres[["Vendedor_key", "Presupuesto"]].drop_duplicates(
        subset=["Vendedor_key"]
    )

    # Cruzo BD con Ciudad-Region y luego con Presupuesto. Esto en el dataframe.
    df = bd.merge(cdrg_min, on="Vendedor_key", how="left")
    df = df.merge(pres_min, on="Vendedor_key", how="left")

    # Me aseguro que ingreso y presupuesto sean numéricos.
    # Si hay ingresos inválidos, los convierto a 0 así no rompo las sumas.
    df["Ingreso Operación"] = pd.to_numeric(
        df["Ingreso Operación"], errors="coerce"
    ).fillna(0.0)
    df["Presupuesto"] = pd.to_numeric(df["Presupuesto"], errors="coerce")

    return df


# ============================================================
# 3) KPIs. Aquí se resuelven los 3 KPIs que me pidieron
# ============================================================

# KPI 1: aqui calcule el ingreso por mes pero ademas muestro el total y tambien un desglose
# por ciudad, region y vendedor para facilitar el análisis.


def kpi_ingreso_mensual(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Genere varias tablas porque en Excel es más fácil hacer un dashboard así
    # cuando ya tengo el nivel listo (total / región / ciudad / vendedor).
    # Total Mensual
    overall = (
        df.groupby("Mes_Period", as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Mensual"})  # type: ignore[arg-type]
        .sort_values("Mes_Period")
    )
    overall["Mes"] = overall["Mes_Period"].astype(str)
    overall = overall[["Mes", "Ingreso Mensual"]]
    #  Mensual por Region
    by_region = (
        df.groupby(["Mes_Period", "REGION"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Mensual"})  # type: ignore[arg-type]
        .sort_values(["Mes_Period", "REGION"])
    )
    by_region["Mes"] = by_region["Mes_Period"].astype(str)
    by_region = by_region[["Mes", "REGION", "Ingreso Mensual"]]
    # Mensual por ciudad
    by_city = (
        df.groupby(["Mes_Period", "REGION", "CIUDAD"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Mensual"})  # type: ignore[arg-type]
        .sort_values(["Mes_Period", "REGION", "CIUDAD"])
    )
    by_city["Mes"] = by_city["Mes_Period"].astype(str)
    by_city = by_city[["Mes", "REGION", "CIUDAD", "Ingreso Mensual"]]
    # Mensual por vendedor
    by_seller = (
        df.groupby(["Mes_Period", "Vendedor_key", "REGION", "CIUDAD"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Mensual"})  # type: ignore[arg-type]
        .sort_values(["Mes_Period", "Vendedor_key"])
    )
    by_seller["Mes"] = by_seller["Mes_Period"].astype(str)
    by_seller = by_seller[
        ["Mes", "Vendedor_key", "REGION", "CIUDAD", "Ingreso Mensual"]
    ]

    return {
        "KPI1_Mensual_Total": overall,
        "KPI1_Mensual_Region": by_region,
        "KPI1_Mensual_Ciudad": by_city,
        "KPI1_Mensual_Vendedor": by_seller,
    }


# KPI 2: aquí voy a mostrar el ingreso trimestral + % crecimiento vs trimestre anterior.


def kpi_crecimiento_trimestral(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # La lógica que use fue: sumo por trimestre y luego comparo contra el trimestre anterior.
    # Este al igual que el KPI 1 lo desgloso pero solo por Región y Vendedor. Esto para tener info
    # adicional a la mano para el dashboard
    # Total Trimestral
    overall = (
        df.groupby("Trimestre_Period", as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Trimestral"})  # type: ignore[arg-type]
        .sort_values("Trimestre_Period")
    )
    # Aqui saque el porcentaje y hago lo mismo para region y vendedor con pct_change
    overall["% Crecimiento Trimestral"] = (
        overall["Ingreso Trimestral"].pct_change() * 100
    )
    overall["Trimestre"] = overall["Trimestre_Period"].astype(str)
    overall = overall[["Trimestre", "Ingreso Trimestral", "% Crecimiento Trimestral"]]
    # Trimestral por Region
    by_region = (
        df.groupby(["Trimestre_Period", "REGION"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Trimestral"})  # type: ignore[arg-type]
        .sort_values(["REGION", "Trimestre_Period"])
    )
    by_region["% Crecimiento Trimestral"] = (
        by_region.groupby("REGION")["Ingreso Trimestral"].pct_change() * 100
    )
    by_region["Trimestre"] = by_region["Trimestre_Period"].astype(str)
    by_region = by_region[
        ["Trimestre", "REGION", "Ingreso Trimestral", "% Crecimiento Trimestral"]
    ]
    # Trimestral por Vendedor
    by_seller = (
        df.groupby(
            ["Trimestre_Period", "Vendedor_key", "REGION", "CIUDAD"], as_index=False
        )["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Trimestral"})  # type: ignore[arg-type]
        .sort_values(["Vendedor_key", "Trimestre_Period"])
    )
    by_seller["% Crecimiento Trimestral"] = (
        by_seller.groupby("Vendedor_key")["Ingreso Trimestral"].pct_change() * 100
    )
    by_seller["Trimestre"] = by_seller["Trimestre_Period"].astype(str)
    by_seller = by_seller[
        [
            "Trimestre",
            "Vendedor_key",
            "REGION",
            "CIUDAD",
            "Ingreso Trimestral",
            "% Crecimiento Trimestral",
        ]
    ]

    return {
        "KPI2_Trimestral_Total": overall,
        "KPI2_Trimestral_Region": by_region,
        "KPI2_Trimestral_Vendedor": by_seller,
    }


# KPI 3: proyección anual para el cierre del año (promedio mensual * 12)
# KPI 4: cumplimiento vs presupuesto por vendedor (y agredo tambien el cumplimiento proyectado)


def kpi_proyeccion(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Primero veo cuántos meses reales tengo en el histórico, ya que por lo que vi no estan los 12 meses.
    meses_con_datos = int(df["Mes_Period"].nunique())
    if meses_con_datos <= 0:
        raise ValueError("No hay meses con datos para calcular proyección.")

    # Con eso calculo el promedio mensual y lo llevo a 12 meses, para tener la proyección anual.
    ingreso_acumulado = float(df["Ingreso Operación"].sum())
    ingreso_promedio = ingreso_acumulado / meses_con_datos
    proyeccion_anual = ingreso_promedio * 12

    total = pd.DataFrame(
        [
            {
                "Mes inicial": str(df["Mes_Period"].min()),
                "Mes final": str(df["Mes_Period"].max()),
                "Meses con datos": meses_con_datos,
                "Ingreso acumulado": ingreso_acumulado,
                "Promedio mensual": ingreso_promedio,
                "Proyección anual": proyeccion_anual,
            }
        ]
    )
    # Proyección y cumplimiento por vendedor. Esto lo saque ya que da información adicional sobre
    # como va cada vendedor basado en su presupuesto y que ingreso a tenido
    by_seller = df.groupby(["Vendedor_key", "REGION", "CIUDAD"], as_index=False).agg(
        Ingreso_Acumulado=("Ingreso Operación", "sum"),
        Meses_Con_Datos=("Mes_Period", "nunique"),
        # Presupuesto: uso "first" porque en la base ya lo crucé por vendedor.
        # Entonces debería ser el mismo valor repetido en todas las filas de ese vendedor.
        Presupuesto=("Presupuesto", "first"),
    )
    # Aquí saco el promedio mensual de cada vendedor, y también detecta anomalias con el presupuesto
    by_seller["Promedio_Mensual"] = by_seller["Ingreso_Acumulado"] / by_seller[
        "Meses_Con_Datos"
    ].replace(0, np.nan)
    by_seller["Proyección_Anual"] = by_seller["Promedio_Mensual"] * 12
    by_seller["% Cumplimiento Presupuesto"] = (
        by_seller["Ingreso_Acumulado"] / by_seller["Presupuesto"] * 100
    )
    by_seller["% Cumplimiento Proyectado"] = (
        by_seller["Proyección_Anual"] / by_seller["Presupuesto"] * 100
    )
    by_seller = by_seller.sort_values("% Cumplimiento Presupuesto", ascending=False)

    return {
        "KPI3_Proyeccion_Total": total,
        "KPI4_Cumplimiento_Vendedor": by_seller,
    }


# ============================================================
# 4) Ejercicio 3 análisis libre
# ============================================================

# Ejercicio 3, aquí estoy analizando la cartera, concentración, reactivación de clientes para cada vendedor.
# Mi idea aquí es sacar insights que ayuden a la gestión comercial y toma de decisiones, no solo KPIs.
# Y considero que estos que identifique pueden ayudar mucho a tomar buenas decisiones basadas en información.
# Tambien puedo determinar el riesgo de ciertas carteras, volumen de clientes, etc.


def ejercicio_3_analisis_libre(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # 1) Primero saco el total de ingreso por vendedor
    vendedor_acumulado = (
        df.groupby(["Vendedor_key", "CIUDAD", "REGION"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Acumulado"})  # type: ignore[arg-type]
    )

    # 2) Luego bajo el nivel a vendedor-cliente para ver aportes por cliente
    cartera_vendedor_cliente = (
        df.groupby(["Vendedor_key", "No. Cliente"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Ingreso Cliente"})  # type: ignore[arg-type]
    ).merge(
        vendedor_acumulado[["Vendedor_key", "Ingreso Acumulado"]],
        on="Vendedor_key",
        how="left",
    )

    # 3) Con el total del vendedor, calculo qué % representa cada cliente
    cartera_vendedor_cliente["% del Vendedor"] = cartera_vendedor_cliente[
        "Ingreso Cliente"
    ] / cartera_vendedor_cliente["Ingreso Acumulado"].replace(0, np.nan)

    # 4) Utilice el indicador de concentración HHI. Para poder determinar si el ingreso de un
    # vendedor depende de pocos clientes o si esta bien distribuido. Asi es mas facil conocer
    # el riesgo de su cartera.
    # Si está alto es decir arriba de 0.18, significa que el vendedor depende de pocos clientes y es riesgoso.
    # Ejemplo súper simple: si 1 cliente aporta 100% -> 1^2 = 1 (máxima concentración).
    # Si tengo 10 clientes iguales (10% cada uno) -> 10*(0.1^2) = 0.10 (más distribuido).
    concentracion_hhi = (
        cartera_vendedor_cliente.groupby("Vendedor_key", as_index=False)[
            "% del Vendedor"
        ]
        .apply(lambda s: float((s.fillna(0) ** 2).sum()))
        .rename(columns={"% del Vendedor": "HHI Concentración"})
    )

    # 5) Ranking por cliente para poder sumar el Top 5, con esto puedo determinar si un vendedor
    # es dependiente de ciertos clientes, esto ayuda a determinar el riesgo de su cartera.
    cartera_vendedor_cliente["Rank Cliente"] = cartera_vendedor_cliente.groupby(
        "Vendedor_key"
    )["Ingreso Cliente"].rank(method="first", ascending=False)

    top5_ingreso = (
        cartera_vendedor_cliente[cartera_vendedor_cliente["Rank Cliente"] <= 5]
        .groupby("Vendedor_key", as_index=False)["Ingreso Cliente"]
        .sum()
        .rename(columns={"Ingreso Cliente": "Ingreso Top 5 Clientes"})
    )

    # 6) Tamaño de cartera: cuántos clientes únicos maneja cada vendedor para conocer el volumen de
    # clientes que maneja cada uno
    clientes_unicos = (
        cartera_vendedor_cliente.groupby("Vendedor_key", as_index=False)["No. Cliente"]
        .nunique()
        .rename(columns={"No. Cliente": "Clientes Únicos"})
    )

    # 7) Armo una tabla resumen por vendedor
    cartera_vendedor = (
        vendedor_acumulado.merge(clientes_unicos, on="Vendedor_key", how="left")
        .merge(top5_ingreso, on="Vendedor_key", how="left")
        .merge(concentracion_hhi, on="Vendedor_key", how="left")
    )

    cartera_vendedor["Ingreso Top 5 Clientes"] = cartera_vendedor[
        "Ingreso Top 5 Clientes"
    ].fillna(0.0)
    cartera_vendedor["% Ingreso Top 5"] = (
        cartera_vendedor["Ingreso Top 5 Clientes"]
        / cartera_vendedor["Ingreso Acumulado"].replace(0, np.nan)
        * 100
    )
    # 8) Cree una bandera que da una alerta para priorizar seguimiento en caso de que el 70% del
    # ingreso de un vendedor corresponda solo del Top 5 de sus clientes ya que es de alto riesgo
    cartera_vendedor["Riesgo (Concentración)"] = np.where(
        cartera_vendedor["% Ingreso Top 5"] >= 70, "ALTO", "NORMAL"
    )
    cartera_vendedor = cartera_vendedor.sort_values("% Ingreso Top 5", ascending=False)

    # 9) Reactivación: busco clientes de alto valor que llevan días sin comprar, esto permite tomar
    # decisiones como por ejemplo ponerse en contacto con el cliente para reactivar compras.
    fecha_max = df["Fecha Operación"].max()
    clientes_detalle = df.groupby(["Vendedor_key", "No. Cliente"], as_index=False).agg(
        Ingreso_Total=("Ingreso Operación", "sum"),
        Compras=("Guia", "count"),
        Última_Compra=("Fecha Operación", "max"),
        Primera_Compra=("Fecha Operación", "min"),
        Ciudad=("CIUDAD", "first"),
        Región=("REGION", "first"),
    )
    clientes_detalle["Días desde última compra"] = (
        fecha_max - clientes_detalle["Última_Compra"]
    ).dt.days

    # Defini a clientes de Alto valor como los que son el top 20% de ingreso dentro de cada vendedor.
    # Lo hago por vendedor ya que cada cartera es distinta.
    umbral = clientes_detalle.groupby("Vendedor_key")["Ingreso_Total"].transform(
        lambda s: s.quantile(0.80)
    )
    clientes_detalle["Alto_Valor"] = clientes_detalle["Ingreso_Total"] >= umbral

    # Aquí defino a un cliente dormido como uno que lleva 30+ días sin compra (solo con fines prácticos),
    #  por lo que una acción quizás sea necesaria.
    clientes_a_reactivar = (
        clientes_detalle[
            (clientes_detalle["Alto_Valor"])
            & (clientes_detalle["Días desde última compra"] >= 30)
        ]
        .sort_values(
            ["Días desde última compra", "Ingreso_Total"], ascending=[False, False]
        )
        .drop(columns=["Alto_Valor"])
    )

    return {
        "E3_Cartera_Vendedor": cartera_vendedor,
        "E3_Reactivar_Clientes": clientes_a_reactivar,
    }


# ============================================================
# 5) Export a Excel (entregable final)
# ============================================================


def safe_sheet_name(name: str, used: set[str]) -> str:
    # Excel limita el nombre de hoja a 31 caracteres. Aquí yo me aseguro de: recortar a 31 y
    # evitar nombres repetidos
    base = name.strip().replace("/", "_")
    base = base[:31] if len(base) > 31 else base
    if base not in used:
        return base
    # Si ya existe solo uso sufijos hasta que no exista
    for i in range(1, 1000):
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in used:
            return candidate
    raise ValueError("No fue posible generar un nombre de hoja único para Excel.")


def export_excel(tables: dict[str, pd.DataFrame], output_path: Path) -> None:
    # Exporto todas las tablas a un Excel con las pestañas que ya defini antes.
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        used: set[str] = set()
        for key, df in tables.items():
            sheet = safe_sheet_name(key, used)
            used.add(sheet)
            df.to_excel(writer, sheet_name=sheet, index=False)


def build_all_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Armo un diccionario {nombre_pestaña: dataframe} para exportar fácil. Así despues me facilita
    # quitar o agregar pestañas sin dañar nada
    tables: dict[str, pd.DataFrame] = {}
    tables["Base_Cruzada"] = df.drop(
        columns=["Mes_Period", "Trimestre_Period"], errors="ignore"
    )
    tables.update(kpi_ingreso_mensual(df))
    tables.update(kpi_crecimiento_trimestral(df))
    tables.update(kpi_proyeccion(df))
    tables.update(ejercicio_3_analisis_libre(df))
    return tables


def run_selftest() -> None:
    # Esto solo hace pruebas rápidas para asegurar que lo básico funciona.
    assert normalize_vendor_name("  Juan  Pérez ") == "JUAN PÉREZ"
    assert (
        normalize_vendor_name("001  Juan Pérez", remove_leading_digits=True)
        == "JUAN PÉREZ"
    )
    assert presupuesto_to_common_key("PEREZ JUAN") == "JUAN PEREZ"
    used: set[str] = set()
    a = safe_sheet_name("A" * 40, used)
    used.add(a)
    b = safe_sheet_name("A" * 40, used)
    assert a != b
    assert len(a) <= 31 and len(b) <= 31


def parse_args(argv: list[str]) -> argparse.Namespace:
    # Uso Argumentos para poder correr el script desde terminal sin tocar el código.
    # Lo dejo así porque en un entorno de empresa es más práctico:
    # corren 1 comando y les genero el Excel final. Y asi esto que hice funciona para cualquier
    # base de datos con la misma estructura, la automatización se vuelve repetible y escalable
    p = argparse.ArgumentParser()
    p.add_argument(
        "--input", required=True, help="Ruta al archivo Excel de entrada (.xlsx)"
    )
    p.add_argument(
        "--output",
        default="CasoPractico_Analisis.xlsx",
        help="Ruta del Excel de salida",
    )
    p.add_argument(
        "--quiet", action="store_true", help="No imprimir la ruta del archivo de salida"
    )
    p.add_argument(
        "--selftest", action="store_true", help="Ejecuta pruebas rápidas y termina"
    )
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    # Este solo es el punto de entrada: aquí conecto todo el flujo
    args = parse_args(argv)

    # Selftest prueba rápida cuando se corre desde la terminal. No analiza el Excel, solo verifica
    # que funciones clave estén funcionando.
    if args.selftest:
        run_selftest()
        return 0

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {input_path}")

    # 1) Cargo hojas del Excel original (BD, Ciudad-Region, Presupuesto)
    bd, cdrg, pres = load_sheets(input_path)

    # 2) Preparo base final (limpieza + cruces)
    df = prepare_base(bd, cdrg, pres)

    # 3) Calculo KPIs + ejercicio 3
    tables = build_all_tables(df)

    output_path = Path(args.output).expanduser().resolve()
    # 4) Exporto todo a Excel (este es el archivo que voy a entregar)
    export_excel(tables, output_path)

    # Si no usan --quiet, imprimo la ruta para que sea fácil encontrar el archivo.
    # (Así la persona que lo corre sabe exactamente dónde quedó el Excel final.)
    if not args.quiet:
        print(str(output_path))
    return 0


if __name__ == "__main__":
    # Con esto se puede ejecutar el script de forma directa
    raise SystemExit(main(sys.argv[1:]))
