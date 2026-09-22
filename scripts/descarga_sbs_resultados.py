#!/usr/bin/env python3
"""
Descarga los reportes B-2201 de la SBS, mes a mes, y extrae el Estado de
Ganancias y Pérdidas por empresa bancaria: desde "INGRESOS FINANCIEROS"
hasta "UTILIDAD ( PÉRDIDA ) NETA" o "RESULTADO NETO DEL EJERCICIO"
(inclusive), lo que aparezca en cada archivo.

Formato real del archivo (verificado contra B-2201-jl2026.XLS y
B-2201-ag2012.XLS):
- El Estado de Ganancias y Pérdidas esta en una hoja APARTE del Balance
  General, dentro del mismo libro. El nombre de esa hoja varia con el
  tiempo: archivos recientes usan "2"; archivos antiguos usan "06-EGP",
  "06-EGP (P)" o "EGP" (con variantes de espacios/parentesis).
  _elegir_hoja_egp() normaliza el nombre y elige la hoja correcta.
- Layout identico al del Balance General: cada empresa bancaria ocupa 3
  columnas contiguas (Moneda Nacional, Moneda Extranjera, Total), con el
  nombre del banco en la fila justo encima de "MN"/"ME"/"TOTAL", y
  columnas separadoras en blanco entre bloques de bancos.
- La columna con el nombre de la cuenta no siempre es la A (puede estar
  corrida a la B u otra); se detecta dinamicamente buscando el texto
  "INGRESOS FINANCIEROS" en las columnas anteriores al primer bloque de
  banco, igual que se hace con "TOTAL ACTIVO"/"OBLIGACIONES CON EL
  PUBLICO" en los otros dos scripts hermanos (descarga_sbs_activos.py y
  descarga_sbs_pasivos.py).
- La etiqueta de la ultima fila del estado de resultados cambio de
  nombre en algun momento del historico: en 2012 es "UTILIDAD ( PÉRDIDA )
  NETA" (con espacios dentro del parentesis) y en 2026 es "RESULTADO NETO
  DEL EJERCICIO". Se buscan ambas variantes y se usa la que aparezca.
- NO se asume que la cantidad/orden de cuentas ni el orden de bancos sea
  igual entre meses distintos, ni que coincida con el de Activo/Pasivo del
  mismo archivo: cada mes (y cada hoja) se procesa de forma independiente,
  con sus propias cuentas tal cual vienen. El consolidado final copia cada
  mes con su PROPIO encabezado repetido justo antes de sus datos (ver
  guardar_excel_por_bloques()), sin alinear cuentas de distintos meses por
  nombre.

IMPORTANTE sobre la descarga:
- El host `intranet2.sbs.gob.pe` no es accesible desde el entorno de este
  agente (bloqueado por politica de red del sandbox). El parseo de arriba
  SI fue validado con archivos reales que subio el usuario, pero la
  descarga automatica debe correrse desde una maquina con acceso real a
  esa intranet (red de la SBS / VPN).

Uso:
    python descarga_sbs_resultados.py                       # descarga todo el rango 2002-01 a 2026-07
    python descarga_sbs_resultados.py --start 2020-01 --end 2021-12
    python descarga_sbs_resultados.py --inspect 2026-07      # solo descarga y vuelca el crudo de un mes
    python descarga_sbs_resultados.py --local-file ruta.xls --fecha 2026-07  # parsea un archivo ya descargado, sin red
"""

from __future__ import annotations

import argparse
import io
import re
import sys
import time
import unicodedata
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter, Retry

BASE_URL = "https://intranet2.sbs.gob.pe/estadistica/financiera"
CODIGO_REPORTE = "B-2201"

MESES = {
    1: ("Enero", ["en"]),
    2: ("Febrero", ["fe"]),
    3: ("Marzo", ["mr", "ma"]),
    4: ("Abril", ["ab", "ap"]),
    5: ("Mayo", ["my", "ma"]),
    6: ("Junio", ["jn", "ju"]),
    7: ("Julio", ["jl"]),
    8: ("Agosto", ["ag"]),
    9: ("Setiembre", ["se", "st", "sp"]),
    10: ("Octubre", ["oc", "ot"]),
    11: ("Noviembre", ["nv", "no"]),
    12: ("Diciembre", ["dc", "di"]),
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

RAW_DIR = Path("data/raw/sbs_b2201")
PROCESSED_DIR = Path("data/processed")

PALABRA_INICIO = "INGRESOS FINANCIEROS"
# Cualquiera de las dos variantes marca el final del bloque (inclusive).
PALABRAS_FIN = ("UTILIDAD ( PERDIDA ) NETA", "RESULTADO NETO DEL EJERCICIO")


def _sin_tildes(texto) -> str:
    texto = str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    ).upper().strip()


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    retries = Retry(total=4, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    return s


def construir_urls_candidatas(year: int, month: int) -> list[str]:
    nombre_mes, abreviados = MESES[month]
    anios = [str(year), f"{year % 100:02d}"]
    return [
        f"{BASE_URL}/{year}/{nombre_mes}/{CODIGO_REPORTE}-{abv}{anio}.XLS"
        for abv in abreviados
        for anio in anios
    ]


def descargar_mes(session: requests.Session, year: int, month: int, cache_dir: Path) -> Path | None:
    """Descarga el archivo del mes (usa cache local si ya existe). Devuelve la ruta local o None si fallo."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    destino = cache_dir / f"{year}-{month:02d}.xls"
    if destino.exists() and destino.stat().st_size > 0:
        return destino

    for url in construir_urls_candidatas(year, month):
        try:
            resp = session.get(url, timeout=30)
        except requests.RequestException as exc:
            print(f"  [WARN] {url} -> error de red: {exc}", file=sys.stderr)
            continue
        if resp.status_code == 200 and len(resp.content) > 1000:
            destino.write_bytes(resp.content)
            print(f"  OK  {year}-{month:02d}: {url}")
            return destino
        print(f"  [skip] {url} -> HTTP {resp.status_code}")
    print(f"  [FALLA] No se pudo descargar {year}-{month:02d} con ninguna variante de nombre.", file=sys.stderr)
    return None


OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _elegir_hoja_egp(nombres_hoja: list[str]) -> str:
    """
    Elige la hoja del Estado de Ganancias y Perdidas: "2" en archivos
    recientes, "06-EGP", "06-EGP (P)" o "EGP" en antiguos (con variantes de
    espacios/parentesis). Se compara normalizando a mayusculas y sin
    caracteres no alfanumericos.
    """
    normalizados = {n: re.sub(r"[^A-Z0-9]", "", n.upper()) for n in nombres_hoja}
    for nombre, norm in normalizados.items():
        if norm == "2":
            return nombre
    for nombre, norm in normalizados.items():
        if norm.startswith("06EGP") or norm.startswith("EGP"):
            return nombre
    print(
        f"  [WARN] No se reconocio el nombre de hoja del Estado de Ganancias y Perdidas entre "
        f"{nombres_hoja}; se usa la ultima hoja ('{nombres_hoja[-1]}'). Verificar manualmente.",
        file=sys.stderr,
    )
    return nombres_hoja[-1]


def _leer_hoja_egp(path: Path) -> pd.DataFrame:
    """Ubica y lee la hoja del Estado de Ganancias y Perdidas, sin importar
    si el archivo es xlsx real, xls binario clasico, o HTML disfrazado de
    .xls, ni como se llame la hoja."""
    contenido = path.read_bytes()

    if contenido[:2] == b"PK":
        engine = "openpyxl"
    elif contenido[:8] == OLE2_MAGIC:
        engine = "xlrd"
    else:
        cabecera = contenido[:512].lstrip().lower()
        if cabecera.startswith(b"<html") or b"<table" in cabecera or cabecera.startswith(b"<!doctype"):
            tablas = pd.read_html(io.BytesIO(contenido), header=None)
            # Un archivo HTML disfrazado normalmente trae una tabla por hoja
            # impresa; si hay mas de una, la 2da suele ser el EGP (misma
            # convencion que las hojas "1"/"2" de un libro real).
            return tablas[1] if len(tablas) > 1 else tablas[0]
        engine = None

    xl = pd.ExcelFile(io.BytesIO(contenido), engine=engine)
    hoja = _elegir_hoja_egp(xl.sheet_names)
    return xl.parse(hoja, header=None)


def _buscar_texto_en_columna(df: pd.DataFrame, col: int, texto_normalizado, desde_fila: int = 0):
    """Devuelve el indice de la primera fila (desde `desde_fila`) donde la
    columna `col`, normalizada, coincide con el texto buscado (un string, o
    una tupla de variantes aceptadas). None si no aparece."""
    columna = df.iloc[desde_fila:, col].map(_sin_tildes)
    if isinstance(texto_normalizado, tuple):
        coincidencias = columna[columna.isin(texto_normalizado)]
    else:
        coincidencias = columna[columna == texto_normalizado]
    if coincidencias.empty:
        return None
    return coincidencias.index[0]


def _ubicar_bloque_resultados(df: pd.DataFrame) -> tuple[int, int, int, int, int]:
    """
    Devuelve (fila_banco, fila_submoneda, fila_inicio, fila_fin, col_cuenta)
    para el bloque del Estado de Ganancias y Perdidas (desde
    PALABRA_INICIO hasta la primera de PALABRAS_FIN que aparezca).
    """
    fila_submoneda = None
    for r in range(min(15, len(df))):
        valores = df.iloc[r].astype(str).str.strip().str.upper()
        if (valores == "MN").sum() >= 2 and (valores == "ME").sum() >= 2:
            fila_submoneda = r
            break
    if fila_submoneda is None:
        raise ValueError("No se encontro la fila de encabezado MN/ME/TOTAL en las primeras 15 filas")

    fila_banco = fila_submoneda - 1

    submoneda = df.iloc[fila_submoneda].astype(str).str.strip().str.upper()
    primer_col_banco = next((c for c in range(df.shape[1]) if submoneda.iloc[c] == "MN"), None)
    if primer_col_banco is None:
        raise ValueError("No se encontro ninguna columna 'MN' en la fila de sub-encabezado")

    col_cuenta = None
    fila_inicio = None
    for c in range(primer_col_banco):
        fila = _buscar_texto_en_columna(df, c, PALABRA_INICIO, desde_fila=fila_submoneda)
        if fila is not None:
            col_cuenta = c
            fila_inicio = fila
            break
    if col_cuenta is None:
        raise ValueError(f"No se encontro '{PALABRA_INICIO}' en ninguna columna 0..{primer_col_banco - 1}")

    fila_fin = _buscar_texto_en_columna(df, col_cuenta, PALABRAS_FIN, desde_fila=fila_inicio + 1)
    if fila_fin is None:
        raise ValueError(f"No se encontro ninguna variante de {PALABRAS_FIN} despues de la fila {fila_inicio}")

    return fila_banco, fila_submoneda, fila_inicio, fila_fin, col_cuenta


def extraer_resultados(df: pd.DataFrame, fecha) -> pd.DataFrame:
    """
    Extrae, de la hoja del Estado de Ganancias y Perdidas, el bloque desde
    "INGRESOS FINANCIEROS" hasta "UTILIDAD ( PERDIDA ) NETA" / "RESULTADO
    NETO DEL EJERCICIO" (inclusive) por empresa bancaria, en formato largo:
    fecha, banco, es_total, cuenta, moneda_nacional, moneda_extranjera, total
    """
    fila_banco, fila_submoneda, fila_inicio, fila_fin, col_cuenta = _ubicar_bloque_resultados(df)

    fila_datos_fin = fila_fin + 1  # inclusive

    submoneda = df.iloc[fila_submoneda].astype(str).str.strip().str.upper()
    nombres_banco = df.iloc[fila_banco]

    cuentas_crudo = df.iloc[fila_inicio:fila_datos_fin, col_cuenta]
    # Mismo criterio que en activos/pasivos: las filas separadoras se
    # identifican por el TEXTO de la columna de cuenta (blanco para todos
    # los bancos por igual), no por los valores numericos de cada banco.
    filas_validas = cuentas_crudo.notna() & (cuentas_crudo.astype(str).str.strip() != "")
    cuentas = cuentas_crudo[filas_validas].astype(str).str.strip()

    bloques = []
    ncols = df.shape[1]
    for c in range(ncols):
        if submoneda.iloc[c] != "MN":
            continue
        if c + 2 >= ncols or submoneda.iloc[c + 1] != "ME" or submoneda.iloc[c + 2] != "TOTAL":
            continue
        banco = nombres_banco.iloc[c]
        if not isinstance(banco, str) or not banco.strip():
            continue
        banco = re.sub(r"\s+", " ", banco.strip())

        sub = df.iloc[fila_inicio:fila_datos_fin, [c, c + 1, c + 2]].copy()
        sub.columns = ["moneda_nacional", "moneda_extranjera", "total"]
        sub = sub[filas_validas.values]
        sub.insert(0, "cuenta", cuentas.values)
        sub.insert(0, "es_total", banco.lower().startswith("total"))
        sub.insert(0, "banco", banco)
        sub.insert(0, "fecha", fecha)
        bloques.append(sub)

    if not bloques:
        raise ValueError("No se detecto ningun bloque de banco (MN/ME/TOTAL) en el Estado de Ganancias y Perdidas")

    resultado = pd.concat(bloques, ignore_index=True)
    for col in ("moneda_nacional", "moneda_extranjera", "total"):
        resultado[col] = pd.to_numeric(resultado[col], errors="coerce")
    return resultado


def a_formato_ancho(largo: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el formato largo (una fila por banco+cuenta) al formato ancho:
    una fila por banco x moneda (MN/ME/TOTAL) y una columna por cuenta.
    Se arma por POSICION (no por nombre de cuenta), por si algun nombre se
    repite dentro del mismo mes.
    """
    orden_bancos = list(dict.fromkeys(largo["banco"]))
    primer_banco = largo.loc[largo["banco"] == orden_bancos[0]]
    orden_cuentas = primer_banco["cuenta"].tolist()

    metas = []
    filas_valores = []
    for banco in orden_bancos:
        bloque = largo.loc[largo["banco"] == banco].reset_index(drop=True)
        if bloque["cuenta"].tolist() != orden_cuentas:
            raise ValueError(
                f"El banco '{banco}' no tiene las mismas cuentas, en el mismo orden, que '{orden_bancos[0]}'"
            )
        fecha = bloque["fecha"].iloc[0]
        es_total = bloque["es_total"].iloc[0]
        for col_origen, moneda in (
            ("moneda_nacional", "MN"),
            ("moneda_extranjera", "ME"),
            ("total", "TOTAL"),
        ):
            metas.append((fecha, banco, es_total, moneda))
            filas_valores.append(bloque[col_origen].tolist())

    meta = pd.DataFrame(metas, columns=["fecha", "banco", "es_total", "moneda"])
    valores = pd.DataFrame(filas_valores, columns=orden_cuentas)
    ancho = pd.concat([meta, valores], axis=1)

    ancho["banco"] = pd.Categorical(ancho["banco"], categories=orden_bancos, ordered=True)
    ancho["moneda"] = pd.Categorical(ancho["moneda"], categories=["MN", "ME", "TOTAL"], ordered=True)
    ancho = ancho.sort_values(["fecha", "banco", "moneda"], kind="stable").reset_index(drop=True)
    return ancho


def ultimo_dia_mes(year: int, month: int):
    """Fecha (date) del ultimo dia del mes."""
    return (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).date()


def guardar_excel(df: pd.DataFrame, destino: Path) -> None:
    """Guarda en .xlsx con la columna 'fecha' como fecha real formateada mmm-aa (ej. jul-26)."""
    df.to_excel(destino, index=False, sheet_name="Resultados")

    from openpyxl import load_workbook

    wb = load_workbook(destino)
    ws = wb["Resultados"]
    col_fecha = df.columns.get_loc("fecha") + 1
    for fila in range(2, ws.max_row + 1):
        ws.cell(row=fila, column=col_fecha).number_format = "mmm-yy"
    ws.freeze_panes = "A2"
    ws.column_dimensions[ws.cell(row=1, column=col_fecha).column_letter].width = 10
    ws.column_dimensions[ws.cell(row=1, column=df.columns.get_loc("banco") + 1).column_letter].width = 32
    wb.save(destino)


def guardar_excel_por_bloques(piezas: list[pd.DataFrame], destino: Path) -> None:
    """
    Guarda el consolidado multi-mes copiando cada mes con su PROPIO
    encabezado de cuentas justo antes de sus datos, uno debajo del otro
    (sin alinear cuentas de distintos meses por nombre).
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    fila = 0
    for df in piezas:
        fila += 1
        for col_idx, nombre_col in enumerate(df.columns, start=1):
            ws.cell(row=fila, column=col_idx, value=nombre_col)
        for _, registro in df.iterrows():
            fila += 1
            for col_idx, valor in enumerate(registro, start=1):
                ws.cell(row=fila, column=col_idx, value=valor)
            ws.cell(row=fila, column=1).number_format = "mmm-yy"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 32
    wb.save(destino)


def rango_meses(inicio: str, fin: str):
    y0, m0 = (int(x) for x in inicio.split("-"))
    y1, m1 = (int(x) for x in fin.split("-"))
    y, m = y0, m0
    while (y, m) <= (y1, m1):
        yield y, m
        m += 1
        if m > 12:
            m = 1
            y += 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--start", default="2002-01", help="Mes inicial YYYY-MM (default 2002-01)")
    ap.add_argument("--end", default="2026-07", help="Mes final YYYY-MM (default 2026-07, cierre julio 2026)")
    ap.add_argument("--sleep", type=float, default=1.0, help="Segundos de espera entre descargas")
    ap.add_argument("--inspect", metavar="YYYY-MM", help="Descarga y vuelca el crudo de un solo mes para calibrar la logica de recorte")
    ap.add_argument("--local-file", metavar="RUTA", help="Parsea un archivo ya descargado localmente (sin red), usado junto con --fecha")
    ap.add_argument("--fecha", metavar="YYYY-MM", help="Fecha a usar junto con --local-file")
    ap.add_argument("--outdir", default=str(PROCESSED_DIR), help="Carpeta de salida para los CSV procesados")
    args = ap.parse_args()

    salida = Path(args.outdir)
    salida.mkdir(parents=True, exist_ok=True)

    if args.local_file:
        if not args.fecha:
            print("--local-file requiere --fecha YYYY-MM", file=sys.stderr)
            sys.exit(1)
        year, month = (int(x) for x in args.fecha.split("-"))
        df_crudo = _leer_hoja_egp(Path(args.local_file))
        resultados = extraer_resultados(df_crudo, ultimo_dia_mes(year, month))
        ancho = a_formato_ancho(resultados)
        destino = salida / f"resultados_{args.fecha}.xlsx"
        guardar_excel(ancho, destino)
        print(f"Listo: {len(ancho)} filas guardadas en {destino}")
        print(ancho.head(9).to_string())
        return

    session = _session()

    if args.inspect:
        year, month = (int(x) for x in args.inspect.split("-"))
        path = descargar_mes(session, year, month, RAW_DIR)
        if not path:
            sys.exit(1)
        df = _leer_hoja_egp(path)
        volcado = salida / f"inspect_resultados_{year}-{month:02d}.csv"
        df.to_csv(volcado, index=False)
        print(f"Crudo volcado en {volcado} ({df.shape[0]} filas x {df.shape[1]} cols). Revisalo para ajustar las constantes del script.")
        return

    piezas = []
    meses_omitidos = []
    for year, month in rango_meses(args.start, args.end):
        etiqueta = f"{year}-{month:02d}"
        print(f"Procesando {etiqueta}...")
        path = descargar_mes(session, year, month, RAW_DIR)
        if not path:
            meses_omitidos.append((etiqueta, "no se pudo descargar"))
            continue
        try:
            df_crudo = _leer_hoja_egp(path)
            resultados = extraer_resultados(df_crudo, ultimo_dia_mes(year, month))
            ancho = a_formato_ancho(resultados)
        except Exception as exc:
            print(f"  [ERROR] No se pudo parsear {etiqueta}: {exc}", file=sys.stderr)
            meses_omitidos.append((etiqueta, str(exc)))
            time.sleep(args.sleep)
            continue

        # Cada mes aporta sus propias cuentas, en su propio orden y
        # cantidad, sin exigir que coincidan con las de otros meses. Se
        # apila tal cual, con su propio encabezado repetido antes de sus
        # datos (ver guardar_excel_por_bloques()).
        piezas.append(ancho)
        time.sleep(args.sleep)

    if not piezas:
        print("No se logro procesar ningun mes.", file=sys.stderr)
        sys.exit(1)

    total_filas = sum(len(p) for p in piezas)

    if meses_omitidos:
        print(f"\n[RESUMEN] {len(meses_omitidos)} mes(es) NO quedaron en el consolidado:", file=sys.stderr)
        for etiqueta, motivo in meses_omitidos:
            print(f"   - {etiqueta}: {motivo}", file=sys.stderr)

    destino = salida / "resultados_bancos_2002_2026.xlsx"
    guardar_excel_por_bloques(piezas, destino)
    print(f"\nListo: {len(piezas)} mes(es), {total_filas} filas de datos guardadas en {destino}")


if __name__ == "__main__":
    main()
