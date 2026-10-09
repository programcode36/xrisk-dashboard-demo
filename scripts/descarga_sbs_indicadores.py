#!/usr/bin/env python3
"""
Descarga los reportes B-2401 (Indicadores Financieros por Empresa
Bancaria) de la SBS, mes a mes, y transpone cada variable con su propio
dato por empresa bancaria.

Formato real del archivo (verificado contra B-2401-en2002.XLS y
B-2401-jl2026.XLS):
- Cada archivo tiene UNA sola hoja (a diferencia de B-2201, que tiene 2).
  El nombre de esa hoja varia con el tiempo ("14-Indicadores" en 2002,
  "3" en 2026) asi que simplemente se usa la primera y unica hoja del
  libro, sin necesidad de reconocer un nombre especifico.
- A diferencia del Balance General, aca cada banco ocupa UNA sola columna
  (son ratios/porcentajes, no hay desglose MN/ME/TOTAL).
- Hay demasiados bancos para una sola fila impresa, asi que la tabla se
  repite en un segundo bloque de columnas en la MISMA fila (con una
  columna en blanco de separacion y una columna de etiqueta duplicada,
  igual que el patron ya visto en B-2201 con la columna "Activo"
  repetida). Se ignora esa columna de etiqueta duplicada: solo se usan
  las columnas donde la fila de nombres de banco realmente trae un banco.
- Las variables estan agrupadas bajo titulos de seccion (SOLVENCIA,
  CALIDAD DE ACTIVOS, EFICIENCIA Y GESTION, RENTABILIDAD, LIQUIDEZ, etc.)
  y al final suele haber notas al pie. Estos titulos de grupo y las notas
  tienen texto en la columna de "variable" pero NINGUN dato numerico en
  ninguna columna de banco esa misma fila. Esa es la señal que se usa
  para DISTINGUIR una variable real (se copia) de un titulo de grupo o
  nota al pie (se descarta), en vez de buscar nombres de variable
  especificos: el usuario confirmo que el listado de variables cambia de
  un mes a otro en 24 años de historico, asi que no hay un nombre fijo de
  inicio/fin que buscar como en los otros reportes.
- NO se asume que la cantidad/orden/nombre de las variables sea igual
  entre meses distintos: cada mes aporta sus propias variables, tal como
  vienen en su archivo. El consolidado final copia cada mes con su PROPIO
  encabezado repetido justo antes de sus datos (ver
  guardar_excel_por_bloques()), igual que en los scripts hermanos
  (descarga_sbs_activos.py, descarga_sbs_pasivos.py,
  descarga_sbs_resultados.py).

IMPORTANTE sobre la descarga:
- El host `intranet2.sbs.gob.pe` no es accesible desde el entorno de este
  agente (bloqueado por politica de red del sandbox). El parseo de arriba
  SI fue validado con archivos reales que subio el usuario, pero la
  descarga automatica debe correrse desde una maquina con acceso real a
  esa intranet (red de la SBS / VPN).

Uso:
    python descarga_sbs_indicadores.py                       # descarga todo el rango 2002-01 a 2026-07
    python descarga_sbs_indicadores.py --start 2020-01 --end 2021-12
    python descarga_sbs_indicadores.py --inspect 2026-07      # solo descarga y vuelca el crudo de un mes
    python descarga_sbs_indicadores.py --local-file ruta.xls --fecha 2026-07  # parsea un archivo ya descargado, sin red
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
CODIGO_REPORTE = "B-2401"

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

RAW_DIR = Path("data/raw/sbs_b2401")
PROCESSED_DIR = Path("data/processed")

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


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


def _leer_hoja_indicadores(path: Path) -> pd.DataFrame:
    """Lee la unica hoja del libro de indicadores, sin importar si el
    archivo es xlsx real, xls binario clasico, o HTML disfrazado de .xls."""
    contenido = path.read_bytes()

    if contenido[:2] == b"PK":
        engine = "openpyxl"
    elif contenido[:8] == OLE2_MAGIC:
        engine = "xlrd"
    else:
        cabecera = contenido[:512].lstrip().lower()
        if cabecera.startswith(b"<html") or b"<table" in cabecera or cabecera.startswith(b"<!doctype"):
            tablas = pd.read_html(io.BytesIO(contenido), header=None)
            return tablas[0]
        engine = None

    xl = pd.ExcelFile(io.BytesIO(contenido), engine=engine)
    if len(xl.sheet_names) != 1:
        print(
            f"  [WARN] Se esperaba 1 sola hoja y el libro tiene {len(xl.sheet_names)} "
            f"({xl.sheet_names}); se usa la primera.",
            file=sys.stderr,
        )
    return xl.parse(xl.sheet_names[0], header=None)


def _ubicar_bancos(df: pd.DataFrame) -> tuple[int, list[int], int]:
    """
    Devuelve (fila_banco, columnas_banco, col_cuenta).

    fila_banco es la fila con los nombres de banco: se detecta como la
    primera fila (dentro de las primeras 10) cuya columna 0 esta vacia
    pero que trae 2 o mas celdas de texto en el resto de columnas (los
    titulos/fecha de las filas de arriba tienen texto solo en la columna
    0). columnas_banco son las columnas de esa fila que efectivamente
    traen un nombre de banco (se excluye asi la columna de etiqueta
    duplicada que aparece cuando la tabla se repite en un segundo bloque
    de columnas en la misma fila, ademas de la columna en blanco que las
    separa). col_cuenta es la columna inmediatamente anterior a la primera
    columna de banco (normalmente la A).
    """
    fila_banco = None
    for r in range(min(10, len(df))):
        col0 = df.iloc[r, 0]
        resto_textos = sum(1 for v in df.iloc[r, 1:] if isinstance(v, str) and v.strip())
        if not isinstance(col0, str) and resto_textos >= 2:
            fila_banco = r
            break
    if fila_banco is None:
        raise ValueError("No se encontro la fila con los nombres de banco en las primeras 10 filas")

    columnas_banco = [
        c for c in range(1, df.shape[1])
        if isinstance(df.iloc[fila_banco, c], str) and df.iloc[fila_banco, c].strip()
    ]
    if not columnas_banco:
        raise ValueError(f"La fila {fila_banco} no trae ningun nombre de banco")

    col_cuenta = min(columnas_banco) - 1
    return fila_banco, columnas_banco, col_cuenta


def extraer_indicadores(df: pd.DataFrame, fecha) -> pd.DataFrame:
    """
    Extrae, de la unica hoja del reporte, cada variable (indicador) con su
    propio dato por empresa bancaria, en formato largo:
    fecha, banco, es_total, variable, valor

    Una fila se copia como variable solo si trae al menos un dato numerico
    en alguna columna de banco; de lo contrario se descarta (titulo de
    grupo como "SOLVENCIA"/"CALIDAD DE ACTIVOS", nota al pie, o fila en
    blanco).
    """
    fila_banco, columnas_banco, col_cuenta = _ubicar_bancos(df)
    nombres_banco = df.iloc[fila_banco]

    filas = []
    for r in range(fila_banco + 1, len(df)):
        nombre_var = df.iloc[r, col_cuenta]
        if not isinstance(nombre_var, str) or not nombre_var.strip():
            continue
        valores_fila = df.iloc[r, columnas_banco]
        if valores_fila.notna().sum() == 0:
            continue  # titulo de grupo, nota al pie, o similar: no es una variable real
        nombre_var = re.sub(r"\s+", " ", nombre_var.strip())
        for c in columnas_banco:
            banco = nombres_banco.iloc[c]
            banco = re.sub(r"\s+", " ", str(banco).strip())
            filas.append((fecha, banco, banco.lower().startswith("total"), nombre_var, df.iloc[r, c]))

    if not filas:
        raise ValueError("No se detecto ninguna variable con datos por banco en la hoja")

    resultado = pd.DataFrame(filas, columns=["fecha", "banco", "es_total", "variable", "valor"])
    resultado["valor"] = pd.to_numeric(resultado["valor"], errors="coerce")
    return resultado


def a_formato_ancho(largo: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el formato largo (una fila por banco+variable) al formato
    ancho: una fila por banco y una columna por variable. Se arma por
    POSICION (no por nombre de variable), por si algun nombre se repitiera
    dentro del mismo mes.
    """
    orden_bancos = list(dict.fromkeys(largo["banco"]))
    primer_banco = largo.loc[largo["banco"] == orden_bancos[0]]
    orden_variables = primer_banco["variable"].tolist()

    metas = []
    filas_valores = []
    for banco in orden_bancos:
        bloque = largo.loc[largo["banco"] == banco].reset_index(drop=True)
        if bloque["variable"].tolist() != orden_variables:
            raise ValueError(
                f"El banco '{banco}' no tiene las mismas variables, en el mismo orden, que '{orden_bancos[0]}'"
            )
        metas.append((bloque["fecha"].iloc[0], banco, bloque["es_total"].iloc[0]))
        filas_valores.append(bloque["valor"].tolist())

    meta = pd.DataFrame(metas, columns=["fecha", "banco", "es_total"])
    valores = pd.DataFrame(filas_valores, columns=orden_variables)
    ancho = pd.concat([meta, valores], axis=1)

    ancho["banco"] = pd.Categorical(ancho["banco"], categories=orden_bancos, ordered=True)
    ancho = ancho.sort_values(["fecha", "banco"], kind="stable").reset_index(drop=True)
    return ancho


def ultimo_dia_mes(year: int, month: int):
    """Fecha (date) del ultimo dia del mes."""
    return (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).date()


def guardar_excel(df: pd.DataFrame, destino: Path) -> None:
    """Guarda en .xlsx con la columna 'fecha' como fecha real formateada mmm-aa (ej. jul-26)."""
    df.to_excel(destino, index=False, sheet_name="Indicadores")

    from openpyxl import load_workbook

    wb = load_workbook(destino)
    ws = wb["Indicadores"]
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
    encabezado de variables justo antes de sus datos, uno debajo del otro
    (sin alinear variables de distintos meses por nombre, ya que el
    listado de indicadores puede cambiar de un mes a otro).
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Indicadores"

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
        df_crudo = _leer_hoja_indicadores(Path(args.local_file))
        indicadores = extraer_indicadores(df_crudo, ultimo_dia_mes(year, month))
        ancho = a_formato_ancho(indicadores)
        destino = salida / f"indicadores_{args.fecha}.xlsx"
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
        df = _leer_hoja_indicadores(path)
        volcado = salida / f"inspect_indicadores_{year}-{month:02d}.csv"
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
            df_crudo = _leer_hoja_indicadores(path)
            indicadores = extraer_indicadores(df_crudo, ultimo_dia_mes(year, month))
            ancho = a_formato_ancho(indicadores)
        except Exception as exc:
            print(f"  [ERROR] No se pudo parsear {etiqueta}: {exc}", file=sys.stderr)
            meses_omitidos.append((etiqueta, str(exc)))
            time.sleep(args.sleep)
            continue

        # Cada mes aporta sus propias variables, en su propio orden y
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

    destino = salida / "indicadores_bancos_2002_2026.xlsx"
    guardar_excel_por_bloques(piezas, destino)
    print(f"\nListo: {len(piezas)} mes(es), {total_filas} filas de datos guardadas en {destino}")


if __name__ == "__main__":
    main()
