#!/usr/bin/env python3
"""
Descarga los reportes B-2201 (Balance General) de la SBS, mes a mes, y
extrae unicamente el bloque de ACTIVO por empresa bancaria (la parte de
la hoja "Balance General" antes de que empiece "Pasivo").

Formato real del archivo (verificado contra B-2201-jl2026.XLS, julio 2026):
- Es un .xlsx real (Excel 2007+, zip) aunque la extension diga .XLS. En
  archivos antiguos (verificar 2002-2015) puede venir como .xls binario
  clasico (BIFF) o incluso HTML disfrazado de .xls; el script detecta el
  formato por firma de bytes y usa el engine correcto en cada caso.
- El nombre de la hoja de Balance General varia con el tiempo: archivos
  recientes usan "1"; archivos antiguos usan "05-BG", "05-BG (P)" o "BG"
  (con variantes de espacios/parentesis). _elegir_hoja_balance() normaliza
  el nombre y elige la hoja correcta automaticamente.
- Layout ancho: cada empresa bancaria ocupa 3 columnas contiguas
  (Moneda Nacional, Moneda Extranjera, Total), con el nombre del banco
  en la fila justo encima de la sub-cabecera "MN"/"ME"/"TOTAL". Hay
  columnas separadoras en blanco entre bloques de bancos. La fila donde
  empieza este encabezado (y por lo tanto donde arranca el bloque de
  cuentas) varia mes a mes (fila 4 a 8 aprox.); _ubicar_encabezados() la
  busca dinamicamente en vez de asumir una fila fija.
- La columna 0 trae el nombre de la cuenta contable (DISPONIBLE, FONDOS
  INTERBANCARIOS, ..., TOTAL ACTIVO) y se repite igual al inicio de cada
  bloque de bancos (son solo para lectura visual al imprimir). Dentro de
  un mismo mes hay nombres que se repiten (ej. "Otros" y "Provisiones"
  aparecen dos veces cada uno en secciones distintas) — ver
  a_formato_ancho(). NO se asume que la cantidad ni el orden de cuentas
  sea igual entre meses distintos: cada mes se procesa con sus propias
  cuentas, tal como vienen en su archivo, sin compararlas contra otros
  meses ni descartar ninguna.
- Debajo de "TOTAL ACTIVO" sigue, en la misma hoja, el bloque de Pasivo.

IMPORTANTE sobre la descarga:
- El host `intranet2.sbs.gob.pe` no es accesible desde el entorno de este
  agente (bloqueado por politica de red del sandbox). El parseo de arriba
  SI fue validado con un archivo real que subio el usuario, pero la
  descarga automatica debe correrse desde una maquina con acceso real a
  esa intranet (red de la SBS / VPN). Por lo tanto, correr el rango
  2002-01 a 2026-07 completo (~295 meses) y revisar la salida es tarea
  del usuario, no de este agente.
- Los codigos de mes en el nombre de archivo (ej. "jl" = julio) pueden no
  ser consistentes en todo el historico 2002-2026; el script prueba varias
  variantes conocidas por mes, pero no esta garantizado que cubran todos
  los años (avisar si algun mes da 404 con todas las variantes).
- El consolidado final apila cada mes con sus propias cuentas (columnas),
  sin exigir que coincidan en nombre ni en cantidad con las de otro mes:
  donde un mes no tiene una cuenta que sí trae otro, esa celda queda vacia.

Uso:
    python descarga_sbs_activos.py                       # descarga todo el rango 2002-01 a 2026-07
    python descarga_sbs_activos.py --start 2020-01 --end 2021-12
    python descarga_sbs_activos.py --inspect 2026-07      # solo descarga y vuelca el crudo de un mes
    python descarga_sbs_activos.py --local-file ruta.xls --fecha 2026-07  # parsea un archivo ya descargado, sin red
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

# Nombre de carpeta (mes en español, tal como aparece en la URL) y
# variantes conocidas del abreviado de 2 letras usado en el nombre de
# archivo. Se intentan en orden hasta que una responda con un archivo valido.
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
    """
    El nombre real observado (B-2201-jl2026.XLS, julio 2026) usa el año
    completo de 4 digitos, no 2 como se asumio originalmente. Por las
    dudas de que meses mas antiguos del historico usen el formato de 2
    digitos, se prueban ambas variantes.
    """
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

# Nombres de hoja observados para el Balance General a lo largo del historico:
# archivos recientes usan "1"; archivos antiguos usan "05-BG", "05-BG (P)"
# o simplemente "BG" (con variantes de espacios/parentesis). Se compara
# normalizando a mayusculas y sin caracteres no alfanumericos.
def _elegir_hoja_balance(nombres_hoja: list[str]) -> str:
    normalizados = {n: re.sub(r"[^A-Z0-9]", "", n.upper()) for n in nombres_hoja}
    for nombre, norm in normalizados.items():
        if norm == "1":
            return nombre
    for nombre, norm in normalizados.items():
        if norm.startswith("05BG") or norm.startswith("BG"):
            return nombre
    # Ninguna variante conocida encontrada: usar la primera hoja y avisar.
    print(
        f"  [WARN] No se reconocio el nombre de hoja del Balance General entre {nombres_hoja}; "
        f"se usa la primera hoja ('{nombres_hoja[0]}'). Verificar manualmente.",
        file=sys.stderr,
    )
    return nombres_hoja[0]


def _leer_hoja_balance(path: Path) -> pd.DataFrame:
    """Ubica y lee la hoja de Balance General, sin importar si el archivo es
    xlsx real, xls binario clasico, o HTML disfrazado de .xls (variantes
    vistas historicamente en publicaciones de la SBS), ni como se llame la
    hoja ("1" en archivos recientes, "05-BG"/"05-BG (P)" en antiguos)."""
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
        engine = None  # dejar que pandas intente detectar el engine

    xl = pd.ExcelFile(io.BytesIO(contenido), engine=engine)
    hoja = _elegir_hoja_balance(xl.sheet_names)
    return xl.parse(hoja, header=None)


def _ubicar_encabezados(df: pd.DataFrame) -> tuple[int, int, int, int]:
    """
    Devuelve (fila_banco, fila_submoneda, fila_total_activo, col_cuenta).

    fila_submoneda es la fila con las etiquetas 'MN' / 'ME' / 'TOTAL' que
    aparecen 3 veces por banco. fila_banco es la fila inmediatamente
    anterior, con el nombre de cada banco. col_cuenta es la columna con el
    nombre de la cuenta contable (normalmente la A, pero en algunos meses
    esta corrida a la B u otra) y fila_total_activo es la fila donde esa
    columna dice 'TOTAL ACTIVO'. col_cuenta se detecta buscando, entre las
    columnas anteriores al primer bloque de banco, cual de ellas contiene
    el texto 'TOTAL ACTIVO'.
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
    fila_total_activo = None
    for c in range(primer_col_banco):
        columna = df.iloc[:, c].astype(str).map(_sin_tildes)
        coincidencias = columna[columna == "TOTAL ACTIVO"]
        if not coincidencias.empty:
            col_cuenta = c
            fila_total_activo = coincidencias.index[0]
            break
    if col_cuenta is None:
        raise ValueError(
            f"No se encontro la fila 'TOTAL ACTIVO' en ninguna de las columnas 0..{primer_col_banco - 1}"
        )

    return fila_banco, fila_submoneda, fila_total_activo, col_cuenta


def extraer_activos(df: pd.DataFrame, fecha: str) -> pd.DataFrame:
    """
    Extrae, de la hoja de Balance General, el bloque de Activo por empresa
    bancaria y lo devuelve en formato largo:
    fecha, banco, es_total, cuenta, moneda_nacional, moneda_extranjera, total
    """
    fila_banco, fila_submoneda, fila_total_activo, col_cuenta = _ubicar_encabezados(df)

    fila_datos_ini = fila_submoneda + 2  # hay una fila en blanco entre el encabezado y los datos
    fila_datos_fin = fila_total_activo + 1  # inclusive

    submoneda = df.iloc[fila_submoneda].astype(str).str.strip().str.upper()
    nombres_banco = df.iloc[fila_banco]

    cuentas_crudo = df.iloc[fila_datos_ini:fila_datos_fin, col_cuenta]
    # Las filas separadoras (subtitulos en blanco dentro del bloque de cuentas)
    # se identifican por el TEXTO de la columna de cuenta (que esta en blanco
    # para todos los bancos por igual), NO por los valores numericos de cada
    # banco: un banco puede reportar blanco/guion en una cuenta real que si
    # usa (no es separador), y filtrar por eso desalinearia la cantidad de
    # cuentas entre bancos del mismo mes.
    # OJO: no comparar via astype(str) == "nan": con pandas >= 2.x/3.x y el
    # dtype "str" nativo, un NaN convertido a texto ya no compara igual a la
    # cadena "nan" (aunque su repr se vea igual). Hay que chequear con
    # notna() sobre el valor original.
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

        sub = df.iloc[fila_datos_ini:fila_datos_fin, [c, c + 1, c + 2]].copy()
        sub.columns = ["moneda_nacional", "moneda_extranjera", "total"]
        sub = sub[filas_validas.values]
        sub.insert(0, "cuenta", cuentas.values)
        sub.insert(0, "es_total", banco.lower().startswith("total"))
        sub.insert(0, "banco", banco)
        sub.insert(0, "fecha", fecha)
        bloques.append(sub)

    if not bloques:
        raise ValueError("No se detecto ningun bloque de banco (MN/ME/TOTAL) en la hoja")

    resultado = pd.concat(bloques, ignore_index=True)
    for col in ("moneda_nacional", "moneda_extranjera", "total"):
        resultado[col] = pd.to_numeric(resultado[col], errors="coerce")
    return resultado


def a_formato_ancho(largo: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el formato largo (una fila por banco+cuenta) al formato ancho
    de trabajo: una fila por banco x moneda (MN/ME/TOTAL) y una columna por
    cuenta contable, replicando el criterio manual del usuario (transponer,
    quitar columnas/filas en blanco, dejar 3 filas -MN/ME/TOTAL- por banco).

    OJO: varias cuentas se repiten de nombre (ej. "Otros" aparece bajo
    DISPONIBLE y de nuevo bajo Creditos Vigentes; "Provisiones" aparece
    bajo Inversiones y de nuevo bajo Creditos). Por eso NO se puede pivotear
    agrupando por el texto de "cuenta" (un pivot_table fusionaria ambas
    filas en una sola columna y descartaria un valor). En su lugar se arma
    la tabla por posicion: cada banco aporta sus cuentas en el mismo orden
    de filas del Excel original, y esa lista (con nombres repetidos y
    todo) se usa tal cual como encabezado de columnas.
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
    """Fecha (date) del ultimo dia del mes, como en la planilla del usuario."""
    return (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).date()


def guardar_excel(df: pd.DataFrame, destino: Path) -> None:
    """Guarda en .xlsx con la columna 'fecha' como fecha real formateada mmm-aa (ej. jul-26)."""
    df.to_excel(destino, index=False, sheet_name="Activos")

    from openpyxl import load_workbook

    wb = load_workbook(destino)
    ws = wb["Activos"]
    col_fecha = df.columns.get_loc("fecha") + 1
    for fila in range(2, ws.max_row + 1):
        ws.cell(row=fila, column=col_fecha).number_format = "mmm-yy"
    ws.freeze_panes = "A2"
    ws.column_dimensions[ws.cell(row=1, column=col_fecha).column_letter].width = 10
    ws.column_dimensions[ws.cell(row=1, column=df.columns.get_loc("banco") + 1).column_letter].width = 32
    wb.save(destino)


def _dedup_columnas(columnas: list) -> list:
    """
    Vuelve unicas (temporalmente) las etiquetas repetidas de una lista de
    columnas, agregando un sufijo interno "__dupN" a partir de la 2da
    aparicion. Necesario porque pandas no permite concatenar DataFrames de
    distinta forma cuando alguno tiene columnas duplicadas (lanza
    "Reindexing only valid with uniquely valued Index objects").
    """
    contador: dict = {}
    resultado = []
    for c in columnas:
        contador[c] = contador.get(c, 0) + 1
        resultado.append(c if contador[c] == 1 else f"{c}__dup{contador[c]}")
    return resultado


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
        df_crudo = _leer_hoja_balance(Path(args.local_file))
        activos = extraer_activos(df_crudo, ultimo_dia_mes(year, month))
        ancho = a_formato_ancho(activos)
        destino = salida / f"activos_{args.fecha}.xlsx"
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
        df = _leer_hoja_balance(path)
        volcado = salida / f"inspect_{year}-{month:02d}.csv"
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
            df_crudo = _leer_hoja_balance(path)
            activos = extraer_activos(df_crudo, ultimo_dia_mes(year, month))
            ancho = a_formato_ancho(activos)
        except Exception as exc:
            print(f"  [ERROR] No se pudo parsear {etiqueta}: {exc}", file=sys.stderr)
            meses_omitidos.append((etiqueta, str(exc)))
            time.sleep(args.sleep)
            continue

        # Cada mes aporta sus propias cuentas, en su propio orden y cantidad,
        # sin exigir que coincidan con las de otros meses (el usuario
        # confirmo que el listado de cuentas puede cambiar en 24 años de
        # historico). Antes de apilar se renombran las columnas duplicadas
        # DENTRO de este mismo mes (ej. "Otros" x2) con un sufijo interno,
        # porque pd.concat no admite ejes con etiquetas repetidas cuando los
        # meses no tienen exactamente las mismas columnas; el sufijo se
        # quita de nuevo al final, sobre el consolidado ya armado.
        ancho.columns = _dedup_columnas(list(ancho.columns))
        piezas.append(ancho)
        time.sleep(args.sleep)

    if not piezas:
        print("No se logro procesar ningun mes.", file=sys.stderr)
        sys.exit(1)

    consolidado = pd.concat(piezas, ignore_index=True, sort=False)
    consolidado.columns = [re.sub(r"__dup\d+$", "", c) if isinstance(c, str) else c for c in consolidado.columns]

    if meses_omitidos:
        print(f"\n[RESUMEN] {len(meses_omitidos)} mes(es) NO quedaron en el consolidado:", file=sys.stderr)
        for etiqueta, motivo in meses_omitidos:
            print(f"   - {etiqueta}: {motivo}", file=sys.stderr)

    destino = salida / "activos_bancos_2002_2026.xlsx"
    guardar_excel(consolidado, destino)
    print(f"\nListo: {len(consolidado)} filas guardadas en {destino}")


if __name__ == "__main__":
    main()
