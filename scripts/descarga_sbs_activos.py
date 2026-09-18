#!/usr/bin/env python3
"""
Descarga los reportes B-2201 (Balance General) de la SBS, mes a mes, y
extrae unicamente el bloque de ACTIVO por empresa bancaria (pagina 1,
primeras ~50-60 filas, antes de que empiece PASIVO).

IMPORTANTE - leer antes de usar:
- El host `intranet2.sbs.gob.pe` no es accesible desde este entorno (esta
  bloqueado por la politica de red del sandbox y probablemente solo
  responde dentro de la red de la SBS / Peru). Este script NO fue probado
  contra un archivo real: se escribio con la mejor informacion disponible
  sobre el formato tipico de estos reportes (tabla HTML servida con
  extension .XLS, o binario .xls antiguo). Ejecuta primero el modo
  --inspect (ver mas abajo) para validar la deteccion de filas/columnas
  contra un archivo real y ajustar las constantes marcadas con TODO si
  hace falta.
- Los codigos de mes en el nombre de archivo (ej. "jl" = julio) pueden
  no ser consistentes en todo el historico 2016-2026. El script prueba
  varias variantes conocidas por mes y usa la primera que responda 200
  con contenido valido.

Uso:
    python descarga_sbs_activos.py                       # descarga todo el rango 2016-01 a 2026-07
    python descarga_sbs_activos.py --start 2020-01 --end 2021-12
    python descarga_sbs_activos.py --inspect 2026-07      # solo descarga y vuelca el crudo de un mes, para calibrar
"""

from __future__ import annotations

import argparse
import io
import sys
import time
from dataclasses import dataclass
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
    9: ("Septiembre", ["se", "st", "sp"]),
    10: ("Octubre", ["oc", "ot"]),
    11: ("Noviembre", ["nv", "no"]),
    12: ("Diciembre", ["dc", "di"]),
}

# Tope de filas a inspeccionar en la pagina 1 antes de recortar por PASIVO.
MAX_FILAS_ACTIVO = 60

# Palabras clave (en mayusculas, sin tildes) que marcan el fin del bloque
# de activos / inicio del de pasivos en la primera columna.
PALABRAS_CORTE_PASIVO = ("PASIVO",)
PALABRA_TOTAL_ACTIVO = "TOTAL ACTIVO"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

RAW_DIR = Path("data/raw/sbs_b2201")
PROCESSED_DIR = Path("data/processed")


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    retries = Retry(total=4, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    return s


def construir_urls_candidatas(year: int, month: int) -> list[str]:
    nombre_mes, abreviados = MESES[month]
    yy = f"{year % 100:02d}"
    return [
        f"{BASE_URL}/{year}/{nombre_mes}/{CODIGO_REPORTE}-{abv}{yy}.XLS"
        for abv in abreviados
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


def _leer_primera_tabla(path: Path) -> pd.DataFrame:
    """Lee la pagina 1 del reporte, sin importar si es HTML disfrazado de .xls o un binario real."""
    contenido = path.read_bytes()
    cabecera = contenido[:512].lstrip().lower()

    if cabecera.startswith(b"<html") or b"<table" in cabecera or cabecera.startswith(b"<!doctype"):
        tablas = pd.read_html(io.BytesIO(contenido), header=None)
        return tablas[0]

    # xls binario clasico (BIFF) o xlsx moderno
    try:
        return pd.read_excel(path, sheet_name=0, header=None, engine="xlrd")
    except Exception:
        return pd.read_excel(path, sheet_name=0, header=None)


@dataclass
class BloqueActivos:
    fecha: str
    encabezados: list[str]
    tabla: pd.DataFrame


def extraer_bloque_activos(df: pd.DataFrame, fecha: str) -> BloqueActivos:
    """
    Ubica, dentro de las primeras MAX_FILAS_ACTIVO filas, el sub-bloque de
    ACTIVO por empresa bancaria y lo recorta antes de que aparezca PASIVO.

    Supuesto (a validar con --inspect contra un archivo real): la primera
    columna trae el nombre de la cuenta contable y las columnas siguientes
    una por cada empresa bancaria, con 1-2 filas de encabezado con el
    nombre del banco.
    """
    col0 = df.iloc[:, 0].astype(str).str.upper().str.strip()
    col0_sin_tildes = (
        col0.str.normalize("NFKD").str.encode("ascii", "ignore").str.decode("ascii")
    )

    limite = min(MAX_FILAS_ACTIVO, len(df))
    ventana = col0_sin_tildes.iloc[:limite]

    fin = limite
    idx_total = ventana[ventana.str.contains(PALABRA_TOTAL_ACTIVO, na=False)]
    if not idx_total.empty:
        fin = idx_total.index[0] + 1  # incluir la fila de TOTAL ACTIVO
    else:
        for palabra in PALABRAS_CORTE_PASIVO:
            idx_pasivo = ventana[ventana.str.contains(palabra, na=False)]
            if not idx_pasivo.empty:
                fin = idx_pasivo.index[0]
                break

    # TODO: calibrar cuantas filas de encabezado hay realmente (aqui se
    # asume 1 fila de encabezado con el nombre de cada banco en la fila 0).
    fila_encabezado = 0
    encabezados = df.iloc[fila_encabezado].astype(str).str.strip().tolist()

    inicio_datos = fila_encabezado + 1
    tabla = df.iloc[inicio_datos:fin].reset_index(drop=True)
    tabla.columns = encabezados

    return BloqueActivos(fecha=fecha, encabezados=encabezados, tabla=tabla)


def a_formato_largo(bloque: BloqueActivos) -> pd.DataFrame:
    """Convierte el bloque ancho (cuenta x banco) a formato largo: fecha, cuenta, banco, valor."""
    tabla = bloque.tabla.copy()
    col_cuenta = tabla.columns[0]
    tabla = tabla.rename(columns={col_cuenta: "cuenta"})

    largo = tabla.melt(id_vars="cuenta", var_name="banco", value_name="valor")
    largo["valor"] = pd.to_numeric(largo["valor"], errors="coerce")
    largo["fecha"] = bloque.fecha
    largo = largo.dropna(subset=["valor"])
    return largo[["fecha", "banco", "cuenta", "valor"]]


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
    ap.add_argument("--start", default="2016-01", help="Mes inicial YYYY-MM (default 2016-01)")
    ap.add_argument("--end", default="2026-07", help="Mes final YYYY-MM (default 2026-07, cierre julio 2026)")
    ap.add_argument("--sleep", type=float, default=1.0, help="Segundos de espera entre descargas")
    ap.add_argument("--inspect", metavar="YYYY-MM", help="Descarga y vuelca el crudo de un solo mes para calibrar la logica de recorte")
    ap.add_argument("--outdir", default=str(PROCESSED_DIR), help="Carpeta de salida para los CSV procesados")
    args = ap.parse_args()

    session = _session()
    salida = Path(args.outdir)
    salida.mkdir(parents=True, exist_ok=True)

    if args.inspect:
        year, month = (int(x) for x in args.inspect.split("-"))
        path = descargar_mes(session, year, month, RAW_DIR)
        if not path:
            sys.exit(1)
        df = _leer_primera_tabla(path)
        volcado = salida / f"inspect_{year}-{month:02d}.csv"
        df.to_csv(volcado, index=False)
        print(f"Crudo volcado en {volcado} ({df.shape[0]} filas x {df.shape[1]} cols). Revisalo para ajustar las constantes del script.")
        return

    piezas = []
    for year, month in rango_meses(args.start, args.end):
        fecha = f"{year}-{month:02d}"
        print(f"Procesando {fecha}...")
        path = descargar_mes(session, year, month, RAW_DIR)
        if not path:
            continue
        try:
            df_crudo = _leer_primera_tabla(path)
            bloque = extraer_bloque_activos(df_crudo, fecha)
            piezas.append(a_formato_largo(bloque))
        except Exception as exc:
            print(f"  [ERROR] No se pudo parsear {fecha}: {exc}", file=sys.stderr)
        time.sleep(args.sleep)

    if not piezas:
        print("No se logro procesar ningun mes.", file=sys.stderr)
        sys.exit(1)

    consolidado = pd.concat(piezas, ignore_index=True)
    destino = salida / "activos_bancos_2016_2026.csv"
    consolidado.to_csv(destino, index=False)
    print(f"\nListo: {len(consolidado)} filas guardadas en {destino}")


if __name__ == "__main__":
    main()
