#!/usr/bin/env python3
"""
Descarga el reporte "Tasa de Interés Promedio del Sistema Bancario"
(tasas activas por tipo de crédito y por empresa bancaria) desde el sitio
PUBLICO de la SBS:
https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B

A diferencia de los reportes B-2201/B-2401 (archivos .XLS fijos en la
intranet), esta es una pagina web dinamica tipo ASP.NET (Telerik
RadControls). Verificado contra el codigo fuente real de la pagina (HTML
guardado por el usuario, fecha 07/10/2026):

- El boton "Exportar" (id ctl00_cphContent_btnExportar) dispara un
  POSTBACK COMPLETO (no AJAX parcial) que hace que el servidor responda
  directamente con el archivo Excel. Por eso se puede replicar con una
  simple peticion POST (requests), sin necesitar Selenium/navegador.
- El formulario trae los campos ocultos tipicos de ASP.NET (__VIEWSTATE,
  __VIEWSTATEGENERATOR, __EVENTVALIDATION, etc.) que cambian en cada
  sesion/carga de pagina, asi que SIEMPRE hay que hacer primero un GET
  para obtener un formulario fresco antes de cada POST.
- Campos relevantes para elegir fecha y moneda:
    ctl00$cphContent$rdpDate              -> fecha ISO (YYYY-MM-DD)
    ctl00$cphContent$rdpDate$dateInput    -> fecha visible (DD/MM/YYYY)
    ctl00$cphContent$hdTipoMoneda         -> "MN" o "ME"
    ctl00$cphContent$hdTipoEntidad        -> "B" (Bancos; coincide con
                                              el parametro ?tip=B de la URL)
    ctl00$cphContent$btnExportar          -> "Exportar" (nombre/valor del
                                              boton que se "presiona")
- IMPORTANTE: a diferencia de los otros scripts (B-2201/B-2401), este
  sitio es PUBLICO (www.sbs.gob.pe), no la intranet. Debería ser
  accesible desde cualquier conexion normal a internet, sin necesitar la
  VPN/red institucional.

Uso:
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026 --moneda ME
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026 --moneda ambas
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter, Retry

URL = "https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": URL,
}

PROCESSED_DIR = Path("data/processed")


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    retries = Retry(total=4, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    return s


def _campos_formulario(html: str) -> dict:
    """
    Extrae todos los campos <input> del formulario (hidden, text, submit)
    como un diccionario name -> value. Se usa para partir de un formulario
    "fresco" (con __VIEWSTATE/__EVENTVALIDATION validos de esa carga de
    pagina) y solo sobreescribir los campos que nos interesan, en vez de
    armar el payload a mano y arriesgarnos a que falte algun campo oculto
    que el servidor espera.
    """
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", attrs={"name": "aspnetForm"}) or soup.find("form")
    if form is None:
        raise ValueError("No se encontro el formulario 'aspnetForm' en la pagina")

    campos = {}
    for inp in form.find_all("input"):
        nombre = inp.get("name")
        if not nombre:
            continue
        tipo = (inp.get("type") or "text").lower()
        if tipo == "submit":
            continue  # solo se envia el boton que se "presiona", se agrega aparte
        campos[nombre] = inp.get("value", "")
    return campos


def exportar_tasas(session: requests.Session, fecha: str, moneda: str) -> requests.Response:
    """
    fecha: 'DD/MM/YYYY'
    moneda: 'MN' o 'ME'
    Devuelve la respuesta cruda del servidor (el archivo exportado).
    """
    dia, mes, anio = fecha.split("/")
    fecha_iso = f"{anio}-{mes}-{dia}"

    resp_get = session.get(URL, timeout=30)
    resp_get.raise_for_status()
    campos = _campos_formulario(resp_get.text)

    campos["ctl00$cphContent$rdpDate"] = fecha_iso
    campos["ctl00$cphContent$rdpDate$dateInput"] = fecha
    campos["ctl00$cphContent$hdTipoMoneda"] = moneda
    campos["ctl00$cphContent$hdTipoEntidad"] = "B"
    campos["ctl00$cphContent$btnExportar"] = "Exportar"

    resp_post = session.post(URL, data=campos, timeout=60)
    resp_post.raise_for_status()
    return resp_post


def _nombre_archivo(resp: requests.Response, fecha: str, moneda: str) -> str:
    """Intenta tomar el nombre de archivo del header Content-Disposition;
    si no viene, arma uno propio a partir de la fecha y moneda pedidas."""
    disposicion = resp.headers.get("Content-Disposition", "")
    m = re.search(r'filename="?([^";]+)"?', disposicion)
    if m:
        return m.group(1).strip()

    extension = ".xls"
    tipo_contenido = resp.headers.get("Content-Type", "")
    if "openxmlformats" in tipo_contenido:
        extension = ".xlsx"

    dia, mes, anio = fecha.split("/")
    return f"tasas_activas_bancos_{moneda}_{anio}-{mes}-{dia}{extension}"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fecha", required=True, metavar="DD/MM/YYYY", help="Fecha a consultar, ej. 07/10/2026")
    ap.add_argument(
        "--moneda",
        choices=["MN", "ME", "ambas"],
        default="MN",
        help="MN = Moneda Nacional, ME = Moneda Extranjera, ambas = descarga las dos (default MN)",
    )
    ap.add_argument("--outdir", default=str(PROCESSED_DIR), help="Carpeta de salida")
    args = ap.parse_args()

    if not re.match(r"^\d{2}/\d{2}/\d{4}$", args.fecha):
        print("La fecha debe tener el formato DD/MM/YYYY, ej. 07/10/2026", file=sys.stderr)
        sys.exit(1)

    salida = Path(args.outdir)
    salida.mkdir(parents=True, exist_ok=True)

    monedas = ["MN", "ME"] if args.moneda == "ambas" else [args.moneda]

    session = _session()
    for moneda in monedas:
        print(f"Exportando tasas activas ({moneda}) al {args.fecha}...")
        try:
            resp = exportar_tasas(session, args.fecha, moneda)
        except requests.RequestException as exc:
            print(f"  [ERROR] Fallo la peticion: {exc}", file=sys.stderr)
            continue

        tipo_contenido = resp.headers.get("Content-Type", "")
        if "text/html" in tipo_contenido and len(resp.content) < 20000:
            # Probablemente no exporto el archivo, sino que devolvio la
            # pagina de nuevo (ej. fecha invalida, sin datos para ese dia,
            # o cambio algun campo del formulario). Se guarda igual para
            # poder diagnosticar que paso.
            print(
                f"  [WARN] La respuesta parece ser HTML, no un archivo Excel "
                f"(Content-Type: {tipo_contenido}). Revisa el archivo guardado.",
                file=sys.stderr,
            )

        nombre = _nombre_archivo(resp, args.fecha, moneda)
        destino = salida / nombre
        destino.write_bytes(resp.content)
        print(f"  OK: {destino} ({len(resp.content)} bytes)")


if __name__ == "__main__":
    main()
