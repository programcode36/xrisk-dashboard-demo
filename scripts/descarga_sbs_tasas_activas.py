#!/usr/bin/env python3
"""
Descarga el reporte "Tasa de Interés Promedio del Sistema Bancario"
(tasas activas por tipo de crédito y por empresa bancaria) desde el sitio
PUBLICO de la SBS:
https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B

A diferencia de los reportes B-2201/B-2401 (archivos .XLS fijos en la
intranet), esta es una pagina web dinamica tipo ASP.NET (Telerik
RadControls), y ademas esta protegida por Incapsula (firewall/anti-bots):
una peticion simple con `requests` (sin pasar antes por un navegador real)
recibe una pagina de bloqueo de ~800 bytes en vez del contenido real,
aunque el status HTTP sea 200. Verificado con una prueba real del
usuario.

Por eso el script funciona en DOS ETAPAS:
1. Abre un navegador real (Selenium + Chrome) para cargar la pagina. El
   navegador ejecuta el JavaScript del reto de Incapsula de forma normal
   y el sitio le entrega cookies validas de "no soy un bot". Se toman
   esas cookies (y el mismo User-Agent del navegador, Incapsula tambien
   valida que coincida) y se cierra el navegador: ya no hace falta para
   el resto.
2. Con esas cookies, se arma una sesion de `requests` normal y se replica
   el boton "Exportar" (id ctl00_cphContent_btnExportar), que en el HTML
   real dispara un POSTBACK COMPLETO de ASP.NET (no AJAX parcial): el
   servidor responde directamente con el archivo Excel. Se hace un GET
   fresco antes de cada POST para tomar un __VIEWSTATE/__EVENTVALIDATION
   validos (cambian en cada carga de pagina), y se sobreescriben solo los
   campos que nos interesan:
     ctl00$cphContent$rdpDate              -> fecha ISO (YYYY-MM-DD)
     ctl00$cphContent$rdpDate$dateInput    -> fecha visible (DD/MM/YYYY)
     ctl00$cphContent$hdTipoMoneda         -> "MN" o "ME"
     ctl00$cphContent$hdTipoEntidad        -> "B" (Bancos; coincide con
                                               el parametro ?tip=B de la URL)
     ctl00$cphContent$btnExportar          -> "Exportar"

Requisitos: Chrome instalado (Selenium >= 4.6 descarga el chromedriver
compatible automaticamente, no hace falta instalarlo a mano).

Uso:
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026 --moneda ME
    python descarga_sbs_tasas_activas.py --fecha 07/10/2026 --moneda ambas
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter, Retry

URL = "https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B"

PROCESSED_DIR = Path("data/processed")


def _resolver_reto_incapsula(url: str, espera_seg: float = 8.0) -> tuple[list[dict], str]:
    """
    Abre Chrome real con Selenium, carga la pagina, y espera a que el
    reto JS de Incapsula se resuelva (deja un par de segundos de margen
    despues de que el elemento esperado aparezca). Devuelve las cookies
    de la sesion y el User-Agent real del navegador, para reutilizarlos
    en una sesion de `requests` normal.
    """
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    opciones = webdriver.ChromeOptions()
    opciones.add_argument("--window-size=1280,900")
    # Sin --headless a proposito: algunos sistemas anti-bots distinguen
    # Chrome headless del normal y lo bloquean con mas facilidad.

    driver = webdriver.Chrome(options=opciones)
    try:
        driver.get(url)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "ctl00_cphContent_btnExportar"))
        )
        time.sleep(espera_seg)  # margen extra para que termine de asentar la cookie de Incapsula
        cookies = driver.get_cookies()
        user_agent = driver.execute_script("return navigator.userAgent")
        return cookies, user_agent
    finally:
        driver.quit()


def _session_con_cookies(cookies: list[dict], user_agent: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": user_agent, "Referer": URL})
    for c in cookies:
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"))
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
        raise ValueError(
            "No se encontro el formulario 'aspnetForm' en la pagina (probablemente "
            "Incapsula volvio a bloquear la peticion; revisa el archivo .html guardado)."
        )

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

    print("Abriendo Chrome para pasar la verificacion anti-bots de la SBS (Incapsula)...")
    try:
        cookies, user_agent = _resolver_reto_incapsula(URL)
    except Exception as exc:
        print(f"[ERROR] No se pudo abrir/usar Chrome con Selenium: {exc}", file=sys.stderr)
        print("Verifica que Chrome este instalado y que 'pip install selenium' se haya hecho bien.", file=sys.stderr)
        sys.exit(1)
    print(f"  Cookies obtenidas: {len(cookies)}. Continuando sin el navegador...")

    session = _session_con_cookies(cookies, user_agent)
    for moneda in monedas:
        print(f"Exportando tasas activas ({moneda}) al {args.fecha}...")
        try:
            resp = exportar_tasas(session, args.fecha, moneda)
        except requests.RequestException as exc:
            print(f"  [ERROR] Fallo la peticion: {exc}", file=sys.stderr)
            continue
        except ValueError as exc:
            print(f"  [ERROR] {exc}", file=sys.stderr)
            continue

        tipo_contenido = resp.headers.get("Content-Type", "")
        if "text/html" in tipo_contenido and len(resp.content) < 20000:
            # Probablemente no exporto el archivo, sino que devolvio la
            # pagina de nuevo (ej. fecha invalida, sin datos para ese dia,
            # o la cookie de Incapsula ya expiro). Se guarda igual para
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
