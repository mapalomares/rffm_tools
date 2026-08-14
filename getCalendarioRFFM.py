#!/usr/bin/env python3
"""
getCalendarioRFFM.py - v1.0

Descarga el calendario de una competición de rffm.es y lo exporta a Excel
con las columnas: Jornada, Fecha, Local, Visitante.
Las filas del equipo configurado en [highlight] se marcan con color.

Uso: python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
"""

import argparse
import configparser
import json
import logging
import os
import re
import sys
from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONF = os.path.join(SCRIPT_DIR, "getCalendarioRFFM.conf")
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) getCalendarioRFFM/1.0"

log = logging.getLogger("getCalendarioRFFM")


def load_config(conf_path):
    if not os.path.isfile(conf_path):
        print(f"[ERROR] Fichero de configuración no encontrado: {conf_path}")
        sys.exit(1)
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(conf_path, encoding="utf-8")
    return cfg


def setup_logging(cfg):
    folder = resolve_path(cfg.get("logs", "folder", fallback="logs"))
    os.makedirs(folder, exist_ok=True)
    level = getattr(logging, cfg.get("logs", "level", fallback="INFO").upper(), logging.INFO)
    logfile = os.path.join(folder, f"getCalendarioRFFM_{datetime.now():%Y%m%d_%H%M%S}.log")
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(logfile, encoding="utf-8"),
        ],
    )
    log.info(f"Log: {logfile}")


def resolve_path(path):
    return path if os.path.isabs(path) else os.path.join(SCRIPT_DIR, path)


def fetch_calendar(url, timeout):
    """Descarga la página y extrae el bloque JSON __NEXT_DATA__ con el calendario."""
    log.info(f"Descargando: {url}")
    resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
    resp.raise_for_status()

    match = re.search(
        r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>',
        resp.text,
        re.DOTALL,
    )
    if not match:
        log.error("No se ha encontrado el bloque __NEXT_DATA__ en la página.")
        sys.exit(2)

    data = json.loads(match.group(1))
    calendar = data.get("props", {}).get("pageProps", {}).get("calendar")
    if not calendar or not calendar.get("rounds"):
        log.error("La respuesta no contiene calendario (revisa los parámetros de la URL).")
        sys.exit(2)
    return calendar


def extract_rows(calendar):
    """Devuelve una lista de dicts con jornada, fecha, local y visitante."""
    rows = []
    for rnd in calendar.get("rounds", []):
        jornada = rnd.get("codjornada") or rnd.get("jornada", "")
        for match in rnd.get("equipos", []):
            rows.append(
                {
                    "jornada": int(jornada) if str(jornada).isdigit() else jornada,
                    "fecha": parse_date(match.get("fecha", "")),
                    "local": match.get("equipo_local", "").strip(),
                    "visitante": match.get("equipo_visitante", "").strip(),
                }
            )
    log.info(f"Partidos extraídos: {len(rows)} en {len(calendar.get('rounds', []))} jornadas")
    return rows


def parse_date(value):
    """Convierte 'dd-mm-yyyy' en date; si no es válida, devuelve el texto original."""
    try:
        return datetime.strptime(value.strip(), "%d-%m-%Y").date()
    except (ValueError, AttributeError):
        return value


def slugify(text):
    text = re.sub(r"[^\w\s-]", "", str(text), flags=re.UNICODE).strip()
    return re.sub(r"[\s]+", "_", text)


def write_excel(rows, calendar, cfg):
    folder = resolve_path(cfg.get("output", "folder", fallback="output"))
    os.makedirs(folder, exist_ok=True)
    filename = cfg.get("output", "filename").format(
        competicion=slugify(calendar.get("competicion", "competicion")),
        grupo=slugify(calendar.get("grupo", "grupo")),
        temporada=slugify(calendar.get("temporada", "")),
    )
    path = os.path.join(folder, filename)

    highlight_team = cfg.get("highlight", "team", fallback="").strip().upper()
    fill = PatternFill("solid", fgColor=cfg.get("highlight", "color", fallback="FFFFF2CC"))

    wb = Workbook()
    ws = wb.active
    ws.title = cfg.get("output", "sheet_name", fallback="Calendario")

    headers = ["Jornada", "Fecha", "Local", "Visitante"]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="FF305496")
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    round_border = Border(bottom=Side(style="double", color="FF000000"))

    marked = 0
    for idx, row in enumerate(rows):
        ws.append([row["jornada"], row["fecha"], row["local"], row["visitante"]])
        r = ws.max_row
        ws.cell(row=r, column=2).number_format = "DD/MM/YYYY"
        if highlight_team and (
            highlight_team in row["local"].upper() or highlight_team in row["visitante"].upper()
        ):
            marked += 1
            for col in range(1, len(headers) + 1):
                ws.cell(row=r, column=col).fill = fill
                ws.cell(row=r, column=col).font = Font(bold=True)

        is_last_of_round = idx + 1 == len(rows) or rows[idx + 1]["jornada"] != row["jornada"]
        if is_last_of_round:
            for col in range(1, len(headers) + 1):
                ws.cell(row=r, column=col).border = round_border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{ws.max_row}"
    for col, width in enumerate([10, 14, 42, 42], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    wb.save(path)
    log.info(f"Filas resaltadas para '{cfg.get('highlight', 'team')}': {marked}")
    log.info(f"Excel generado: {path}")
    return path


def main():
    parser = argparse.ArgumentParser(description="Exporta el calendario RFFM a Excel.")
    parser.add_argument("--conf", default=DEFAULT_CONF, help="Ruta del fichero .conf")
    parser.add_argument("--url", help="Sobrescribe la URL del .conf")
    parser.add_argument("--team", help="Sobrescribe el equipo a resaltar")
    args = parser.parse_args()

    cfg = load_config(args.conf)
    setup_logging(cfg)

    if args.url:
        cfg.set("source", "url", args.url)
    if args.team:
        cfg.set("highlight", "team", args.team)

    calendar = fetch_calendar(
        cfg.get("source", "url"), cfg.getint("source", "timeout", fallback=30)
    )
    log.info(
        f"Competición: {calendar.get('competicion')} | {calendar.get('grupo')} "
        f"| Temporada {calendar.get('temporada')}"
    )
    rows = extract_rows(calendar)
    write_excel(rows, calendar, cfg)


if __name__ == "__main__":
    try:
        main()
    except requests.RequestException as exc:
        log.error(f"Error de red: {exc}")
        sys.exit(3)
