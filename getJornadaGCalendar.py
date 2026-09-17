#!/usr/bin/env python3
"""
getJornadaGCalendar.py - v1.0

Autor: Miguel Angel Palomares
Fecha: 2026-09-16
Version: 1.0

Descripcion:
    Genera un CSV listo para cargar en Google Calendar con los partidos de la
    primera jornada pendiente (fecha >= hoy) de la competicion del equipo
    elegido. Los equipos configurados se leen del fichero .conf y se ofrecen
    por consola para que el usuario seleccione uno.

Parametros de entrada:
    - Ninguno obligatorio. Opcionalmente --conf <ruta> para usar otro fichero
      de configuracion (por defecto getJornadaGCalendar.conf).

Parametros de salida:
    - output/gcalendar_<equipo>_J<jornada>_<fecha>.csv con las columnas
      Asunto, Inicio, Duracion, Lugar, Color, Calendario.
    - logs/getJornadaGCalendar_<timestamp>.log

Uso: python getJornadaGCalendar.py
"""

import argparse
import configparser
import csv
import json
import logging
import os
import re
import sys
import unicodedata
from datetime import date, datetime

import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONF = os.path.join(SCRIPT_DIR, "getJornadaGCalendar.conf")
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) getJornadaGCalendar/1.0"
HEADERS = ["Asunto", "Inicio", "Duracion", "Lugar", "Color", "Calendario"]
TEAM_PREFIX = "equipo:"
REQUIRED_KEYS = ("temporada", "tipojuego", "competicion", "grupo")
MINOR_WORDS = {"de", "del", "y", "en"}
VOWELS = re.compile(r"[AEIOU]")

log = logging.getLogger("getJornadaGCalendar")


def load_config(conf_path):
    if not os.path.isfile(conf_path):
        print(f"[ERROR] Fichero de configuracion no encontrado: {conf_path}")
        sys.exit(1)
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(conf_path, encoding="utf-8")
    return cfg


def resolve_path(path):
    return path if os.path.isabs(path) else os.path.join(SCRIPT_DIR, path)


def setup_logging(cfg):
    folder = resolve_path(cfg.get("logs", "folder", fallback="logs"))
    os.makedirs(folder, exist_ok=True)
    level = getattr(logging, cfg.get("logs", "level", fallback="INFO").upper(), logging.INFO)
    logfile = os.path.join(folder, f"getJornadaGCalendar_{datetime.now():%Y%m%d_%H%M%S}.log")
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(logfile, encoding="utf-8"),
        ],
    )
    log.info(f"Log: {logfile}")


def load_teams(cfg):
    """Equipos configurados: seccion [equipo:<nombre>] con todas las claves obligatorias."""
    teams = []
    for section in cfg.sections():
        if not section.lower().startswith(TEAM_PREFIX):
            continue
        name = section[len(TEAM_PREFIX):].strip()
        data = {k: cfg.get(section, k, fallback="").strip() for k in REQUIRED_KEYS}
        if not name or not all(data.values()):
            log.warning(f"Equipo incompleto, se omite: [{section}]")
            continue
        data["nombre"] = name
        data["calendario"] = cfg.get(section, "calendario", fallback=name).strip()
        teams.append(data)
    return sorted(teams, key=lambda t: t["nombre"])


def choose_team(teams):
    print("\nEquipos configurados:")
    for i, team in enumerate(teams, start=1):
        print(f"  {i:2}. {team['nombre']}")

    if len(teams) == 1:
        print(f"\nSolo hay un equipo configurado: {teams[0]['nombre']}")
        return teams[0]

    while True:
        try:
            answer = input(f"\nEquipo [1-{len(teams)}]: ").strip()
        except EOFError:
            log.info("Entrada no interactiva: se usa el primer equipo configurado.")
            return teams[0]
        if answer.isdigit() and 1 <= int(answer) <= len(teams):
            return teams[int(answer) - 1]
        matches = [t for t in teams if answer and answer.upper() in t["nombre"].upper()]
        if len(matches) == 1:
            return matches[0]
        print("Opcion no valida.")


def build_url(cfg, team):
    base = cfg.get("general", "url_base", fallback="https://www.rffm.es/competicion/calendario")
    query = "&".join(f"{k}={team[k]}" for k in REQUIRED_KEYS)
    return f"{base}?{query}"


def fetch_calendar(url, timeout, verify=True):
    """Descarga la pagina y extrae el bloque JSON __NEXT_DATA__ con el calendario."""
    log.info(f"Descargando: {url}")
    resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout, verify=verify)
    resp.raise_for_status()

    match = re.search(
        r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>',
        resp.text,
        re.DOTALL,
    )
    if not match:
        log.error("No se ha encontrado el bloque __NEXT_DATA__ en la pagina.")
        sys.exit(2)

    data = json.loads(match.group(1))
    calendar = data.get("props", {}).get("pageProps", {}).get("calendar")
    if not calendar or not calendar.get("rounds"):
        log.error("La respuesta no contiene calendario (revisa los parametros del equipo).")
        sys.exit(2)
    return calendar


def parse_date(value):
    try:
        return datetime.strptime(str(value).strip(), "%d-%m-%Y").date()
    except ValueError:
        return None


def round_date(rnd):
    """Fecha de referencia de la jornada: la mas temprana de sus partidos."""
    dates = [d for d in (parse_date(m.get("fecha", "")) for m in rnd.get("equipos", [])) if d]
    return min(dates) if dates else None


def next_round(calendar, today=None):
    """Primera jornada con fecha igual o posterior a hoy."""
    today = today or date.today()
    candidates = [(round_date(r), r) for r in calendar.get("rounds", [])]
    pending = sorted((d, r) for d, r in candidates if d and d >= today)
    if not pending:
        log.error("No hay jornadas pendientes con fecha posterior a hoy.")
        sys.exit(3)
    return pending[0][1], pending[0][0]


def strip_accents(text):
    return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")


def half_minutes(cfg, competicion):
    """Minutos por parte segun la categoria que aparece en el nombre de la competicion."""
    name = strip_accents(competicion or "").lower()
    for key in cfg.options("duraciones"):
        if key == "defecto":
            continue
        if key in name:
            return cfg.getint("duraciones", key)
    log.warning(f"Categoria no reconocida en '{competicion}': se usa la duracion por defecto.")
    return cfg.getint("duraciones", "defecto", fallback=45)


def titlecase_team(name):
    """Capitaliza el nombre del equipo respetando siglas ('A.D.') y sufijos (\"'B'\")."""
    words = []
    for word in name.split():
        core = word.strip("'")
        if "." in word:
            words.append(word)
        elif words and word.lower() in MINOR_WORDS:
            words.append(word.lower())
        elif core.isupper() and (
            len(core) == 1 or re.fullmatch(r"[IVXLC]+", core) or not VOWELS.search(core)
        ):
            words.append(word)
        else:
            words.append(re.sub(r"\w+", lambda m: m.group(0).capitalize(), word))
    return " ".join(words)


def clean_place(campo):
    """Quita los sufijos tecnicos del nombre del campo, p.ej. '(HA)(HA)'."""
    return re.sub(r"\s*\(HA\)", "", campo or "", flags=re.IGNORECASE).strip()


def build_rows(cfg, calendar, rnd, team):
    total = 2 * half_minutes(cfg, calendar.get("competicion", "")) + cfg.getint(
        "general", "descanso", fallback=10
    )
    default_time = cfg.get("general", "hora_defecto", fallback="10:00").strip()
    color_default = cfg.get("general", "color_defecto", fallback="Grafito").strip()
    color_team = cfg.get("general", "color_equipo", fallback="Azulon").strip()
    key = team["nombre"].upper()

    rows = []
    sin_hora = []
    for match in rnd.get("equipos", []):
        local = match.get("equipo_local", "").strip()
        visitante = match.get("equipo_visitante", "").strip()
        mine = key in (local.upper(), visitante.upper())

        if mine:
            subject = f"{local.upper()} - {visitante.upper()}"
        else:
            subject = f"{titlecase_team(local)} - {titlecase_team(visitante)}"

        hora = (match.get("hora") or "").strip()
        if not hora:
            hora = default_time
            sin_hora.append(subject)
        fecha = parse_date(match.get("fecha", ""))
        inicio = f"{fecha:%d/%m/%Y} {hora}" if fecha else f"{match.get('fecha', '')} {hora}"

        rows.append(
            {
                "Asunto": subject,
                "Inicio": inicio,
                "Duracion": total,
                "Lugar": clean_place(match.get("campo", "")),
                "Color": color_team if mine else color_default,
                "Calendario": team["calendario"],
            }
        )
    log.info(f"Partidos de la jornada: {len(rows)} (duracion {total} min)")
    return rows, sin_hora


def confirm_without_time(sin_hora, default_time):
    """Avisa de los partidos sin hora publicada y pide confirmacion antes de guardar."""
    if not sin_hora:
        return True

    log.warning(f"{len(sin_hora)} partido(s) sin hora publicada en la RFFM:")
    for subject in sin_hora:
        log.warning(f"  - {subject}")
    print(f"\nSe les asignara la hora por defecto: {default_time}")

    try:
        answer = input("Generar el CSV igualmente? [s/N]: ").strip().lower()
    except EOFError:
        log.info("Entrada no interactiva: se genera el CSV con la hora por defecto.")
        return True
    return answer in ("s", "si", "sí", "y", "yes")


def slugify(text):
    text = strip_accents(str(text))
    text = re.sub(r"[^\w\s-]", "", text).strip()
    return re.sub(r"\s+", "_", text)


def write_csv(cfg, rows, team, rnd, fecha):
    folder = resolve_path(cfg.get("output", "folder", fallback="output"))
    os.makedirs(folder, exist_ok=True)
    filename = cfg.get("output", "filename", fallback="gcalendar_{equipo}_J{jornada}_{fecha}.csv").format(
        equipo=slugify(team["nombre"]),
        jornada=rnd.get("codjornada", ""),
        fecha=f"{fecha:%Y%m%d}" if fecha else "",
    )
    path = os.path.join(folder, filename)
    delimiter = cfg.get("output", "delimiter", fallback=",")

    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=HEADERS, delimiter=delimiter)
        writer.writeheader()
        writer.writerows(rows)
    log.info(f"CSV generado: {path}")
    return path


def parse_verify(value):
    """true/false o ruta a un bundle PEM corporativo."""
    raw = value.strip()
    if raw.lower() in ("true", "1", "yes"):
        return True
    if raw.lower() in ("false", "0", "no"):
        log.warning("Verificacion TLS desactivada (verify = false).")
        requests.packages.urllib3.disable_warnings()
        return False
    return resolve_path(raw)


def main():
    parser = argparse.ArgumentParser(description="CSV de la proxima jornada para Google Calendar")
    parser.add_argument("--conf", default=DEFAULT_CONF, help="Ruta del fichero .conf")
    args = parser.parse_args()

    cfg = load_config(args.conf)
    setup_logging(cfg)

    teams = load_teams(cfg)
    if not teams:
        log.error("No hay equipos configurados en el fichero .conf.")
        sys.exit(1)

    team = choose_team(teams)
    log.info(f"Equipo seleccionado: {team['nombre']}")

    verify = parse_verify(cfg.get("general", "verify", fallback="true"))

    calendar = fetch_calendar(
        build_url(cfg, team), cfg.getint("general", "timeout", fallback=30), verify
    )
    rnd, fecha = next_round(calendar)
    log.info(
        f"Competicion: {calendar.get('competicion')} - {calendar.get('grupo')} "
        f"({calendar.get('temporada')}) | Jornada {rnd.get('jornada')}"
    )

    rows, sin_hora = build_rows(cfg, calendar, rnd, team)
    if not confirm_without_time(sin_hora, cfg.get("general", "hora_defecto", fallback="10:00")):
        log.info("Operacion cancelada por el usuario: no se ha generado el CSV.")
        return
    write_csv(cfg, rows, team, rnd, fecha)


if __name__ == "__main__":
    main()
