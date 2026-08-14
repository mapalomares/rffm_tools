# getCalendarioRFFM

## Español

**Propósito**: descarga el calendario de una competición de la RFFM (rffm.es) y lo exporta a Excel con las columnas `Jornada`, `Fecha`, `Local`, `Visitante`. Las filas del equipo indicado en la configuración se resaltan con color y negrita.

### Requisitos

```powershell
pip install requests openpyxl
```

### Uso

```powershell
# URL completa como parámetro; el script pregunta qué equipo resaltar
python getCalendarioRFFM.py "https://www.rffm.es/competicion/calendario?temporada=22&tipojuego=1&competicion=26737828&grupo=26737840"

# Usando la URL del .conf
python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
```

Al arrancar muestra los equipos del grupo numerados: introduce el número, parte del nombre, `0` para no resaltar ninguno, o Enter para aceptar el equipo por defecto (`*`, el del `.conf`).

Opcionales:

- `--url <URL>`: alternativa a la URL posicional.
- `--team "<NOMBRE EQUIPO>"`: fija el equipo a resaltar y omite la pregunta (coincidencia parcial).
- `--no-prompt`: no pregunta; usa el equipo del `.conf`.
- `--conf <ruta>`: fichero de configuración alternativo.

### Configuración (`getCalendarioRFFM.conf`)

| Sección | Clave | Descripción |
|---|---|---|
| `source` | `url` | URL del calendario (`temporada`, `tipojuego`, `competicion`, `grupo`) |
| `source` | `timeout` | Timeout de descarga en segundos |
| `output` | `folder` | Carpeta de salida (por defecto `output`) |
| `output` | `filename` | Nombre del Excel; admite `{competicion}`, `{grupo}`, `{temporada}` |
| `output` | `sheet_name` | Nombre de la hoja |
| `highlight` | `team` | Equipo resaltado por defecto en la pregunta interactiva |
| `highlight` | `color` | Color ARGB del relleno (ej. `FFFFF2CC`) |
| `logs` | `folder` / `level` | Carpeta y nivel de log |

### Salidas

- `output/calendario_<competicion>_<grupo>_<temporada>.xlsx`
- `logs/getCalendarioRFFM_<timestamp>.log`

### Notas

Los datos se leen del bloque JSON `__NEXT_DATA__` que la web incrusta en el HTML; no requiere navegador ni JavaScript.

---

## English

**Purpose**: downloads a RFFM (rffm.es) competition schedule and exports it to Excel with the columns `Jornada` (matchday), `Fecha` (date), `Local` (home), `Visitante` (away). Rows of the configured team are highlighted in colour and bold.

### Requirements

```powershell
pip install requests openpyxl
```

### Usage

```powershell
# Full URL as argument; the script asks which team to highlight
python getCalendarioRFFM.py "https://www.rffm.es/competicion/calendario?temporada=22&tipojuego=1&competicion=26737828&grupo=26737840"

# Using the URL from the .conf file
python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
```

It prints the numbered list of teams in the group: type the number, part of the name, `0` for none, or press Enter to accept the default team (`*`, taken from the `.conf`).

Optional: `--url <URL>` (alternative to the positional URL), `--team "<TEAM NAME>"` (skips the prompt, partial match), `--no-prompt` (uses the team from the `.conf`), `--conf <path>`.

### Configuration (`getCalendarioRFFM.conf`)

Same keys as described in the Spanish table above: source URL/timeout, output folder/filename/sheet, highlight team/colour and log folder/level.

### Outputs

- `output/calendario_<competition>_<group>_<season>.xlsx`
- `logs/getCalendarioRFFM_<timestamp>.log`

### Notes

Data is parsed from the `__NEXT_DATA__` JSON block embedded in the page HTML; no browser or JavaScript execution is required.
