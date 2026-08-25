# getCalendarioRFFM

## Español

**Propósito**: descarga el calendario de una competición de la RFFM (rffm.es) y lo exporta a Excel y PDF con las columnas `Jornada`, `Fecha`, `Local`, `Visitante`. Las filas del equipo elegido se resaltan con color y negrita, y cada jornada se separa con una línea doble.

### Requisitos

```powershell
pip install -r requirements.txt
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
- `--no-pdf`: genera solo el Excel.
- `--conf <ruta>`: fichero de configuración alternativo.

### Configuración (`getCalendarioRFFM.conf`)

| Sección | Clave | Descripción |
|---|---|---|
| `source` | `url` | URL del calendario (`temporada`, `tipojuego`, `competicion`, `grupo`) |
| `source` | `timeout` | Timeout de descarga en segundos |
| `output` | `folder` | Carpeta de salida (por defecto `output`) |
| `output` | `filename` | Nombre del Excel; admite `{competicion}`, `{grupo}`, `{temporada}` |
| `output` | `pdf` / `pdf_filename` | Generar PDF (true/false) y nombre del fichero |
| `output` | `sheet_name` | Nombre de la hoja |
| `highlight` | `team` | Equipo resaltado por defecto en la pregunta interactiva |
| `highlight` | `color` | Color ARGB del relleno (ej. `FFFFF2CC`) |
| `watermark` | `enabled` | Activa la marca de agua del PDF |
| `watermark` | `team_match` | Solo se aplica si el equipo elegido contiene este texto |
| `watermark` | `image` | Ruta de la imagen (ej. `images/Santa_Barbara_logo.png`) |
| `watermark` | `opacity` | Opacidad de la marca de agua (0-1) |
| `logs` | `folder` / `level` | Carpeta y nivel de log |

### Salidas

- `output/calendario_<competicion>_<grupo>_<temporada>.xlsx`
- `output/calendario_<competicion>_<grupo>_<temporada>.pdf`
- `logs/getCalendarioRFFM_<timestamp>.log`

### Notas

Los datos se leen del bloque JSON `__NEXT_DATA__` que la web incrusta en el HTML; no requiere navegador ni JavaScript.

La marca de agua solo se aplica al PDF: se escala al ancho de la tabla y se repite en vertical hasta donde termina la tabla (Excel no soporta marcas de agua reales).

---

## English

**Purpose**: downloads a RFFM (rffm.es) competition schedule and exports it to Excel and PDF with the columns `Jornada` (matchday), `Fecha` (date), `Local` (home), `Visitante` (away). Rows of the chosen team are highlighted in colour and bold, and each matchday is separated by a double line.

### Requirements

```powershell
pip install -r requirements.txt
```

### Usage

```powershell
# Full URL as argument; the script asks which team to highlight
python getCalendarioRFFM.py "https://www.rffm.es/competicion/calendario?temporada=22&tipojuego=1&competicion=26737828&grupo=26737840"

# Using the URL from the .conf file
python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
```

It prints the numbered list of teams in the group: type the number, part of the name, `0` for none, or press Enter to accept the default team (`*`, taken from the `.conf`).

Optional: `--url <URL>` (alternative to the positional URL), `--team "<TEAM NAME>"` (skips the prompt, partial match), `--no-prompt` (uses the team from the `.conf`), `--no-pdf` (Excel only), `--conf <path>`.

### Configuration (`getCalendarioRFFM.conf`)

Same keys as described in the Spanish table above: source URL/timeout, output folder/filename/sheet, highlight team/colour and log folder/level.

### Outputs

- `output/calendario_<competition>_<group>_<season>.xlsx`
- `output/calendario_<competition>_<group>_<season>.pdf`
- `logs/getCalendarioRFFM_<timestamp>.log`

### Notes

Data is parsed from the `__NEXT_DATA__` JSON block embedded in the page HTML; no browser or JavaScript execution is required.

The watermark applies to the PDF only: it is scaled to the table width and tiled vertically until the table ends (Excel has no real watermark support).

---

# getEquiposRFFM

## Español

**Propósito**: dada la URL de la clasificación de un grupo de la RFFM (rffm.es), recorre todos los equipos (siguiendo el enlace a la ficha de cada uno) y recoge `Equipo`, `Localidad`, `Terreno de juego` y `Equipación` (camiseta/pantalón/medias). El resultado se exporta a Excel.

### Requisitos

```powershell
pip install -r requirements.txt
```

### Uso

```powershell
# URL completa como parámetro (el parámetro 'jornada' es irrelevante y puede omitirse)
python getEquiposRFFM.py "https://www.rffm.es/competicion/clasificaciones?temporada=22&tipojuego=1&competicion=26737828&grupo=26737840"

# Usando la URL del .conf
python getEquiposRFFM.py --conf getEquiposRFFM.conf
```

Opcionales:

- `--url <URL>`: alternativa a la URL posicional.
- `--conf <ruta>`: fichero de configuración alternativo.

### Configuración (`getEquiposRFFM.conf`)

| Sección | Clave | Descripción |
|---|---|---|
| `source` | `url` | URL de la clasificación (`temporada`, `tipojuego`, `competicion`, `grupo`) |
| `source` | `timeout` | Timeout de descarga en segundos |
| `source` | `delay` | Pausa entre la descarga de cada ficha de equipo |
| `output` | `folder` | Carpeta de salida (por defecto `output`) |
| `output` | `filename` | Nombre del Excel; admite `{competicion}`, `{grupo}`, `{temporada}` |
| `output` | `sheet_name` | Nombre de la hoja |
| `logs` | `folder` / `level` | Carpeta y nivel de log |

### Salidas

- `output/equipos_<competicion>_<grupo>_<temporada>.xlsx`
- `logs/getEquiposRFFM_<timestamp>.log`

### Notas

Los datos se leen del bloque JSON `__NEXT_DATA__` de la página de clasificación y de la ficha de cada equipo (`https://www.rffm.es/fichaequipo/<codigo>`); no requiere navegador ni JavaScript.

El campo `Terreno de juego` se construye a partir del campo `campo` de la ficha (formato `LOCALIDAD - NOMBRE (ANOTACIÓN)`), mostrando `Nombre (Localidad)`. La `Equipación` usa siempre la primera equipación (principal) tal como aparece en la ficha del equipo.

---

## English

**Purpose**: given the URL of a RFFM (rffm.es) group standings page, iterates through every team (following the link to each team's detail page) and collects `Equipo` (team), `Localidad` (town), `Terreno de juego` (home ground) and `Equipación` (kit: shirt/shorts/socks). The result is exported to Excel.

### Requirements

```powershell
pip install -r requirements.txt
```

### Usage

```powershell
# Full URL as argument (the 'jornada' parameter is irrelevant and can be omitted)
python getEquiposRFFM.py "https://www.rffm.es/competicion/clasificaciones?temporada=22&tipojuego=1&competicion=26737828&grupo=26737840"

# Using the URL from the .conf file
python getEquiposRFFM.py --conf getEquiposRFFM.conf
```

Optional: `--url <URL>` (alternative to the positional URL), `--conf <path>`.

### Configuration (`getEquiposRFFM.conf`)

Same keys as described in the Spanish table above: source URL/timeout/delay, output folder/filename/sheet, and log folder/level.

### Outputs

- `output/equipos_<competition>_<group>_<season>.xlsx`
- `logs/getEquiposRFFM_<timestamp>.log`

### Notes

Data is parsed from the `__NEXT_DATA__` JSON block on the standings page and on each team's detail page (`https://www.rffm.es/fichaequipo/<id>`); no browser or JavaScript execution is required.

The `Terreno de juego` field is built from the ground's `campo` value (format `TOWN - NAME (NOTE)`), rendered as `Name (Town)`. `Equipación` always uses the first (main) kit as listed on the team's page.
