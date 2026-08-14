# getCalendarioRFFM

## Español

**Propósito**: descarga el calendario de una competición de la RFFM (rffm.es) y lo exporta a Excel con las columnas `Jornada`, `Fecha`, `Local`, `Visitante`. Las filas del equipo indicado en la configuración se resaltan con color y negrita.

### Requisitos

```powershell
pip install requests openpyxl
```

### Uso

```powershell
python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
```

Opcionales:

- `--url <URL>`: sobrescribe la URL del `.conf`.
- `--team "<NOMBRE EQUIPO>"`: sobrescribe el equipo a resaltar (coincidencia parcial).

### Configuración (`getCalendarioRFFM.conf`)

| Sección | Clave | Descripción |
|---|---|---|
| `source` | `url` | URL del calendario (`temporada`, `tipojuego`, `competicion`, `grupo`) |
| `source` | `timeout` | Timeout de descarga en segundos |
| `output` | `folder` | Carpeta de salida (por defecto `output`) |
| `output` | `filename` | Nombre del Excel; admite `{competicion}`, `{grupo}`, `{temporada}` |
| `output` | `sheet_name` | Nombre de la hoja |
| `highlight` | `team` | Equipo a resaltar (coincidencia parcial, sin distinguir mayúsculas) |
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
python getCalendarioRFFM.py --conf getCalendarioRFFM.conf
```

Optional: `--url <URL>` overrides the configured URL, `--team "<TEAM NAME>"` overrides the highlighted team (partial match).

### Configuration (`getCalendarioRFFM.conf`)

Same keys as described in the Spanish table above: source URL/timeout, output folder/filename/sheet, highlight team/colour and log folder/level.

### Outputs

- `output/calendario_<competition>_<group>_<season>.xlsx`
- `logs/getCalendarioRFFM_<timestamp>.log`

### Notes

Data is parsed from the `__NEXT_DATA__` JSON block embedded in the page HTML; no browser or JavaScript execution is required.
