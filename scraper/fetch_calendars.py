"""Calendarios oficiales de publicaciones macro: INDEC, BCRA, ARCA.

Genera scraper/calendar_oficial.json con eventos {fecha, fuente, label, periodo, serie_key}.
NO estima fechas: lee los calendarios publicados por los organismos.

Fuentes:
- INDEC: PDFs semestrales en /ftp/cuadros/publicaciones/calendario_{1,2}sem{YYYY}.pdf
- BCRA: HTML de https://www.bcra.gob.ar/calendario-de-informes/
- ARCA (recaudación): regla fija (primer día hábil del mes siguiente)
- Hacienda IMIG (resultado fiscal): la nota oficial INDEC del IMIG se publica unos
  días después del informe Hacienda. Usamos el INDEC IMIG como proxy (sale en el
  calendario INDEC como "Informe de Avance del Nivel de Actividad" cuando aplica).
"""
from __future__ import annotations
import json
import re
from datetime import datetime, timezone, timedelta, date
from pathlib import Path

import requests

try:
    import pypdf
except ImportError:
    pypdf = None

HERE = Path(__file__).parent
OUT = HERE / "calendar_oficial.json"

UA = {"User-Agent": "Mozilla/5.0"}

MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]
MESES_HEADER = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

# Mapeo nombre publicación INDEC → clave de serie del scraper
SERIE_MAP_INDEC = [
    # (regex, serie_key)
    (re.compile(r"\b(IPC)\b|precios al consumidor", re.I),                "ipc"),
    (re.compile(r"SIPM|sistema de .ndices de precios mayoristas", re.I),  "ipim"),
    (re.compile(r"EMAE|estimador mensual de actividad", re.I),            "emae"),
    (re.compile(r"IPI manufacturero|producci.n industrial manufacturer", re.I), "ipi"),
    (re.compile(r"actividad de la construcci.n", re.I),                   "isac"),
    (re.compile(r"UCII|capacidad instalada", re.I),                       "uci"),
    (re.compile(r"intercambio comercial argentino|^ICA|\bICA\b", re.I),   "bc"),
    (re.compile(r"encuesta de supermercados", re.I),                      "super"),
    (re.compile(r"autoservicios mayoristas", re.I),                       "mayor"),
    (re.compile(r"turismo internacional", re.I),                          "turismo"),
    (re.compile(r".ndice de salarios", re.I),                             "salarios"),
    (re.compile(r"mercado de trabajo|EPH.*trimestre|condiciones de vida", re.I), "eph"),
    (re.compile(r"informe de avance del nivel de actividad", re.I),       "pbi"),
    (re.compile(r"balanza de pagos", re.I),                               "bdp"),
]

def _map_serie(label: str) -> str | None:
    for rx, key in SERIE_MAP_INDEC:
        if rx.search(label):
            return key
    return None

def _periodo_indec(texto: str) -> str | None:
    """Extrae el período del label, ej 'Marzo de 2026' → '2026-03'."""
    m = re.search(r"\b(" + "|".join(MESES_ES) + r")\s+de\s+(\d{4})\b", texto, re.I)
    if m:
        mes_idx = MESES_ES.index(m.group(1).lower()) + 1
        return f"{m.group(2)}-{mes_idx:02d}"
    m = re.search(r"(primer|segundo|tercer|cuarto)\s+trimestre\s+de\s+(\d{4})", texto, re.I)
    if m:
        qmap = {"primer":1, "segundo":2, "tercer":3, "cuarto":4}
        return f"{qmap[m.group(1).lower()]}°T {m.group(2)}"
    m = re.search(r"a.o\s+(\d{4})", texto, re.I)
    if m:
        return m.group(1)
    return None

def _clean_text(t: str) -> str:
    """Limpia texto del PDF (espacios extra, encoding raro de ñ/í)."""
    # pypdf devuelve algunos chars mal; los reemplazamos
    t = t.replace("­", "")  # soft hyphen
    return " ".join(t.split())

def parse_indec_calendar(pdf_bytes: bytes, year: int) -> list[dict]:
    """Parsea un PDF semestral de INDEC. Devuelve lista de eventos."""
    if not pypdf:
        raise ImportError("pypdf no instalado")
    reader = pypdf.PdfReader_from_bytes(pdf_bytes) if hasattr(pypdf, "PdfReader_from_bytes") else None
    if reader is None:
        import io
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join(p.extract_text() for p in reader.pages)
    # Estructura: línea con mes solo + bloques de "DD DOW publicación"
    eventos = []
    # Encontrar bloques por mes
    # Algunas líneas tienen "4\n5\nJU\nVI\npub1\npub2" mezclado. Robusto: parsear linea por linea
    current_month = None
    current_day = None
    buffer_lines = []

    def flush(day, month, lines):
        if not day or not month or not lines:
            return
        try:
            d = date(year, month, day)
        except ValueError:
            return
        text_block = _clean_text(" ".join(lines))
        # Puede contener múltiples publicaciones (separadas por ". " inicio de mayúscula).
        # Simplificación: cada flush = 1 evento. Si vemos múltiples títulos, generamos varios.
        # Heurística: dividir por puntos seguido de espacio + Mayúscula NO precedida de número
        # Para evitar complicar: una sola entrada por (day, month) y todas las pubs concatenadas.
        # Mejor: segmentar por patrones conocidos.
        items = _segment_publications(text_block)
        for item in items:
            serie = _map_serie(item)
            per = _periodo_indec(item) or ""
            eventos.append({
                "fecha": d.isoformat(),
                "fuente": "INDEC",
                "label": item.strip().rstrip("."),
                "periodo": per,
                "serie_key": serie,
            })

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        # Header de mes?
        if line in MESES_HEADER:
            flush(current_day, current_month, buffer_lines)
            buffer_lines = []
            current_day = None
            current_month = MESES_HEADER.index(line) + 1
            continue
        # Patrón "DD DOW resto..." o "DD" solo (multilinea formato del PDF)
        m = re.match(r"^(\d{1,2})\s+(LU|MA|MI|JU|VI|SA|DO)\s+(.+)$", line)
        if m:
            flush(current_day, current_month, buffer_lines)
            buffer_lines = [m.group(3)]
            current_day = int(m.group(1))
            continue
        # Patrón "DD" solo (cuando viene en columna)
        m2 = re.match(r"^(\d{1,2})$", line)
        if m2 and current_month:
            flush(current_day, current_month, buffer_lines)
            buffer_lines = []
            current_day = int(m2.group(1))
            continue
        # Patrón "LU/MA/..." solo
        if line in ("LU","MA","MI","JU","VI","SA","DO"):
            continue
        # Línea de continuación: agregar al buffer
        if current_day:
            buffer_lines.append(line)

    flush(current_day, current_month, buffer_lines)
    return eventos

def _segment_publications(text: str) -> list[str]:
    """Divide un bloque de publicaciones múltiples en items individuales.

    INDEC pone varias publicaciones separadas; cada una termina en "{Mes} de {YYYY}"
    o "trimestre de {YYYY}" o "año {YYYY}". Usamos esos sufijos como delimitadores.
    """
    if not text.strip():
        return []
    # Patrón delimitador: termina en fecha
    pat = re.compile(
        r"(.+?(?:" + "|".join(MESES_ES) + r"|trimestre|a.o|semestre|expectativas\s+\w+-\w+)\s+(?:de\s+)?\d{4}\b\.?)",
        re.I
    )
    items = pat.findall(text)
    if items:
        return [i.strip() for i in items if len(i.strip()) > 10]
    # Fallback: 1 solo item
    return [text.strip()]

def fetch_indec_calendars(year: int) -> list[dict]:
    eventos = []
    for sem in (1, 2):
        url = f"https://www.indec.gob.ar/ftp/cuadros/publicaciones/calendario_{sem}sem{year}.pdf"
        try:
            r = requests.get(url, timeout=30, headers=UA)
            r.raise_for_status()
            eventos.extend(parse_indec_calendar(r.content, year))
            print(f"OK INDEC calendario_{sem}sem{year}.pdf: {len(eventos)} eventos acumulados")
        except Exception as e:
            print(f"WARN no pude bajar INDEC {sem}sem{year}: {e}")
    return eventos

_MES_ABREV = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
              "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}

def fetch_bcra_calendar() -> list[dict]:
    """Scrape la página HTML del calendario BCRA.

    Estructura: tabla con filas <td>{label}</td><td>{DD mes YYYY}</td>
    Ej: <td>Informe sobre Bancos</td><td>23 ene 2026</td>
    """
    eventos = []
    url = "https://www.bcra.gob.ar/calendario-de-informes/"
    try:
        r = requests.get(url, timeout=20, headers=UA)
        r.raise_for_status()
        html = r.text
        pat = re.compile(
            r"<td[^>]*>\s*([^<]{5,150}?)\s*</td>\s*<td[^>]*>\s*(\d{1,2})\s+(\w{3})\s+(\d{4})\s*</td>",
            re.I
        )
        for m in pat.finditer(html):
            label = _clean_text(m.group(1))
            day = int(m.group(2))
            mes_abrev = m.group(3).lower()[:3]
            year = int(m.group(4))
            if mes_abrev not in _MES_ABREV:
                continue
            try:
                d = date(year, _MES_ABREV[mes_abrev], day)
            except ValueError:
                continue
            serie_key = None
            if re.search(r"informe sobre bancos", label, re.I):
                serie_key = "mora"
            eventos.append({
                "fecha": d.isoformat(),
                "fuente": "BCRA",
                "label": label,
                "periodo": "",
                "serie_key": serie_key,
            })
        print(f"OK BCRA calendario: {len(eventos)} eventos")
    except Exception as e:
        print(f"WARN no pude scrapear BCRA: {e}")
    return eventos

def primer_dia_habil(y: int, m: int) -> date:
    d = date(y, m, 1)
    while d.weekday() >= 5:  # sábado=5, domingo=6
        d += timedelta(days=1)
    return d

def fetch_arca_calendar(hoy: date, meses_adelante: int = 3) -> list[dict]:
    """ARCA publica recaudación el primer día hábil del mes siguiente al de referencia."""
    eventos = []
    for offset in range(0, meses_adelante + 1):
        y, m = hoy.year, hoy.month + offset
        while m > 12: m -= 12; y += 1
        pub_y, pub_m = y, m + 1
        if pub_m > 12: pub_m = 1; pub_y += 1
        fecha_pub = primer_dia_habil(pub_y, pub_m)
        periodo_ref = f"{y:04d}-{m:02d}"
        eventos.append({
            "fecha": fecha_pub.isoformat(),
            "fuente": "ARCA",
            "label": "Recaudación tributaria nacional",
            "periodo": periodo_ref,
            "serie_key": "rec",
        })
    return eventos

def fetch_hacienda_fiscal_calendar(hoy: date, meses_adelante: int = 3) -> list[dict]:
    """Hacienda publica IMIG (resultado fiscal SPNF) ~día 20-22 del mes siguiente.

    No hay calendario oficial publicado online; usamos día 22 como referencia
    histórica (cuando llegue el dato real con Last-Modified, la fecha se corrige).
    """
    eventos = []
    for offset in range(0, meses_adelante + 1):
        y, m = hoy.year, hoy.month - 1 + offset
        while m <= 0: m += 12; y -= 1
        pub_y, pub_m = y, m + 1
        if pub_m > 12: pub_m = 1; pub_y += 1
        try:
            fecha_pub = date(pub_y, pub_m, 22)
        except ValueError:
            continue
        if fecha_pub < hoy:
            continue
        eventos.append({
            "fecha": fecha_pub.isoformat(),
            "fuente": "Hacienda",
            "label": "Resultado fiscal SPNF (IMIG)",
            "periodo": f"{y:04d}-{m:02d}",
            "serie_key": "fiscal",
        })
    return eventos

def main():
    hoy = datetime.now(timezone.utc).date()
    year = hoy.year
    eventos = []
    eventos += fetch_indec_calendars(year)
    if hoy.month >= 10:  # cerca de fin de año, también el siguiente
        eventos += fetch_indec_calendars(year + 1)
    eventos += fetch_bcra_calendar()
    eventos += fetch_arca_calendar(hoy)
    eventos += fetch_hacienda_fiscal_calendar(hoy)
    eventos.sort(key=lambda e: e["fecha"])
    OUT.write_text(json.dumps({"generated_at": hoy.isoformat(), "eventos": eventos},
                              ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"OK {OUT.name}: {len(eventos)} eventos totales")

if __name__ == "__main__":
    main()
