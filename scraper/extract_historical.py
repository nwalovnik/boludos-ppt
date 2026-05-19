"""
Extrae el _B64 del HTML actual y lo guarda como historical.json.
Se corre UNA SOLA VEZ para preservar el dato histórico baked en el HTML original.
Despues de esto, el scraper diario carga historical.json + datos frescos de APIs.
"""
import base64
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HTML = ROOT / "Tablero Macro 2.4.26.html"
OUT = Path(__file__).resolve().parent / "historical.json"

def main():
    html = HTML.read_text(encoding="utf-8")
    m = re.search(r'const\s+_B64\s*=\s*"([A-Za-z0-9+/=]+)"', html)
    if not m:
        raise SystemExit("No se encontró _B64 en el HTML")
    b64 = m.group(1)
    raw = base64.b64decode(b64).decode("utf-8")
    data = json.loads(raw)

    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK - {len(data)} series guardadas en {OUT}")
    for k, v in data.items():
        if isinstance(v, list):
            print(f"  {k}: {len(v)} registros")

if __name__ == "__main__":
    main()
