"""Wrapper atómico del pipeline: scrape → build_slides.

Cualquier cambio que afecte data.json debe correr a través de este script para
que el PPTX semanal quede sincronizado. El workflow daily-update.yml ya hace
esta secuencia; este wrapper es para corridas locales.

Uso:
    python scraper/update_all.py
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent


def run(name: str, *args: str) -> int:
    path = HERE / name
    print(f"\n=== {name} ===")
    return subprocess.run([sys.executable, str(path), *args]).returncode


def main() -> int:
    # 1) Calendarios oficiales (INDEC PDF + BCRA HTML + ARCA/Hacienda)
    run("fetch_calendars.py")
    # 2) Scraping de datos
    rc_scrape = run("scrape.py", *sys.argv[1:])
    if rc_scrape != 0:
        print(f"!! scrape.py terminó con código {rc_scrape}; sigo con build_slides igual.")
    # 3) PPTX semanal — puede salir 1 si no hay pubs de esta semana, no es fatal
    run("build_slides.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
