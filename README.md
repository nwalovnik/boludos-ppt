# Boludos del PPT — Macro Argentina

Dashboard de indicadores macroeconómicos argentinos con actualización diaria automática.

## Cómo funciona

```
[GitHub Action diario 08:30 ART] → [scraper/scrape.py] → [data.json] → [git commit + push] → [Netlify rebuild]
```

El `scraper/scrape.py` carga `scraper/historical.json` (bedrock histórico desde 2014)
y lo actualiza con datos frescos de:

- **INDEC** (`apis.datos.gob.ar`) — IPC, EMAE, IPI, ISAC, ICA, EPH, RIPTE, recaudación, base monetaria
- **BCRA** (`api.bcra.gob.ar/v4`) — Reservas, TC, BM diaria, variación IPC (más al día que el catálogo INDEC)
- **ArgentinaDatos** — EMBI riesgo país, cotizaciones blue/MEP/CCL
- **Bluelytics** — TC blue de respaldo

El HTML del tablero (`Tablero Macro 2.4.26.html`) hace `fetch("data.json")` al cargar.
El `_B64` baked dentro del HTML queda como fallback por si `data.json` no existe.

## Correr local

```bash
cd scraper
pip install -r requirements.txt
python scrape.py
```

Output: `data.json` en la raíz del repo.

## Layout

```
.
├── .github/workflows/daily-update.yml  # cron diario
├── scraper/
│   ├── scrape.py                       # scraper principal
│   ├── extract_historical.py           # one-shot: extrae _B64 → historical.json
│   ├── historical.json                 # bedrock histórico (3700+ registros desde 2014)
│   └── requirements.txt
├── data.json                           # output del scraper (regenerado diariamente)
├── index.html                          # portal
├── Tablero Macro 2.4.26.html           # tablero macro (lee data.json)
├── Tablero Despidos.html               # tablero despidos (separado)
├── _despidos_data.js
├── netlify.toml
└── README.md
```

## Deploy

Netlify conectado al repo. Cada push a `main` redespliega automáticamente.

Para forzar actualización fuera del cron: pestaña **Actions** del repo → workflow
"Update macro data daily" → **Run workflow**.
