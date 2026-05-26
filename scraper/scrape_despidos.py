"""
Scraper diario de despidos/cierres argentinos via Google News RSS.

Carga bedrock historico (historical_despidos.json, 867+ eventos desde dic-2023)
y suma noticias frescas. Output: despidos.json en raiz del repo.

Mejoras vs version legacy:
  - Baja el cuerpo del articulo (no solo el titulo RSS) para extraer empleados/fecha.
  - Blacklist agresiva: politicos, conceptos, ubicaciones, medios.
  - ALIAS_EMPRESAS expandido (~150 entradas argentinas).
  - Extraccion de empresa prioriza mayusculas sostenidas / nombres entre comillas.
  - Fecha real del articulo via <meta article:published_time> > <time> > pubDate RSS.
  - Dedupe con canonicalizacion ANTES de comparar.

Uso:
    python scrape_despidos.py            # corre todo, guarda despidos.json
    python scrape_despidos.py --dry-run  # corre pero no guarda
    python scrape_despidos.py --max-fetch 80  # cap de articulos a bajar (default 60)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

import requests

try:
    from lxml import html as lxml_html
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

# ──────────────────────────────────────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────────────────────────────────────
SCRAPER_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRAPER_DIR.parent
HISTORICAL = SCRAPER_DIR / "historical_despidos.json"
OUT_PATH = REPO_ROOT / "despidos.json"

# ──────────────────────────────────────────────────────────────────────────────
# Parametros
# ──────────────────────────────────────────────────────────────────────────────
MAX_EMPLEADOS_AUTO = 10000   # cap para evitar articulos resumen ("100 mil despidos en 2024")
FECHA_MINIMA = date(2024, 1, 1)
MAX_ARTICULOS_FETCH = 150    # cap de articulos a bajar por corrida (rate limiting)
TIMEOUT_RSS = 15
TIMEOUT_ARTICULO = 10
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/121.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-AR,es;q=0.9,en;q=0.5",
}

# ──────────────────────────────────────────────────────────────────────────────
# Blacklist: no son nombres de empresa
# ──────────────────────────────────────────────────────────────────────────────
# Politicos / funcionarios frecuentes en titulares
BLACKLIST_POLITICOS = {
    "milei", "javier milei", "karina milei", "caputo", "luis caputo", "santiago caputo",
    "sturzenegger", "federico sturzenegger", "patricia bullrich", "bullrich",
    "petri", "luis petri", "francos", "guillermo francos", "adorni", "manuel adorni",
    "pettovello", "sandra pettovello", "cuneo libarona", "cunero libarona",
    "ferraro", "ferraro maximiliano", "menem", "martin menem", "lemoine",
    "macri", "mauricio macri", "kicillof", "axel kicillof", "magario", "veronica magario",
    "massa", "sergio massa", "cristina kirchner", "cfk", "nestor kirchner",
    "scioli", "alberto fernandez", "fernandez", "kulfas", "moroni", "claudio moroni",
    "carlos castagneto", "castagneto", "ricardo quintela", "kicillof", "vidal",
    "rodriguez larreta", "larreta", "jorge macri", "santilli", "diego santilli",
    "lopez murphy", "espert", "jose luis espert", "bregman", "myriam bregman",
    "grabois", "juan grabois", "del caño", "nicolas del caño",
    "wado de pedro", "de pedro", "rossi", "agustin rossi", "manzur",
    "alfonsin", "raul alfonsin", "menem", "carlos menem",
    "trump", "donald trump", "lula", "boric", "milei karina",
    # provinciales/gobernadores comunes
    "morales", "gerardo morales", "uñac", "jaldo", "torres", "ignacio torres",
    "weretilneck", "frigerio", "saenz", "gustavo saenz",
    "ziliotto", "sergio ziliotto", "valdes", "gustavo valdes",
    "passalacqua", "herrera ahuad", "perotti", "omar perotti",
    "schiaretti", "juan schiaretti", "llaryora", "martin llaryora",
    "jorge capitanich", "capitanich", "zdero",
}

# Conceptos / sustantivos comunes que se confunden con empresa
BLACKLIST_CONCEPTOS = {
    "reforma", "reforma laboral", "ley bases", "ley", "decreto", "dnu",
    "justicia", "gobierno", "gobierno nacional", "estado", "estado nacional",
    "nacion", "argentina", "republica argentina", "casa rosada",
    "congreso", "diputados", "senado", "camara", "camara de diputados",
    "ejecutivo", "poder ejecutivo", "judicial", "poder judicial",
    "el", "la", "los", "las", "un", "una", "del", "este", "esta", "esto",
    "otro", "otra", "otros", "otras", "varios", "varias", "muchos", "muchas",
    "nuevo", "nueva", "nuevos", "nuevas", "primer", "segundo", "tercero",
    "denuncian", "advierten", "anuncian", "confirman", "rechazan", "alertan",
    "echaron", "echo", "echan", "despidieron", "suspendieron", "cesantearon",
    "cerraron", "quebraron", "anunciaron", "recortaron",
    "despido", "despidos", "cierre", "cierres", "echo personal", "echó personal",
    "quiebra", "quiebras", "suspension", "suspensión", "suspensiones",
    "desvinculacion", "desvinculaciones", "cesanteo", "cesanteos",
    "fuero", "fuero del trabajo", "industricidio", "industricidios",
    "trabajadores", "empleados", "empresa", "empresas", "industria", "industrias",
    "planta", "fabrica", "establecimiento", "comercio", "negocio",
    "rubro", "sector", "actividad", "crisis", "ola", "oleada", "ronda",
    "panorama", "contexto", "informe", "balance", "resumen", "ranking",
    "alerta", "tras", "pese", "ante", "frente",
    "quebro", "cierra", "cerro", "despide", "suspende",
    "economista", "consultora", "indec", "fmi", "tesoro",
    "policia", "gendarmeria", "ejercito", "fuerzas armadas", "fuerza",
    "iglesia", "vaticano", "papa", "francisco",
    "argentina hoy", "argentina sin", "hoy", "ayer", "mañana", "manana",
    # frases sueltas que aparecen
    "sin identificar", "desconocida", "varias empresas", "multiples empresas",
    "diversas firmas", "diferentes companias",
    # palabras frecuentes en titulares que no son empresas
    "canal", "pagina", "página", "historica", "histórica", "historico", "histórico",
    "caso", "momento", "futuro", "presente", "pasado",
    "capital humano", "ministerio capital humano",  # ministerio, no empresa
    "luz", "luz azul", "punto", "linea", "línea",
    "industricidio", "industricidios",
    "fuero", "fuero del trabajo", "trabajo",
}

# Provincias y ciudades que NO son nombre de empresa cuando aparecen solas
BLACKLIST_UBICACIONES = {
    "buenos aires", "ciudad de buenos aires", "caba", "gran buenos aires",
    "cordoba", "santa fe", "rosario", "mendoza", "tucuman", "salta",
    "entre rios", "chaco", "corrientes", "misiones", "jujuy", "san juan",
    "rio negro", "neuquen", "formosa", "chubut", "la pampa",
    "santiago del estero", "san luis", "catamarca", "la rioja",
    "santa cruz", "tierra del fuego", "patagonia", "cuyo", "noa", "nea",
    "la plata", "mar del plata", "bahia blanca", "tandil", "junin",
    "pilar", "tigre", "san isidro", "vicente lopez", "san martin",
    "moron", "lomas de zamora", "quilmes", "berazategui", "florencio varela",
    "avellaneda", "lanus", "almirante brown", "esteban echeverria",
    "matanza", "la matanza", "merlo", "moreno", "ituzaingo", "hurlingham",
    "trelew", "comodoro rivadavia", "ushuaia", "rio gallegos",
    "villa mercedes", "rio cuarto", "villa carlos paz", "san rafael",
    "concordia", "parana", "resistencia", "posadas", "san salvador de jujuy",
    "general roca", "viedma", "bariloche", "san carlos de bariloche",
}

# Medios (ampliar respecto del legacy)
BLACKLIST_MEDIOS = {
    "infobae", "clarin", "clarín", "la nacion", "la nación", "nacion", "nación",
    "telam", "télam", "el cronista", "cronista", "ambito", "ámbito",
    "pagina 12", "pagina/12", "página/12", "página 12", "p/12",
    "perfil", "diario perfil", "minutouno", "minuto uno", "infonews",
    "mdz", "mdzol", "la capital", "la voz", "la voz del interior",
    "el destape", "el destape web", "tsn", "tn", "todo noticias",
    "c5n", "a24", "lt10", "letra p", "el cohete a la luna",
    "baenegocios", "bae", "iprofesional", "el economista", "ieco",
    "lavoz", "rionegro", "rio negro", "diario rio negro", "diario río negro",
    "el patagonico", "patagonico", "el dia", "el día",
    "la tecla", "infocielo", "infobae", "noticias argentinas",
    "noticias mercado", "diario popular", "popular", "cronica",
    "redaccion", "agencia", "corresponsal", "diario digital",
    "primera linea", "uno", "diario uno", "uno santa fe",
    "data clave", "letrap", "agencia noticias argentinas",
    "noticias", "ambito financiero", "diariohoy", "diario hoy",
    "iprofesional", "el litoral", "litoral", "data24", "infogei",
    "diario digital conclusion", "diario digital conclusión",
}

BLACKLIST = (
    BLACKLIST_POLITICOS | BLACKLIST_CONCEPTOS |
    BLACKLIST_UBICACIONES | BLACKLIST_MEDIOS
)

# ──────────────────────────────────────────────────────────────────────────────
# ALIAS_EMPRESAS expandido (~150 entradas)
# ──────────────────────────────────────────────────────────────────────────────
ALIAS_EMPRESAS = {
    # Estado nacional / organismos
    "smn":              "Servicio Meteorologico Nacional",
    "indec":            "INDEC",
    "anses":            "ANSES",
    "afip":             "AFIP",
    "arca":             "ARCA",
    "conicet":          "CONICET",
    "invap":            "INVAP",
    "inti":             "INTI",
    "inta":             "INTA",
    "senasa":           "SENASA",
    "anmat":            "ANMAT",
    "ina":              "INA",
    "pami":             "PAMI",
    "incucai":          "INCUCAI",
    "enacom":           "ENACOM",
    "enarsa":           "ENARSA",
    "iea":              "IEASA",
    "ieasa":            "IEASA",
    "afsca":            "AFSCA",
    "aysa":             "AySA",
    "nucleoelectrica":  "Nucleoelectrica Argentina",
    "casa de moneda":   "Casa de Moneda",
    "vialidad nacional": "Vialidad Nacional",
    "dnv":              "Vialidad Nacional",
    "secretaria de trabajo": "Secretaria de Trabajo",
    "ministerio de trabajo": "Ministerio de Trabajo",
    "ministerio de salud": "Ministerio de Salud",
    "ministerio de capital humano": "Ministerio de Capital Humano",
    "tv publica":       "TV Publica",
    "radio nacional":   "Radio Nacional",
    "ypf":              "YPF",
    "trenes argentinos": "Trenes Argentinos",
    "sofse":            "Trenes Argentinos (SOFSE)",
    "belgrano cargas":  "Belgrano Cargas",
    "correo argentino": "Correo Argentino",
    "banco nacion":     "Banco Nacion",
    "banco nación":     "Banco Nacion",
    "banco provincia":  "Banco Provincia",
    "bna":              "Banco Nacion",
    "bapro":            "Banco Provincia",
    "banco ciudad":     "Banco Ciudad",
    "banco central":    "Banco Central",
    "bcra":             "Banco Central",
    "bice":             "BICE",
    # Aerolineas / transporte
    "aerolineas":           "Aerolineas Argentinas",
    "aerolíneas":           "Aerolineas Argentinas",
    "aerolineas argentinas": "Aerolineas Argentinas",
    "intercargo":           "Intercargo",
    "aeropuertos argentina 2000": "Aeropuertos Argentina 2000",
    "aa2000":               "Aeropuertos Argentina 2000",
    "flybondi":             "Flybondi",
    "jetsmart":             "JetSmart",
    "andes lineas aereas":  "Andes Lineas Aereas",
    # Automotrices
    "toyota":               "Toyota Argentina",
    "ford":                 "Ford Argentina",
    "volkswagen":           "Volkswagen Argentina",
    "vw":                   "Volkswagen Argentina",
    "general motors":       "General Motors Argentina",
    "gm":                   "General Motors Argentina",
    "stellantis":           "Stellantis Argentina",
    "fiat":                 "Stellantis Argentina",
    "peugeot":              "Stellantis Argentina",
    "renault":              "Renault Argentina",
    "honda":                "Honda Argentina",
    "iveco":                "Iveco Argentina",
    "scania":               "Scania Argentina",
    "mercedes benz":        "Mercedes-Benz Argentina",
    "mercedes-benz":        "Mercedes-Benz Argentina",
    "nissan":               "Nissan Argentina",
    "fate":                 "FATE",
    "pirelli":              "Pirelli",
    "bridgestone":          "Bridgestone Argentina",
    # Siderurgica / metalurgica
    "techint":              "Techint",
    "ternium":              "Ternium Argentina",
    "tenaris":              "Tenaris",
    "siderar":              "Ternium Siderar",
    "acindar":              "Acindar",
    "aluar":                "Aluar",
    "fate ":                "FATE",
    # Construccion / cemento
    "loma negra":           "Loma Negra",
    "holcim":               "Holcim Argentina",
    "cementos avellaneda":  "Cementos Avellaneda",
    "petersen thiele cruz": "Petersen Thiele y Cruz",
    "techint ingenieria":   "Techint",
    # Energia / oil & gas
    "pan american energy":  "Pan American Energy",
    "pae":                  "Pan American Energy",
    "pampa energia":        "Pampa Energia",
    "vista":                "Vista Energy",
    "tecpetrol":            "Tecpetrol",
    "shell":                "Shell Argentina",
    "axion":                "Axion Energy",
    "raizen":               "Raizen",
    "edesur":               "Edesur",
    "edenor":               "Edenor",
    "metrogas":             "Metrogas",
    "naturgy":              "Naturgy",
    "transener":            "Transener",
    "tgs":                  "TGS",
    "tgn":                  "TGN",
    "central puerto":       "Central Puerto",
    # Alimentacion
    "arcor":                "Arcor",
    "molinos":              "Molinos Rio de la Plata",
    "molinos rio de la plata": "Molinos Rio de la Plata",
    "ledesma":              "Ledesma",
    "mastellone":           "Mastellone Hermanos",
    "la serenisima":        "Mastellone Hermanos",
    "sancor":               "SanCor",
    "sancor seguros":       "SanCor Seguros",
    "danone":               "Danone Argentina",
    "nestle":               "Nestle Argentina",
    "kraft heinz":          "Kraft Heinz",
    "mondelez":             "Mondelez",
    "pepsico":              "PepsiCo Argentina",
    "coca cola":            "Coca-Cola",
    "coca-cola":            "Coca-Cola",
    "coca-cola femsa":      "Coca-Cola FEMSA",
    "georgalos":            "Georgalos",
    "felfort":              "Felfort",
    "havanna":              "Havanna",
    "bagley":               "Bagley",
    "bimbo":                "Bimbo Argentina",
    "fargo":                "Fargo",
    "lactear":              "Lactear",
    "lactalis":             "Lactalis Argentina",
    "verónica":             "Veronica",
    "veronica":             "Veronica",
    "swift":                "Swift",
    "paty":                 "Paty",
    "vicentin":             "Vicentin",
    "aceitera general deheza": "Aceitera General Deheza",
    "agd":                  "Aceitera General Deheza",
    "cargill":              "Cargill Argentina",
    "bunge":                "Bunge Argentina",
    "louis dreyfus":        "Louis Dreyfus",
    # Textil / calzado
    "alpargatas":           "Alpargatas",
    "dass":                 "DASS",
    "john foos":            "John Foos",
    "topper":               "Topper",
    "puma":                 "Puma Argentina",
    "adidas":               "Adidas Argentina",
    "nike":                 "Nike Argentina",
    "levis":                "Levi's Argentina",
    "levi's":               "Levi's Argentina",
    "tn platex":            "TN Platex",
    "ritex":                "Ritex",
    "santista":             "Santista Textil",
    # Retail / supermercados / electro
    "carrefour":            "Carrefour",
    "walmart":              "Walmart",
    "changomas":            "ChangoMas",
    "dia":                  "Dia Argentina",
    "coto":                 "Coto",
    "cencosud":             "Cencosud",
    "disco":                "Disco",
    "jumbo":                "Jumbo",
    "vea":                  "Vea",
    "la anonima":           "La Anonima",
    "garbarino":            "Garbarino",
    "fravega":              "Fravega",
    "frávega":              "Fravega",
    "musimundo":            "Musimundo",
    "easy":                 "Easy",
    "sodimac":              "Sodimac",
    "falabella":            "Falabella",
    "mercado libre":        "Mercado Libre",
    "meli":                 "Mercado Libre",
    "tiendanube":           "Tiendanube",
    "rappi":                "Rappi",
    "pedidosya":            "PedidosYa",
    "globant":              "Globant",
    "despegar":             "Despegar",
    # Bancos / financieras
    "banco galicia":        "Banco Galicia",
    "banco macro":          "Banco Macro",
    "banco santander":      "Banco Santander",
    "banco bbva":           "BBVA",
    "bbva":                 "BBVA",
    "banco patagonia":      "Banco Patagonia",
    "banco supervielle":    "Banco Supervielle",
    "banco hipotecario":    "Banco Hipotecario",
    "banco itau":           "Banco Itau",
    "icbc":                 "ICBC Argentina",
    "hsbc":                 "HSBC Argentina",
    "credicoop":            "Banco Credicoop",
    # Telco / tech
    "telecom":              "Telecom Argentina",
    "telecom argentina":    "Telecom Argentina",
    "claro argentina":      "Claro Argentina",
    "personal flow":        "Personal (Telecom)",
    "movistar":             "Movistar Argentina",
    "telefonica":           "Telefonica",
    "telefónica":           "Telefonica",
    "directv":              "DirecTV Argentina",
    "cablevision":          "Cablevision",
    # Otros
    "bioceres":             "Bioceres",
    "celulosa":             "Celulosa Argentina",
    "celulosa argentina":   "Celulosa Argentina",
    "papel prensa":         "Papel Prensa",
    "siat":                 "SIAT",
    "lumilagro":            "Lumilagro",
    "secco":                "Secco",
    "iram":                 "IRAM",
    "ipra":                 "IPRA",
    # Salud
    "swiss medical":        "Swiss Medical",
    "osde":                 "OSDE",
    "galeno":               "Galeno",
    "medife":               "Medife",
    "omint":                "OMINT",
    "hospital italiano":    "Hospital Italiano",
    "hospital aleman":      "Hospital Aleman",
    "hospital britanico":   "Hospital Britanico",
    "fleni":                "FLENI",
    # Construccion / desarrolladoras
    "iecsa":                "IECSA",
    "techint construccion": "Techint",
    "supercemento":         "Supercemento",
    "esuco":                "ESUCO",
    "rovella":              "Rovella Carranza",
    "milicic":              "Milicic",
    "homecenter sodimac":   "Sodimac",
}

# ──────────────────────────────────────────────────────────────────────────────
# Keywords
# ──────────────────────────────────────────────────────────────────────────────
KEYWORDS_RELEVANTES = [
    "despido", "despidos", "despedido", "despedidos", "despide", "despidio",
    "despidieron", "echaron", "echo a", "echó a",
    "cierre", "cerro", "cerró", "cerraron", "cierra", "cerrara", "cerrará",
    "quiebra", "quiebras", "quebró", "quebro", "concurso de acreedores",
    "suspension", "suspensión", "suspensiones", "suspendidos", "suspendido",
    "desvinculacion", "desvinculación", "desvinculaciones", "desvinculado", "desvinculados",
    "cesanteo", "cesanteos", "cesanteados",
    "reduccion de personal", "recorte de personal", "ajuste de personal",
    "retiro voluntario", "retiros voluntarios",
    "deja de producir", "dejo de producir", "dejó de producir", "cierre de planta",
    "traslada produccion", "traslada producción", "traslado de produccion",
    "presento concurso", "presentó concurso", "preventivo de crisis",
    "sin trabajo", "sin empleo", "perdida de empleo", "pérdida de empleo",
    "perdida de puestos", "pérdida de puestos",
    "echo personal", "echó personal", "achica personal", "achicó personal",
]

KEYWORDS_SUSPENSION = [
    "suspension", "suspensión", "suspensiones", "suspendido", "suspendidos",
    "suspende", "suspendieron", "licencia sin goce", "preventivo de crisis",
]
KEYWORDS_CIERRE = [
    "cierre", "cerro", "cerró", "cerraron", "cerrara", "cerrará", "cierra",
    "quiebra", "quiebras", "quebro", "quebró", "concurso de acreedores",
    "liquidacion", "liquidación", "liquido", "liquidó",
    "deja de producir", "dejo de producir", "dejó de producir",
    "cierre de planta", "traslada produccion", "traslada producción",
    "preventivo de crisis",
]

# Articulos resumen / agregados → rechazar
KEYWORDS_RESUMEN = [
    "balance anual", "acumulado", "total de despidos", "miles de despidos",
    "oleada de despidos", "ola de despidos", "record de despidos", "récord de despidos",
    "año de despidos", "ano de despidos", "meses de despidos",
    "primer semestre", "segundo semestre", "primer trimestre", "ultimo trimestre",
    "ranking ", "informe anual", "informe mensual", "reporte mensual",
    "historico ", "histórico ", "estadistica", "estadística",
    "segun el indec", "según el indec", "segun indec",
    "indice de desempleo", "índice de desempleo", "tasa de desempleo",
    "mercado laboral", "situacion laboral", "situación laboral",
    "empleo en argentina", "panorama laboral", "crisis laboral",
    "los despidos del mes", "los despidos de", "cuantos despidos", "cuántos despidos",
    "lista de despidos", "listado de despidos", "todas las empresas que",
]

# Consultas RSS
CONSULTAS_RSS = [
    "despidos trabajadores Argentina",
    "despido empresa Argentina",
    "suspension trabajadores Argentina",
    "suspensiones empresa Argentina",
    "cierre empresa Argentina",
    "cierre fabrica Argentina",
    "cierre planta Argentina",
    "quiebra empresa Argentina",
    "concurso acreedores empresa Argentina",
    "desvinculaciones Argentina",
    "recorte personal empresa Argentina",
    "retiro voluntario empresa Argentina",
    "deja producir Argentina",
    "preventivo de crisis Argentina",
    "echaron trabajadores Argentina",
]

# Provincias canonicas
PROVINCIAS_CANON = {
    "buenos aires": "Buenos Aires",
    "caba": "CABA", "ciudad de buenos aires": "CABA", "ciudad autonoma": "CABA",
    "cordoba": "Cordoba", "córdoba": "Cordoba",
    "rosario": "Santa Fe", "santa fe": "Santa Fe",
    "mendoza": "Mendoza", "tucuman": "Tucuman", "tucumán": "Tucuman",
    "salta": "Salta", "entre rios": "Entre Rios", "entre ríos": "Entre Rios",
    "chaco": "Chaco", "corrientes": "Corrientes", "misiones": "Misiones",
    "jujuy": "Jujuy", "san juan": "San Juan",
    "rio negro": "Rio Negro", "río negro": "Rio Negro",
    "neuquen": "Neuquen", "neuquén": "Neuquen",
    "formosa": "Formosa", "chubut": "Chubut", "la pampa": "La Pampa",
    "santiago del estero": "Santiago del Estero",
    "san luis": "San Luis", "catamarca": "Catamarca", "la rioja": "La Rioja",
    "santa cruz": "Santa Cruz", "tierra del fuego": "Tierra del Fuego",
}

# Rubros canonicos basados en keywords del texto
RUBRO_KEYWORDS = {
    "Automotriz":      ["automotriz", "autopart", "fabrica de auto", "terminal automotriz",
                        "neumatic", "neumátic", "armado de auto"],
    "Metalurgica":     ["metalurgic", "siderurgic", "acero", "aluminio", "fundicion", "fundición"],
    "Textil, Calzado": ["textil", "calzado", "zapateri", "indumentaria", "confeccion", "confección"],
    "Construccion":    ["construccion", "construcción", "obra publica", "obra pública", "cemento",
                        "ladrillo", "viviendas", "albañil"],
    "Alimentacion":    ["alimenticia", "alimento", "lactea", "láctea", "frigorific", "panaderi",
                        "panaderí", "harinera", "molino", "carne", "azucar", "azúcar"],
    "Comercio":        ["supermercado", "comercio", "retail", "tienda", "shopping", "local comercial"],
    "Transporte":      ["transporte", "logistic", "logístic", "camioner", "ferroviari", "aerolinea",
                        "aerolínea", "linea aerea"],
    "Energia":         ["petrolera", "gas natural", "energia", "energía", "electric", "hidrocarburo",
                        "oil", "refineri"],
    "Agro":            ["agro", "rural", "cosecha", "siembra", "ganaderia", "ganadería", "sojer",
                        "agricultor"],
    "Medios de comunicacion": ["medio de comunicacion", "diario", "radio ", "tv ", "tve", "periodist"],
    "Educacion":       ["educacion", "educación", "escuela", "colegio", "universidad", "docente"],
    "Salud":           ["hospital", "clinic", "sanatorio", "salud ", "medic", "obra social"],
    "Tecnologia":      ["tecnologi", "tecnologí", "software", "tech ", "fintech", "startup", "it ",
                        "informatic", "informátic"],
    "Estado/Publico":  ["estatal", "publico", "público", "ministerio", "secretaria", "secretaría",
                        "municipio", "intendencia", "organismo"],
}

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def log(msg: str, level: str = "info") -> None:
    ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
    icon = {"info": "  ", "ok": "OK", "warn": "!!", "err": "XX"}.get(level, "  ")
    print(f"[{ts}] {icon} {msg}", flush=True)


def limpiar_texto(t: str | None) -> str:
    if not t:
        return ""
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def normalizar_key(nombre: str) -> str:
    """Clave de comparacion: lowercase, sin sufijos legales, sin articulos."""
    if not nombre:
        return ""
    n = nombre.lower().strip()
    # Quitar acentos basicos (mantener ñ)
    repl = str.maketrans({"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u"})
    n = n.translate(repl)
    # Quitar sufijos legales
    n = re.sub(r"\s+(s\.?a\.?u?\.?|s\.?r\.?l\.?|s\.?a\.?s\.?|ltda?\.?|inc\.?|llc\.?)$", "", n)
    n = re.sub(r"^(el|la|los|las|el\s+grupo|grupo)\s+", "", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


def es_blacklist(nombre: str) -> bool:
    """True si nombre cae en blacklist (no es empresa real)."""
    if not nombre:
        return True
    k = normalizar_key(nombre)
    if not k or len(k) < 3:
        return True
    if k in BLACKLIST:
        return True
    # Tambien matchea prefijos/sufijos comunes
    for bl in BLACKLIST:
        if k == bl or k.startswith(bl + " ") or k.endswith(" " + bl):
            return True
    return False


# ──────────────────────────────────────────────────────────────────────────────
# Fechas
# ──────────────────────────────────────────────────────────────────────────────
_MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def parsear_fecha_rss(fecha_str: str) -> date | None:
    if not fecha_str or not fecha_str.strip():
        return None
    s = fecha_str.strip()
    try:
        dt = parsedate_to_datetime(s)
        if dt is not None:
            return dt.date() if hasattr(dt, "date") else dt
    except (TypeError, ValueError, IndexError):
        pass
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ",
                "%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def detectar_anio_viejo(texto: str) -> int | None:
    """Anio mas antiguo mencionado en el texto, si es < FECHA_MINIMA.year."""
    pat_mes = r"\b(" + "|".join(_MESES_ES.keys()) + r")\s+(?:de\s+)?(\d{4})\b"
    pat_anio = r"\b(?:en|durante|del?|desde|hasta|para)\s+(\d{4})\b"
    candidatos = set()
    for pat in (pat_mes, pat_anio):
        for m in re.finditer(pat, texto, re.IGNORECASE):
            try:
                y = int(m.group(len(m.groups())))
                if 1990 < y < 2030:
                    candidatos.add(y)
            except Exception:
                pass
    viejos = [y for y in candidatos if y < FECHA_MINIMA.year]
    return min(viejos) if viejos else None


def validar_fecha_evento(f: date | None, texto: str, hoy: date | None = None) -> tuple[bool, str]:
    hoy = hoy or date.today()
    if f is None:
        return False, "fecha no parseable"
    if f > hoy + timedelta(days=2):
        return False, f"fecha futura ({f.isoformat()})"
    if f < FECHA_MINIMA:
        return False, f"fecha < FECHA_MINIMA ({f.isoformat()})"
    anio_viejo = detectar_anio_viejo(texto)
    if anio_viejo is not None and abs(f.year - anio_viejo) >= 2:
        return False, f"contenido refiere {anio_viejo} (pubDate {f.year}): historico"
    return True, "ok"


# ──────────────────────────────────────────────────────────────────────────────
# Extraccion de campos del articulo
# ──────────────────────────────────────────────────────────────────────────────
VERBOS_ACCION = (r"(?:despide|despidi[óo]|despiden|suspende|suspendi[óo]|suspenden|"
                 r"cierra|cerr[óo]|cierran|anuncia|anunci[óo]|anuncian|recorta|recort[óo]|"
                 r"recortan|confirma|confirm[óo]|confirman|quiebra|quebr[óo]|"
                 r"liquida|liquid[óo]|echa|ech[óo]|echan)")


def _limpiar_candidato_empresa(frase: str) -> str:
    """Limpia prefijos y sufijos genericos de un candidato a empresa."""
    palabras = frase.split()
    GENERICOS_LIMITE = {"el", "la", "los", "las", "un", "una", "de", "del", "al",
                        "en", "su", "sus", "que", "lo", "tras", "ante",
                        "trabajadores", "empleados", "operarios", "personal",
                        "empresa", "fabrica", "fábrica", "planta", "industria",
                        "gigante", "firma", "compania", "compañia", "compañía"}
    while palabras:
        p = re.sub(r"[^a-záéíóúñA-ZÁÉÍÓÚÑ]", "", palabras[0]).lower()
        if p in GENERICOS_LIMITE:
            palabras.pop(0)
        else:
            break
    while palabras:
        p = re.sub(r"[^a-záéíóúñA-ZÁÉÍÓÚÑ]", "", palabras[-1]).lower()
        if p in GENERICOS_LIMITE:
            palabras.pop()
        else:
            break
    return " ".join(palabras).strip(" .,;:-")


def extraer_empresa(titulo: str, cuerpo: str = "") -> str:
    """
    Estrategia (de mas a menos confiable):
      0. Match con ALIAS_EMPRESAS via substring exacto (alta señal).
      1. Token en MAYUSCULAS sostenidas (>=3 chars, alfanum, no en blacklist).
      2. Sujeto directo del verbo de accion: 'Ford despide...'
      3. Frase entre comillas que contenga capitalizada.
      4. "en/de NombreEmpresa" antes de verbo o punctuacion.
      5. Secuencia capitalizada mas larga, descontando blacklist.
    """
    if not titulo:
        return "Sin identificar"
    t = titulo.strip()
    # Sacar nombre del medio al final
    t = re.sub(r"\s*[-–|]\s*[A-ZÁÉÍÓÚÑ][A-Za-záéíóúñÁÉÍÓÚÑ\s.]{2,30}$", "", t).strip()

    texto_busqueda = t + " " + (cuerpo[:500] if cuerpo else "")

    # 0. Alias por substring: solo aplica a alias multi-palabra o lo suficientemente
    # distintivos (>=6 chars). Acronimos cortos se manejan en el paso 1.
    t_low = texto_busqueda.lower()
    alias_keys = sorted(ALIAS_EMPRESAS.keys(), key=len, reverse=True)
    for alias in alias_keys:
        if " " not in alias and len(alias) < 6:
            continue
        if re.search(r"\b" + re.escape(alias) + r"\b", t_low):
            return ALIAS_EMPRESAS[alias]

    # 1. Mayusculas sostenidas en el titulo (acronimos como YPF, FATE, ARCOR, SMN)
    for m in re.finditer(r"\b([A-ZÁÉÍÓÚÑ]{3,15})\b", t):
        token = m.group(1)
        if token in {"DESPIDO", "DESPIDOS", "CIERRE", "CIERRES", "ALERTA",
                     "CRISIS", "PARO", "PROTESTA", "MILEI", "TRUMP", "PBA",
                     "CABA", "INDEC", "AFIP", "ANSES"}:
            # ANSES/INDEC/AFIP SON empresas → no descartar
            if token in {"INDEC", "AFIP", "ANSES", "ARCA", "CONICET", "INVAP",
                         "INTI", "INTA", "SENASA", "ANMAT", "PAMI", "INCUCAI",
                         "ENACOM", "ENARSA", "IEASA", "YPF", "BCRA", "BICE",
                         "FATE", "DASS", "ICBC", "HSBC", "BBVA", "AGD",
                         "TGS", "TGN", "PAE"}:
                if not es_blacklist(token):
                    return ALIAS_EMPRESAS.get(token.lower(), token)
            continue
        # Acronimo conocido
        if token.lower() in ALIAS_EMPRESAS:
            return ALIAS_EMPRESAS[token.lower()]
        # Acronimo desconocido pero corto + valido: aceptar
        if 3 <= len(token) <= 6 and not es_blacklist(token):
            return token

    # 2. Sujeto del verbo de accion al inicio
    m = re.match(
        r"^([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñÁÉÍÓÚÑ0-9\s&.\-]{1,50}?)\s+" + VERBOS_ACCION,
        t
    )
    if m:
        cand = _limpiar_candidato_empresa(m.group(1))
        if cand and not es_blacklist(cand) and len(cand) >= 3:
            return cand

    # 3. Entre comillas (alta señal en titulares)
    QUOTES_OPEN = "«\"'‘“"
    QUOTES_CLOSE = "»\"'’”"
    quote_pat = "[" + re.escape(QUOTES_OPEN) + "]" \
        + r"([A-ZÁÉÍÓÚÑ][^" + re.escape(QUOTES_OPEN + QUOTES_CLOSE) + r"]{2,50}?)" \
        + "[" + re.escape(QUOTES_CLOSE) + "]"
    for m in re.finditer(quote_pat, t):
        cand = _limpiar_candidato_empresa(m.group(1))
        if cand and not es_blacklist(cand):
            return cand

    # 4. "en/de NombreEmpresa" antes de verbo
    for pat in [
        r"(?:en|de)\s+(?:el\s+|la\s+|los\s+|las\s+)?([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñÁÉÍÓÚÑ\s&.\-]{2,50}?)"
        r"(?:\s*[:,\-]|\s+" + VERBOS_ACCION + r"|\s*$)",
    ]:
        m = re.search(pat, t)
        if m:
            cand = _limpiar_candidato_empresa(m.group(1))
            if cand and not es_blacklist(cand) and len(cand) >= 3:
                return cand

    # 5. Secuencia capitalizada mas larga, sin blacklist
    candidatos = []
    for m in re.finditer(
        r"\b([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ]+(?:\s+(?:de|del|la|el|los|las|y|&)?\s*"
        r"[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ]+){0,3})", t
    ):
        cand = _limpiar_candidato_empresa(m.group(1))
        if cand and not es_blacklist(cand) and len(cand) >= 4:
            candidatos.append(cand)
    if candidatos:
        # Preferir el mas largo
        return max(candidatos, key=len)

    return "Sin identificar"


def extraer_empleados(titulo: str, cuerpo: str = "") -> int:
    """Extrae cantidad de empleados afectados. Cap en MAX_EMPLEADOS_AUTO.

    Estrategia:
      1. Buscar primero en titulo (mas confiable: si el titular dice 200, son 200).
      2. Si no hay, buscar en cuerpo (primer parrafo da el dato la mayoria de las veces).
      3. Rechazar numeros >MAX_EMPLEADOS_AUTO (articulos resumen).
    """
    PATRONES = [
        # "200 trabajadores/empleados/operarios"
        r"(\d[\d.,]*)\s*(?:trabajadores|empleados|operarios|personas|puestos|familias|obreros)",
        # "200 despidos/suspensiones/desvinculaciones"
        r"(\d[\d.,]*)\s*(?:despidos|suspensiones|desvinculaciones|cesanteos)",
        # "despidio/despidio a 200" - acepta hasta una preposicion antes del numero
        r"despid[oió]+\s+(?:a\s+)?(\d[\d.,]*)",
        r"suspendi[oó]+\s+(?:a\s+)?(\d[\d.,]*)",
        r"ech[oó]+\s+(?:a\s+)?(\d[\d.,]*)",
        r"cesantea[rdo]+\s+(?:a\s+)?(\d[\d.,]*)",
        # "deja sin trabajo a 200"
        r"sin\s+(?:trabajo|empleo)\s+a\s+(\d[\d.,]*)",
        # "afecta a 200"
        r"afecta(?:r|ra|ria|rán|ran)?\s+a\s+(\d[\d.,]*)",
    ]
    for texto in (titulo, cuerpo[:1500] if cuerpo else ""):
        if not texto:
            continue
        for pat in PATRONES:
            m = re.search(pat, texto, re.IGNORECASE)
            if not m:
                continue
            val = m.group(1).replace(".", "").replace(",", "")
            try:
                n = int(val)
                if 1 <= n <= MAX_EMPLEADOS_AUTO:
                    return n
            except ValueError:
                pass
    return 1


def detectar_tipo(texto: str) -> str:
    t = texto.lower()
    if any(k in t for k in KEYWORDS_CIERRE):
        return "cierre"
    if any(k in t for k in KEYWORDS_SUSPENSION):
        return "suspensión"
    return "despido"


def detectar_provincia(texto: str) -> str:
    t = texto.lower()
    for k, v in PROVINCIAS_CANON.items():
        if re.search(r"\b" + re.escape(k) + r"\b", t):
            return v
    return "Por determinar"


def detectar_rubro(texto: str) -> str:
    t = texto.lower()
    for rubro, kws in RUBRO_KEYWORDS.items():
        for kw in kws:
            if kw in t:
                return rubro
    return "Otro"


def detectar_sector(empresa: str, texto: str) -> str:
    """Privado vs Publico. Heuristica: si empresa esta en ALIAS publicos, publico."""
    PUBLICAS = {
        "ANSES", "AFIP", "ARCA", "INDEC", "CONICET", "INVAP", "INTI", "INTA",
        "SENASA", "ANMAT", "PAMI", "INCUCAI", "ENACOM", "ENARSA", "IEASA",
        "AySA", "Nucleoelectrica Argentina", "Vialidad Nacional",
        "Casa de Moneda", "TV Publica", "Radio Nacional", "YPF",
        "Trenes Argentinos", "Belgrano Cargas", "Correo Argentino",
        "Banco Nacion", "Banco Provincia", "Banco Ciudad", "Banco Central",
        "Aerolineas Argentinas", "Intercargo", "Servicio Meteorologico Nacional",
        "Ministerio de Trabajo", "Ministerio de Salud", "Ministerio de Capital Humano",
        "Secretaria de Trabajo", "BICE", "Trenes Argentinos (SOFSE)",
    }
    if empresa in PUBLICAS:
        return "Sector Público"
    if any(k in texto.lower() for k in ["estatal", "publico", "público", "ministerio",
                                         "municipalidad", "organismo", "secretaria de estado"]):
        return "Sector Público"
    return "Sector Privado"


# ──────────────────────────────────────────────────────────────────────────────
# Fetch articulo
# ──────────────────────────────────────────────────────────────────────────────
def fetch_articulo(url: str) -> dict:
    """Baja la URL, sigue redirects, extrae body + fecha publicacion.

    Devuelve {'cuerpo': str, 'fecha_publicacion': date|None, 'url_final': str}.
    Soft-fail: si algo rompe, devuelve cuerpo vacio.
    """
    out = {"cuerpo": "", "fecha_publicacion": None, "url_final": url}
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT_ARTICULO, allow_redirects=True)
        if r.status_code != 200 or not r.text:
            return out
        out["url_final"] = r.url
        html_text = r.text

        # Fecha de publicacion via meta tags
        for pat in [
            r'<meta[^>]+property=["\']article:published_time["\'][^>]+content=["\']([^"\']+)["\']',
            r'<meta[^>]+name=["\']pubdate["\'][^>]+content=["\']([^"\']+)["\']',
            r'<meta[^>]+name=["\']publish[_-]?date["\'][^>]+content=["\']([^"\']+)["\']',
            r'<meta[^>]+itemprop=["\']datePublished["\'][^>]+content=["\']([^"\']+)["\']',
            r'<time[^>]+datetime=["\']([^"\']+)["\']',
        ]:
            m = re.search(pat, html_text, re.IGNORECASE)
            if m:
                fp = parsear_fecha_rss(m.group(1))
                if fp:
                    out["fecha_publicacion"] = fp
                    break

        # Extraer body
        if HAS_LXML:
            try:
                doc = lxml_html.fromstring(html_text)
                # Limpiar scripts, styles, nav, footer
                for tag in doc.xpath("//script | //style | //nav | //footer | //aside | //iframe"):
                    tag.getparent().remove(tag)
                # Buscar <article>, <main>, o el body completo
                article = doc.xpath("//article")
                if article:
                    text = article[0].text_content()
                else:
                    main = doc.xpath("//main")
                    if main:
                        text = main[0].text_content()
                    else:
                        # Fallback: meta og:description + body principal
                        m = re.search(
                            r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)["\']',
                            html_text, re.IGNORECASE)
                        og_desc = m.group(1) if m else ""
                        body = doc.xpath("//body")
                        text = og_desc + " " + (body[0].text_content() if body else "")
                out["cuerpo"] = re.sub(r"\s+", " ", text).strip()[:5000]
            except Exception:
                out["cuerpo"] = limpiar_texto(html_text)[:5000]
        else:
            # Sin lxml, fallback a meta og:description
            m = re.search(
                r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)["\']',
                html_text, re.IGNORECASE)
            out["cuerpo"] = m.group(1) if m else ""
    except Exception as e:
        log(f"fetch_articulo error {url[:60]}: {e}", "warn")
    return out


def fetch_rss(query: str) -> bytes | None:
    q = urllib.parse.quote(query)
    url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=AR&ceid=AR:es-419"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": HEADERS["User-Agent"]})
        with urllib.request.urlopen(req, timeout=TIMEOUT_RSS) as r:
            return r.read()
    except Exception as e:
        log(f"RSS error ({query}): {e}", "warn")
        return None


# ──────────────────────────────────────────────────────────────────────────────
# Relevancia / dedupe
# ──────────────────────────────────────────────────────────────────────────────
def es_relevante(titulo: str, descripcion: str = "") -> bool:
    texto = (titulo + " " + descripcion).lower()
    if not any(k in texto for k in KEYWORDS_RELEVANTES):
        return False
    if any(k in texto for k in KEYWORDS_RESUMEN):
        return False
    return True


STOPWORDS = {
    "el", "la", "los", "las", "un", "una", "de", "del", "al", "en", "y", "o", "a",
    "que", "con", "por", "para", "su", "sus", "se", "le", "les", "es", "son",
    "fue", "han", "hay", "pero", "como", "mas", "más", "sin", "sobre", "entre",
    "despido", "despidos", "suspension", "suspensión", "suspensiones",
    "cierre", "cierres", "quiebra", "argentina", "empresa", "empresas",
    "trabajadores", "trabajador", "empleados", "empleado", "personal",
    "nacional", "gobierno", "fabrica", "fábrica", "industrial",
}


def palabras_clave(texto: str) -> set[str]:
    palabras = re.findall(r"[a-záéíóúñA-ZÁÉÍÓÚÑ]{4,}", texto.lower())
    return {p for p in palabras if p not in STOPWORDS}


def son_duplicados(a: str, b: str, umbral: int = 4) -> bool:
    return len(palabras_clave(a) & palabras_clave(b)) >= umbral


def deduplicar(eventos: list[dict]) -> list[dict]:
    """
    Pasada 1: por (empresa_normalizada, YYYY-MM).
    Pasada 2: por similitud de titulo en el mismo mes.
    Conserva siempre el evento con mayor cantidad de empleados.
    """
    grupos = {}
    for ev in eventos:
        mes = ev.get("fecha", "0000-00")[:7]
        emp_key = normalizar_key(ev.get("empresa", ""))
        key = (emp_key, mes)
        if key not in grupos or ev.get("empleados", 0) > grupos[key].get("empleados", 0):
            grupos[key] = ev
    resultado = list(grupos.values())

    # Pasada 2: similitud cross-empresa en el mismo mes
    por_mes = defaultdict(list)
    for ev in resultado:
        mes = ev.get("fecha", "0000-00")[:7]
        por_mes[mes].append(ev)
    final = []
    for evs in por_mes.values():
        usados = [False] * len(evs)
        for i in range(len(evs)):
            if usados[i]:
                continue
            grupo = [evs[i]]
            for j in range(i + 1, len(evs)):
                if usados[j]:
                    continue
                if son_duplicados(evs[i].get("comentario", ""), evs[j].get("comentario", "")):
                    grupo.append(evs[j])
                    usados[j] = True
            mejor = max(grupo, key=lambda e: e.get("empleados", 0))
            final.append(mejor)
    return final


def canonicalizar(eventos: list[dict]) -> list[dict]:
    """Aplica ALIAS_EMPRESAS a todos los eventos via normalizar_key."""
    for ev in eventos:
        nombre = ev.get("empresa", "").strip()
        if not nombre or nombre == "Sin identificar":
            continue
        k = normalizar_key(nombre)
        if k in ALIAS_EMPRESAS:
            ev["empresa"] = ALIAS_EMPRESAS[k]
            continue
        # Prefix match para casos como "Toyota Argentina S.A." → "Toyota Argentina"
        for alias, canon in ALIAS_EMPRESAS.items():
            if len(alias) >= 5 and (k == alias or k.startswith(alias + " ")):
                ev["empresa"] = canon
                break
    return eventos


# ──────────────────────────────────────────────────────────────────────────────
# Scraping principal
# ──────────────────────────────────────────────────────────────────────────────
def scraping_rss(vistos: set[str], max_fetch: int) -> tuple[list[dict], dict]:
    """Devuelve (nuevos_eventos, stats)."""
    nuevos = []
    rechazos = Counter()
    fetched = 0
    for query in CONSULTAS_RSS:
        log(f"Consultando: {query}")
        data = fetch_rss(query)
        if not data:
            continue
        try:
            root = ET.fromstring(data)
        except Exception as e:
            log(f"  Parse error: {e}", "warn")
            continue
        channel = root.find("channel")
        if channel is None:
            continue
        for item in channel.findall("item"):
            link = limpiar_texto(item.findtext("link", ""))
            if not link or link in vistos:
                continue
            titulo = limpiar_texto(item.findtext("title", ""))
            desc = limpiar_texto(item.findtext("description", ""))
            pub = item.findtext("pubDate", "")

            if not es_relevante(titulo, desc):
                vistos.add(link)
                rechazos["no_relevante"] += 1
                continue

            # Bajar cuerpo del articulo si tenemos cuota
            cuerpo = ""
            fecha_articulo = None
            if fetched < max_fetch:
                art = fetch_articulo(link)
                cuerpo = art["cuerpo"]
                fecha_articulo = art["fecha_publicacion"]
                if art["url_final"] != link:
                    vistos.add(art["url_final"])
                fetched += 1
                time.sleep(0.3)  # rate limiting cortes

            # Fecha: priorizar la del articulo, fallback al pubDate del RSS
            fecha = fecha_articulo or parsear_fecha_rss(pub)
            texto_validacion = titulo + " " + desc + " " + cuerpo[:1000]
            ok_fecha, razon = validar_fecha_evento(fecha, texto_validacion)
            if not ok_fecha:
                rechazos[razon.split("(")[0].strip()] += 1
                vistos.add(link)
                continue

            empresa = extraer_empresa(titulo, cuerpo + " " + desc)
            empleados = extraer_empleados(titulo, (desc + " " + cuerpo) if cuerpo else desc)
            tipo = detectar_tipo(titulo + " " + desc + " " + cuerpo[:500])
            provincia = detectar_provincia(titulo + " " + desc + " " + cuerpo[:1000])
            rubro = detectar_rubro(titulo + " " + cuerpo[:1000])
            sector = detectar_sector(empresa, titulo + " " + cuerpo[:500])
            cerro = tipo == "cierre"

            nuevos.append({
                "fecha":      fecha.isoformat(),
                "empresa":    empresa,
                "rubro":      rubro,
                "empleados":  empleados,
                "comentario": titulo,
                "provincia":  provincia,
                "municipio":  "",
                "estado":     "Pendiente revisión",
                "sector":     sector,
                "cerro":      cerro,
                "tipo":       tipo,
                "fuente":     "auto-rss",
                "url":        link,
            })
            vistos.add(link)
    log(f"Articulos bajados: {fetched}")
    log(f"Nuevos antes dedup: {len(nuevos)}")
    nuevos = deduplicar(canonicalizar(nuevos))
    log(f"Nuevos despues dedup: {len(nuevos)}", "ok")
    if rechazos:
        log("Rechazos:")
        for razon, n in sorted(rechazos.items(), key=lambda x: -x[1]):
            log(f"  - {razon}: {n}")
    return nuevos, dict(rechazos)


# ──────────────────────────────────────────────────────────────────────────────
# Procesamiento de una URL puntual (modo --url)
# ──────────────────────────────────────────────────────────────────────────────
def procesar_url_unica(url: str) -> list[dict]:
    """Baja la URL, extrae titulo + campos, devuelve [evento] o []."""
    art = fetch_articulo(url)
    cuerpo = art.get("cuerpo", "")
    if not cuerpo:
        log(f"URL sin cuerpo extraible: {url[:80]}", "warn")
        return []

    # Titulo via <title> o og:title
    titulo = ""
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT_ARTICULO, allow_redirects=True)
        m = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']+)["\']',
                      r.text, re.IGNORECASE)
        if m:
            titulo = limpiar_texto(m.group(1))
        else:
            m = re.search(r"<title>([^<]+)</title>", r.text, re.IGNORECASE)
            if m:
                titulo = limpiar_texto(m.group(1))
    except Exception as e:
        log(f"No pude bajar titulo: {e}", "warn")

    if not titulo:
        titulo = cuerpo[:200]

    fecha = art.get("fecha_publicacion") or date.today()
    texto_validacion = titulo + " " + cuerpo[:1000]
    ok_fecha, razon = validar_fecha_evento(fecha, texto_validacion)
    if not ok_fecha:
        log(f"URL rechazada por fecha: {razon}", "warn")
        return []

    empresa = extraer_empresa(titulo, cuerpo)
    empleados = extraer_empleados(titulo, cuerpo)
    tipo = detectar_tipo(titulo + " " + cuerpo[:500])
    provincia = detectar_provincia(titulo + " " + cuerpo[:1000])
    rubro = detectar_rubro(titulo + " " + cuerpo[:1000])
    sector = detectar_sector(empresa, titulo + " " + cuerpo[:500])

    evento = {
        "fecha":      fecha.isoformat(),
        "empresa":    empresa,
        "rubro":      rubro,
        "empleados":  empleados,
        "comentario": titulo,
        "provincia":  provincia,
        "municipio":  "",
        "estado":     "Pendiente revisión",
        "sector":     sector,
        "cerro":      tipo == "cierre",
        "tipo":       tipo,
        "fuente":     "auto-link",
        "url":        url,
    }
    log(f"URL procesada: empresa={empresa} | empleados={empleados} | tipo={tipo}", "ok")
    return [evento]


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="No guarda despidos.json")
    parser.add_argument("--max-fetch", type=int, default=MAX_ARTICULOS_FETCH,
                        help=f"Cap de articulos a bajar (default {MAX_ARTICULOS_FETCH})")
    parser.add_argument("--skip-rss", action="store_true",
                        help="Skip scraping, solo regenera desde bedrock (debug)")
    parser.add_argument("--url", type=str, default=None,
                        help="Procesar solo esta URL (modo manual)")
    args = parser.parse_args(argv)

    log("=" * 60)
    log("Monitor de Despidos - Scraping diario")
    log("=" * 60)

    # 1. Cargar bedrock
    if not HISTORICAL.exists():
        log(f"Bedrock no encontrado: {HISTORICAL}", "err")
        return 1
    with open(HISTORICAL, "r", encoding="utf-8") as f:
        bedrock = json.load(f)
    eventos_previos = bedrock.get("eventos", [])
    log(f"Bedrock: {len(eventos_previos)} eventos")

    # URLs vistas: derivar de los eventos previos
    vistos = {e.get("url", "") for e in eventos_previos if e.get("url")}
    vistos.discard("")
    log(f"URLs ya vistas: {len(vistos)}")

    # 2. Scraping
    if args.url:
        log(f"MODO --url: procesando {args.url}")
        nuevos = procesar_url_unica(args.url)
    elif args.skip_rss:
        log("SKIP RSS (--skip-rss)")
        nuevos = []
    else:
        nuevos, _stats = scraping_rss(vistos, max_fetch=args.max_fetch)

    # 3. Combinar
    AUTO_FUENTES = {"auto-rss", "auto-link"}
    eventos_auto_previos = [e for e in eventos_previos if e.get("fuente") in AUTO_FUENTES]
    eventos_manual = [e for e in eventos_previos if e.get("fuente") not in AUTO_FUENTES]

    # Re-procesar eventos auto-rss viejos con el extractor nuevo (limpia basura legacy).
    # Conserva 'empleados' si era >1 (asumimos que fue extraido bien); recalcula si era 1.
    relaboreados = 0
    for ev in eventos_auto_previos:
        coment = ev.get("comentario", "")
        if not coment:
            continue
        empresa_old = ev.get("empresa", "")
        empresa_new = extraer_empresa(coment, "")
        if empresa_new != "Sin identificar" and (
            empresa_old == "Sin identificar" or
            normalizar_key(empresa_old) in BLACKLIST or
            es_blacklist(empresa_old)
        ):
            ev["empresa"] = empresa_new
            relaboreados += 1
        elif empresa_new == "Sin identificar" and es_blacklist(empresa_old):
            ev["empresa"] = "Sin identificar"
            relaboreados += 1
        # Empleados: si era el default 1, intentar recalcular del comentario
        if ev.get("empleados") == 1:
            new_emp = extraer_empleados(coment, "")
            if new_emp > 1:
                ev["empleados"] = new_emp
    log(f"Eventos auto-rss del bedrock re-procesados: {relaboreados}")

    todos_auto = deduplicar(canonicalizar(eventos_auto_previos + nuevos))

    # Filtrado de fechas final
    hoy_iso = date.today().isoformat()
    manana_iso = (date.today() + timedelta(days=2)).isoformat()
    def _ok_fecha(e):
        f = e.get("fecha", "")
        return isinstance(f, str) and len(f) >= 10 and FECHA_MINIMA.isoformat() <= f <= manana_iso
    antes = len(todos_auto)
    todos_auto = [e for e in todos_auto if _ok_fecha(e)]
    if antes != len(todos_auto):
        log(f"Filtrado fechas: removidos {antes - len(todos_auto)}")

    # 4. Output
    todos = eventos_manual + todos_auto
    desde = (date.today() - timedelta(days=7)).isoformat()
    ultimos_7 = [e for e in todos
                 if e.get("fuente") in AUTO_FUENTES and e.get("fecha", "") >= desde]
    out = {
        "updated":    datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "nuevos_hoy": len(nuevos),
        "eventos":    todos,
        "nuevos":     ultimos_7,
    }

    if args.dry_run:
        log("DRY RUN - no se guarda")
        log(f"Stats: total={len(todos)} | manual={len(eventos_manual)} | "
            f"auto={len(todos_auto)} | nuevos_hoy={len(nuevos)} | ultimos_7={len(ultimos_7)}")
        return 0

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=None)
    log(f"Output: {OUT_PATH} ({OUT_PATH.stat().st_size // 1024} KB)", "ok")
    log(f"Total: {len(todos)} | nuevos hoy: {len(nuevos)} | ultimos 7d: {len(ultimos_7)}", "ok")

    # 5. Persistir bedrock actualizado (solo eventos auto-rss; manual queda intacto)
    bedrock["eventos"] = todos
    with open(HISTORICAL, "w", encoding="utf-8") as f:
        json.dump(bedrock, f, ensure_ascii=False, indent=None)
    log(f"Bedrock actualizado: {len(todos)} eventos", "ok")

    return 0


if __name__ == "__main__":
    sys.exit(main())
