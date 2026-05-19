"""
Scraper diario de indicadores macro argentinos.

Carga el bedrock histórico (historical.json) y sobreescribe los registros recientes
con datos frescos de las APIs públicas. Output: data.json (en raíz del repo).

Fuentes:
  - INDEC (apis.datos.gob.ar/series)        → IPC, EMAE, IPI, ISAC, ICA, EPH, RIPTE, recaudación, BM
  - BCRA  (api.bcra.gob.ar/estadisticas/v4) → Reservas, TC oficial/mayorista, BM diaria, var IPC
  - ArgentinaDatos                          → EMBI riesgo país, cotizaciones (blue/mep/ccl)
  - Bluelytics                              → TC blue actual (respaldo)

Uso:
    python scrape.py            # corre todo, guarda data.json
    python scrape.py --dry-run  # corre todo pero no guarda

Diseño: idempotente, no rompe si una fuente falla (loggea y sigue).
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

# ──────────────────────────────────────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────────────────────────────────────
SCRAPER_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRAPER_DIR.parent
HISTORICAL = SCRAPER_DIR / "historical.json"
OUT_PATH = REPO_ROOT / "data.json"

# ──────────────────────────────────────────────────────────────────────────────
# Endpoints
# ──────────────────────────────────────────────────────────────────────────────
INDEC_API = "https://apis.datos.gob.ar/series/api/series/"
BCRA_API = "https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/"
ARG_DATOS = "https://api.argentinadatos.com/v1"
BLUELYTICS = "https://api.bluelytics.com.ar/v2/latest"

# IDs INDEC (mismas que el HTML original)
INDEC_SERIES = {
    "ipc":       "148.3_INIVELNAL_DICI_M_26",
    "emae_orig": "143.3_NO_PR_2004_A_21",
    "emae_dest": "143.3_NO_PR_2004_A_31",
    "ipi_orig":  "453.1_SERIE_ORIGNAL_0_0_14_46",
    "ipi_dest":  "453.1_SERIE_DESEADA_0_0_24_58",
    "isac":      "33.2_ISAC_SIN_EDAD_0_M_23_56",
    "ica_expo":  "74.3_IET_0_M_16",
    "ica_pp":    "74.3_IEPP_0_M_35",
    "ica_moi":   "74.3_IEMOI_0_M_46",
    "ica_cye":   "74.3_IECE_0_M_35",
    "eph_des":   "45.2_ECTDT_0_T_33",
    "ripte":     "158.1_REPTE_0_0_5",
    "rec":       "172.3_TL_RECAION_M_0_0_17",
    "bm":        "90.1_BMT_0_0_20",
}

# Variables BCRA usadas
BCRA_VARS = {
    "reservas":   1,    # USD millones, diaria
    "tc_minor":   4,    # TC minorista vendedor, diaria
    "tc_mayor":   5,    # TC mayorista de referencia, diaria
    "bm_diaria":  15,   # Base monetaria diaria (M$)
    "ipc_vm":     27,   # Variación mensual IPC (este SÍ tiene abril 2026 cuando INDEC API aún no)
    "ipc_via":    28,   # Variación interanual IPC
}

TIMEOUT = 30
RETRIES = 3
HEADERS = {"Accept-Language": "es-AR", "User-Agent": "boludos-ppt-scraper/1.0"}

# ──────────────────────────────────────────────────────────────────────────────
# Helpers de logging
# ──────────────────────────────────────────────────────────────────────────────
def log(msg: str, level: str = "info") -> None:
    ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
    icon = {"info": "  ", "ok": "OK", "warn": "!!", "err": "XX"}.get(level, "  ")
    print(f"[{ts}] {icon} {msg}", flush=True)

# ──────────────────────────────────────────────────────────────────────────────
# HTTP helpers
# ──────────────────────────────────────────────────────────────────────────────
def get_json(url: str, params: dict | None = None, headers: dict | None = None) -> Any:
    """GET con retries y backoff exponencial."""
    last_err = None
    for attempt in range(RETRIES):
        try:
            r = requests.get(url, params=params, headers=headers or HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            last_err = e
            if attempt < RETRIES - 1:
                time.sleep(2 ** attempt)
    raise RuntimeError(f"GET {url} fallido tras {RETRIES} intentos: {last_err}")

# ──────────────────────────────────────────────────────────────────────────────
# Fetchers por fuente
# ──────────────────────────────────────────────────────────────────────────────
def fetch_indec(series_ids: list[str] | str, last: int = 60) -> list[list]:
    ids = ",".join(series_ids) if isinstance(series_ids, list) else series_ids
    data = get_json(INDEC_API, params={"ids": ids, "last": last, "format": "json"})
    rows = data.get("data", [])
    # API ya devuelve ASC (oldest first). Forzamos sort por las dudas.
    return sorted(rows, key=lambda r: r[0] or "")

def fetch_bcra(var_id: int, limit: int = 365) -> list[dict]:
    """Devuelve lista de {fecha, valor} ordenada cronológicamente."""
    data = get_json(f"{BCRA_API}{var_id}", params={"limit": limit})
    detalle = data.get("results", [{}])[0].get("detalle", [])
    return list(reversed(detalle))

def fetch_embi_historico() -> list[dict]:
    """Devuelve serie completa EMBI desde 1999."""
    return get_json(f"{ARG_DATOS}/finanzas/indices/riesgo-pais")

def fetch_embi_ultimo() -> dict | None:
    """La API a veces omite 'valor' en los registros recientes; /ultimo siempre lo trae."""
    try:
        return get_json(f"{ARG_DATOS}/finanzas/indices/riesgo-pais/ultimo")
    except Exception as e:
        log(f"EMBI /ultimo falló: {e}", "warn")
        return None

def fetch_cotizaciones() -> dict[str, list[dict]]:
    """Blue, MEP, CCL históricos de ArgentinaDatos."""
    out = {}
    for tipo in ("blue", "bolsa", "contadoconliqui"):
        try:
            data = get_json(f"{ARG_DATOS}/cotizaciones/dolares/{tipo}")
            out[tipo] = data
        except Exception as e:
            log(f"Cotización {tipo} falló: {e}", "warn")
            out[tipo] = []
    return out

def fetch_bluelytics() -> dict | None:
    try:
        return get_json(BLUELYTICS)
    except Exception as e:
        log(f"Bluelytics falló: {e}", "warn")
        return None

# ──────────────────────────────────────────────────────────────────────────────
# Upsert helpers
# ──────────────────────────────────────────────────────────────────────────────
def upsert_by_f(arr: list[dict], record: dict) -> bool:
    """Inserta o actualiza por clave 'f'. Devuelve True si fue insert."""
    f = record["f"]
    for i, r in enumerate(arr):
        if r.get("f") == f:
            # actualizar solo campos nuevos / cambiados; preservar lo que ya estaba
            for k, v in record.items():
                if v is not None:
                    r[k] = v
            r.pop("proj", None)  # si lo estamos actualizando con datos reales, ya no es proyección
            return False
    arr.append(record)
    arr.sort(key=lambda x: x.get("f", ""))
    return True

def ym(fecha: str) -> str:
    """YYYY-MM-DD → YYYY-MM"""
    return fecha[:7]

# ──────────────────────────────────────────────────────────────────────────────
# Procesadores por serie
# ──────────────────────────────────────────────────────────────────────────────
def update_ipc(D: dict) -> int:
    """IPC: usa INDEC para nivel n; recalcula vm y via."""
    arr = D.setdefault("ipc", [])
    nuevos = 0

    # 1) Traer nivel desde INDEC (datos oficiales con desglose si vienen)
    indec_last = None
    try:
        rows = fetch_indec(INDEC_SERIES["ipc"], last=24)
        for row in rows:
            if not row[0] or row[1] is None:
                continue
            f = ym(row[0])
            n = round(row[1], 4)
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "n": n})
                nuevos += 1
            else:
                # Si era proyección, ahora es real: limpiamos proj y reemplazamos n
                rec["n"] = n
                rec.pop("proj", None)
        if rows:
            indec_last = rows[-1][0][:7]
            log(f"IPC INDEC: {nuevos} períodos nuevos (último {indec_last})", "ok")
    except Exception as e:
        log(f"IPC INDEC API falló: {e}", "warn")

    # 2) Completar meses recientes que INDEC aún no publica vía BCRA Var 27 (vm) y 28 (via)
    #    La BCRA suele tener el dato 1 mes antes que el API datos.gob.ar
    try:
        vm_rows = fetch_bcra(BCRA_VARS["ipc_vm"], limit=24)
        via_rows = fetch_bcra(BCRA_VARS["ipc_via"], limit=24)
        via_by_month = {ym(v["fecha"]): float(v["valor"]) for v in via_rows if v.get("fecha")}
        # solo procesar meses POSTERIORES al último de INDEC, o que estén marcados como proj
        for vm_row in vm_rows[-6:]:
            fecha = vm_row.get("fecha", "")
            if not fecha:
                continue
            f = ym(fecha)
            vm = float(vm_row["valor"])
            via = via_by_month.get(f)
            rec = next((x for x in arr if x.get("f") == f), None)
            es_posterior_indec = indec_last is None or f > indec_last
            es_proj = rec is not None and rec.get("proj")
            if rec is None or es_proj:
                if not es_posterior_indec and not es_proj:
                    continue  # INDEC ya tiene este mes con datos reales, no tocamos
                # Necesitamos calcular n desde el mes anterior REAL (no proj)
                prev_real = None
                for x in reversed(arr):
                    if x.get("f", "") < f and x.get("n") and not x.get("proj"):
                        prev_real = x
                        break
                if prev_real:
                    n_calc = round(prev_real["n"] * (1 + vm / 100), 4)
                    if rec is None:
                        arr.append({"f": f, "n": n_calc, "vm": round(vm, 2),
                                    "via": round(via, 2) if via is not None else None})
                        nuevos += 1
                    else:
                        # reemplazar proyección con dato real BCRA
                        rec.clear()
                        rec.update({"f": f, "n": n_calc, "vm": round(vm, 2)})
                        if via is not None:
                            rec["via"] = round(via, 2)
                    log(f"IPC {f}: completado vía BCRA ({vm:.2f}% mensual, n={n_calc:.2f})", "ok")
    except Exception as e:
        log(f"IPC BCRA Var27/28 falló: {e}", "warn")

    # 3) Recalcular vm/via para todos los registros con n
    arr.sort(key=lambda x: x.get("f", ""))
    idx = {x["f"]: i for i, x in enumerate(arr) if "f" in x}
    for r in arr:
        if r.get("n") is None or r.get("proj"):
            continue
        f = r["f"]
        prev_f = prev_month(f)
        py_f = prev_year(f)
        if prev_f in idx:
            p = arr[idx[prev_f]]
            if p.get("n") and not p.get("proj"):
                r["vm"] = round((r["n"] / p["n"] - 1) * 100, 2)
        if py_f in idx:
            p12 = arr[idx[py_f]]
            if p12.get("n") and not p12.get("proj"):
                r["via"] = round((r["n"] / p12["n"] - 1) * 100, 2)

    return nuevos

def update_emae(D: dict) -> int:
    arr = D.setdefault("emae", [])
    nuevos = 0
    try:
        rows = fetch_indec([INDEC_SERIES["emae_orig"], INDEC_SERIES["emae_dest"]], last=36)
        for row in rows:
            if not row[0]:
                continue
            f = ym(row[0])
            orig = row[1]
            dest = row[2]
            if orig is None:
                continue
            rec = next((x for x in arr if x.get("f") == f), None)
            new = {
                "f": f,
                "orig": round(orig, 2),
                "dest": round(dest, 2) if dest is not None else None,
            }
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update({k: v for k, v in new.items() if v is not None})
        # recalcular via (interanual sobre orig) y vm (mensual sobre dest)
        arr.sort(key=lambda x: x.get("f", ""))
        idx = {x["f"]: i for i, x in enumerate(arr)}
        for i, r in enumerate(arr):
            py = prev_year(r["f"])
            if py in idx and arr[idx[py]].get("orig"):
                r["via"] = round((r["orig"] / arr[idx[py]]["orig"] - 1) * 100, 2)
            if i > 0 and r.get("dest") and arr[i - 1].get("dest"):
                r["vm"] = round((r["dest"] / arr[i - 1]["dest"] - 1) * 100, 2)
        log(f"EMAE: {nuevos} períodos nuevos (último {rows[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"EMAE falló: {e}", "warn")
    return nuevos

def update_ipi_isac(D: dict) -> int:
    nuevos = 0
    # IPI
    try:
        rows = fetch_indec([INDEC_SERIES["ipi_orig"], INDEC_SERIES["ipi_dest"]], last=36)
        arr = D.setdefault("ipi", [])
        for row in rows:
            if not row[0] or row[1] is None:
                continue
            f = ym(row[0])
            rec = next((x for x in arr if x.get("f") == f), None)
            new = {"f": f, "orig": round(row[1], 2)}
            if row[2] is not None:
                new["dest"] = round(row[2], 2)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update({k: v for k, v in new.items() if v is not None})
        # recalc via/vm
        recalc_via_vm(arr, key_orig="orig", key_dest="dest")
        log(f"IPI: actualizado (último {rows[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"IPI falló: {e}", "warn")
    # ISAC
    try:
        rows = fetch_indec(INDEC_SERIES["isac"], last=36)
        arr = D.setdefault("isac", [])
        for row in rows:
            if not row[0] or row[1] is None:
                continue
            f = ym(row[0])
            rec = next((x for x in arr if x.get("f") == f), None)
            new = {"f": f, "dest": round(row[1], 2)}
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update(new)
        arr.sort(key=lambda x: x["f"])
        for i, r in enumerate(arr):
            if i > 0 and r.get("dest") and arr[i - 1].get("dest"):
                r["vm"] = round((r["dest"] / arr[i - 1]["dest"] - 1) * 100, 2)
        log(f"ISAC: actualizado (último {rows[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"ISAC falló: {e}", "warn")
    return nuevos

def update_bcra_monthly(D: dict) -> int:
    """Reservas, TC, BM — todas mensuales, último día observado del mes."""
    nuevos = 0

    # Reservas (mensual)
    try:
        rows = fetch_bcra(BCRA_VARS["reservas"], limit=365)
        arr = D.setdefault("reservas", [])
        # Agrupar por mes, quedarse con el último día
        by_month = {}
        for r in rows:
            f = ym(r["fecha"])
            by_month[f] = r  # va sobreescribiendo, queda el último
        for f, r in by_month.items():
            v = round(float(r["valor"]))
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "v": v})
                nuevos += 1
            else:
                rec["v"] = v
        arr.sort(key=lambda x: x["f"])
        log(f"Reservas BCRA: último {rows[-1]['fecha']} = USD {rows[-1]['valor']:.0f}M", "ok")
    except Exception as e:
        log(f"Reservas BCRA falló: {e}", "warn")

    # TC Oficial (mensual)
    try:
        rows = fetch_bcra(BCRA_VARS["tc_minor"], limit=365)
        arr = D.setdefault("tc", [])
        by_month = {}
        for r in rows:
            by_month[ym(r["fecha"])] = r
        for f, r in by_month.items():
            v = round(float(r["valor"]), 2)
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "of": v})
            else:
                rec["of"] = v
        arr.sort(key=lambda x: x["f"])
        # Recalcular variaciones
        for i, r in enumerate(arr):
            if i > 0 and r.get("of") and arr[i - 1].get("of"):
                r["vm_of"] = round((r["of"] / arr[i - 1]["of"] - 1) * 100, 2)
            py = prev_year(r["f"])
            py_rec = next((x for x in arr if x.get("f") == py), None)
            if py_rec and py_rec.get("of"):
                r["via_of"] = round((r["of"] / py_rec["of"] - 1) * 100, 2)
        log(f"TC Oficial BCRA: último {rows[-1]['fecha']} = ${rows[-1]['valor']:.2f}", "ok")
    except Exception as e:
        log(f"TC Oficial BCRA falló: {e}", "warn")

    # Base monetaria (mensual)
    try:
        rows = fetch_bcra(BCRA_VARS["bm_diaria"], limit=365)
        arr = D.setdefault("bm", [])
        by_month = {}
        for r in rows:
            by_month[ym(r["fecha"])] = r
        for f, r in by_month.items():
            v = round(float(r["valor"]))
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "tot": v})
            else:
                rec["tot"] = v
        arr.sort(key=lambda x: x["f"])
        log(f"Base monetaria BCRA: último {rows[-1]['fecha']}", "ok")
    except Exception as e:
        log(f"BM BCRA falló: {e}", "warn")

    return nuevos

def update_embi(D: dict) -> int:
    arr = D.setdefault("embi", [])
    nuevos = 0
    try:
        rp = fetch_embi_historico()
        by_month = {}
        for x in rp:
            if not x.get("fecha"):
                continue
            v = x.get("valor")
            if v is None:
                continue
            by_month[ym(x["fecha"])] = v
        for f, v in by_month.items():
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "v": v})
                nuevos += 1
            else:
                rec["v"] = v
        arr.sort(key=lambda x: x["f"])
        # Completar último con /ultimo (la API a veces omite valor en los más recientes)
        ult = fetch_embi_ultimo()
        if ult and ult.get("valor"):
            f = ym(ult["fecha"])
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "v": ult["valor"]})
                nuevos += 1
            else:
                rec["v"] = ult["valor"]
            arr.sort(key=lambda x: x["f"])
            log(f"EMBI: último {ult['fecha']} = {ult['valor']} bps", "ok")
    except Exception as e:
        log(f"EMBI falló: {e}", "warn")
    return nuevos

def update_simple_monthly(D: dict, key: str, series_id: str, val_key: str, label: str) -> int:
    """Helper para series mensuales simples con un solo valor."""
    arr = D.setdefault(key, [])
    nuevos = 0
    try:
        rows = fetch_indec(series_id, last=36)
        for row in rows:
            if not row[0] or row[1] is None:
                continue
            f = ym(row[0])
            v = round(float(row[1]), 2)
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, val_key: v})
                nuevos += 1
            else:
                rec[val_key] = v
        arr.sort(key=lambda x: x["f"])
        log(f"{label}: actualizado (último {rows[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"{label} falló: {e}", "warn")
    return nuevos

# ──────────────────────────────────────────────────────────────────────────────
# Utilidades de fecha
# ──────────────────────────────────────────────────────────────────────────────
def prev_month(f: str) -> str:
    y, m = int(f[:4]), int(f[5:7])
    if m == 1:
        return f"{y - 1:04d}-12"
    return f"{y:04d}-{m - 1:02d}"

def prev_year(f: str) -> str:
    return f"{int(f[:4]) - 1:04d}-{f[5:7]}"

def recalc_via_vm(arr: list[dict], key_orig: str = "orig", key_dest: str = "dest") -> None:
    arr.sort(key=lambda x: x.get("f", ""))
    idx = {x["f"]: i for i, x in enumerate(arr) if "f" in x}
    for r in arr:
        py = prev_year(r["f"])
        if py in idx and arr[idx[py]].get(key_orig):
            try:
                r["via"] = round((r[key_orig] / arr[idx[py]][key_orig] - 1) * 100, 2)
            except (KeyError, TypeError):
                pass
        i = idx[r["f"]]
        if i > 0 and r.get(key_dest) and arr[i - 1].get(key_dest):
            try:
                r["vm"] = round((r[key_dest] / arr[i - 1][key_dest] - 1) * 100, 2)
            except (KeyError, TypeError):
                pass

# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="No escribe data.json")
    args = ap.parse_args()

    log(f"Cargando bedrock desde {HISTORICAL.name}")
    D = json.loads(HISTORICAL.read_text(encoding="utf-8"))
    log(f"Bedrock: {len(D)} series, {sum(len(v) for v in D.values() if isinstance(v, list))} registros")

    # Actualizar cada bloque (cada uno maneja sus propios errores)
    update_ipc(D)
    update_emae(D)
    update_ipi_isac(D)
    update_bcra_monthly(D)
    update_embi(D)
    update_simple_monthly(D, "ripte", INDEC_SERIES["ripte"], "nom", "RIPTE")
    update_simple_monthly(D, "rec",   INDEC_SERIES["rec"],   "tot", "Recaudación")

    # Metadata
    D["_meta"] = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "scraper_version": "1.0",
        "sources": ["INDEC", "BCRA", "ArgentinaDatos", "Bluelytics"],
    }

    if args.dry_run:
        log("DRY RUN — no escribo data.json")
        log(f"Sumario: {len(D) - 1} series (excluye _meta)")
    else:
        OUT_PATH.write_text(json.dumps(D, ensure_ascii=False, indent=None, separators=(",", ":")),
                            encoding="utf-8")
        size_kb = OUT_PATH.stat().st_size / 1024
        log(f"OK — data.json escrito ({size_kb:.1f} KB)", "ok")

    return 0

if __name__ == "__main__":
    sys.exit(main())
