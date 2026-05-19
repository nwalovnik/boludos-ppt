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
import io
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
import xlrd  # noqa: E402  (.xls 97-2003 lectura)

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

# IDs INDEC (mismas que el HTML original + nuevas)
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
    # === nuevos ===
    "sal_priv_reg":  "149.1_SOR_PRIADO_OCTU_0_25",   # Índice salarios privado registrado
    "sal_priv_no":   "149.1_SOR_PRIADO_OCTU_0_28",   # Índice salarios privado no registrado
    "sal_pub":       "149.1_SOR_PUBICO_OCTU_0_14",   # Índice salarios público
    "empleo_sipa":   "151.1_TL_ESTADAD_2012_M_20",   # Total trabajadores SIPA (con est) — en miles
    "empleo_priv":   "151.1_AARIADODAD_2012_M_31",   # Asalariados sector privado (con est) — en miles
    "turismo_rec":   "322.3_TURISMO_REIVO__17",      # Turismo receptivo (Ezeiza+Aeroparque)
    "turismo_em":    "322.3_TURISMO_EMIVO__15",      # Turismo emisivo
}

# UCI sectorial (INDEC publica por sector; nivel general "tot" sale del PDF y lo dejamos del histórico)
UCI_SECTORES = {
    "text":   "31.3_UPT_2004_M_23",    # textiles
    "quim":   "31.3_USPQ_2004_M_34",   # químicos
    "metal":  "31.3_UIMB_2004_M_33",   # metales básicos
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
# Procesadores adicionales: UCI, Salarios, Empleo, Turismo, MERVAL
# ──────────────────────────────────────────────────────────────────────────────
def update_uci_sectorial(D: dict) -> int:
    """UCI por sector (textiles, químicos, metales básicos). Nivel general 'tot' del histórico."""
    arr = D.setdefault("uci", [])
    nuevos = 0
    # Mapeo nuestro key → series INDEC
    rows_por_key = {}
    for key, series_id in UCI_SECTORES.items():
        try:
            data = fetch_indec(series_id, last=36)
            for row in data:
                if not row[0] or row[1] is None:
                    continue
                f = ym(row[0])
                rows_por_key.setdefault(f, {})[key] = round(float(row[1]), 1)
        except Exception as e:
            log(f"UCI {key} falló: {e}", "warn")

    for f, vals in rows_por_key.items():
        rec = next((x for x in arr if x.get("f") == f), None)
        if rec is None:
            new = {"f": f, **vals}
            arr.append(new)
            nuevos += 1
        else:
            rec.update(vals)
    arr.sort(key=lambda x: x["f"])
    if rows_por_key:
        ult = max(rows_por_key.keys())
        log(f"UCI sectorial: actualizado (último {ult})", "ok")
    return nuevos

def update_salarios(D: dict) -> int:
    """Salarios privado registrado / no registrado / público. Schema: {f, is_r, real_pub, real_priv}."""
    arr = D.setdefault("salarios", [])
    nuevos = 0
    try:
        ids = [INDEC_SERIES["sal_priv_reg"], INDEC_SERIES["sal_priv_no"], INDEC_SERIES["sal_pub"]]
        data = fetch_indec(ids, last=24)
        for row in data:
            if not row[0]:
                continue
            f = ym(row[0])
            is_r = row[1]      # privado registrado (lo usamos como índice principal)
            priv_no = row[2]
            pub = row[3]
            if is_r is None and priv_no is None and pub is None:
                continue
            rec = next((x for x in arr if x.get("f") == f), None)
            new = {"f": f}
            if is_r is not None:
                new["is_r"] = round(float(is_r), 2)
            if priv_no is not None:
                new["real_priv"] = round(float(priv_no), 2)
            if pub is not None:
                new["real_pub"] = round(float(pub), 2)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update(new)
        arr.sort(key=lambda x: x["f"])
        if data:
            log(f"Salarios: actualizado (último {data[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"Salarios falló: {e}", "warn")
    return nuevos

def update_empleo(D: dict) -> int:
    """Empleo SIPA total + privado. Schema histórico: {f, tot, priv} (valores en unidades, no miles).

    IMPORTANTE: INDEC devuelve estas series en MILES de trabajadores. El bedrock está en
    UNIDADES (ej. tot=10982740 = 10.9M). Multiplicamos por 1000 para mantener consistencia.
    """
    arr = D.setdefault("trabajo", [])
    nuevos = 0
    try:
        ids = [INDEC_SERIES["empleo_sipa"], INDEC_SERIES["empleo_priv"]]
        data = fetch_indec(ids, last=48)
        for row in data:
            if not row[0]:
                continue
            f = ym(row[0])
            tot_miles = row[1]
            priv_miles = row[2]
            if tot_miles is None and priv_miles is None:
                continue
            rec = next((x for x in arr if x.get("f") == f), None)
            new = {"f": f}
            if tot_miles is not None:
                new["tot"] = int(round(float(tot_miles) * 1000))  # miles → unidades
            if priv_miles is not None:
                new["priv"] = int(round(float(priv_miles) * 1000))
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update({k: v for k, v in new.items() if k != "f"})
        arr.sort(key=lambda x: x["f"])
        # Limpiar registro corrupto histórico "2019-23" (era ruido del baked) y outliers
        arr[:] = [r for r in arr
                  if r.get("f", "").startswith("20")
                  and "-" in r["f"]
                  and len(r["f"]) == 7
                  and r.get("tot", 0) > 1000000]  # filtra el 2012-01 raro con tot=4.8M
        if data:
            log(f"Empleo SIPA: actualizado (último {data[-1][0][:7]}, tot+priv en unidades)", "ok")
    except Exception as e:
        log(f"Empleo SIPA falló: {e}", "warn")
    return nuevos

def update_turismo(D: dict) -> int:
    """Turismo emisivo + receptivo. Schema: {f, rec, em, sal}."""
    arr = D.setdefault("turismo", [])
    nuevos = 0
    try:
        ids = [INDEC_SERIES["turismo_rec"], INDEC_SERIES["turismo_em"]]
        data = fetch_indec(ids, last=36)
        for row in data:
            if not row[0]:
                continue
            f = ym(row[0])
            rec_v, em_v = row[1], row[2]
            if rec_v is None and em_v is None:
                continue
            rec_existing = next((x for x in arr if x.get("f") == f), None)
            new = {"f": f}
            if rec_v is not None:
                new["rec"] = int(round(float(rec_v)))
            if em_v is not None:
                new["em"] = int(round(float(em_v)))
            if "rec" in new and "em" in new:
                new["sal"] = new["rec"] - new["em"]
            if rec_existing is None:
                arr.append(new)
                nuevos += 1
            else:
                rec_existing.update(new)
        arr.sort(key=lambda x: x["f"])
        if data:
            log(f"Turismo: actualizado (último {data[-1][0][:7]})", "ok")
    except Exception as e:
        log(f"Turismo falló: {e}", "warn")
    return nuevos

def update_daily_series(D: dict) -> int:
    """Series diarias (DD): rrii_d, tc_d, embi_d, merval_d. Las KPIs del HTML leen de estas."""
    nuevos = 0

    # === Reservas diaria (BCRA Var 1) — últimos ~120 días ===
    try:
        rows = fetch_bcra(BCRA_VARS["reservas"], limit=180)
        arr = D.setdefault("rrii_d", [])
        existing = {r["f"]: i for i, r in enumerate(arr) if "f" in r}
        for r in rows:
            f = r["fecha"]
            v = round(float(r["valor"]))
            if f in existing:
                arr[existing[f]]["v"] = v
            else:
                arr.append({"f": f, "v": v})
                existing[f] = len(arr) - 1
        # Trim a últimos 180 dias
        arr.sort(key=lambda x: x["f"])
        D["rrii_d"] = arr[-180:]
        if rows:
            log(f"Reservas diaria: último {rows[-1]['fecha']} = USD {rows[-1]['valor']:.0f}M", "ok")
    except Exception as e:
        log(f"Reservas diaria falló: {e}", "warn")

    # === TC diaria (BCRA Var 4 = minorista) ===
    try:
        rows = fetch_bcra(BCRA_VARS["tc_minor"], limit=180)
        arr = D.setdefault("tc_d", [])
        existing = {r["f"]: i for i, r in enumerate(arr) if "f" in r}
        for r in rows:
            f = r["fecha"]
            v = round(float(r["valor"]), 2)
            if f in existing:
                arr[existing[f]]["of"] = v
            else:
                arr.append({"f": f, "of": v, "mep": None, "ccl": None, "blue": None})
                existing[f] = len(arr) - 1
        arr.sort(key=lambda x: x["f"])

        # Enriquecer con MEP / CCL / Blue desde ArgentinaDatos cotizaciones
        cotiz = fetch_cotizaciones()
        for tipo_key, dd_field in (("bolsa", "mep"), ("contadoconliqui", "ccl"), ("blue", "blue")):
            for x in cotiz.get(tipo_key, []):
                if not x.get("fecha"):
                    continue
                f = x["fecha"]
                v = x.get("venta") or x.get("compra")
                if v is None:
                    continue
                if f in existing:
                    arr[existing[f]][dd_field] = round(float(v), 2)
        # Bluelytics como respaldo del blue de hoy
        bl = fetch_bluelytics()
        if bl:
            f_today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
            rec = next((x for x in arr if x.get("f") == f_today), None)
            if rec is None:
                rec = {"f": f_today, "of": None, "mep": None, "ccl": None, "blue": None}
                arr.append(rec)
            rec["blue"] = rec.get("blue") or round(float(bl["blue"]["value_sell"]), 2)
            rec["of"]   = rec.get("of")   or round(float(bl["oficial"]["value_sell"]), 2)
            arr.sort(key=lambda x: x["f"])
        D["tc_d"] = arr[-180:]
        log(f"TC diaria: último {arr[-1]['f']} | of={arr[-1].get('of')} blue={arr[-1].get('blue')}", "ok")
    except Exception as e:
        log(f"TC diaria falló: {e}", "warn")

    # === EMBI diaria (ArgentinaDatos riesgo-pais) — últimos 365 días ===
    try:
        rp = fetch_embi_historico()
        arr = D.setdefault("embi_d", [])
        existing = {r["f"]: i for i, r in enumerate(arr) if "f" in r}
        for x in rp:
            if not x.get("fecha"):
                continue
            v = x.get("valor")
            if v is None:
                continue
            f = x["fecha"][:10]
            if f in existing:
                arr[existing[f]]["v"] = v
            else:
                arr.append({"f": f, "v": v})
                existing[f] = len(arr) - 1
        ult = fetch_embi_ultimo()
        if ult and ult.get("valor"):
            f = ult["fecha"][:10]
            if f in existing:
                arr[existing[f]]["v"] = ult["valor"]
            else:
                arr.append({"f": f, "v": ult["valor"]})
        arr.sort(key=lambda x: x["f"])
        D["embi_d"] = arr[-365:]
        log(f"EMBI diaria: último {D['embi_d'][-1]['f']} = {D['embi_d'][-1]['v']} bps", "ok")
    except Exception as e:
        log(f"EMBI diaria falló: {e}", "warn")

    # === MERVAL diaria (Yahoo Finance) — últimos 365 días ===
    try:
        url = "https://query1.finance.yahoo.com/v8/finance/chart/%5EMERV"
        data = get_json(url, params={"interval": "1d", "range": "2y"},
                        headers={**HEADERS, "User-Agent": "Mozilla/5.0"})
        result = data["chart"]["result"][0]
        ts = result["timestamp"]
        closes = result["indicators"]["quote"][0]["close"]
        arr = D.setdefault("merval_d", [])
        existing = {r["f"]: i for i, r in enumerate(arr) if "f" in r}
        for t, c in zip(ts, closes):
            if c is None:
                continue
            f = datetime.fromtimestamp(t, tz=timezone.utc).strftime("%Y-%m-%d")
            v = round(float(c) / 1000, 2)  # mismas unidades que el bedrock (miles)
            if f in existing:
                arr[existing[f]]["v"] = v
            else:
                arr.append({"f": f, "v": v})
                existing[f] = len(arr) - 1
        arr.sort(key=lambda x: x["f"])
        D["merval_d"] = arr[-365:]
        if D["merval_d"]:
            log(f"MERVAL diaria: último {D['merval_d'][-1]['f']} = {D['merval_d'][-1]['v']}K", "ok")
    except Exception as e:
        log(f"MERVAL diaria falló: {e}", "warn")

    return nuevos

def update_merval(D: dict) -> int:
    """MERVAL via Yahoo Finance directo (sin proxy CORS porque corremos en GitHub Actions)."""
    arr = D.setdefault("merval", [])
    nuevos = 0
    try:
        url = "https://query1.finance.yahoo.com/v8/finance/chart/%5EMERV"
        data = get_json(url, params={"interval": "1d", "range": "5y"},
                        headers={**HEADERS, "User-Agent": "Mozilla/5.0"})
        result = data["chart"]["result"][0]
        ts = result["timestamp"]
        closes = result["indicators"]["quote"][0]["close"]
        # Agrupar por mes, último día
        by_month = {}
        for t, c in zip(ts, closes):
            if c is None:
                continue
            fecha = datetime.fromtimestamp(t, tz=timezone.utc).strftime("%Y-%m")
            by_month[fecha] = c  # se sobreescribe, queda el último
        for f, c in by_month.items():
            v = round(float(c) / 1000, 1)  # convertir a "miles de puntos" como el histórico
            rec = next((x for x in arr if x.get("f") == f), None)
            if rec is None:
                arr.append({"f": f, "v": v})
                nuevos += 1
            else:
                rec["v"] = v
        arr.sort(key=lambda x: x["f"])
        if by_month:
            ult = max(by_month.keys())
            log(f"MERVAL Yahoo: actualizado (último {ult} = {round(by_month[ult]/1000, 1)}K)", "ok")
    except Exception as e:
        log(f"MERVAL Yahoo falló: {e}", "warn")
    return nuevos

# ──────────────────────────────────────────────────────────────────────────────
# INDEC XLS oficiales (cuando la API series está atrasada vs los Excel publicados)
# ──────────────────────────────────────────────────────────────────────────────
MONTHS_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
INDEC_XLS_BASE = "https://www.indec.gob.ar/ftp/cuadros/economia"

def download_xls(url: str) -> xlrd.book.Book:
    """Descarga un .xls de INDEC y devuelve el workbook."""
    r = requests.get(url, timeout=TIMEOUT, headers={**HEADERS, "User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    return xlrd.open_workbook(file_contents=r.content)

def _month_num(cell_str: str) -> int | None:
    """'Enero' / 'Enero*' → 1, 'Febrero' → 2, etc."""
    s = str(cell_str).strip().rstrip("*").strip()
    return (MONTHS_ES.index(s) + 1) if s in MONTHS_ES else None

def _safe_float(v) -> float | None:
    try:
        if v == "" or v == "///" or v is None:
            return None
        return float(v)
    except (ValueError, TypeError):
        return None

def parse_emae_xls(wb: xlrd.book.Book) -> list[dict]:
    """sh_emae_mensual_base2004.xls → [{f, orig, via, dest, vm}]"""
    sh = wb.sheet_by_name("EMAE")
    out, current_year = [], None
    for r in range(sh.nrows):
        y_cell = sh.cell_value(r, 0)
        m_cell = sh.cell_value(r, 1)
        # Año: float (2004.0) o str ("2026*")
        if isinstance(y_cell, float) and y_cell > 1990:
            current_year = int(y_cell)
        elif isinstance(y_cell, str):
            s = y_cell.strip().rstrip("*").strip()
            if s.isdigit() and 1990 < int(s) < 2100:
                current_year = int(s)
        if not current_year:
            continue
        month_num = _month_num(m_cell) if m_cell else None
        if not month_num:
            continue
        orig = _safe_float(sh.cell_value(r, 2))
        via  = _safe_float(sh.cell_value(r, 3))
        dest = _safe_float(sh.cell_value(r, 4))
        vm   = _safe_float(sh.cell_value(r, 5))
        if orig is None:
            continue
        out.append({
            "f": f"{current_year:04d}-{month_num:02d}",
            "orig": round(orig, 2),
            "via":  round(via, 2)  if via  is not None else None,
            "dest": round(dest, 2) if dest is not None else None,
            "vm":   round(vm, 2)   if vm   is not None else None,
        })
    return out

def parse_ipi_xls(wb: xlrd.book.Book) -> list[dict]:
    """sh_ipi_manufacturero_YYYY.xls (Cuadro 1) → [{f, orig, via}]"""
    try:
        sh = wb.sheet_by_name("Cuadro 1")
    except xlrd.biffh.XLRDError:
        sh = wb.sheet_by_index(1)
    out, current_year = [], None
    for r in range(sh.nrows):
        y_cell = sh.cell_value(r, 1)
        m_cell = sh.cell_value(r, 2)
        if isinstance(y_cell, float) and y_cell > 1990:
            current_year = int(y_cell)
        elif isinstance(y_cell, str):
            s = y_cell.strip().rstrip("*").strip()
            if s.isdigit() and 1990 < int(s) < 2100:
                current_year = int(s)
        if not current_year:
            continue
        month_num = _month_num(m_cell) if m_cell else None
        if not month_num:
            continue
        orig = _safe_float(sh.cell_value(r, 3))
        via  = _safe_float(sh.cell_value(r, 4))
        if orig is None:
            continue
        out.append({
            "f": f"{current_year:04d}-{month_num:02d}",
            "orig": round(orig, 2),
            "via":  round(via, 2) if via is not None else None,
        })
    return out

# Mapeo columnas UCI → keys de nuestro schema
UCI_COL_MAP = {
    1: "tot",     # Nivel general
    2: "alim",    # Productos alimenticios
    3: "tabaco",  # Productos del tabaco
    4: "text",    # Productos textiles
    5: "papel",
    6: "impr",    # Edición/impresión
    7: "refin",   # Refinación
    8: "quim",    # Sustancias químicas
    9: "caucho",
    10: "min",    # Productos minerales
    11: "metal",  # Industrias metálicas básicas
    12: "auto",   # Industria automotriz
    13: "metalm", # Metalmecánica
}

def parse_uci_xls(wb: xlrd.book.Book) -> list[dict]:
    """sh_capacidad_MM_YY.xls → [{f, tot, alim, ..., metalm}]"""
    sh = wb.sheet_by_index(0)
    out, current_year = [], None
    for r in range(sh.nrows):
        col0 = sh.cell_value(r, 0)
        if isinstance(col0, str):
            s = col0.strip()
            # "Año 2026" o "Año 2025*"
            if s.startswith("Año "):
                ystr = s[4:].strip().rstrip("*").strip()
                if ystr.isdigit():
                    current_year = int(ystr)
                continue
            month_num = _month_num(s)
            if not month_num or not current_year:
                continue
            rec = {"f": f"{current_year:04d}-{month_num:02d}"}
            for col, key in UCI_COL_MAP.items():
                v = _safe_float(sh.cell_value(r, col))
                if v is not None:
                    rec[key] = round(v, 1)
            if "tot" in rec:
                out.append(rec)
    return out

def find_uci_xls_url() -> str | None:
    """sh_capacidad_MM_YY.xls — prueba del mes actual hacia atrás (máx 6 meses)."""
    today = datetime.now(timezone.utc)
    for delta in range(0, 7):
        y, m = today.year, today.month - delta
        while m <= 0:
            m += 12
            y -= 1
        url = f"{INDEC_XLS_BASE}/sh_capacidad_{m:02d}_{y % 100:02d}.xls"
        try:
            r = requests.head(url, timeout=10, headers={**HEADERS, "User-Agent": "Mozilla/5.0"})
            if r.status_code == 200 and "excel" in r.headers.get("Content-Type", "").lower():
                return url
        except Exception:
            continue
    return None

# Mapeo EMAE actividad: cols del XLS → keys del schema bedrock D.emae[].s
EMAE_S_COLS = {
    2:  "Agro",      # A - Agricultura
    3:  "Pesca",     # B - Pesca
    4:  "Mineria",   # C - Minería
    5:  "Industria", # D - Industria manufacturera
    6:  "EGA",       # E - Electricidad/Gas/Agua
    7:  "Const",     # F - Construcción
    8:  "Comer",     # G - Comercio
    9:  "Hoteles",   # H - Hoteles y restaurantes
    10: "Transp",    # I - Transporte y comunicaciones
    11: "Financ",    # J - Intermediación financiera
    12: "Inmob",     # K - Actividades inmobiliarias
    13: "AdmPub",    # L - Administración pública
    14: "Ensen",     # M - Enseñanza
    15: "Salud_s",   # N - Salud
}

def parse_emae_actividad_xls(wb: xlrd.book.Book) -> list[dict]:
    """sh_emae_actividad_base2004.xls (hoja 'Tabla Letras') → [{f, s:{Agro,Pesca,...}}]"""
    sh = wb.sheet_by_name("Tabla Letras")
    out, current_year = [], None
    for r in range(sh.nrows):
        y_cell = sh.cell_value(r, 0)
        m_cell = sh.cell_value(r, 1)
        if isinstance(y_cell, float) and y_cell > 1990:
            current_year = int(y_cell)
        elif isinstance(y_cell, str):
            s = y_cell.strip().rstrip("*").strip()
            if s.isdigit() and 1990 < int(s) < 2100:
                current_year = int(s)
        if not current_year:
            continue
        month_num = _month_num(m_cell) if m_cell else None
        if not month_num:
            continue
        sectores = {}
        for col, name in EMAE_S_COLS.items():
            v = _safe_float(sh.cell_value(r, col))
            if v is not None:
                sectores[name] = round(v, 2)
        if sectores:
            out.append({"f": f"{current_year:04d}-{month_num:02d}", "s": sectores})
    return out

# Mapeo IPI Cuadro 5: cols del XLS (impares = nivel del sector) → keys de D.ipi[].r
IPI_R_COLS = {
    5:  "Alimentos",
    7:  "Tabaco",
    9:  "Textiles",
    11: "Vestim",
    13: "Madera",
    15: "Petroleo",
    17: "Quimica",
    19: "Caucho",
    21: "Minerales",
    23: "MetalBase",
    25: "ProdMetal",
    27: "Maquinaria",
    29: "OtrosEq",
    31: "Automotriz",
    33: "OtroTransp",
    35: "Muebles",
}

def parse_ipi_sectores_xls(wb: xlrd.book.Book) -> list[dict]:
    """sh_ipi_manufacturero_YYYY.xls (Cuadro 5) → [{f, r:{Alimentos,Tabaco,...}}]"""
    try:
        sh = wb.sheet_by_name("Cuadro 5")
    except xlrd.biffh.XLRDError:
        return []
    out, current_year = [], None
    for r in range(sh.nrows):
        y_cell = sh.cell_value(r, 1)
        m_cell = sh.cell_value(r, 2)
        if isinstance(y_cell, float) and y_cell > 1990:
            current_year = int(y_cell)
        elif isinstance(y_cell, str):
            s = y_cell.strip().rstrip("*").strip()
            if s.isdigit() and 1990 < int(s) < 2100:
                current_year = int(s)
        if not current_year:
            continue
        month_num = _month_num(m_cell) if m_cell else None
        if not month_num:
            continue
        ramas = {}
        for col, name in IPI_R_COLS.items():
            v = _safe_float(sh.cell_value(r, col))
            if v is not None:
                ramas[name] = round(v, 2)
        if ramas:
            out.append({"f": f"{current_year:04d}-{month_num:02d}", "r": ramas})
    return out

def update_emae_sectores_xls(D: dict) -> int:
    """Completa D.emae[].s desde el XLS de EMAE actividad."""
    arr = D.setdefault("emae", [])
    nuevos = 0
    try:
        wb = download_xls(f"{INDEC_XLS_BASE}/sh_emae_actividad_base2004.xls")
        rows = parse_emae_actividad_xls(wb)
        for new in rows:
            rec = next((x for x in arr if x.get("f") == new["f"]), None)
            if rec is None:
                # No deberíamos llegar acá si update_emae_xls corrió antes, pero por las dudas
                arr.append(new)
                nuevos += 1
            else:
                rec["s"] = new["s"]
        if rows:
            log(f"EMAE sectores (XLS oficial): actualizado (último {rows[-1]['f']})", "ok")
    except Exception as e:
        log(f"EMAE sectores XLS falló: {e}", "warn")
    return nuevos

def update_ipi_sectores_xls(D: dict) -> int:
    """Completa D.ipi[].r desde el Cuadro 5 del XLS de IPI."""
    arr = D.setdefault("ipi", [])
    nuevos = 0
    today = datetime.now(timezone.utc)
    wb = None
    for year in (today.year, today.year - 1):
        try:
            wb = download_xls(f"{INDEC_XLS_BASE}/sh_ipi_manufacturero_{year}.xls")
            break
        except Exception:
            continue
    if wb is None:
        return 0
    try:
        rows = parse_ipi_sectores_xls(wb)
        for new in rows:
            rec = next((x for x in arr if x.get("f") == new["f"]), None)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec["r"] = new["r"]
        if rows:
            log(f"IPI sectores (XLS Cuadro 5): actualizado (último {rows[-1]['f']})", "ok")
    except Exception as e:
        log(f"IPI sectores XLS falló: {e}", "warn")
    return nuevos

def update_emae_xls(D: dict) -> int:
    """EMAE oficial desde el XLS de INDEC (más actualizado que la API series)."""
    arr = D.setdefault("emae", [])
    nuevos = 0
    try:
        wb = download_xls(f"{INDEC_XLS_BASE}/sh_emae_mensual_base2004.xls")
        rows = parse_emae_xls(wb)
        for new in rows:
            rec = next((x for x in arr if x.get("f") == new["f"]), None)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                # No pisar la subdivisión sectorial 's' si la tiene del bedrock
                for k, v in new.items():
                    if v is not None and k != "s":
                        rec[k] = v
        arr.sort(key=lambda x: x.get("f", ""))
        if rows:
            log(f"EMAE (XLS oficial): actualizado (último {rows[-1]['f']})", "ok")
    except Exception as e:
        log(f"EMAE XLS falló: {e}", "warn")
    return nuevos

def update_ipi_xls(D: dict) -> int:
    """IPI oficial desde el XLS de INDEC. Prueba YYYY actual y anterior si 404."""
    arr = D.setdefault("ipi", [])
    nuevos = 0
    today = datetime.now(timezone.utc)
    wb = None
    for year in (today.year, today.year - 1):
        url = f"{INDEC_XLS_BASE}/sh_ipi_manufacturero_{year}.xls"
        try:
            wb = download_xls(url)
            break
        except Exception:
            continue
    if wb is None:
        log("IPI XLS no encontrado", "warn")
        return 0
    try:
        rows = parse_ipi_xls(wb)
        for new in rows:
            rec = next((x for x in arr if x.get("f") == new["f"]), None)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                for k, v in new.items():
                    if v is not None:
                        rec[k] = v
        arr.sort(key=lambda x: x.get("f", ""))
        # Recalcular vm (mensual sobre dest) si tenemos serie desestacionalizada
        for i, r in enumerate(arr):
            if i > 0 and r.get("dest") and arr[i-1].get("dest"):
                r["vm"] = round((r["dest"] / arr[i-1]["dest"] - 1) * 100, 2)
        if rows:
            log(f"IPI (XLS oficial): actualizado (último {rows[-1]['f']})", "ok")
    except Exception as e:
        log(f"IPI XLS parse falló: {e}", "warn")
    return nuevos

def update_uci_xls(D: dict) -> int:
    """UCI Nivel General + sectorial desde el XLS oficial de INDEC."""
    arr = D.setdefault("uci", [])
    nuevos = 0
    url = find_uci_xls_url()
    if not url:
        log("UCI XLS no encontrado (probados últimos 6 meses)", "warn")
        return 0
    try:
        wb = download_xls(url)
        rows = parse_uci_xls(wb)
        for new in rows:
            rec = next((x for x in arr if x.get("f") == new["f"]), None)
            if rec is None:
                arr.append(new)
                nuevos += 1
            else:
                rec.update({k: v for k, v in new.items() if v is not None})
        arr.sort(key=lambda x: x.get("f", ""))
        if rows:
            log(f"UCI (XLS oficial {url.split('/')[-1]}): actualizado (último {rows[-1]['f']})", "ok")
    except Exception as e:
        log(f"UCI XLS parse falló: {e}", "warn")
    return nuevos

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
    # EMAE / IPI / UCI: primero API series (rápido), después XLS oficial (más fresco — pisa la API)
    update_emae(D)
    update_emae_xls(D)
    update_emae_sectores_xls(D)
    update_ipi_isac(D)
    update_ipi_xls(D)
    update_ipi_sectores_xls(D)
    update_uci_xls(D)
    update_bcra_monthly(D)
    update_embi(D)
    update_simple_monthly(D, "ripte", INDEC_SERIES["ripte"], "nom", "RIPTE")
    update_simple_monthly(D, "rec",   INDEC_SERIES["rec"],   "tot", "Recaudación")
    update_uci_sectorial(D)
    update_salarios(D)
    update_empleo(D)
    update_turismo(D)
    update_merval(D)
    update_daily_series(D)

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
