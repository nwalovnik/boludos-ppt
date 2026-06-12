"""
Edita o elimina un evento de despidos por ID en el bedrock.

Uso:
    python edit_event.py --id ABC123 --action delete
    python edit_event.py --id ABC123 --action edit --patch '{"empleados": 250, "empresa": "Lear"}'

Despues de modificar el bedrock, hay que correr `scrape_despidos.py --skip-rss`
para regenerar despidos.json. El workflow dispatch_edit.yml hace ambos pasos.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SCRAPER_DIR = Path(__file__).resolve().parent
BEDROCK = SCRAPER_DIR / "historical_despidos.json"

# Campos que se permite editar desde la UI
CAMPOS_EDITABLES = {
    "empresa", "rubro", "empleados", "comentario", "provincia",
    "municipio", "estado", "sector", "cerro", "tipo", "fecha",
}


def normalizar_key(nombre: str) -> str:
    if not nombre:
        return ""
    n = nombre.lower().strip()
    repl = str.maketrans({"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u"})
    return n.translate(repl)


def event_id(ev: dict) -> str:
    base = "|".join([
        (ev.get("fecha") or "")[:10],
        normalizar_key(ev.get("empresa") or ""),
        (ev.get("url") or ev.get("comentario") or "")[:120],
    ])
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:12]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--id", required=True, help="ID del evento (hash de 12 chars)")
    parser.add_argument("--action", required=True, choices=["edit", "delete"])
    parser.add_argument("--patch", default="{}",
                        help="JSON con campos a actualizar (solo para edit)")
    args = parser.parse_args(argv)

    target_id = args.id.strip().lower()
    if len(target_id) != 12:
        print(f"ERROR: ID debe ser de 12 chars, recibi {len(target_id)!r}")
        return 1

    if not BEDROCK.exists():
        print(f"ERROR: bedrock no encontrado: {BEDROCK}")
        return 1
    bedrock = json.load(open(BEDROCK, "r", encoding="utf-8"))
    eventos = bedrock.get("eventos", [])

    # Buscar evento por id (regenerar id porque puede no estar persistido)
    idx_target = None
    for i, ev in enumerate(eventos):
        if (ev.get("id") or event_id(ev)) == target_id:
            idx_target = i
            break
    if idx_target is None:
        print(f"ERROR: evento con id={target_id} no encontrado en bedrock")
        return 2

    ev = eventos[idx_target]
    print(f"Evento target: {ev.get('fecha')} | {ev.get('empresa')} | "
          f"empleados={ev.get('empleados')} | fuente={ev.get('fuente')}")

    if args.action == "delete":
        eventos.pop(idx_target)
        print(f"Evento eliminado")
    else:
        try:
            patch = json.loads(args.patch)
        except json.JSONDecodeError as e:
            print(f"ERROR: --patch no es JSON valido: {e}")
            return 1
        if not isinstance(patch, dict):
            print(f"ERROR: --patch debe ser un objeto JSON")
            return 1
        cambios = {}
        for k, v in patch.items():
            if k not in CAMPOS_EDITABLES:
                print(f"  IGNORE campo no editable: {k}")
                continue
            # Cast tipos basicos
            if k == "empleados":
                # null = s/d (sin dato); int >= 0 = valor
                if v is None:
                    pass  # mantener como None
                else:
                    try:
                        v = int(v)
                    except (TypeError, ValueError):
                        print(f"  IGNORE empleados invalido: {v!r}")
                        continue
            elif k == "cerro":
                if isinstance(v, str):
                    v = v.lower() in ("si", "sí", "yes", "true", "1")
                else:
                    v = bool(v)
            cambios[k] = v
        if not cambios:
            print("ERROR: --patch no contiene campos validos")
            return 1
        ev.update(cambios)
        # Marcar como editado manualmente para preservar contra dedupe/reproceso
        ev["edited"] = True
        # Re-calcular id si cambiaron campos que lo afectan
        ev["id"] = event_id(ev)
        print(f"Evento editado: {cambios}")

    bedrock["eventos"] = eventos
    with open(BEDROCK, "w", encoding="utf-8") as f:
        json.dump(bedrock, f, ensure_ascii=False, indent=None)
    print(f"Bedrock guardado: {len(eventos)} eventos")
    return 0


if __name__ == "__main__":
    sys.exit(main())
