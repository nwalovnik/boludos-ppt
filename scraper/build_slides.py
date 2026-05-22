"""
Generador de PPTX con las publicaciones de la semana.

Lee data.json (campo _meta.publicaciones) y genera un slide por publicación
con el estilo del ejemplo "Comex" del usuario:
  - Título grande arriba en mayúsculas, color cyan #00B2C9
  - Texto principal en gris #3F3F3F, font Encode Sans
  - Bullets editoriales con narrativa
  - Footer "Fuente: elaboración propia en base a INDEC/..."

Output: site/publicaciones_semana.pptx
"""
from __future__ import annotations

import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = Path(__file__).resolve().parent.parent
DATA_PATH = ROOT / "data.json"
OUT_PATH = ROOT / "publicaciones_semana.pptx"

# Paleta editorial Comex
GRIS_TX = RGBColor(0x3F, 0x3F, 0x3F)
CYAN    = RGBColor(0x00, 0xB2, 0xC9)
ROJO    = RGBColor(0xB9, 0x1C, 0x1C)
VERDE   = RGBColor(0x15, 0x80, 0x3D)
BLANCO  = RGBColor(0xFF, 0xFF, 0xFF)
FONT    = "Encode Sans"

MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]

def lbl_mes(periodo: str) -> str:
    if not periodo or len(periodo) < 7 or "-" not in periodo:
        return periodo
    y, m = periodo[:4], int(periodo[5:7])
    return f"{MESES_ES[m-1]} de {y}"

def fmt_n(v, decimals=0) -> str:
    if v is None:
        return "—"
    if isinstance(v, (int, float)):
        s = f"{v:,.{decimals}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    return str(v)

def ultimo_lunes() -> datetime:
    """Lunes 00:00 UTC de la semana en curso."""
    d = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    diff = d.weekday()  # 0 = lunes
    return d - timedelta(days=diff)

# ──────────────────────────────────────────────────────────────────────────────
# Templates por serie — generan título + 3-4 bullets de narrativa
# ──────────────────────────────────────────────────────────────────────────────
def gen_ipc(p):
    d = p["datos"]
    titulo = f"INFLACIÓN DE {lbl_mes(p['periodo']).upper()}: {('+' if d.get('vm',0)>=0 else '')}{d.get('vm','—')}%"
    lede = (f"El IPC nivel general aumentó {('+' if d.get('vm',0)>=0 else '')}{d.get('vm','—')}% mensual en "
            f"{lbl_mes(p['periodo'])}, acumulando {d.get('via','—')}% interanual. "
            f"INDEC publicó el indicador con la metodología vigente.")
    bullets = [
        f"Variación mensual: {('+' if d.get('vm',0)>=0 else '')}{d.get('vm','—')}% — IPC nivel general INDEC.",
        f"Variación interanual: {d.get('via','—')}% — acumulado últimos 12 meses.",
    ]
    return titulo, lede, bullets

def gen_ipim(p):
    d = p["datos"]
    titulo = f"INFLACIÓN MAYORISTA {lbl_mes(p['periodo']).upper()}: +{d.get('vm','—')}%"
    lede = (f"El IPIM (Índice de Precios Internos al por Mayor) registró un aumento "
            f"de +{d.get('vm','—')}% mensual en {lbl_mes(p['periodo'])}, "
            f"acumulando {d.get('via','—')}% interanual.")
    bullets = [
        f"Nivel general IPIM: +{d.get('vm','—')}% mensual — INDEC.",
        f"Interanual: {d.get('via','—')}%.",
        "El IPIM suele anticipar el IPC minorista: brechas positivas indican presión sobre góndolas en meses siguientes.",
    ]
    return titulo, lede, bullets

def gen_bc(p):
    d = p["datos"]
    saldo = d.get("saldo", 0)
    expo = d.get("expo", 0)
    impo = d.get("impo_abs") or abs(d.get("impo", 0))
    signo = "SUPERÁVIT" if saldo >= 0 else "DÉFICIT"
    titulo = f"BALANZA COMERCIAL {lbl_mes(p['periodo']).upper()}: {signo} USD {fmt_n(abs(saldo))}M"
    lede = (f"El intercambio comercial argentino en {lbl_mes(p['periodo'])} arrojó un "
            f"{signo.lower()} de USD {fmt_n(abs(saldo))} millones. "
            f"Las exportaciones alcanzaron USD {fmt_n(expo)} millones y "
            f"las importaciones USD {fmt_n(impo)} millones.")
    bullets = [
        f"Exportaciones: USD {fmt_n(expo)} millones.",
        f"Importaciones: USD {fmt_n(impo)} millones.",
        f"Saldo comercial: USD {('+' if saldo>=0 else '')}{fmt_n(saldo)} millones.",
    ]
    return titulo, lede, bullets

def gen_fiscal(p):
    d = p["datos"]
    sp = d.get("sal_prim", 0)
    sf = d.get("sal_fin", 0)
    signo = "SUPERÁVIT" if sp >= 0 else "DÉFICIT"
    titulo = f"RESULTADO FISCAL {lbl_mes(p['periodo']).upper()}: {signo} PRIMARIO ${fmt_n(abs(sp/1000))} MIL MILLONES"
    lede = (f"El Sector Público Nacional No Financiero registró un resultado primario "
            f"de {('+' if sp>=0 else '')}{fmt_n(sp/1000)} mil millones de pesos en "
            f"{lbl_mes(p['periodo'])}. El resultado financiero (neto de intereses) "
            f"alcanzó {('+' if sf>=0 else '')}{fmt_n(sf/1000)} mil millones.")
    bullets = [
        f"Resultado primario: {('+' if sp>=0 else '')}{fmt_n(sp/1000)} mil millones de pesos.",
        f"Resultado financiero: {('+' if sf>=0 else '')}{fmt_n(sf/1000)} mil millones de pesos (neto de intereses).",
        f"Fuente: Secretaría de Hacienda, base caja (Metodología 2017).",
    ]
    return titulo, lede, bullets

def gen_emae(p):
    d = p["datos"]
    via = d.get("via", 0)
    vm = d.get("vm")
    verbo = "CRECIÓ" if via >= 0 else "CAYÓ"
    titulo = f"EMAE {lbl_mes(p['periodo']).upper()}: ACTIVIDAD {verbo} {('+' if via>=0 else '')}{via}% I.A."
    lede = (f"El Estimador Mensual de Actividad Económica (EMAE) registró una variación "
            f"interanual de {('+' if via>=0 else '')}{via}% en {lbl_mes(p['periodo'])}. "
            + (f"En la medición desestacionalizada, el indicador se movió {('+' if vm>=0 else '')}{vm}% respecto al mes anterior." if vm is not None else ""))
    bullets = [
        f"Variación interanual: {('+' if via>=0 else '')}{via}%.",
    ]
    if vm is not None:
        bullets.append(f"Variación mensual desestacionalizada: {('+' if vm>=0 else '')}{vm}%.")
    bullets.append("Serie INDEC, base 2004=100. Mide el nivel general de actividad económica argentina.")
    return titulo, lede, bullets

def gen_ipi(p):
    d = p["datos"]
    via = d.get("via", 0)
    verbo = "CRECIÓ" if via >= 0 else "CAYÓ"
    titulo = f"INDUSTRIA {lbl_mes(p['periodo']).upper()}: IPI {verbo} {('+' if via>=0 else '')}{via}% I.A."
    lede = (f"El Índice de Producción Industrial Manufacturero (IPI) registró una "
            f"variación interanual de {('+' if via>=0 else '')}{via}% en {lbl_mes(p['periodo'])}.")
    bullets = [
        f"Variación interanual: {('+' if via>=0 else '')}{via}%.",
        "Serie INDEC, base 2004=100. Mide la producción manufacturera (16 ramas).",
    ]
    return titulo, lede, bullets

def gen_super(p):
    d = p["datos"]
    via = d.get("via_real", 0)
    vm = d.get("vm_dest")
    verbo = "CRECIERON" if via >= 0 else "CAYERON"
    titulo = f"SUPERMERCADOS {lbl_mes(p['periodo']).upper()}: VENTAS {verbo} {('+' if via>=0 else '')}{via}% I.A. REAL"
    lede = (f"Las ventas de supermercados a precios constantes registraron una variación "
            f"interanual real de {('+' if via>=0 else '')}{via}% en {lbl_mes(p['periodo'])}. "
            + (f"En la serie desestacionalizada, se movieron {('+' if vm>=0 else '')}{vm}% respecto al mes anterior." if vm is not None else ""))
    bullets = [
        f"Variación interanual real: {('+' if via>=0 else '')}{via}%.",
    ]
    if vm is not None:
        bullets.append(f"Variación mensual desestacionalizada: {('+' if vm>=0 else '')}{vm}%.")
    if d.get("acum_real") is not None:
        bullets.append(f"Variación acumulada del año: {('+' if d['acum_real']>=0 else '')}{d['acum_real']}%.")
    return titulo, lede, bullets

def gen_salarios(p):
    d = p["datos"]
    titulo = f"SALARIOS {lbl_mes(p['periodo']).upper()}: ÍNDICE TOTAL REGISTRADO {fmt_n(d.get('is_r',0))}"
    lede = (f"El Índice de Salarios INDEC alcanzó un nivel total de {fmt_n(d.get('is_r',0))} "
            f"para el sector registrado en {lbl_mes(p['periodo'])}. "
            f"El privado registrado llegó a {fmt_n(d.get('real_priv',0))} y el público a {fmt_n(d.get('real_pub',0))}.")
    bullets = [
        f"Total registrado: {fmt_n(d.get('is_r','—'))}",
        f"Privado registrado: {fmt_n(d.get('real_priv','—'))}",
        f"Sector público: {fmt_n(d.get('real_pub','—'))}",
    ]
    if d.get("no_reg") is not None:
        bullets.append(f"No registrado: {fmt_n(d['no_reg'])}")
    return titulo, lede, bullets

def gen_mora(p):
    d = p["datos"]
    fam = d.get("fam", 0)
    emp = d.get("emp", 0)
    titulo = f"MORA BANCARIA {lbl_mes(p['periodo']).upper()}: FAMILIAS {fam}% · EMPRESAS {emp}%"
    lede = (f"La irregularidad de cartera del sistema financiero alcanzó {fam}% en familias "
            f"y {emp}% en empresas en {lbl_mes(p['periodo'])} (BCRA Informe sobre Bancos).")
    bullets = [
        f"Familias: {fam}% del total de préstamos en situación irregular.",
        f"Empresas: {emp}%.",
        "Fuente: BCRA Informe sobre Bancos, Anexo Cuadro 'Calidad de Cartera'.",
    ]
    return titulo, lede, bullets

def gen_default(p):
    d = p["datos"]
    titulo = f"{p['label'].upper()} {lbl_mes(p['periodo']).upper()}"
    lede = f"Se publicó la actualización mensual de {p['label']} para {lbl_mes(p['periodo'])}. Fuente: {p.get('fuente','INDEC')}."
    bullets = [f"{k}: {fmt_n(v, 2)}" for k, v in list(d.items())[:5]]
    return titulo, lede, bullets

TEMPLATES = {
    "ipc": gen_ipc, "ipim": gen_ipim, "ipib": gen_ipim, "ipp": gen_ipim,
    "bc": gen_bc, "fiscal": gen_fiscal, "emae": gen_emae,
    "ipi": gen_ipi, "isac": gen_ipi, "uci": gen_ipi,
    "super": gen_super, "mayor": gen_super,
    "salarios": gen_salarios, "mora": gen_mora,
}

# ──────────────────────────────────────────────────────────────────────────────
# Construcción de cada slide
# ──────────────────────────────────────────────────────────────────────────────
def build_slide(prs, pub):
    """Genera un slide con título cyan + lede + bullets, footer fuente."""
    serie = pub.get("serie", "")
    gen = TEMPLATES.get(serie, gen_default)
    titulo, lede, bullets = gen(pub)

    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    # Fondo blanco implícito. Banner superior cyan sutil.
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.3))
    banner.fill.solid()
    banner.fill.fore_color.rgb = CYAN
    banner.line.fill.background()

    # Título
    tit_left = Inches(0.4)
    tit_top = Inches(0.55)
    tit_w = prs.slide_width - Inches(0.8)
    tit_h = Inches(1.4)
    tit_box = slide.shapes.add_textbox(tit_left, tit_top, tit_w, tit_h)
    tf = tit_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    p_tit = tf.paragraphs[0]
    p_tit.alignment = PP_ALIGN.LEFT
    r_tit = p_tit.add_run()
    r_tit.text = titulo
    r_tit.font.name = FONT
    r_tit.font.size = Pt(28)
    r_tit.font.bold = True
    r_tit.font.color.rgb = CYAN

    # Periodo + fuente bajo el título
    sub_top = tit_top + tit_h + Inches(0.05)
    sub_box = slide.shapes.add_textbox(tit_left, sub_top, tit_w, Inches(0.3))
    sp = sub_box.text_frame.paragraphs[0]
    sp.alignment = PP_ALIGN.LEFT
    sr = sp.add_run()
    sr.text = f"{pub.get('fuente','INDEC')} · período {pub.get('periodo','')}"
    sr.font.name = FONT
    sr.font.size = Pt(11)
    sr.font.italic = True
    sr.font.color.rgb = GRIS_TX

    # Lede (párrafo principal)
    lede_top = sub_top + Inches(0.4)
    lede_box = slide.shapes.add_textbox(tit_left, lede_top, tit_w, Inches(1.5))
    ltf = lede_box.text_frame
    ltf.word_wrap = True
    ltf.margin_left = ltf.margin_right = Emu(0)
    lp = ltf.paragraphs[0]
    lp.alignment = PP_ALIGN.JUSTIFY
    lr = lp.add_run()
    lr.text = lede
    lr.font.name = FONT
    lr.font.size = Pt(14)
    lr.font.color.rgb = GRIS_TX

    # Bullets
    bul_top = lede_top + Inches(1.7)
    bul_box = slide.shapes.add_textbox(tit_left, bul_top, tit_w, Inches(3.5))
    btf = bul_box.text_frame
    btf.word_wrap = True
    btf.margin_left = ltf.margin_right = Emu(0)
    for i, b in enumerate(bullets):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        r = p.add_run()
        r.text = "●  " + b
        r.font.name = FONT
        r.font.size = Pt(14)
        r.font.color.rgb = GRIS_TX
        p.space_after = Pt(8)

    # Footer "Fuente"
    footer_box = slide.shapes.add_textbox(
        tit_left, prs.slide_height - Inches(0.5), tit_w, Inches(0.3)
    )
    fp = footer_box.text_frame.paragraphs[0]
    fp.alignment = PP_ALIGN.LEFT
    fr = fp.add_run()
    fr.text = f"Fuente: elaboración propia en base a {pub.get('fuente','INDEC')}"
    fr.font.name = FONT
    fr.font.size = Pt(10)
    fr.font.italic = True
    fr.font.color.rgb = GRIS_TX

def build_cover(prs, n_pubs, lunes):
    """Slide portada."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Fondo cyan
    fondo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    fondo.fill.solid()
    fondo.fill.fore_color.rgb = CYAN
    fondo.line.fill.background()

    # Título grande
    titulo = slide.shapes.add_textbox(Inches(0.6), Inches(2.0), prs.slide_width - Inches(1.2), Inches(1.5))
    tf = titulo.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "PANORAMA MACRO"
    r.font.name = FONT
    r.font.size = Pt(54)
    r.font.bold = True
    r.font.color.rgb = BLANCO

    sub = slide.shapes.add_textbox(Inches(0.6), Inches(3.4), prs.slide_width - Inches(1.2), Inches(1.0))
    p2 = sub.text_frame.paragraphs[0]
    r2 = p2.add_run()
    r2.text = f"Publicaciones de la semana del {lunes.strftime('%d de ')}" + MESES_ES[lunes.month - 1] + lunes.strftime(' de %Y')
    r2.font.name = FONT
    r2.font.size = Pt(20)
    r2.font.color.rgb = BLANCO

    sub2 = slide.shapes.add_textbox(Inches(0.6), Inches(4.0), prs.slide_width - Inches(1.2), Inches(0.5))
    p3 = sub2.text_frame.paragraphs[0]
    r3 = p3.add_run()
    r3.text = f"{n_pubs} publicaciones · datos oficiales INDEC / BCRA / Hacienda"
    r3.font.name = FONT
    r3.font.size = Pt(14)
    r3.font.italic = True
    r3.font.color.rgb = BLANCO

def main():
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    pubs = (data.get("_meta") or {}).get("publicaciones") or []
    if not pubs:
        print("Sin publicaciones tracked en _meta.publicaciones. No genero PPTX.")
        return 1
    lunes = ultimo_lunes()
    lunes_iso = lunes.isoformat()
    semana = [p for p in pubs if (p.get("publicado_at") or p.get("detectado_at","")) >= lunes_iso]
    if not semana:
        print("Sin publicaciones esta semana. No genero PPTX.")
        return 1
    # Ordenar por publicado_at ascendente (lunes arriba)
    semana.sort(key=lambda p: p.get("publicado_at") or p.get("detectado_at",""))

    # 16:9 widescreen
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    build_cover(prs, len(semana), lunes)
    for pub in semana:
        build_slide(prs, pub)

    prs.save(str(OUT_PATH))
    print(f"OK — {OUT_PATH} ({OUT_PATH.stat().st_size/1024:.1f} KB, {len(semana)+1} slides)")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
