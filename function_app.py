import logging
import json
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
import re

import azure.functions as func
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.enum.section import WD_ORIENT
from groq import Groq

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

# ═══════════════════════════════════════════════════════
# PROMPTS
# ═══════════════════════════════════════════════════════

SYSTEM_PROMPT = """Eres una experta en Educación Parvularia chilena. Generas planificaciones semanales co-construidas y diversificadas para Nivel Medio Menor (3 años) siguiendo las Bases Curriculares de Educación Parvularia 2018.

Responde SOLO con un objeto JSON válido sin texto adicional, sin bloques de código markdown.
El JSON debe tener claves: lunes, martes, miercoles, jueves, viernes. Cada una con exactamente estos campos:
{
  "titulo_dia": "Lunes 9 de junio: Titulo creativo de la experiencia",
  "inicio": "Texto del momento de inicio 2-3 oraciones motivadoras con presentacion de materiales en primera persona de la educadora",
  "escenario_1_titulo": "Titulo descriptivo escenario 1",
  "escenario_1_desc": "Descripcion detallada con mediacion y cierre 4-5 oraciones",
  "escenario_2_titulo": "Titulo descriptivo escenario 2",
  "escenario_2_desc": "Descripcion detallada con mediacion y cierre 4-5 oraciones",
  "escenario_3_titulo": "Titulo descriptivo escenario 3",
  "escenario_3_desc": "Descripcion detallada con mediacion y cierre 4-5 oraciones",
  "materiales_1": "material1, material2, material3, material4",
  "materiales_2": "material1, material2, material3",
  "materiales_3": "material1, material2, material3",
  "que_haran_ninos": "Descripcion breve en 2-3 oraciones de lo que haran los ninos ese dia, lenguaje simple para familias",
  "que_haran_familias": "Sugerencia concreta de actividad familiar en casa relacionada con el tema, 2-3 oraciones",
  "que_necesitamos": "Material o preparation que debe enviar la familia, o 'Utilizaremos recursos que se encuentran en el centro educativo.'",
  "indicador_oa1": "Indicador observable 1 del OA comienza con verbo",
  "indicador_oa2": "Indicador observable 2 del OA comienza con verbo",
  "indicador_oat": "Indicador observable del OAT comienza con verbo"
}

IMPORTANTE: Usa solo comillas dobles. No uses saltos de linea dentro de los valores. Cada valor debe ser texto continuo en una sola linea."""


def build_prompt(data: dict) -> str:
    dias_texto = []
    for dia in data["dias"]:
        dias_texto.append(f"""{dia['nombre']}:
- Tema: {dia['tema']}
- Ambito: {dia['ambito']}
- Nucleo: {dia['oa']['nucleo']}
- OA: {dia['oa']['num']} - {dia['oa']['texto']}
- OAT Ambito: Desarrollo Personal y Social
- OAT Nucleo: {dia['oat']['nucleo']}
- OAT: {dia['oat']['num']} - {dia['oat']['texto']}""")
    return f"""Genera la planificacion semanal completa para Nivel Medio Menor.
Educadora: {data['educadora']}
Tecnicos: {data['tecnicos']}
Semana: {data['fecha_inicio']} al {data['fecha_fin']}

{chr(10).join(dias_texto)}

Genera contenido pedagogico rico y concreto apropiado para ninos de 3 anios.
IMPORTANTE: Responde SOLO con JSON valido. Sin markdown. Sin saltos de linea dentro de los valores."""


# ═══════════════════════════════════════════════════════
# UTILIDADES
# ═══════════════════════════════════════════════════════

CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Groq-Key"
}

COLOR_HEADER = "ADDB7B"
COLOR_GRAY   = "F2F2F2"
COLOR_WHITE  = "FFFFFF"

DIAS_KEYS  = ["lunes", "martes", "miercoles", "jueves", "viernes"]
DIAS_NAMES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
ESCALA     = ("Escala de apreciación:\n"
              "L: Logrado  ML: Medianamente Logrado\n"
              "IC: Iniciando Conducta  N/O: No observado  A: Ausente")
ROL_TEXTO  = ("- Los niños/as escogen libremente su lugar de juego y material.\n"
              "- Los niños y niñas exploran y manipulan el material libremente.\n"
              "- Participan activamente durante la experiencia de aprendizaje.\n"
              "- Conversan y comparten sus ideas con sus pares y con los adultos.")


def clean_json(raw: str) -> dict:
    text = raw.strip().replace("```json", "").replace("```", "").strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group()

    result = []
    in_string = False
    escape = False
    for ch in text:
        if escape:
            result.append(ch)
            escape = False
        elif ch == '\\':
            result.append(ch)
            escape = True
        elif ch == '"':
            in_string = not in_string
            result.append(ch)
        elif in_string:
            if ch in ('\n', '\r'):
                result.append(' ')
            elif ch == '\t':
                result.append(' ')
            elif ord(ch) < 0x20:
                pass
            else:
                result.append(ch)
        else:
            result.append(ch)

    return json.loads(''.join(result))


def set_cell_bg(cell, hex_color: str):
    shading = parse_xml(
        f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:val="clear" w:color="auto" w:fill="{hex_color}"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)


def cell_write(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.paragraphs[0].clear()
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return para


def cell_add(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
    para = cell.add_paragraph()
    para.alignment = align
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return para


def call_groq(groq_key: str, body: dict) -> dict:
    client = Groq(api_key=groq_key)
    resp   = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        max_tokens=6000,
        temperature=0.7,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": build_prompt(body)}
        ]
    )
    return clean_json(resp.choices[0].message.content)


# ═══════════════════════════════════════════════════════
# BUILDER PLANIFICACIÓN
# ═══════════════════════════════════════════════════════

def build_planificacion(data: dict, ai: dict) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.orientation   = WD_ORIENT.LANDSCAPE
    section.page_width    = Cm(27.94)
    section.page_height   = Cm(21.59)
    section.top_margin    = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin   = Cm(1.27)
    section.right_margin  = Cm(1.27)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title_p.add_run("PLANIFICACIÓN CO-CONSTRUIDA Y DIVERSIFICADA")
    r.bold = True; r.font.size = Pt(13)

    info1 = doc.add_paragraph()
    r = info1.add_run("\tNivel: "); r.bold = True; r.font.size = Pt(11)
    info1.add_run("Medio Menor").font.size = Pt(11)
    r2 = info1.add_run("                                                 Fecha: ")
    r2.bold = True; r2.font.size = Pt(11)
    info1.add_run(f"{data['fecha_inicio']} al {data['fecha_fin']}").font.size = Pt(11)

    info2 = doc.add_paragraph()
    r3 = info2.add_run("\tEducadora: "); r3.bold = True; r3.font.size = Pt(11)
    info2.add_run(data['educadora']).font.size = Pt(11)
    r4 = info2.add_run("                                 Técnicos: ")
    r4.bold = True; r4.font.size = Pt(11)
    info2.add_run(data['tecnicos']).font.size = Pt(11)

    doc.add_paragraph()

    C = [Cm(3.74), Cm(4.78), Cm(9.21), Cm(9.21), Cm(7.79), Cm(7.79), Cm(7.79)]
    table = doc.add_table(rows=13, cols=7)
    table.style = "Table Grid"
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = C[i]

    # FILA 0 — Header OA
    r0 = table.rows[0]
    for i in range(7):
        set_cell_bg(r0.cells[i], COLOR_HEADER)
    cell_write(r0.cells[0], "ÁMBITO", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_write(r0.cells[1], "NÚCLEO", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    oa_hdr = r0.cells[2].merge(r0.cells[3])
    set_cell_bg(oa_hdr, COLOR_HEADER)
    cell_write(oa_hdr, "OBJETIVO DE APRENDIZAJE", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    oat_hdr = r0.cells[4].merge(r0.cells[5]).merge(r0.cells[6])
    set_cell_bg(oat_hdr, COLOR_HEADER)
    cell_write(oat_hdr, "OBJETIVO DE APRENDIZAJE TRANSVERSAL", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    # FILAS 1-5 — OA por día
    for idx, key in enumerate(DIAS_KEYS):
        fd  = data["dias"][idx]
        row = table.rows[idx + 1]

        c0 = row.cells[0]
        c0.paragraphs[0].clear()
        rb = c0.paragraphs[0].add_run(f"{DIAS_NAMES[idx]}: ")
        rb.bold = True; rb.font.size = Pt(9)
        c0.paragraphs[0].add_run(fd["ambito"]).font.size = Pt(9)

        cell_write(row.cells[1], fd["oa"]["nucleo"], size=9)

        oa_c = row.cells[2].merge(row.cells[3])
        cell_write(oa_c, f"{fd['oa']['num']}: {fd['oa']['texto']}", size=9)

        oat_c = row.cells[4].merge(row.cells[5]).merge(row.cells[6])
        oat_c.paragraphs[0].clear()
        ra = oat_c.paragraphs[0].add_run("Ámbito: ")
        ra.bold = True; ra.font.size = Pt(9)
        oat_c.paragraphs[0].add_run("Desarrollo Personal y Social").font.size = Pt(9)
        pn = oat_c.add_paragraph()
        rn = pn.add_run("Núcleo: "); rn.bold = True; rn.font.size = Pt(9)
        pn.add_run(fd["oat"]["nucleo"]).font.size = Pt(9)
        po = oat_c.add_paragraph()
        ro = po.add_run(f"{fd['oat']['num']}: "); ro.bold = True; ro.font.size = Pt(9)
        po.add_run(fd["oat"]["texto"]).font.size = Pt(9)

    # FILA 6 — Header metodología
    r6 = table.rows[6]
    for i in range(7):
        set_cell_bg(r6.cells[i], COLOR_HEADER)
    cell_write(r6.cells[0], "ROL PROTAGÓNICO DEL NIÑO Y LA NIÑA", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    met_hdr = r6.cells[1].merge(r6.cells[2])
    set_cell_bg(met_hdr, COLOR_HEADER)
    cell_write(met_hdr, "SUGERENCIA METODOLÓGICA", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    mat_hdr = r6.cells[3].merge(r6.cells[4])
    set_cell_bg(mat_hdr, COLOR_HEADER)
    cell_write(mat_hdr, "RECURSOS/MATERIALES", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_write(r6.cells[5], "EVALUACIÓN", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_write(r6.cells[6], "PARTICIPACIÓN DE LA FAMILIA", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    # FILAS 7-11 — Metodología por día
    for idx, key in enumerate(DIAS_KEYS):
        fd  = data["dias"][idx]
        aid = ai.get(key, {})
        row = table.rows[idx + 7]

        c0 = row.cells[0]
        c0.paragraphs[0].clear()
        for i, line in enumerate(ROL_TEXTO.split("\n")):
            p = c0.paragraphs[0] if i == 0 else c0.add_paragraph()
            p.add_run(line).font.size = Pt(9)

        mc = row.cells[1].merge(row.cells[2])
        mc.paragraphs[0].clear()
        rt = mc.paragraphs[0].add_run(aid.get("titulo_dia", f"{DIAS_NAMES[idx]}: {fd['tema']}"))
        rt.bold = True; rt.font.size = Pt(9)
        p_ini = mc.add_paragraph()
        ri = p_ini.add_run("Inicio: "); ri.bold = True; ri.font.size = Pt(9)
        p_ini.add_run(aid.get("inicio", "")).font.size = Pt(9)
        for s in ["1","2","3"]:
            pe = mc.add_paragraph()
            re_r = pe.add_run(f"Escenario {s}: {aid.get(f'escenario_{s}_titulo','')}")
            re_r.bold = True; re_r.font.size = Pt(9)
            pd = mc.add_paragraph()
            pd.add_run(aid.get(f"escenario_{s}_desc","")).font.size = Pt(9)

        matc = row.cells[3].merge(row.cells[4])
        matc.paragraphs[0].clear()
        first = True
        for s in ["1","2","3"]:
            p_sh = matc.paragraphs[0] if first else matc.add_paragraph()
            first = False
            rsh = p_sh.add_run(f"Escenario {s}:"); rsh.bold = True; rsh.font.size = Pt(9)
            for mat in [m.strip() for m in aid.get(f"materiales_{s}","").split(",") if m.strip()]:
                matc.add_paragraph().add_run(f"- {mat}").font.size = Pt(9)

        evc = row.cells[5]
        evc.paragraphs[0].clear()
        evc.paragraphs[0].add_run(ESCALA).font.size = Pt(8)
        ph = evc.add_paragraph()
        ph.add_run("Indicadores OA (Niños y Niñas):").bold = True
        ph.runs[0].font.size = Pt(9)
        for n, field in [("1","indicador_oa1"),("2","indicador_oa2")]:
            evc.add_paragraph().add_run(f"{n}. {aid.get(field,'')}").font.size = Pt(9)
        poh = evc.add_paragraph()
        poh.add_run("Indicador OAT:").bold = True
        poh.runs[0].font.size = Pt(9)
        evc.add_paragraph().add_run(f"1. {aid.get('indicador_oat','')}").font.size = Pt(9)

        famc = row.cells[6]
        famc.paragraphs[0].clear()
        famc.paragraphs[0].add_run(aid.get("que_haran_familias","")).font.size = Pt(9)

    # FILA 12 — Observación
    r12 = table.rows[12]
    obs = r12.cells[0]
    for i in range(1, 7):
        obs = obs.merge(r12.cells[i])
    obs.paragraphs[0].clear()
    ro = obs.paragraphs[0].add_run("Observación:     ")
    ro.bold = True; ro.font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════
# BUILDER INFO FAMILIA
# ═══════════════════════════════════════════════════════

def build_info_familia(data: dict, ai: dict) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.orientation   = WD_ORIENT.LANDSCAPE
    section.page_width    = Cm(27.94)
    section.page_height   = Cm(21.59)
    section.top_margin    = Cm(2.25)
    section.bottom_margin = Cm(2.50)
    section.left_margin   = Cm(2.50)
    section.right_margin  = Cm(2.50)

    # Título
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = title_p.add_run(f'"Organizando mi semana del {data["fecha_inicio"]} al {data["fecha_fin"]}"')
    rt.bold = True; rt.font.size = Pt(12)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.add_run("Medio Menor 2026").font.size = Pt(12)

    doc.add_paragraph()

    # Tabla 4 filas x 5 columnas
    table = doc.add_table(rows=4, cols=5)
    table.style = "Table Grid"

    col_w = Cm(3.74)
    for row in table.rows:
        for cell in row.cells:
            cell.width = col_w

    # Extraer números de días desde las fechas
    from datetime import datetime, timedelta
    try:
        fecha_ini = datetime.strptime(data["fecha_inicio"], "%Y-%m-%d")
        fechas = [(fecha_ini + timedelta(days=i)) for i in range(5)]
        dias_labels = [f"{DIAS_NAMES[i].upper()} {fechas[i].day}" for i in range(5)]
    except Exception:
        dias_labels = [f"{n.upper()}" for n in DIAS_NAMES]

    # FILA 0 — Header días (verde)
    for i, label in enumerate(dias_labels):
        cell = table.rows[0].cells[i]
        set_cell_bg(cell, COLOR_HEADER)
        cell_write(cell, label, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    # FILA 1 — ¿Qué harán los niños/as?
    for i, key in enumerate(DIAS_KEYS):
        aid  = ai.get(key, {})
        cell = table.rows[1].cells[i]
        set_cell_bg(cell, COLOR_GRAY)
        cell.paragraphs[0].clear()
        rl = cell.paragraphs[0].add_run("¿Qué harán los niños/as? ")
        rl.bold = True; rl.font.size = Pt(9)
        cell.paragraphs[0].add_run(aid.get("que_haran_ninos","")).font.size = Pt(9)

    # FILA 2 — ¿Qué harán las familias?
    for i, key in enumerate(DIAS_KEYS):
        aid  = ai.get(key, {})
        cell = table.rows[2].cells[i]
        set_cell_bg(cell, COLOR_GRAY)
        cell.paragraphs[0].clear()
        rl = cell.paragraphs[0].add_run("¿Qué harán las familias? ")
        rl.bold = True; rl.font.size = Pt(9)
        cell.paragraphs[0].add_run(aid.get("que_haran_familias","")).font.size = Pt(9)

    # FILA 3 — ¿Qué necesitamos?
    for i, key in enumerate(DIAS_KEYS):
        aid  = ai.get(key, {})
        cell = table.rows[3].cells[i]
        set_cell_bg(cell, COLOR_GRAY)
        cell.paragraphs[0].clear()
        rl = cell.paragraphs[0].add_run("¿Qué necesitamos? ")
        rl.bold = True; rl.font.size = Pt(9)
        necesitamos = aid.get("que_necesitamos",
                              "Utilizaremos recursos que se encuentran en el centro educativo.")
        cell.paragraphs[0].add_run(necesitamos).font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════
# ENDPOINTS
# ═══════════════════════════════════════════════════════

@app.route(route="generate", methods=["GET","POST","OPTIONS"])
def generate(req: func.HttpRequest) -> func.HttpResponse:
    """Genera y devuelve el JSON con el contenido para preview."""
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS)

    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse("Body JSON invalido", status_code=400, headers=CORS)

    groq_key = req.headers.get("X-Groq-Key") or os.environ.get("GROQ_API_KEY","")
    if not groq_key:
        return func.HttpResponse("Falta Groq API Key", status_code=401, headers=CORS)

    try:
        ai_data = call_groq(groq_key, body)
    except Exception as e:
        logging.error(f"Groq/JSON error: {e}")
        return func.HttpResponse(f"Error generando contenido: {e}", status_code=500, headers=CORS)

    out_hdrs = {**CORS, "Content-Type": "application/json"}
    return func.HttpResponse(
        body=json.dumps({"ok": True, "data": ai_data}, ensure_ascii=False),
        status_code=200,
        headers=out_hdrs
    )


@app.route(route="download-planificacion", methods=["POST","OPTIONS"])
def download_planificacion(req: func.HttpRequest) -> func.HttpResponse:
    """Recibe form data + ai_data y devuelve el .docx de planificación."""
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS)

    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse("Body JSON invalido", status_code=400, headers=CORS)

    try:
        docx_bytes = build_planificacion(body["form"], body["ai"])
    except Exception as e:
        logging.error(f"DOCX planificacion error: {e}")
        return func.HttpResponse(f"Error generando planificacion: {e}", status_code=500, headers=CORS)

    fecha    = body["form"].get("fecha_inicio","semana")
    out_hdrs = {**CORS,
                "Content-Disposition": f'attachment; filename="Planificacion_{fecha}.docx"'}
    return func.HttpResponse(
        body=docx_bytes, status_code=200,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=out_hdrs
    )


@app.route(route="download-familia", methods=["POST","OPTIONS"])
def download_familia(req: func.HttpRequest) -> func.HttpResponse:
    """Recibe form data + ai_data y devuelve el .docx de info familia."""
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS)

    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse("Body JSON invalido", status_code=400, headers=CORS)

    try:
        docx_bytes = build_info_familia(body["form"], body["ai"])
    except Exception as e:
        logging.error(f"DOCX familia error: {e}")
        return func.HttpResponse(f"Error generando info familia: {e}", status_code=500, headers=CORS)

    fecha    = body["form"].get("fecha_inicio","semana")
    out_hdrs = {**CORS,
                "Content-Disposition": f'attachment; filename="InfoFamilia_{fecha}.docx"'}
    return func.HttpResponse(
        body=docx_bytes, status_code=200,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=out_hdrs
    )



NINOS = [
    "RIHANNY ABOGADO","MAXIMILIANO AGÜERO","AMARO ALONSO","PAZ AMPUERO",
    "OLIVIA ANDRES","ANNY AULAR","CHRIS CARDENAS","ADIEL CARVAJAL",
    "GADIEL CARVAJAL","AMARU CASTRO","AURORA CERNA","MÍA CONTRERAS",
    "SOFÍA CORNEJO","ELOÍSA CORRALES","INTI FREDES","MÁXIMO GARCÍA",
    "DYLAN GELVIS","NICOLÁS GONZÁLEZ","TOMÁS GONZÁLEZ","EZEQUIEL HERRERA",
    "NOAH LARA","THIAGO LIZANO","BASTIÁN QUISPE","LUCIANO ROCCO",
    "MATEO RODRÍGUEZ","BRUNO SAAVEDRA","MIRANDA SAENZ","ARLETH SUÁREZ",
    "EMILIANO VARAS","ÓSCAR VILLASECA"
]

DIAS_KEYS  = ["lunes","martes","miercoles","jueves","viernes"]
DIAS_NAMES = ["Lunes","Martes","Miércoles","Jueves","Viernes"]
ESCALA_VALS = ["L","ML","IC","N/O","A"]

# Colores
C_HEADER   = "ADDB7B"   # verde claro (igual al docx)
C_HEADER2  = "185FA5"   # azul
C_GRAY     = "F2F2F2"
C_WHITE    = "FFFFFF"
C_L        = "C6EFCE"   # verde logrado
C_ML       = "FFEB9C"   # amarillo med logrado
C_IC       = "FFCC99"   # naranja iniciando
C_NO       = "D9D9D9"   # gris no observado
C_A        = "FCB4B4"   # rojo ausente

def hfill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def hfont(bold=False, size=9, color="000000"):
    return Font(bold=bold, size=size, color=color, name="Arial")

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_mid():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def build_evaluacion(form: dict, ai: dict) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # quitar hoja default

    fecha_ini = form.get("fecha_inicio","")
    fecha_fin = form.get("fecha_fin","")
    educadora = form.get("educadora","Giovanna Marino")
    tecnicos  = form.get("tecnicos","Carol Miranda – Helen Oyarzún")

    # ══════════════════════════════════════
    # HOJA PORTADA
    # ══════════════════════════════════════
    ws_p = wb.create_sheet("Portada")
    ws_p.column_dimensions["A"].width = 35
    ws_p.column_dimensions["B"].width = 50

    def portada_row(r, label, value, bold_val=False):
        ws_p.cell(r,1, label).font     = hfont(bold=True, size=10)
        ws_p.cell(r,1).alignment       = left_mid()
        ws_p.cell(r,2, value).font     = hfont(bold=bold_val, size=10)
        ws_p.cell(r,2).alignment       = left_mid()

    # Título
    ws_p.merge_cells("A1:B1")
    t = ws_p.cell(1,1,"EVALUACIÓN DE APRENDIZAJES")
    t.font = Font(bold=True, size=14, name="Arial", color="185FA5")
    t.alignment = center()
    ws_p.row_dimensions[1].height = 28

    ws_p.merge_cells("A2:B2")
    s = ws_p.cell(2,1,f"Planificación Co-construida y Diversificada — Semana del {fecha_ini} al {fecha_fin}")
    s.font = hfont(size=11)
    s.alignment = center()

    ws_p.merge_cells("A3:B3")
    s2 = ws_p.cell(3,1,"Nivel Medio Menor — Sala Cuna y Jardín Infantil Claudio Arrau")
    s2.font = hfont(size=10)
    s2.alignment = center()

    ws_p.cell(5,1).fill = hfill(C_HEADER)
    ws_p.merge_cells("A5:B5")
    ws_p.cell(5,1,"IDENTIFICACIÓN").font = hfont(bold=True, size=10, color="000000")
    ws_p.cell(5,1).alignment = center()

    portada_row(6,  "Nivel:",                    "Medio Menor (2 a 3 años)")
    portada_row(7,  "Educadora de Párvulos:",     educadora)
    portada_row(8,  "Técnicas en Ed. Parvularia:",tecnicos)
    portada_row(9,  "Período de evaluación:",     f"{fecha_ini} al {fecha_fin}")
    portada_row(10, "Total niños/as evaluados:",  str(len(NINOS)))
    portada_row(11, "Instrumento:",               "Escala de apreciación por indicador (OA y OAT)")

    # Escala
    ws_p.cell(13,1).fill = hfill(C_HEADER)
    ws_p.merge_cells("A13:B13")
    ws_p.cell(13,1,"ESCALA DE APRECIACIÓN").font = hfont(bold=True, size=10)
    ws_p.cell(13,1).alignment = center()

    escala_rows = [
        ("L",   "Logrado",                "El indicador se observa de forma autónoma y constante.", C_L),
        ("ML",  "Medianamente Logrado",   "Se observa con apoyo verbal o gestual del adulto.", C_ML),
        ("IC",  "Iniciando Conducta",     "Se observan acercamientos iniciales o esporádicos.", C_IC),
        ("N/O", "No Observado",           "No se logró observar durante la experiencia.", C_NO),
        ("A",   "Ausente",                "El niño/a no asistió a la jornada.", C_A),
    ]
    for i, (cod, sig, desc, col) in enumerate(escala_rows):
        r = 14 + i
        ws_p.cell(r,1, f"{cod} — {sig}").font = hfont(bold=True, size=9)
        ws_p.cell(r,1).fill = hfill(col)
        ws_p.cell(r,1).alignment = left_mid()
        ws_p.cell(r,2, desc).font = hfont(size=9)
        ws_p.cell(r,2).alignment = left_mid()

    # Estructura
    ws_p.cell(20,1).fill = hfill(C_HEADER)
    ws_p.merge_cells("A20:B20")
    ws_p.cell(20,1,"ESTRUCTURA DEL LIBRO").font = hfont(bold=True, size=10)
    ws_p.cell(20,1).alignment = center()

    hojas_info = [(DIAS_NAMES[i], form["dias"][i].get("tema",""), ai.get(DIAS_KEYS[i],{}).get("titulo_dia","")) for i in range(5)]
    for i, (dia, tema, titulo) in enumerate(hojas_info):
        r = 21 + i
        ws_p.cell(r,1, dia).font = hfont(bold=True, size=9)
        ws_p.cell(r,1).alignment = left_mid()
        ws_p.cell(r,2, titulo or tema).font = hfont(size=9)
        ws_p.cell(r,2).alignment = left_mid()

    ws_p.cell(27,1, "Resumen Semanal").font  = hfont(bold=True, size=9)
    ws_p.cell(27,2, "Consolidado por niño/a y por indicador. Cálculos automáticos.").font = hfont(size=9)
    ws_p.cell(28,1, "Resumen por Núcleo").font  = hfont(bold=True, size=9)
    ws_p.cell(28,2, "Distribución de logros por ámbito y núcleo BCEP.").font = hfont(size=9)

    # ══════════════════════════════════════
    # HOJAS DIARIAS
    # ══════════════════════════════════════
    day_sheets = {}
    for d_idx in range(5):
        key    = DIAS_KEYS[d_idx]
        nombre = DIAS_NAMES[d_idx]
        fd     = form["dias"][d_idx]
        aid    = ai.get(key, {})
        tema   = aid.get("titulo_dia") or fd.get("tema","")

        oa     = fd.get("oa") or {}
        oat    = fd.get("oat") or {}
        ind1   = aid.get("indicador_oa1","")
        ind2   = aid.get("indicador_oa2","")
        ind_oat= aid.get("indicador_oat","")

        fecha_dia = f"{nombre} {fecha_ini}"
        sheet_name = f"{nombre[:4]} {d_idx+1}"
        ws = wb.create_sheet(sheet_name)
        day_sheets[d_idx] = sheet_name

        # Anchos
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 32
        ws.column_dimensions["D"].width = 32
        ws.column_dimensions["E"].width = 32
        ws.column_dimensions["F"].width = 30

        # FILA 1 — Título
        ws.merge_cells("A1:F1")
        ws.cell(1,1, f"EVALUACIÓN — {fecha_dia}: \"{tema}\"")
        ws.cell(1,1).font      = Font(bold=True, size=12, name="Arial", color="185FA5")
        ws.cell(1,1).alignment = center()
        ws.cell(1,1).fill      = hfill("EFF6FF")
        ws.row_dimensions[1].height = 22

        # FILAS 2-7 — Metadata OA y OAT
        meta = [
            ("Ámbito OA:",  oa.get("ambito", fd.get("ambito",""))),
            ("Núcleo OA:",  oa.get("nucleo","")),
            ("OA principal:", f"{oa.get('num','')}: {oa.get('texto','')}"),
            ("Ámbito OAT:", "Desarrollo Personal y Social"),
            ("Núcleo OAT:", oat.get("nucleo","")),
            ("OAT (transversal):", f"{oat.get('num','')}: {oat.get('texto','')}"),
        ]
        for i, (lbl, val) in enumerate(meta):
            r = i + 2
            ws.merge_cells(f"A{r}:B{r}")
            ws.cell(r,1, lbl).font = hfont(bold=True, size=9)
            ws.cell(r,1).alignment = left_mid()
            ws.merge_cells(f"C{r}:F{r}")
            ws.cell(r,3, val).font = hfont(size=9)
            ws.cell(r,3).alignment = left_mid()

        # FILA 8 — Header columnas
        headers = ["N°","Niño/a","INDICADOR OA 1","INDICADOR OA 2","INDICADOR OAT","Observaciones"]
        fills_h = [C_HEADER,C_HEADER,C_HEADER,C_HEADER,C_HEADER,"185FA5"]
        colors_h= ["000000","000000","000000","000000","000000","FFFFFF"]
        for ci, (h, f, c) in enumerate(zip(headers, fills_h, colors_h)):
            cell = ws.cell(8, ci+1, h)
            cell.font      = hfont(bold=True, size=9, color=c)
            cell.fill      = hfill(f)
            cell.alignment = center()
            cell.border    = thin_border()
        ws.row_dimensions[8].height = 18

        # FILA 9 — Sub-header indicadores
        ws.cell(9,1,"").border = thin_border()
        ws.cell(9,2,"").border = thin_border()
        inds = [ind1, ind2, ind_oat]
        for ci, ind in enumerate(inds):
            cell = ws.cell(9, ci+3, f"OA-Ind {ci+1}: {ind}" if ci < 2 else f"OAT-Ind 1: {ind}")
            cell.font      = hfont(size=8)
            cell.fill      = hfill(C_GRAY)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border    = thin_border()
        ws.cell(9,6,"").fill   = hfill(C_GRAY)
        ws.cell(9,6).border    = thin_border()
        ws.row_dimensions[9].height = 50

        # Validación de datos para celdas de escala
        dv = DataValidation(type="list", formula1='"L,ML,IC,N/O,A"', allow_blank=True)
        dv.error      = "Ingrese: L, ML, IC, N/O o A"
        dv.errorTitle = "Valor inválido"
        dv.prompt     = "Escala: L / ML / IC / N/O / A"
        ws.add_data_validation(dv)

        # FILAS 10-39 — Niños
        for n_idx, nombre_nino in enumerate(NINOS):
            r = n_idx + 10
            ws.cell(r,1, n_idx+1).font      = hfont(size=9)
            ws.cell(r,1).alignment          = center()
            ws.cell(r,1).border             = thin_border()
            ws.cell(r,2, nombre_nino).font  = hfont(size=9)
            ws.cell(r,2).alignment          = left_mid()
            ws.cell(r,2).border             = thin_border()
            ws.row_dimensions[r].height     = 16

            for ci in range(3,6):
                cell = ws.cell(r, ci, "")
                cell.font      = hfont(bold=True, size=10)
                cell.alignment = center()
                cell.border    = thin_border()
                dv.add(f"{get_column_letter(ci)}{r}")
                # Formato condicional por código
                col_letter = get_column_letter(ci)
                for cod, color in [("L",C_L),("ML",C_ML),("IC",C_IC),("N/O",C_NO),("A",C_A)]:
                    ws.conditional_formatting.add(
                        f"{col_letter}{r}",
                        CellIsRule(operator="equal", formula=[f'"{cod}"'], fill=hfill(color))
                    )

            ws.cell(r,6,"").font      = hfont(size=8)
            ws.cell(r,6).alignment    = left_mid()
            ws.cell(r,6).border       = thin_border()

        # FILAS TOTALES (40-46)
        tot_labels = [
            ("TOTAL Logrado (L)",            "L"),
            ("TOTAL Medianamente Logrado (ML)","ML"),
            ("TOTAL Iniciando Conducta (IC)", "IC"),
            ("TOTAL No Observado (N/O)",      "N/O"),
            ("TOTAL Ausentes (A)",            "A"),
        ]
        for t_idx, (lbl, cod) in enumerate(tot_labels):
            r = 40 + t_idx
            ws.merge_cells(f"A{r}:B{r}")
            cell = ws.cell(r,1, lbl)
            cell.font = hfont(bold=True, size=9)
            cell.fill = hfill(C_GRAY)
            cell.alignment = left_mid()
            cell.border = thin_border()
            for ci in range(3,6):
                col = get_column_letter(ci)
                formula = f'=COUNTIF({col}10:{col}39,"{cod}")'
                tc = ws.cell(r, ci, formula)
                tc.font = hfont(bold=True, size=9)
                tc.alignment = center()
                tc.border = thin_border()
                tc.fill = hfill(C_GRAY)
            ws.cell(r,6).border = thin_border()

        # % LOGRO
        r_logro = 45
        ws.merge_cells(f"A{r_logro}:B{r_logro}")
        ws.cell(r_logro,1,"% LOGRO (L+ML)").font = hfont(bold=True, size=9, color="FFFFFF")
        ws.cell(r_logro,1).fill = hfill("185FA5")
        ws.cell(r_logro,1).alignment = left_mid()
        ws.cell(r_logro,1).border = thin_border()
        for ci in range(3,6):
            col = get_column_letter(ci)
            # L + ML / total no ausentes * 100
            formula = f'=IFERROR((COUNTIF({col}10:{col}39,"L")+COUNTIF({col}10:{col}39,"ML"))/(COUNTA({col}10:{col}39)-COUNTIF({col}10:{col}39,"A"))*100,0)'
            tc = ws.cell(r_logro, ci, formula)
            tc.number_format = "0.0\"%\""
            tc.font = hfont(bold=True, size=9, color="FFFFFF")
            tc.fill = hfill("185FA5")
            tc.alignment = center()
            tc.border = thin_border()
        ws.cell(r_logro,6).border = thin_border()

    # ══════════════════════════════════════
    # HOJA RESUMEN SEMANAL
    # ══════════════════════════════════════
    ws_r = wb.create_sheet("Resumen Semanal")
    ws_r.column_dimensions["A"].width = 5
    ws_r.column_dimensions["B"].width = 28

    # Header principal
    ws_r.merge_cells("A1:W1")
    ws_r.cell(1,1,"RESUMEN SEMANAL POR NIÑO/A — Conteo de la escala")
    ws_r.cell(1,1).font = Font(bold=True, size=12, name="Arial", color="FFFFFF")
    ws_r.cell(1,1).fill = hfill("185FA5")
    ws_r.cell(1,1).alignment = center()

    # Headers días
    day_cols = [3, 8, 13, 18, 23]  # columna inicio de cada día (5 cols: L,ML,IC,N/O,A)
    for d_idx in range(5):
        start_col = day_cols[d_idx]
        end_col   = start_col + 4
        ws_r.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
        cell = ws_r.cell(2, start_col, DIAS_NAMES[d_idx])
        cell.font = hfont(bold=True, size=9, color="000000")
        cell.fill = hfill(C_HEADER)
        cell.alignment = center()
        cell.border = thin_border()
        for sub_idx, sub in enumerate(["L","ML","IC","N/O","A"]):
            c = ws_r.cell(3, start_col+sub_idx, sub)
            c.font = hfont(bold=True, size=8)
            c.fill = hfill(C_GRAY)
            c.alignment = center()
            c.border = thin_border()

    # Headers fijos
    for ci, hdr in [(1,"N°"),(2,"Niño/a"),(28,"Total Ind."),(29,"% Logro")]:
        ws_r.merge_cells(start_row=2, start_column=ci, end_row=3, end_column=ci)
        cell = ws_r.cell(2, ci, hdr)
        cell.font = hfont(bold=True, size=9, color="FFFFFF")
        cell.fill = hfill("185FA5")
        cell.alignment = center()
        cell.border = thin_border()

    # Filas de niños
    for n_idx, nombre_nino in enumerate(NINOS):
        r = n_idx + 4
        ws_r.cell(r,1, n_idx+1).font = hfont(size=9)
        ws_r.cell(r,1).alignment = center()
        ws_r.cell(r,1).border = thin_border()
        ws_r.cell(r,2, nombre_nino).font = hfont(size=9)
        ws_r.cell(r,2).alignment = left_mid()
        ws_r.cell(r,2).border = thin_border()

        for d_idx in range(5):
            sname = day_sheets[d_idx]
            start_col = day_cols[d_idx]
            nino_row  = n_idx + 10  # fila del niño en hoja diaria
            for sub_idx, cod in enumerate(["L","ML","IC","N/O","A"]):
                # Columnas en hoja diaria: C=3, D=4, E=5 (3 indicadores)
                formula = f"=COUNTIF('{sname}'!C{nino_row}:E{nino_row},\"{cod}\")"
                cell = ws_r.cell(r, start_col+sub_idx, formula)
                cell.font = hfont(size=9)
                cell.alignment = center()
                cell.border = thin_border()

        # Total indicadores evaluados
        total_formula = f"=COUNTA(C{r}:AA{r})-COUNTIF(C{r}:AA{r},\"A\")"
        ws_r.cell(r,28, total_formula).font = hfont(size=9)
        ws_r.cell(r,28).alignment = center()
        ws_r.cell(r,28).border = thin_border()

        # % Logro semanal
        pct_formula = (f"=IFERROR((SUMIF(C{r}:AA{r},\"L\",C{r}:AA{r})*0+COUNTIF(C{r}:AA{r},\"L\")"
                       f"+COUNTIF(C{r}:AA{r},\"ML\"))/AB{r}*100,0)")
        pct_formula = f"=IFERROR((COUNTIF(C{r}:AA{r},\"L\")+COUNTIF(C{r}:AA{r},\"ML\"))/X{r}*100,0)"
        ws_r.cell(r,29, pct_formula).font = hfont(size=9)
        ws_r.cell(r,29).number_format = "0.0\"%\""
        ws_r.cell(r,29).alignment = center()
        ws_r.cell(r,29).border = thin_border()

    # Fila TOTAL NIVEL
    r_tot = len(NINOS) + 4
    ws_r.merge_cells(f"A{r_tot}:B{r_tot}")
    ws_r.cell(r_tot,1,"TOTAL NIVEL").font = hfont(bold=True, size=9, color="FFFFFF")
    ws_r.cell(r_tot,1).fill = hfill("185FA5")
    ws_r.cell(r_tot,1).alignment = center()
    ws_r.cell(r_tot,1).border = thin_border()
    for ci in range(3, 30):
        formula = f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{r_tot-1})"
        cell = ws_r.cell(r_tot, ci, formula)
        cell.font = hfont(bold=True, size=9, color="FFFFFF")
        cell.fill = hfill("185FA5")
        cell.alignment = center()
        cell.border = thin_border()

    # ══════════════════════════════════════
    # HOJA RESUMEN POR NÚCLEO
    # ══════════════════════════════════════
    ws_n = wb.create_sheet("Resumen por Núcleo")
    ws_n.column_dimensions["A"].width = 8
    ws_n.column_dimensions["B"].width = 35
    ws_n.column_dimensions["C"].width = 30
    ws_n.column_dimensions["D"].width = 30
    for ci in range(5,11):
        ws_n.column_dimensions[get_column_letter(ci)].width = 8

    ws_n.merge_cells("A1:J1")
    ws_n.cell(1,1,"RESUMEN POR ÁMBITO Y NÚCLEO BCEP — Distribución de logros semanales")
    ws_n.cell(1,1).font = Font(bold=True, size=12, name="Arial", color="FFFFFF")
    ws_n.cell(1,1).fill = hfill("185FA5")
    ws_n.cell(1,1).alignment = center()

    for ci, hdr in enumerate(["N°","Día y experiencia","Ámbito BCEP","Núcleo BCEP","L","ML","IC","N/O","A","% Logro"],1):
        cell = ws_n.cell(2, ci, hdr)
        cell.font = hfont(bold=True, size=9, color="000000")
        cell.fill = hfill(C_HEADER)
        cell.alignment = center()
        cell.border = thin_border()

    for d_idx in range(5):
        key  = DIAS_KEYS[d_idx]
        fd   = form["dias"][d_idx]
        aid  = ai.get(key,{})
        oa   = fd.get("oa") or {}
        oat  = fd.get("oat") or {}
        tema = aid.get("titulo_dia") or fd.get("tema","")
        sname = day_sheets[d_idx]

        # OA row
        r_oa = 3 + d_idx*2
        ws_n.cell(r_oa,1,f"{d_idx+1}.OA").font = hfont(size=9)
        ws_n.cell(r_oa,1).alignment = center()
        ws_n.cell(r_oa,1).border = thin_border()
        ws_n.cell(r_oa,2,f"{DIAS_NAMES[d_idx]} — {tema}").font = hfont(size=9)
        ws_n.cell(r_oa,2).alignment = left_mid()
        ws_n.cell(r_oa,2).border = thin_border()
        ws_n.cell(r_oa,3,oa.get("ambito",fd.get("ambito",""))).font = hfont(size=9)
        ws_n.cell(r_oa,3).alignment = left_mid()
        ws_n.cell(r_oa,3).border = thin_border()
        ws_n.cell(r_oa,4,oa.get("nucleo","")).font = hfont(size=9)
        ws_n.cell(r_oa,4).alignment = left_mid()
        ws_n.cell(r_oa,4).border = thin_border()

        for sub_idx, cod in enumerate(["L","ML","IC","N/O","A"]):
            # Suma de los 2 indicadores OA (columnas C y D de la hoja diaria)
            formula = f"=COUNTIF('{sname}'!C10:D39,\"{cod}\")"
            cell = ws_n.cell(r_oa, 5+sub_idx, formula)
            cell.font = hfont(size=9)
            cell.alignment = center()
            cell.border = thin_border()

        pct = f"=IFERROR((E{r_oa}+F{r_oa})/(E{r_oa}+F{r_oa}+G{r_oa}+H{r_oa})*100,0)"
        ws_n.cell(r_oa,10,pct).font = hfont(size=9)
        ws_n.cell(r_oa,10).number_format = "0.0\"%\""
        ws_n.cell(r_oa,10).alignment = center()
        ws_n.cell(r_oa,10).border = thin_border()

        # OAT row
        r_oat = r_oa + 1
        ws_n.cell(r_oat,1,f"{d_idx+1}.OAT").font = hfont(size=9)
        ws_n.cell(r_oat,1).alignment = center()
        ws_n.cell(r_oat,1).border = thin_border()
        ws_n.cell(r_oat,2,f"{DIAS_NAMES[d_idx]} — {tema}").font = hfont(size=9)
        ws_n.cell(r_oat,2).alignment = left_mid()
        ws_n.cell(r_oat,2).border = thin_border()
        ws_n.cell(r_oat,3,"Desarrollo Personal y Social (OAT)").font = hfont(size=9)
        ws_n.cell(r_oat,3).alignment = left_mid()
        ws_n.cell(r_oat,3).border = thin_border()
        ws_n.cell(r_oat,4,oat.get("nucleo","")).font = hfont(size=9)
        ws_n.cell(r_oat,4).alignment = left_mid()
        ws_n.cell(r_oat,4).border = thin_border()

        for sub_idx, cod in enumerate(["L","ML","IC","N/O","A"]):
            formula = f"=COUNTIF('{sname}'!E10:E39,\"{cod}\")"
            cell = ws_n.cell(r_oat, 5+sub_idx, formula)
            cell.font = hfont(size=9)
            cell.alignment = center()
            cell.border = thin_border()

        pct_oat = f"=IFERROR((E{r_oat}+F{r_oat})/(E{r_oat}+F{r_oat}+G{r_oat}+H{r_oat})*100,0)"
        ws_n.cell(r_oat,10,pct_oat).font = hfont(size=9)
        ws_n.cell(r_oat,10).number_format = "0.0\"%\""
        ws_n.cell(r_oat,10).alignment = center()
        ws_n.cell(r_oat,10).border = thin_border()

    # Fila TOTAL
    r_tot_n = 13
    ws_n.merge_cells(f"A{r_tot_n}:D{r_tot_n}")
    ws_n.cell(r_tot_n,1,"TOTAL").font = hfont(bold=True, size=9, color="FFFFFF")
    ws_n.cell(r_tot_n,1).fill = hfill("185FA5")
    ws_n.cell(r_tot_n,1).alignment = center()
    ws_n.cell(r_tot_n,1).border = thin_border()
    for ci in range(5,11):
        col = get_column_letter(ci)
        formula = f"=SUM({col}3:{col}{r_tot_n-1})"
        cell = ws_n.cell(r_tot_n, ci, formula)
        cell.font = hfont(bold=True, size=9, color="FFFFFF")
        cell.fill = hfill("185FA5")
        cell.alignment = center()
        cell.border = thin_border()

    # Guardar
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


@app.route(route="download-evaluacion", methods=["POST","OPTIONS"])
def download_evaluacion(req: func.HttpRequest) -> func.HttpResponse:
    """Genera el .xlsx de evaluación semanal."""
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS)
    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse("Body JSON invalido", status_code=400, headers=CORS)
    try:
        xlsx_bytes = build_evaluacion(body["form"], body["ai"])
    except Exception as e:
        logging.error(f"XLSX evaluacion error: {e}")
        return func.HttpResponse(f"Error generando evaluacion: {e}", status_code=500, headers=CORS)
    fecha    = body["form"].get("fecha_inicio","semana")
    out_hdrs = {**CORS,
                "Content-Disposition": f'attachment; filename="Evaluacion_{fecha}.xlsx"'}
    return func.HttpResponse(
        body=xlsx_bytes, status_code=200,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=out_hdrs
    )


# ═══════════════════════════════════════════════════════
# ENDPOINT: SUGERIR IDEAS
# ═══════════════════════════════════════════════════════

SUGGEST_PROMPT = """Eres una experta en Educación Parvularia chilena para Nivel Medio Menor (niños de 3 años).

Dado un tema del día, genera exactamente 5 ideas de experiencias de aprendizaje.
Responde SOLO con JSON válido sin markdown, sin texto adicional.
Formato:
{
  "ideas": [
    {"titulo": "Título corto de 4-6 palabras", "descripcion": "Descripción de 2-3 oraciones explicando la experiencia, materiales y aprendizaje esperado"},
    ...5 ideas en total
  ]
}

IMPORTANTE: Títulos creativos y concretos. Descripciones en lenguaje pedagógico apropiado para párvulos."""


@app.route(route="suggest-ideas", methods=["POST","OPTIONS"])
def suggest_ideas(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS)

    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse("Body JSON invalido", status_code=400, headers=CORS)

    tema = body.get("tema", "")
    dia  = body.get("dia", "")
    groq_key = os.environ.get("GROQ_API_KEY", "")

    if not groq_key:
        return func.HttpResponse("Falta GROQ_API_KEY", status_code=401, headers=CORS)

    try:
        client = Groq(api_key=groq_key)
        resp   = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            max_tokens=1000,
            temperature=0.8,
            messages=[
                {"role": "system", "content": SUGGEST_PROMPT},
                {"role": "user",   "content": f"Día: {dia}\nTema: {tema}\n\nGenera 5 ideas de experiencias de aprendizaje para niños de 3 años."}
            ]
        )
        ideas_data = clean_json(resp.choices[0].message.content)
    except Exception as e:
        logging.error(f"Groq suggest error: {e}")
        return func.HttpResponse(f"Error generando ideas: {e}", status_code=500, headers=CORS)

    out_hdrs = {**CORS, "Content-Type": "application/json"}
    return func.HttpResponse(
        body=json.dumps(ideas_data, ensure_ascii=False),
        status_code=200,
        headers=out_hdrs
    )