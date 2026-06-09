import logging
import json
import os
import io
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