# -*- coding: utf-8 -*-
"""
RNCSSV – Autoavaliação de Sustentabilidade (Python / Streamlit)
Referencial Nacional de Certificação de Sustentabilidade do Setor Vitivinícola
Julho 2025 + Procedimento P51
"""

import json
import os
import io
from datetime import date
from pathlib import Path

import streamlit as st

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

import zipfile
import tempfile
import re

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    try:
        from PyPDF2 import PdfReader
        HAS_PYPDF = True
    except ImportError:
        HAS_PYPDF = False

try:
    import docx as python_docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="RNCSSV – Autoavaliação",
    page_icon="🍇",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Dados
# ---------------------------------------------------------------------------
DATA_PATH = Path(__file__).parent / "indicators.json"
with open(DATA_PATH, encoding="utf-8") as f:
    DATA = {int(k): v for k, v in json.load(f).items()}

DOMAIN_NAMES = {1: "Gestão e Melhoria Contínua", 2: "Ambiental", 3: "Social", 4: "Económico"}

LEVEL_DESC = {
    0: "Não cumprido / sem evidências",
    1: "Evidências básicas + cumprimento legal mínimo",
    2: "Plano/procedimento documentado + objetivos + monitorização",
    3: "Indicadores de desempenho + revisão periódica + melhoria contínua",
}

DOC_TIPS = {
    1: "Registos básicos, fotos, declarações, cumprimento legal.",
    2: "Procedimento/plano escrito, objetivos, responsáveis, registos de implementação.",
    3: "Relatórios de monitorização, KPIs, atas de revisão, evidências de melhoria.",
}

IMPROVE_TIPS = {
    0: "Identifique as práticas atuais e reúna evidências documentais. Garanta o cumprimento legal.",
    1: "Formalize um plano ou procedimento escrito com objetivos, responsáveis e calendário.",
    2: "Implemente monitorização (KPIs), faça revisões periódicas e documente as melhorias.",
}

# ---------------------------------------------------------------------------
# Estado da sessão
# ---------------------------------------------------------------------------
def init_state():
    defaults = {
        "org_name": "",
        "activity": "ambos",
        "answers": {},      # id -> {level, evidence, na}
        "ncs": [],          # list of dicts
        "api_key": "",
        "chat_history": [],
        "docs_text": {},    # domain_id (str) or "all" -> extracted text
        "docs_files": {},   # domain_id -> list of file names
        "docs_images": {},  # domain_id -> list of {name, data}
        "auto_suggestions": {},  # ind_id -> {level, evidence, reasoning}
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def is_visible(ind):
    act = st.session_state.activity
    if act == "ambos":
        return True
    if act == "vinha":
        return ind["scope"] in ("vinha", "ambos")
    if act == "transf":
        return ind["scope"] in ("transf", "ambos")
    return True


def get_answer(ind_id):
    return st.session_state.answers.get(ind_id, {"level": None, "evidence": "", "na": False})


def set_answer(ind_id, **kwargs):
    cur = get_answer(ind_id)
    cur.update(kwargs)
    st.session_state.answers[ind_id] = cur


def calc_score():
    points = 0
    max_points = 0
    ko_missing = []
    for d in range(1, 5):
        for ch in DATA[d]["chapters"]:
            for ind in ch["indicators"]:
                if not is_visible(ind):
                    continue
                ans = get_answer(ind["id"])
                if ans.get("na"):
                    continue
                max_points += 3
                lvl = ans.get("level")
                if lvl is not None and lvl > 0:
                    points += lvl
                if ind["ko"] and (lvl is None or lvl < 1):
                    ko_missing.append(ind["id"])
    pct = round(points / max_points * 100) if max_points else 0
    return points, max_points, pct, ko_missing


def export_json():
    payload = {
        "org_name": st.session_state.org_name,
        "activity": st.session_state.activity,
        "answers": st.session_state.answers,
        "ncs": st.session_state.ncs,
        "date": str(date.today()),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def export_txt():
    pts, mx, pct, ko = calc_score()
    lines = [
        "AUTOAVALIAÇÃO RNCSSV",
        f"Organização: {st.session_state.org_name or '—'}",
        f"Atividade: {st.session_state.activity}",
        f"Data: {date.today().isoformat()}",
        f"Pontuação: {pct}% ({pts} / {mx} pontos)",
        f"KO em falta: {', '.join(ko) if ko else 'Nenhum'}",
        "",
    ]
    for d in range(1, 5):
        lines.append(f"=== DOMÍNIO {d}: {DATA[d]['name']} ===")
        for ch in DATA[d]["chapters"]:
            for ind in ch["indicators"]:
                if not is_visible(ind):
                    continue
                ans = get_answer(ind["id"])
                if not ans.get("level") and not ans.get("na") and not ans.get("evidence"):
                    continue
                ko_tag = " [KO]" if ind["ko"] else ""
                lvl = "N/A" if ans.get("na") else (ans.get("level") if ans.get("level") is not None else "—")
                lines.append(f"\n{ind['id']} {ind['title']}{ko_tag}")
                lines.append(f"  Nível: {lvl}")
                if ans.get("evidence"):
                    lines.append(f"  Evidências: {ans['evidence']}")
        lines.append("")
    if st.session_state.ncs:
        lines.append("=== OM / NÃO CONFORMIDADES ===")
        for n in st.session_state.ncs:
            lines.append(f"\n[{n['type'].upper()}] {n['desc']}")
            if n.get("ind"):
                lines.append(f"  Indicador: {n['ind']}")
            lines.append(f"  Estado: {n['status']} | Prazo: {n.get('deadline') or '—'}")
            if n.get("action"):
                lines.append(f"  Ações: {n['action']}")
    return "\n".join(lines)


def _iter_answered_indicators():
    """Yield (domain, chapter, indicator, answer) for visible indicators with data."""
    for d in range(1, 5):
        for ch in DATA[d]["chapters"]:
            for ind in ch["indicators"]:
                if not is_visible(ind):
                    continue
                ans = get_answer(ind["id"])
                yield d, ch, ind, ans


def build_excel() -> bytes:
    """Gera Excel alinhado com o formulário oficial VINIP03."""
    if not HAS_XLSX:
        raise RuntimeError("Pacote openpyxl não instalado.")

    wb = openpyxl.Workbook()
    header_fill = PatternFill("solid", fgColor="2D6A4F")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    ko_fill = PatternFill("solid", fgColor="FFEDD5")
    title_font = Font(bold=True, size=13, color="2D6A4F")
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")

    pts, mx, pct, ko_list = calc_score()
    act_label = {
        "ambos": "Vinha e Transformação",
        "vinha": "Vinha",
        "transf": "Transformação",
    }.get(st.session_state.activity, st.session_state.activity)
    scope_label = {
        "vinha": "Vinha",
        "transf": "Transformação",
        "ambos": "Vinha e Transformação",
    }

    # --- Folha Autoavaliação (formato VINIP03) ---
    ws = wb.active
    ws.title = "Autoavaliação VINIP03"

    ws["A1"] = "Referencial Nacional de Certificação em Sustentabilidade do Sector Vitivinícola"
    ws["A1"].font = title_font
    ws.merge_cells("A1:I1")
    ws["A2"] = "Formulário de Autoavaliação (compatível VINIP03)"
    ws["A3"] = "EMPRESA / ORGANIZAÇÃO:"
    ws["B3"] = st.session_state.org_name or ""
    ws["A4"] = "Âmbito / Atividade:"
    ws["B4"] = act_label
    ws["A5"] = "Data:"
    ws["B5"] = date.today().isoformat()
    ws["A6"] = "Pontuação obtida:"
    ws["B6"] = f"{pct}%  ({pts} / {mx} pontos)"
    ws["A7"] = "KO em falta:"
    ws["B7"] = ", ".join(ko_list) if ko_list else "Nenhum"
    ws["A8"] = "Nota: KO=1 significa indicador obrigatório (mínimo nível 1). Aplicável S/N conforme filtro de atividade."

    headers = [
        "DOMÍNIO", "CAPÍTULOS", "APLICAÇÃO", "ÍNDICE", "INDICADORES",
        "KO", "Aplicável S/N", "Pontuação", "Comentários / Origem da Informação (Evidências)",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(10, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    row = 11
    for d, ch, ind, ans in _iter_answered_indicators():
        lvl = ans.get("level")
        na = bool(ans.get("na"))
        aplicavel = "N" if na else "S"
        if na:
            pontuacao = ""
        elif lvl is not None and lvl > 0:
            pontuacao = lvl
        elif lvl == 0:
            pontuacao = 0
        else:
            pontuacao = ""

        values = [
            DATA[d]["name"],
            ch["title"],
            scope_label.get(ind["scope"], ind["scope"]),
            ind["id"],
            ind["title"],
            1 if ind["ko"] else "",
            aplicavel,
            pontuacao,
            ans.get("evidence") or "",
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row, col, val)
            cell.border = thin
            cell.alignment = wrap
            if ind["ko"]:
                cell.fill = ko_fill
        row += 1

    # Totais
    row += 1
    ws.cell(row, 7, "TOTAL pontos obtidos").font = Font(bold=True)
    ws.cell(row, 8, pts).font = Font(bold=True)
    row += 1
    ws.cell(row, 7, "TOTAL pontos potenciais").font = Font(bold=True)
    ws.cell(row, 8, mx).font = Font(bold=True)
    row += 1
    ws.cell(row, 7, "% Sustentabilidade").font = Font(bold=True)
    ws.cell(row, 8, f"{pct}%").font = Font(bold=True)
    row += 1
    ws.cell(row, 7, "Nível (A/B/C)").font = Font(bold=True)
    # A>=75, B>=50, C<50 roughly from referential figure - user had levels
    if pct >= 75:
        nivel_final = "A"
    elif pct >= 50:
        nivel_final = "B"
    else:
        nivel_final = "C (abaixo do mínimo 50%)" if pct < 50 else "B"
    ws.cell(row, 8, nivel_final).font = Font(bold=True)

    widths = [28, 28, 18, 10, 48, 5, 12, 10, 55]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A10:I{max(row - 5, 10)}"
    ws.freeze_panes = "A11"
    ws.row_dimensions[10].height = 30

    # --- Folha OM/NC ---
    ws3 = wb.create_sheet("OM_NC")
    nc_headers = ["Tipo", "Indicador", "Descrição", "Estado", "Prazo", "Ações corretivas"]
    for col, h in enumerate(nc_headers, 1):
        cell = ws3.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
    type_labels = {
        "om": "OM – Observação de Melhoria",
        "ncm-min": "NCm – Não Conformidade Menor",
        "ncm": "NCM – Não Conformidade Maior",
        "ncc": "NCC – Não Conformidade Crítica",
    }
    for r, n in enumerate(st.session_state.ncs, start=2):
        vals = [
            type_labels.get(n.get("type"), n.get("type")),
            n.get("ind") or "",
            n.get("desc") or "",
            n.get("status") or "",
            n.get("deadline") or "",
            n.get("action") or "",
        ]
        for c, v in enumerate(vals, 1):
            cell = ws3.cell(r, c, v)
            cell.border = thin
            cell.alignment = wrap
    for i, w in enumerate([28, 12, 50, 12, 14, 40], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # --- Folha Instruções ---
    ws4 = wb.create_sheet("Notas")
    ws4["A1"] = "Notas sobre este relatório"
    ws4["A1"].font = title_font
    notes = [
        "",
        "Este ficheiro foi gerado pela ferramenta de autoavaliação RNCSSV.",
        "Estrutura alinhada com o formulário oficial VINIP03 (Autoavaliação – Referencial Sustentabilidade Vitivinícola).",
        "",
        "Colunas:",
        "• KO = 1 → indicador obrigatório (mínimo Nível 1 para certificação)",
        "• Aplicável S/N → N se marcado como Não aplicável na app",
        "• Pontuação → nível escolhido (0 a 3); vazio se ainda não preenchido",
        "• Comentários / Origem da Informação → texto de evidências escrito na app",
        "",
        "Mínimo para avançar para auditoria (P51): ≥ 50% de pontuação + todos os KO com nível ≥ 1.",
        "Este documento não substitui a auditoria oficial da CERTIS / deliberação ViniPortugal.",
    ]
    for i, line in enumerate(notes, start=2):
        ws4.cell(i, 1, line)
    ws4.column_dimensions["A"].width = 100

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_pdf() -> bytes:
    """Gera relatório PDF com níveis e evidências."""
    if not HAS_PDF:
        raise RuntimeError("Pacote reportlab não instalado. Adicione reportlab ao requirements.txt")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitlePT", parent=styles["Heading1"], fontSize=14, textColor=colors.HexColor("#2D6A4F"), spaceAfter=8
    )
    h2 = ParagraphStyle(
        "H2PT", parent=styles["Heading2"], fontSize=11, textColor=colors.HexColor("#264653"), spaceBefore=10, spaceAfter=4
    )
    body = ParagraphStyle("BodyPT", parent=styles["Normal"], fontSize=8, leading=11)
    small = ParagraphStyle("SmallPT", parent=styles["Normal"], fontSize=7, leading=9, textColor=colors.HexColor("#495057"))

    story = []
    pts, mx, pct, ko = calc_score()
    act_label = {"ambos": "Vinha + Transformação", "vinha": "Apenas Vinha", "transf": "Apenas Transformação"}.get(
        st.session_state.activity, st.session_state.activity
    )

    story.append(Paragraph("Autoavaliação RNCSSV – Relatório", title_style))
    story.append(Paragraph(
        f"<b>Organização:</b> {st.session_state.org_name or '—'} &nbsp;&nbsp; "
        f"<b>Atividade:</b> {act_label} &nbsp;&nbsp; "
        f"<b>Data:</b> {date.today().isoformat()}",
        body,
    ))
    story.append(Paragraph(
        f"<b>Pontuação:</b> {pct}% ({pts} / {mx} pontos) &nbsp;&nbsp; "
        f"<b>KO em falta:</b> {', '.join(ko) if ko else 'Nenhum'}",
        body,
    ))
    story.append(Paragraph(
        "Mínimo para certificação: 50% de pontuação + cumprimento de todos os indicadores KO (nível ≥ 1).",
        small,
    ))
    story.append(Spacer(1, 8))

    # Tabela por domínio
    for d in range(1, 5):
        story.append(Paragraph(f"Domínio {d}: {DATA[d]['name']}", h2))
        data_table = [[
            Paragraph("<b>Cód.</b>", small),
            Paragraph("<b>Indicador</b>", small),
            Paragraph("<b>KO</b>", small),
            Paragraph("<b>Nível</b>", small),
            Paragraph("<b>Evidências / Justificação</b>", small),
        ]]
        for ch in DATA[d]["chapters"]:
            for ind in ch["indicators"]:
                if not is_visible(ind):
                    continue
                ans = get_answer(ind["id"])
                lvl = ans.get("level")
                na = bool(ans.get("na"))
                if lvl is None and not na and not (ans.get("evidence") or "").strip():
                    nivel = "—"
                elif na:
                    nivel = "N/A"
                else:
                    nivel = str(lvl) if lvl is not None else "—"
                evid = (ans.get("evidence") or "").strip() or "—"
                data_table.append([
                    Paragraph(ind["id"], small),
                    Paragraph(ind["title"][:60] + ("…" if len(ind["title"]) > 60 else ""), small),
                    Paragraph("Sim" if ind["ko"] else "Não", small),
                    Paragraph(nivel, small),
                    Paragraph(evid[:300] + ("…" if len(evid) > 300 else ""), small),
                ])

        if len(data_table) == 1:
            story.append(Paragraph("Sem indicadores visíveis para esta atividade.", small))
            continue

        col_w = [1.6 * cm, 5.5 * cm, 1.2 * cm, 1.4 * cm, 8.5 * cm]
        t = Table(data_table, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2D6A4F")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8F9FA")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)

    # OM/NC
    if st.session_state.ncs:
        story.append(PageBreak())
        story.append(Paragraph("Observações de Melhoria e Não Conformidades (P51)", h2))
        type_labels = {
            "om": "OM",
            "ncm-min": "NCm",
            "ncm": "NCM",
            "ncc": "NCC",
        }
        nc_data = [[
            Paragraph("<b>Tipo</b>", small),
            Paragraph("<b>Indicador</b>", small),
            Paragraph("<b>Descrição</b>", small),
            Paragraph("<b>Estado</b>", small),
            Paragraph("<b>Prazo</b>", small),
            Paragraph("<b>Ações</b>", small),
        ]]
        for n in st.session_state.ncs:
            nc_data.append([
                Paragraph(type_labels.get(n.get("type"), n.get("type", "")), small),
                Paragraph(n.get("ind") or "—", small),
                Paragraph((n.get("desc") or "—")[:200], small),
                Paragraph(n.get("status") or "—", small),
                Paragraph(n.get("deadline") or "—", small),
                Paragraph((n.get("action") or "—")[:150], small),
            ])
        t2 = Table(nc_data, colWidths=[1.5*cm, 2*cm, 6*cm, 2*cm, 2.2*cm, 4.5*cm], repeatRows=1)
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#264653")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t2)

    story.append(Spacer(1, 16))
    story.append(Paragraph(
        "Documento gerado pela ferramenta de autoavaliação RNCSSV. Não substitui a auditoria oficial CERTIS / ViniPortugal.",
        small,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()



def extract_text_from_bytes(name: str, data: bytes) -> str:
    """Extrai texto de PDF, DOCX, TXT, Excel e tenta OCR em imagens."""
    lower = name.lower()
    try:
        if lower.endswith((".txt", ".md", ".csv", ".log")):
            for enc in ("utf-8", "latin-1", "cp1252"):
                try:
                    return data.decode(enc)
                except UnicodeDecodeError:
                    continue
            return data.decode("utf-8", errors="ignore")

        if lower.endswith(".pdf") and HAS_PYPDF:
            from io import BytesIO
            reader = PdfReader(BytesIO(data))
            parts = []
            for page in reader.pages[:200]:
                try:
                    parts.append(page.extract_text() or "")
                except Exception:
                    pass
            return "\n".join(parts)

        if lower.endswith(".docx") and HAS_DOCX:
            from io import BytesIO
            document = python_docx.Document(BytesIO(data))
            paras = [p.text for p in document.paragraphs if p.text.strip()]
            # tabelas
            for table in document.tables:
                for row in table.rows:
                    paras.append(" | ".join(c.text.strip() for c in row.cells))
            return "\n".join(paras)

        if lower.endswith(".doc"):
            return f"[Ficheiro .doc antigo: {name}. Converta para .docx ou PDF.]"

        # Excel
        if lower.endswith((".xlsx", ".xlsm")) and HAS_XLSX:
            from io import BytesIO
            wb = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
            parts = []
            for sheet in wb.worksheets:
                parts.append(f"--- Folha: {sheet.title} ---")
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i > 500:
                        parts.append("[... linhas adicionais omitidas ...]")
                        break
                    vals = [str(c) if c is not None else "" for c in row]
                    if any(v.strip() for v in vals):
                        parts.append(" | ".join(vals))
            wb.close()
            return "\n".join(parts)

        if lower.endswith(".xls"):
            return f"[Excel .xls antigo: {name}. Guarde como .xlsx para leitura completa.]"

        # Imagens — OCR local se possível; senão marca para OCR via IA
        if lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp")):
            ocr_text = ""
            if HAS_PIL and HAS_TESSERACT:
                try:
                    from io import BytesIO
                    img = Image.open(BytesIO(data))
                    ocr_text = pytesseract.image_to_string(img, lang="por+eng") or ""
                except Exception as e:
                    ocr_text = f"[OCR local falhou: {e}]"
            if ocr_text.strip() and not ocr_text.startswith("[OCR"):
                return f"[OCR imagem {name}]\n{ocr_text}"
            return f"[IMAGEM_PENDENTE_OCR:{name}]"

    except Exception as e:
        return f"[Erro a ler {name}: {e}]"
    return f"[Tipo de ficheiro não processado: {name}]"


def extract_from_zip(zip_bytes: bytes) -> tuple:
    """Devolve (texto_agregado, lista_nomes, lista_imagens).
    lista_imagens = [{name, data}, ...] para OCR via IA se necessário.
    """
    texts = []
    names = []
    images = []
    with tempfile.TemporaryDirectory() as tmp:
        zpath = Path(tmp) / "upload.zip"
        zpath.write_bytes(zip_bytes)
        with zipfile.ZipFile(zpath, "r") as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                fname = info.filename
                if "__MACOSX" in fname or fname.endswith(".DS_Store"):
                    continue
                base = Path(fname).name
                if base.startswith("."):
                    continue
                try:
                    data = zf.read(info)
                except Exception:
                    continue
                if len(data) > 200 * 1024 * 1024:  # 200 MB por ficheiro dentro do ZIP
                    names.append(fname + " [ficheiro demasiado grande, ignorado]")
                    continue
                names.append(fname)
                lower = base.lower()
                if lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp")):
                    images.append({"name": fname, "data": data})
                chunk = extract_text_from_bytes(base, data)
                if chunk.strip():
                    texts.append(f"\n\n===== FICHEIRO: {fname} =====\n{chunk[:20000]}")
    combined = "\n".join(texts)
    if len(combined) > 400000:
        combined = combined[:150000] + "\n\n[... texto truncado por tamanho ...]"
    return combined, names, images


def ocr_images_with_groq(images: list, api_key: str, max_images: int = 40) -> str:
    """Usa modelo de visão da Groq para extrair texto/descrição de fotos."""
    if not images or not api_key:
        return ""
    try:
        from openai import OpenAI
        import base64
    except ImportError:
        return ""

    client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
    parts = []
    for img in images[:max_images]:  # limite de imagens por avaliação
        name = img["name"]
        data = img["data"]
        lower = name.lower()
        mime = "image/jpeg"
        if lower.endswith(".png"):
            mime = "image/png"
        elif lower.endswith(".webp"):
            mime = "image/webp"
        elif lower.endswith(".gif"):
            mime = "image/gif"
        b64 = base64.b64encode(data).decode("ascii")
        # limitar tamanho — se muito grande, saltar
        if len(b64) > 8_000_000:
            parts.append(f"[Imagem demasiado grande para OCR: {name}]")
            continue
        try:
            resp = client.chat.completions.create(
                model="llama-3.2-11b-vision-preview",
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                "És um assistente de auditoria de sustentabilidade vitivinícola. "
                                "Extrai TODO o texto visível nesta imagem (documentos, certificados, "
                                "rótulos, tabelas, atas, registos). Se for foto de campo/instalação, "
                                "descreve de forma objetiva o que mostra de relevante para evidências. "
                                "Responde em português."
                            ),
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:{mime};base64,{b64}"},
                        },
                    ],
                }],
                temperature=0.1,
                max_tokens=1200,
            )
            content = resp.choices[0].message.content or ""
            parts.append(f"\n\n===== OCR IMAGEM: {name} =====\n{content}")
        except Exception as e:
            # fallback: tentar modelo alternativo ou registar erro
            parts.append(f"[OCR IA falhou para {name}: {e}]")
    return "\n".join(parts)


def auto_evaluate_domain(domain_id: int, api_key: str, docs_text: str) -> dict:
    """
    Pede à IA sugestões de nível + justificação para os indicadores visíveis do domínio.
    Devolve {ind_id: {level, evidence, reasoning}}.
    """
    try:
        from openai import OpenAI
    except ImportError:
        return {"__error__": "Instale o pacote openai."}

    inds = []
    for ch in DATA[domain_id]["chapters"]:
        for ind in ch["indicators"]:
            if not is_visible(ind):
                continue
            inds.append(ind)

    if not inds:
        return {}

    # Processar em lotes de 8 indicadores para caber no contexto
    client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
    results = {}
    batch_size = 8
    docs_snippet = docs_text[:100000] if docs_text else "(Sem documentação carregada para este domínio.)"

    for i in range(0, len(inds), batch_size):
        batch = inds[i:i + batch_size]
        ind_list = "\\n".join(
            f"- {ind['id']} | {'KO' if ind['ko'] else 'opcional'} | {ind['title']}"
            for ind in batch
        )
        prompt = f"""És auditor de sustentabilidade vitivinícola (RNCSSV / VINIP03).
Com base na DOCUMENTAÇÃO abaixo, sugere para cada indicador:
- level: 0, 1, 2 ou 3 (ou null se impossível avaliar)
- na: true só se for claramente não aplicável à atividade
- evidence: frase curta com a origem/evidência encontrada na documentação (ou o que falta)
- reasoning: 1-2 frases a justificar o nível

Critérios de nível:
0 = sem evidências / não cumpre
1 = evidências básicas + cumprimento legal mínimo
2 = plano/procedimento documentado + objetivos + monitorização
3 = indicadores de desempenho + revisão + melhoria contínua

Responde APENAS com um JSON array, sem markdown, no formato:
[{{"id":"1.1.1","level":1,"na":false,"evidence":"...","reasoning":"..."}}, ...]

INDICADORES A AVALIAR:
{ind_list}

DOCUMENTAÇÃO:
{docs_snippet}
"""
        try:
            resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": "Responde apenas com JSON válido (array). Sem texto extra."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
                max_tokens=2500,
            )
            raw = resp.choices[0].message.content.strip()
            # limpar possíveis fences
            raw = re.sub(r"^```(?:json)?\\s*", "", raw)
            raw = re.sub(r"\\s*```$", "", raw)
            arr = json.loads(raw)
            if isinstance(arr, list):
                for item in arr:
                    iid = str(item.get("id", "")).strip()
                    if not iid:
                        continue
                    results[iid] = {
                        "level": item.get("level"),
                        "na": bool(item.get("na")),
                        "evidence": (item.get("evidence") or item.get("reasoning") or "")[:800],
                        "reasoning": (item.get("reasoning") or "")[:500],
                    }
        except Exception as e:
            results["__error__"] = f"Erro no lote {i // batch_size + 1}: {e}"
            break

    return results


def apply_suggestions(suggestions: dict, only_empty: bool = True):
    """Aplica sugestões da IA às answers (confirmação do utilizador)."""
    for iid, sug in suggestions.items():
        if iid.startswith("__"):
            continue
        cur = get_answer(iid)
        if only_empty and (cur.get("level") is not None or cur.get("na") or (cur.get("evidence") or "").strip()):
            continue
        level = sug.get("level")
        na = bool(sug.get("na"))
        if na:
            set_answer(iid, na=True, level=None, evidence=sug.get("evidence") or cur.get("evidence") or "")
        elif level is not None:
            try:
                level = int(level)
            except (TypeError, ValueError):
                continue
            if level < 0 or level > 3:
                continue
            evid = sug.get("evidence") or ""
            reason = sug.get("reasoning") or ""
            text_ev = evid
            if reason and reason not in evid:
                text_ev = f"{evid}\\n[IA: {reason}]".strip()
            set_answer(iid, na=False, level=level, evidence=text_ev[:2000])


def ask_groq(question: str, api_key: str) -> str:
    """Chama a API Groq (compatível OpenAI)."""
    try:
        from openai import OpenAI
    except ImportError:
        return "Erro: instale o pacote openai com:  pip install openai"

    client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
    system = (
        "És um assistente especializado no Referencial Nacional de Certificação de "
        "Sustentabilidade do Setor Vitivinícola (RNCSSV, Julho 2025) e no procedimento P51 da CERTIS. "
        "Ajudas operadores a preparar a autoavaliação: dicas de evidências e documentação por indicador e nível, "
        "como subir de nível, interpretação de KO, OM, NCm, NCM e NCC. "
        "Responde em português de Portugal, de forma prática e objetiva."
    )
    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": question},
            ],
            temperature=0.3,
            max_tokens=1500,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"Erro na API: {e}"


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🍇 RNCSSV")
    st.caption("Autoavaliação · Julho 2025")

    st.session_state.org_name = st.text_input("Organização", st.session_state.org_name)
    st.session_state.activity = st.selectbox(
        "Tipo de atividade",
        options=["ambos", "vinha", "transf"],
        format_func=lambda x: {"ambos": "Vinha + Transformação", "vinha": "Apenas Vinha", "transf": "Apenas Transformação"}[x],
        index=["ambos", "vinha", "transf"].index(st.session_state.activity),
    )

    pts, mx, pct, ko_missing = calc_score()
    st.metric("Sustentabilidade", f"{pct}%", f"{pts} / {mx} pontos")
    if ko_missing:
        st.error(f"⚠ {len(ko_missing)} KO em falta: {', '.join(ko_missing[:8])}{'…' if len(ko_missing)>8 else ''}")
    else:
        st.success("✓ Todos os KO cumpridos (ou N/A)")

    st.divider()
    st.subheader("Exportar relatório")
    org_slug = (st.session_state.org_name or "autoavaliacao").replace(" ", "_")[:30]

    xlsx_data, pdf_data = None, None
    xlsx_err, pdf_err = None, None
    if HAS_XLSX:
        try:
            xlsx_data = build_excel()
        except Exception as e:
            xlsx_err = str(e)
    if HAS_PDF:
        try:
            pdf_data = build_pdf()
        except Exception as e:
            pdf_err = str(e)

    st.download_button(
        "⬇ Excel (níveis + evidências)",
        data=xlsx_data or b"",
        file_name=f"RNCSSV_{org_slug}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not xlsx_data,
        use_container_width=True,
    )
    if xlsx_err:
        st.caption(f"Excel: {xlsx_err}")

    st.download_button(
        "⬇ PDF (relatório)",
        data=pdf_data or b"",
        file_name=f"RNCSSV_{org_slug}_{date.today()}.pdf",
        mime="application/pdf",
        disabled=not pdf_data,
        use_container_width=True,
    )
    if pdf_err:
        st.caption(f"PDF: {pdf_err}")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇ JSON", export_json(), file_name=f"RNCSSV_{date.today()}.json", mime="application/json")
    with col2:
        st.download_button("⬇ Texto", export_txt(), file_name=f"RNCSSV_Resumo_{date.today()}.txt", mime="text/plain")

    uploaded = st.file_uploader("Carregar JSON anterior", type=["json"])
    if uploaded:
        try:
            data = json.load(uploaded)
            st.session_state.org_name = data.get("org_name", "")
            st.session_state.activity = data.get("activity", "ambos")
            st.session_state.answers = data.get("answers", {})
            st.session_state.ncs = data.get("ncs", [])
            st.success("Dados carregados.")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao carregar: {e}")

    if st.button("🗑 Limpar tudo", type="secondary"):
        st.session_state.answers = {}
        st.session_state.ncs = []
        st.session_state.chat_history = []
        st.rerun()

# ---------------------------------------------------------------------------
# Tabs principais
# ---------------------------------------------------------------------------
tab_labels = [f"{d}. {DOMAIN_NAMES[d]}" for d in range(1, 5)] + ["OM / NC", "📁 Docs + Autoavaliação", "🤖 Ajuda IA"]
tabs = st.tabs(tab_labels)

# ---- Domínios 1-4 ----
for di, d in enumerate(range(1, 5)):
    with tabs[di]:
        st.header(f"Domínio {d}: {DATA[d]['name']}")
        for ch in DATA[d]["chapters"]:
            visible_inds = [i for i in ch["indicators"] if is_visible(i)]
            if not visible_inds:
                continue
            with st.expander(f"{ch['id']} – {ch['title']}", expanded=False):
                for ind in visible_inds:
                    ans = get_answer(ind["id"])
                    ko_badge = " 🔴 KO" if ind["ko"] else ""
                    scope_badge = {"vinha": " [Vinha]", "transf": " [Transformação]", "ambos": ""}.get(ind["scope"], "")
                    st.markdown(f"**{ind['id']} – {ind['title']}**{ko_badge}{scope_badge}")

                    c1, c2 = st.columns([1, 2])
                    with c1:
                        na = st.checkbox("Não aplicável", value=ans.get("na", False), key=f"na_{ind['id']}")
                        if na != ans.get("na"):
                            set_answer(ind["id"], na=na, level=None if na else ans.get("level"))
                            st.rerun()

                        if not na:
                            options = [0, 1, 2, 3]
                            current = ans.get("level") if ans.get("level") is not None else 0
                            level = st.radio(
                                "Nível",
                                options=options,
                                format_func=lambda x: f"N{x} – {LEVEL_DESC[x]}",
                                index=options.index(current) if current in options else 0,
                                key=f"lvl_{ind['id']}",
                            )
                            if level != ans.get("level"):
                                set_answer(ind["id"], level=level, na=False)

                    with c2:
                        evidence = st.text_area(
                            "Evidências / Justificação N/A",
                            value=ans.get("evidence", ""),
                            key=f"ev_{ind['id']}",
                            height=80,
                            placeholder="Documentos, registos, fotos, análises…",
                        )
                        if evidence != ans.get("evidence"):
                            set_answer(ind["id"], evidence=evidence)

                        # Dicas
                        lvl = ans.get("level")
                        if not ans.get("na"):
                            if lvl is None or lvl == 0:
                                st.info(f"📄 **Para Nível 1:** {DOC_TIPS[1]}")
                                st.caption(f"🚀 {IMPROVE_TIPS[0]}")
                            elif lvl < 3:
                                st.info(f"📄 **Para Nível {lvl+1}:** {DOC_TIPS[lvl+1]}")
                                st.caption(f"🚀 {IMPROVE_TIPS[lvl]}")
                    st.divider()

# ---- OM / NC ----
with tabs[4]:
    st.header("Observações de Melhoria e Não Conformidades (P51)")
    st.markdown(
        """
| Tipo | Significado | Prazo |
|------|-------------|-------|
| **NCC** | Não Conformidade **Crítica** (ilegal / insustentável / perigosa) | 6 meses · bloqueia certificado |
| **NCM** | Não Conformidade **Maior** (ex.: incumprimento KO) | 6 meses · bloqueia / suspende |
| **NCm** | Não Conformidade **Menor** | 12 meses · pode emitir certificado |
| **OM** | Observação de Melhoria | — |
"""
    )

    with st.form("nc_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            nc_type = st.selectbox(
                "Tipo *",
                ["om", "ncm-min", "ncm", "ncc"],
                format_func=lambda x: {
                    "om": "OM – Observação de Melhoria",
                    "ncm-min": "NCm – Não Conformidade Menor",
                    "ncm": "NCM – Não Conformidade Maior",
                    "ncc": "NCC – Não Conformidade Crítica",
                }[x],
            )
            nc_ind = st.text_input("Indicador relacionado (opcional)", placeholder="Ex: 1.1.1")
        with c2:
            nc_deadline = st.date_input("Prazo de resolução", value=None)
            nc_status = st.selectbox("Estado", ["aberta", "em-curso", "encerrada"])
        nc_desc = st.text_area("Descrição *", placeholder="Descreva a situação encontrada…")
        nc_action = st.text_area("Ações corretivas / Plano de ação")
        submitted = st.form_submit_button("➕ Adicionar registo")
        if submitted:
            if not nc_desc.strip():
                st.warning("A descrição é obrigatória.")
            else:
                st.session_state.ncs.append({
                    "id": str(date.today()) + str(len(st.session_state.ncs)),
                    "type": nc_type,
                    "ind": nc_ind.strip(),
                    "desc": nc_desc.strip(),
                    "deadline": str(nc_deadline) if nc_deadline else "",
                    "status": nc_status,
                    "action": nc_action.strip(),
                })
                st.success("Registo adicionado.")
                st.rerun()

    st.subheader("Registos")
    if not st.session_state.ncs:
        st.info("Ainda não existem registos de OM / NC.")
    else:
        type_labels = {
            "om": "OM",
            "ncm-min": "NCm",
            "ncm": "NCM",
            "ncc": "NCC",
        }
        for i, n in enumerate(st.session_state.ncs):
            color = {"ncc": "🔴", "ncm": "🟠", "ncm-min": "🟡", "om": "🔵"}.get(n["type"], "⚪")
            with st.container():
                cols = st.columns([6, 1])
                with cols[0]:
                    st.markdown(
                        f"{color} **{type_labels.get(n['type'], n['type'])}** — {n['desc']}\n\n"
                        f"Indicador: {n.get('ind') or '—'} · Estado: {n['status']} · Prazo: {n.get('deadline') or '—'}\n\n"
                        f"Ações: {n.get('action') or '—'}"
                    )
                with cols[1]:
                    if st.button("✕", key=f"del_nc_{i}"):
                        st.session_state.ncs.pop(i)
                        st.rerun()
                st.divider()

# ---- Docs + Autoavaliação ----
with tabs[5]:
    st.header("📁 Documentação e Autoavaliação automática")
    st.markdown(
        """
Carregue **pastas ZIP** com a documentação de cada domínio (ou um ZIP geral).
A IA lê o texto dos ficheiros (PDF, DOCX, TXT, etc.), sugere **níveis** e **justificações**.
No fim **confirma ou altera** manualmente em cada indicador.
"""
    )
    st.info(
        "Formatos lidos: PDF, DOCX, TXT, MD, CSV, **Excel (.xlsx)** e **fotos** (OCR). "
        "Limite de upload: **até ~1 GB por ZIP** (ficheiros individuais até 200 MB). "
        "Fotos: OCR local ou via IA Groq. "
        "Textos muito longos são resumidos para a IA (limites da API)."
    )

    api_key_docs = st.text_input(
        "API Key Groq (necessária para autoavaliação)",
        value=st.session_state.api_key,
        type="password",
        key="api_key_docs",
        placeholder="gsk_...",
    )
    if api_key_docs:
        st.session_state.api_key = api_key_docs

    st.subheader("1. Carregar ZIPs de documentação")
    cols = st.columns(5)
    domain_upload_labels = {
        0: ("all", "Geral (todos)"),
        1: (1, "1. Gestão"),
        2: (2, "2. Ambiental"),
        3: (3, "3. Social"),
        4: (4, "4. Económico"),
    }
    for i, (key, label) in domain_upload_labels.items():
        with cols[i]:
            up = st.file_uploader(f"ZIP {label}", type=["zip"], key=f"zip_{key}", help="Até cerca de 1 GB por ZIP (limite do servidor)")
            if up is not None:
                raw = up.read()
                with st.spinner(f"A extrair {label}…"):
                    combined, names, images = extract_from_zip(raw)
                sk = str(key)
                st.session_state.docs_text[sk] = combined
                st.session_state.docs_files[sk] = names
                st.session_state.docs_images[sk] = images
                img_msg = f" · {len(images)} imagens" if images else ""
                st.success(f"{len(names)} ficheiros · {len(combined):,} caracteres{img_msg}")

    # Resumo do que está carregado
    if st.session_state.docs_text:
        st.subheader("Documentação carregada")
        for sk, txt in st.session_state.docs_text.items():
            label = domain_upload_labels.get(
                int(sk) if sk.isdigit() else 0, (sk, sk)
            )[1] if sk != "all" else "Geral (todos)"
            if sk == "all":
                label = "Geral (todos)"
            elif sk.isdigit():
                label = f"{sk}. {DOMAIN_NAMES.get(int(sk), '')}"
            files = st.session_state.docs_files.get(sk, [])
            with st.expander(f"{label} — {len(files)} ficheiros, {len(txt):,} caracteres"):
                st.caption(", ".join(files[:30]) + ("…" if len(files) > 30 else ""))
                st.text_area("Pré-visualização", txt[:3000], height=120, disabled=True, key=f"prev_{sk}")

        if st.button("🗑 Limpar documentação carregada"):
            st.session_state.docs_text = {}
            st.session_state.docs_files = {}
            st.session_state.docs_images = {}
            st.session_state.auto_suggestions = {}
            st.rerun()

    st.subheader("2. Correr autoavaliação com IA")
    only_empty = st.checkbox("Preencher só indicadores ainda vazios", value=True)
    domains_to_run = st.multiselect(
        "Domínios a avaliar",
        options=[1, 2, 3, 4],
        default=[1, 2, 3, 4],
        format_func=lambda d: f"{d}. {DOMAIN_NAMES[d]}",
    )

    if st.button("🚀 Avaliar automaticamente", type="primary", use_container_width=True):
        if not (st.session_state.api_key or "").strip():
            st.error("Indique a API Key Groq.")
        elif not st.session_state.docs_text:
            st.error("Carregue pelo menos um ZIP de documentação.")
        elif not domains_to_run:
            st.error("Selecione pelo menos um domínio.")
        else:
            all_sug = dict(st.session_state.auto_suggestions)
            progress = st.progress(0)
            status = st.empty()
            n = len(domains_to_run)
            for idx_d, d in enumerate(domains_to_run):
                status.write(f"A avaliar domínio {d}. {DOMAIN_NAMES[d]}…")
                # juntar texto geral + do domínio
                parts = []
                if "all" in st.session_state.docs_text:
                    parts.append(st.session_state.docs_text["all"])
                if str(d) in st.session_state.docs_text:
                    parts.append(st.session_state.docs_text[str(d)])
                # OCR de imagens (IA) se ainda não estiver no texto
                imgs = []
                if "all" in st.session_state.docs_images:
                    imgs.extend(st.session_state.docs_images["all"])
                if str(d) in st.session_state.docs_images:
                    imgs.extend(st.session_state.docs_images[str(d)])
                # só imagens ainda pendentes (sem OCR local bem-sucedido)
                pending = [im for im in imgs if im.get("name")]
                if pending and st.session_state.api_key:
                    status.write(f"OCR de {min(len(pending), 15)} imagens (domínio {d})…")
                    ocr_extra = ocr_images_with_groq(pending, st.session_state.api_key.strip())
                    if ocr_extra:
                        parts.append(ocr_extra)
                docs = "\n\n".join(parts) if parts else ""
                if not docs.strip():
                    st.warning(f"Sem documentação específica para o domínio {d} (a usar texto geral se existir).")
                sug = auto_evaluate_domain(d, st.session_state.api_key.strip(), docs)
                if "__error__" in sug:
                    st.error(sug["__error__"])
                for k, v in sug.items():
                    if not k.startswith("__"):
                        all_sug[k] = v
                progress.progress((idx_d + 1) / n)
            st.session_state.auto_suggestions = all_sug
            status.write("Concluído. Reveja as sugestões abaixo e aplique.")
            st.success(f"Sugestões geradas para {len([k for k in all_sug if not k.startswith('__')])} indicadores.")

    if st.session_state.auto_suggestions:
        st.subheader("3. Rever sugestões e confirmar")
        sug = {k: v for k, v in st.session_state.auto_suggestions.items() if not k.startswith("__")}
        # tabela resumo
        rows = []
        for iid, s in sorted(sug.items()):
            rows.append({
                "Indicador": iid,
                "Nível sugerido": "N/A" if s.get("na") else s.get("level"),
                "Justificação": (s.get("reasoning") or s.get("evidence") or "")[:200],
            })
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Aplicar sugestões aos indicadores", type="primary", use_container_width=True):
                apply_suggestions(st.session_state.auto_suggestions, only_empty=only_empty)
                st.success("Sugestões aplicadas. Vá aos separadores dos domínios para confirmar ou ajustar.")
                st.rerun()
        with c2:
            if st.button("🗑 Descartar sugestões", use_container_width=True):
                st.session_state.auto_suggestions = {}
                st.rerun()

        st.caption(
            "Depois de aplicar, percorra os separadores 1–4, valide cada nível e edite as evidências se necessário. "
            "No fim exporte Excel/PDF."
        )

# ---- IA ----
with tabs[6]:
    st.header("🤖 Assistente de Ajuda (IA)")
    st.markdown(
        """
Esta secção usa a **API gratuita da Groq** (modelo Llama) para responder a dúvidas sobre o referencial.

**Como obter uma API key (grátis):**
1. Vá a [https://console.groq.com](https://console.groq.com) e crie conta
2. Em **API Keys** crie uma nova key
3. Cole a key abaixo (fica apenas nesta sessão / neste PC)
"""
    )
    api_key = st.text_input("API Key Groq", value=st.session_state.api_key, type="password", placeholder="gsk_...")
    st.session_state.api_key = api_key

    # Histórico
    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    question = st.chat_input("Faça uma pergunta… Ex: Que evidências preciso para o 2.3.2 no Nível 2?")
    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)

        if not api_key.strip():
            answer = "⚠️ Cole primeiro a sua API Key Groq no campo acima."
        else:
            with st.spinner("A pensar…"):
                answer = ask_groq(question, api_key.strip())

        st.session_state.chat_history.append({"role": "assistant", "content": answer})
        with st.chat_message("assistant"):
            st.markdown(answer)

    st.caption(
        "A key é usada apenas para chamar a API da Groq a partir do seu PC. "
        "Se não quiser usar IA, ignore esta secção — a autoavaliação funciona na mesma."
    )

# ---------------------------------------------------------------------------
# Rodapé
# ---------------------------------------------------------------------------
st.divider()
st.caption(
    "Ferramenta de apoio à autoavaliação · RNCSSV Julho 2025 + P51 · "
    "Não substitui a auditoria oficial CERTIS / ViniPortugal."
)
