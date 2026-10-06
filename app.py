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
tab_labels = [f"{d}. {DOMAIN_NAMES[d]}" for d in range(1, 5)] + ["OM / NC", "🤖 Ajuda IA"]
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

# ---- IA ----
with tabs[5]:
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
