# -*- coding: utf-8 -*-
"""
RNCSSV – Autoavaliação de Sustentabilidade (Python / Streamlit)
Referencial Nacional de Certificação de Sustentabilidade do Setor Vitivinícola
Julho 2025 + Procedimento P51
"""

import json
import os
from datetime import date
from pathlib import Path

import streamlit as st

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
    st.subheader("Dados")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇ JSON", export_json(), file_name=f"RNCSSV_{date.today()}.json", mime="application/json")
    with col2:
        st.download_button("⬇ Resumo", export_txt(), file_name=f"RNCSSV_Resumo_{date.today()}.txt", mime="text/plain")

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
