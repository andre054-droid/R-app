# RNCSSV – Autoavaliação de Sustentabilidade

Aplicação de autoavaliação do **Referencial Nacional de Certificação de Sustentabilidade do Setor Vitivinícola** (Julho 2025) + procedimento P51.

## Funcionalidades

- 87 indicadores dos 4 domínios
- Filtro Vinha / Transformação / Ambos
- Níveis 0–3, evidências e Não aplicável
- Dicas de documentação e melhoria
- Secção OM / NCm / NCM / NCC
- Cálculo automático da % e alerta de KO
- Exportar / carregar JSON
- Chat com IA (API Groq opcional)

## Como correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy online (Streamlit Cloud)

1. Faça fork ou upload deste repositório para o seu GitHub
2. Vá a [share.streamlit.io](https://share.streamlit.io)
3. New app → escolha este repositório → ficheiro `app.py`
4. Deploy

Ficará com um link público do tipo `https://nome-da-app.streamlit.app`
