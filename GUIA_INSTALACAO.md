# Guia de instalação – RNCSSV Autoavaliação (Python)

Aplicação de autoavaliação do **Referencial Nacional de Certificação de Sustentabilidade do Setor Vitivinícola** (Julho 2025).

---

## O que precisa

- Um computador com **Windows**, **Mac** ou **Linux**
- **Python 3.9 ou superior** (se ainda não tiver, veja o Passo 0)
- Ligação à internet **só na instalação** (depois a autoavaliação funciona offline; a IA precisa de internet)

---

## Passo 0 – Verificar se tem Python

### Windows

1. Carregue na tecla **Windows**, escreva `cmd` e abra o **Prompt de Comando**
2. Escreva e carregue Enter:

```text
python --version
```

- Se aparecer algo como `Python 3.11.x` ou `Python 3.12.x` → já tem Python. Passe ao **Passo 1**.
- Se aparecer erro (`não é reconhecido` / `not found`) → instale o Python:

  1. Vá a: https://www.python.org/downloads/
  2. Descarregue a versão mais recente
  3. **Importante:** na instalação, marque a opção **"Add python.exe to PATH"**
  4. Clique em Install
  5. Feche e abra de novo o Prompt de Comando
  6. Confirme com `python --version`

### Mac

No Terminal:

```text
python3 --version
```

Se não tiver, instale a partir de https://www.python.org/downloads/ ou com Homebrew: `brew install python`

### Linux

```text
python3 --version
```

Se precisar: `sudo apt install python3 python3-pip` (Ubuntu/Debian)

---

## Passo 1 – Descarregar a aplicação

1. Descarregue a pasta **rncssv-python** (com estes ficheiros):
   - `app.py`
   - `indicators.json`
   - `requirements.txt`
   - `GUIA_INSTALACAO.md` / `COMO_USAR.txt`
2. Guarde-a num sítio fácil, por exemplo:
   - Windows: `C:\Users\O_Seu_Nome\Documents\rncssv-python`
   - Mac: `/Users/O_Seu_Nome/Documents/rncssv-python`

---

## Passo 2 – Instalar as dependências (só uma vez)

### Windows

1. Abra o **Prompt de Comando** (ou PowerShell)
2. Vá até à pasta da aplicação. Exemplo:

```text
cd Documents\rncssv-python
```

(Se a pasta estiver noutro sítio, use o caminho correto. Pode arrastar a pasta para a janela do cmd depois de escrever `cd `.)

3. Instale os pacotes:

```text
pip install -r requirements.txt
```

Aguarde até terminar (pode demorar 1–2 minutos).  
Se der erro com `pip`, tente:

```text
python -m pip install -r requirements.txt
```

### Mac / Linux

```text
cd ~/Documents/rncssv-python
pip3 install -r requirements.txt
```

(ou `python3 -m pip install -r requirements.txt`)

---

## Passo 3 – Abrir a aplicação

Na **mesma pasta**, execute:

### Windows

```text
streamlit run app.py
```

### Mac / Linux

```text
streamlit run app.py
```

(ou `python3 -m streamlit run app.py`)

O que acontece:
- O terminal mostra algumas mensagens
- O **browser abre automaticamente** em `http://localhost:8501`
- Se não abrir sozinho, copie esse endereço e cole no Chrome / Edge / Firefox

**Para fechar a aplicação:** volte ao terminal e carregue em **Ctrl + C**.

---

## Passo 4 – Usar a aplicação

1. No menu lateral esquerdo, indique o **nome da organização** e o **tipo de atividade** (Vinha / Transformação / Ambos)
2. Nos separadores dos 4 domínios, preencha os indicadores (nível + evidências)
3. No separador **OM / NC**, registe observações e não conformidades
4. No separador **Ajuda IA** (opcional):
   - Crie conta gratuita em https://console.groq.com
   - Gere uma API Key
   - Cole a key na aplicação e faça perguntas

Pode **exportar** JSON ou resumo em texto a qualquer momento (barra lateral).

---

## Usar noutro computador

1. Copie a pasta **rncssv-python** inteira (pen, OneDrive, email, etc.)
2. No outro PC, confirme que tem Python (Passo 0)
3. Na pasta, corra de novo:

```text
pip install -r requirements.txt
streamlit run app.py
```

Não precisa de instalar mais nada além do Python e destes pacotes.

Para levar o **progresso** de um PC para o outro:
- Exporte o JSON na barra lateral
- No outro PC, use “Carregar JSON anterior”

---

## Problemas comuns

| Problema | Solução |
|----------|---------|
| `python` não é reconhecido | Reinstale o Python e marque **Add to PATH**. Ou use `py -m pip` / `py -m streamlit` |
| `pip` não é reconhecido | Use `python -m pip install -r requirements.txt` |
| `streamlit` não é reconhecido | Use `python -m streamlit run app.py` |
| Porta 8501 ocupada | Feche outras janelas do Streamlit ou use `streamlit run app.py --server.port 8502` |
| Erro ao instalar pacotes | Atualize o pip: `python -m pip install --upgrade pip` e tente de novo |
| Browser não abre | Abra manualmente: http://localhost:8501 |

---

## Resumo rápido (Windows)

```text
cd Documents\rncssv-python
pip install -r requirements.txt
streamlit run app.py
```

Depois é só usar no browser.
