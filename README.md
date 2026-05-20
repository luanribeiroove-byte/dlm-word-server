# DLM Word Generator — Servidor de Geração de Relatórios

Servidor Flask que recebe JSON com dados de relatório técnico DLM e devolve um arquivo `.docx` formatado.

## Estrutura

```
dlm-word-server/
├── app.py                     # Servidor Flask (entrypoint)
├── gerar_relatorio.py         # Lógica principal de geração
├── secao4_builder.py          # Construção dinâmica da Seção 4
├── assets/
│   ├── template.docx          # Template Word base
│   └── logo_dlm_transparent.png
├── requirements.txt
├── runtime.txt
├── Procfile
└── README.md
```

## Endpoints

### `GET /` — Health check
Retorna JSON com status do serviço. Útil pro Render verificar se o serviço está vivo.

### `POST /gerar` — Gerar relatório
- **Body:** JSON com os dados (ver formato em `dados_teste.json` no repositório original)
- **Retorna:** arquivo `.docx` (com `Content-Disposition: attachment`)
- **Em caso de erro:** JSON `{"error": "...", "traceback": "..."}` com status 500

### Exemplo de chamada (JavaScript)

```javascript
const dados = {
  cabecalho: { periodo: "março/2026", data_emissao: "25 de março de 2026" },
  introducao: { periodo_intervencoes: "março de 2026", ... },
  inspecoes: [ ... ],
  eficiencia: { houve_campanha: true, ... },
  nao_conformidades: [ ... ],
  consideracoes_paragrafos: [ ... ]
};

const resp = await fetch("https://SEU-SERVICO.onrender.com/gerar", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(dados),
});

const blob = await resp.blob();
// Faça download/preview do blob como .docx
```

## Rodar localmente

```bash
pip install -r requirements.txt
python app.py
# Servidor inicia em http://localhost:5000
```

## Deploy no Render.com

1. Faça upload dessa pasta inteira em um repositório GitHub
2. No Render.com → "New" → "Web Service"
3. Conecte ao repositório
4. Configurações:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --timeout 120`
   - **Python version:** 3.11.9 (do `runtime.txt`)
5. Deploy! Render vai dar uma URL tipo `https://dlm-word-server.onrender.com`

## Nota sobre hibernação

No free tier, o Render hiberna serviços inativos após 15 minutos. A primeira chamada após hibernação demora ~30s. As seguintes são rápidas (~5s).
