# 📊 Extração de Sessões via API e Exportação para Excel

Este projeto em Python realiza a extração de conversas de sessões a partir de uma API, organiza os dados por timestamp, agrupa por remetente (user/agent/function) e exporta os resultados em formato Excel (.xlsx), com estrutura customizada.

---

## ⚙️ Funcionalidades

- 🔐 Consumo de dados via API com autenticação por `x-api-key`
- ⏳ Ordenação das mensagens por `timestamp`
- 🗃️ Reorganização da estrutura de conversa (`user`, `agent`, `function`)
- 📤 Exportação para Excel com colunas específicas e layout ajustado
- 📌 Geração de arquivo com nome baseado no timestamp mais antigo da sessão

---

## 🖥️ Como usar

### 1. Clone o repositório

```bash
git clone https://github.com/WilGabrieel/api-extracao-python.git
cd api-extracao-python