# Sheets AI — Como instalar

## Passo a passo (5 minutos)

### 1. Abrir o Google Apps Script
- Acesse: https://script.google.com
- Clique em "Novo projeto"

### 2. Copiar os arquivos

**Code.gs** (já existe por padrão):
- Apague o conteúdo e cole todo o conteúdo do arquivo `Code.gs`

**Sidebar.html** (criar novo):
- Clique em "+" → "Arquivo HTML"
- Nome: `Sidebar` (sem extensão)
- Cole o conteúdo do arquivo `Sidebar.html`

**appsscript.json** (ativar manifesto):
- Vá em: Configurações do projeto ⚙️ → ative "Mostrar arquivo de manifesto"
- Clique em `appsscript.json` e cole o conteúdo do arquivo `appsscript.json`

### 3. Salvar e autorizar
- Clique em 💾 Salvar
- Clique em ▶ Executar → escolha `onOpen`
- Autorize as permissões quando solicitado

### 4. Testar no Google Sheets
- Abra qualquer planilha no Google Sheets
- No menu: **Extensões → Sheets AI → 🤖 Abrir Assistente AI**
- O painel vai abrir na lateral direita

### 5. Configurar a API Key
- Acesse: https://console.anthropic.com
- Crie uma conta (tem crédito gratuito para começar)
- Gere uma API key
- Cole na tela inicial do add-on

---

## Publicar para todos usarem (opcional)
Para disponibilizar na loja do Google:
- No Apps Script: **Implantar → Novo implantação → Add-on do Editor**
- Seguir o processo de publicação no Google Workspace Marketplace

---

## O que o assistente já faz
- Lê e analisa dados da planilha
- Escreve valores automaticamente
- Aplica fórmulas (SOMA, MÉDIA, CONT.SE, etc.)
- Formata células (cores, negrito, moeda, datas)
- Cria gráficos (barras, linhas, pizza, área)
- Ordena dados
- Responde cálculos e perguntas sobre os dados
