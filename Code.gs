/* =============================================
   Sheets AI — Assistente Inteligente
   Google Apps Script (backend)
   ============================================= */

const MODELO = 'claude-sonnet-4-6';
const API_URL = 'https://api.anthropic.com/v1/messages';

const SYSTEM_PROMPT = `Você é um assistente especialista em planilhas integrado diretamente no Google Sheets do usuário.

Você tem acesso a ferramentas para ler, escrever, formatar dados, criar gráficos e ordenar planilhas.

Regras:
- Use sempre as ferramentas disponíveis para interagir com a planilha. Nunca dê instruções manuais.
- Antes de escrever ou calcular, leia os dados reais da planilha primeiro.
- Quando pedirem cálculos (soma, média, porcentagem etc.), leia os dados e calcule com precisão.
- Responda sempre em português brasileiro, de forma clara e direta.
- Ao concluir uma ação, confirme o que foi feito em uma frase curta.`;

const FERRAMENTAS = [
  {
    name: 'ler_celulas',
    description: 'Lê os valores de um intervalo de células da planilha ativa. Use antes de qualquer operação para entender os dados.',
    input_schema: {
      type: 'object',
      properties: {
        intervalo: { type: 'string', description: 'Intervalo no formato A1: ex "A1:C10", "B2", "A:A", "1:1"' }
      },
      required: ['intervalo']
    }
  },
  {
    name: 'escrever_valores',
    description: 'Escreve valores em células da planilha. Aceita texto, números e datas.',
    input_schema: {
      type: 'object',
      properties: {
        celula_inicio: { type: 'string', description: 'Célula inicial ex: "A1"' },
        valores: {
          type: 'array',
          description: 'Array 2D com os valores. Ex: [["Nome","Valor"],["João",100]]',
          items: { type: 'array' }
        }
      },
      required: ['celula_inicio', 'valores']
    }
  },
  {
    name: 'aplicar_formula',
    description: 'Aplica uma fórmula do Google Sheets em uma célula.',
    input_schema: {
      type: 'object',
      properties: {
        celula: { type: 'string', description: 'Célula destino ex: "D10"' },
        formula: { type: 'string', description: 'Fórmula com "=", ex: "=SOMA(A1:A10)", "=MÉDIA(B2:B20)", "=CONT.SE(A:A,\\"João\\")"' }
      },
      required: ['celula', 'formula']
    }
  },
  {
    name: 'formatar_celulas',
    description: 'Formata células: cor de fundo, cor do texto, negrito, itálico, alinhamento e formato de número.',
    input_schema: {
      type: 'object',
      properties: {
        intervalo: { type: 'string', description: 'Intervalo a formatar ex: "A1:D1"' },
        cor_fundo: { type: 'string', description: 'Cor de fundo em hex ex: "#1B6EC2" azul, "#FF0000" vermelho, "#00FF00" verde' },
        cor_texto: { type: 'string', description: 'Cor do texto em hex ex: "#FFFFFF" branco' },
        negrito: { type: 'boolean', description: 'true para negrito' },
        italico: { type: 'boolean', description: 'true para itálico' },
        alinhamento: { type: 'string', enum: ['left', 'center', 'right'], description: 'Alinhamento horizontal' },
        formato_numero: {
          type: 'string',
          description: 'Formato de número. Exemplos: "R$ #,##0.00" moeda, "dd/mm/yyyy" data, "0.00%" porcentagem, "#,##0" inteiro'
        },
        tamanho_fonte: { type: 'number', description: 'Tamanho da fonte em pontos ex: 12' }
      },
      required: ['intervalo']
    }
  },
  {
    name: 'info_planilha',
    description: 'Retorna informações da planilha atual: nome, tamanho, dados. Use sempre primeiro para entender o contexto.',
    input_schema: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'criar_grafico',
    description: 'Cria um gráfico baseado em um intervalo de dados da planilha.',
    input_schema: {
      type: 'object',
      properties: {
        intervalo: { type: 'string', description: 'Intervalo com dados incluindo cabeçalhos ex: "A1:B10"' },
        tipo: {
          type: 'string',
          enum: ['COLUMN', 'LINE', 'PIE', 'BAR', 'AREA'],
          description: 'COLUMN=barras verticais, LINE=linhas, PIE=pizza, BAR=barras horizontais, AREA=área'
        },
        titulo: { type: 'string', description: 'Título do gráfico' }
      },
      required: ['intervalo', 'tipo']
    }
  },
  {
    name: 'ordenar_dados',
    description: 'Ordena os dados de um intervalo por uma coluna específica.',
    input_schema: {
      type: 'object',
      properties: {
        intervalo: { type: 'string', description: 'Intervalo a ordenar ex: "A1:D100"' },
        coluna: { type: 'number', description: 'Número da coluna no intervalo (1 = primeira)' },
        crescente: { type: 'boolean', description: 'true para crescente, false para decrescente' }
      },
      required: ['intervalo', 'coluna']
    }
  }
];

/* ---- Menu no Google Sheets ---- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('🤖 Abrir Assistente AI', 'abrirSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function abrirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('🤖 Sheets AI')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ---- Gerenciar API Key (salva no servidor, não no navegador) ---- */
function salvarApiKey(key) {
  PropertiesService.getUserProperties().setProperty('CLAUDE_API_KEY', key);
  return true;
}

function temApiKey() {
  return !!PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');
}

/* ---- Processar mensagem (loop completo server-side) ---- */
function processarMensagem(historico) {
  const apiKey = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('API key não configurada');

  let mensagens = historico;
  let acoesRealizadas = [];

  // Loop do agente
  for (let tentativa = 0; tentativa < 10; tentativa++) {
    const resposta = chamarClaude(mensagens, apiKey);

    if (resposta.stop_reason === 'end_turn') {
      const blocoTexto = resposta.content.find(function(b) { return b.type === 'text'; });
      return {
        texto: blocoTexto ? blocoTexto.text : 'Concluído.',
        acoes: acoesRealizadas
      };
    }

    if (resposta.stop_reason === 'tool_use') {
      mensagens.push({ role: 'assistant', content: resposta.content });

      const resultados = [];
      const ferramentasUsadas = resposta.content.filter(function(b) { return b.type === 'tool_use'; });

      for (let i = 0; i < ferramentasUsadas.length; i++) {
        const f = ferramentasUsadas[i];
        acoesRealizadas.push(f.name);

        try {
          const resultado = executarFerramenta(f.name, f.input);
          resultados.push({
            type: 'tool_result',
            tool_use_id: f.id,
            content: JSON.stringify(resultado)
          });
        } catch (err) {
          resultados.push({
            type: 'tool_result',
            tool_use_id: f.id,
            content: 'Erro: ' + err.message,
            is_error: true
          });
        }
      }

      mensagens.push({ role: 'user', content: resultados });
    }
  }

  return { texto: 'Processamento concluído.', acoes: acoesRealizadas };
}

/* ---- Chamar API do Claude ---- */
function chamarClaude(mensagens, apiKey) {
  const payload = {
    model: MODELO,
    max_tokens: 4096,
    system: SYSTEM_PROMPT,
    tools: FERRAMENTAS,
    messages: mensagens
  };

  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(API_URL, options);

  if (response.getResponseCode() !== 200) {
    const erro = JSON.parse(response.getContentText());
    throw new Error(erro.error ? erro.error.message : 'Erro HTTP ' + response.getResponseCode());
  }

  return JSON.parse(response.getContentText());
}

/* ---- Ferramentas do Google Sheets ---- */
function executarFerramenta(nome, params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  switch (nome) {

    case 'ler_celulas': {
      const range = sheet.getRange(params.intervalo);
      return {
        intervalo: params.intervalo,
        valores: range.getValues(),
        formatos: range.getNumberFormats()
      };
    }

    case 'escrever_valores': {
      const inicio = sheet.getRange(params.celula_inicio);
      const linhas = params.valores.length;
      const colunas = params.valores[0].length;
      sheet.getRange(inicio.getRow(), inicio.getColumn(), linhas, colunas)
           .setValues(params.valores);
      return { sucesso: true, celulas: linhas * colunas };
    }

    case 'aplicar_formula': {
      const cell = sheet.getRange(params.celula);
      cell.setFormula(params.formula);
      SpreadsheetApp.flush();
      return { sucesso: true, resultado: cell.getValue() };
    }

    case 'formatar_celulas': {
      const range = sheet.getRange(params.intervalo);
      if (params.cor_fundo)    range.setBackground(params.cor_fundo);
      if (params.cor_texto)    range.setFontColor(params.cor_texto);
      if (params.negrito !== undefined) range.setFontWeight(params.negrito ? 'bold' : 'normal');
      if (params.italico !== undefined) range.setFontStyle(params.italico ? 'italic' : 'normal');
      if (params.alinhamento)  range.setHorizontalAlignment(params.alinhamento);
      if (params.tamanho_fonte) range.setFontSize(params.tamanho_fonte);
      if (params.formato_numero) range.setNumberFormat(params.formato_numero);
      return { sucesso: true };
    }

    case 'info_planilha': {
      const dataRange = sheet.getDataRange();
      return {
        nome: sheet.getName(),
        total_linhas: dataRange.getNumRows(),
        total_colunas: dataRange.getNumColumns(),
        intervalo_usado: dataRange.getA1Notation(),
        previa_dados: dataRange.getValues().slice(0, 8)
      };
    }

    case 'criar_grafico': {
      const range = sheet.getRange(params.intervalo);
      const tipoMap = {
        'COLUMN': Charts.ChartType.COLUMN,
        'LINE':   Charts.ChartType.LINE,
        'PIE':    Charts.ChartType.PIE,
        'BAR':    Charts.ChartType.BAR,
        'AREA':   Charts.ChartType.AREA
      };

      let builder = sheet.newChart()
        .setChartType(tipoMap[params.tipo] || Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(3, 6, 0, 0);

      if (params.titulo) {
        builder = builder.setOption('title', params.titulo);
      }

      sheet.insertChart(builder.build());
      return { sucesso: true, tipo: params.tipo };
    }

    case 'ordenar_dados': {
      const range = sheet.getRange(params.intervalo);
      range.sort({ column: range.getColumn() + params.coluna - 1, ascending: params.crescente !== false });
      return { sucesso: true };
    }

    default:
      throw new Error('Ferramenta desconhecida: ' + nome);
  }
}
