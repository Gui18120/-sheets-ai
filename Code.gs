 var MODELO = 'claude-haiku-4-5';
  var API_URL = 'https://api.anthropic.com/v1/messages';
  var SYSTEM_PROMPT = 'Voce e um assistente especialista em planilhas no Google Sheets. ' +
    'Use sempre as ferramentas para interagir com a planilha. ' +
    'Antes de calcular, leia os dados reais da planilha primeiro. ' +
    'Responda sempre em portugues brasileiro, de forma direta. ' +
    'Ao concluir uma acao, confirme em uma frase curta.';

   var FERRAMENTAS = [
    { name: 'ler_celulas', description: 'Le celulas.', input_schema: { type: 'object', properties: { intervalo: { type:
  'string' } }, required: ['intervalo'] } },
    { name: 'escrever_valores', description: 'Escreve valores.', input_schema: { type: 'object', properties: {
  celula_inicio: { type: 'string' }, valores: { type: 'array', items: { type: 'array' } } }, required: ['celula_inicio',
   'valores'] } },
    { name: 'aplicar_formula', description: 'Aplica formula.', input_schema: { type: 'object', properties: { celula: {
  type: 'string' }, formula: { type: 'string' } }, required: ['celula', 'formula'] } },
    { name: 'formatar_celulas', description: 'Formata celulas.', input_schema: { type: 'object', properties: {
  intervalo: { type: 'string' }, cor_fundo: { type: 'string' }, cor_texto: { type: 'string' }, negrito: { type:
  'boolean' }, italico: { type: 'boolean' }, alinhamento: { type: 'string' }, formato_numero: { type: 'string' },
  tamanho_fonte: { type: 'number' } }, required: ['intervalo'] } },
    { name: 'info_planilha', description: 'Info da planilha.', input_schema: { type: 'object', properties: {} } },
    { name: 'criar_grafico', description: 'Cria grafico.', input_schema: { type: 'object', properties: { intervalo: {
  type: 'string' }, tipo: { type: 'string' }, titulo: { type: 'string' } }, required: ['intervalo', 'tipo'] } },
    { name: 'ordenar_dados', description: 'Ordena dados.', input_schema: { type: 'object', properties: { intervalo: {
  type: 'string' }, coluna: { type: 'number' }, crescente: { type: 'boolean' } }, required: ['intervalo', 'coluna'] } },
    { name: 'listar_abas', description: 'Lista abas.', input_schema: { type: 'object', properties: {} } },
    { name: 'trocar_aba', description: 'Troca aba.', input_schema: { type: 'object', properties: { nome_aba: { type:
  'string' } }, required: ['nome_aba'] } },
    { name: 'criar_aba', description: 'Cria aba.', input_schema: { type: 'object', properties: { nome_aba: { type:
  'string' } }, required: ['nome_aba'] } },
    { name: 'voltar_aba', description: 'Volta aba anterior.', input_schema: { type: 'object', properties: {} } }
  ];
  function onOpen() {
    SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Abrir Assistente AI', 'abrirSidebar')
      .addToUi();
  }

  function onInstall() { onOpen(); }

  function abrirSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Sheets AI')
      .setWidth(340);
    SpreadsheetApp.getUi().showSidebar(html);
  }

  function salvarApiKey(key) {
    PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', key);
    return true;
  }

  function temApiKey() {
    return !!PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  }
 function verificarAlertas() {
    var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) throw new Error('API key nao configurada');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dr = sheet.getDataRange();
    var dados = dr.getValues();
    var mensagem = 'Analise esta planilha e encontre problemas nos dados:\n\n';
    mensagem += JSON.stringify(dados.slice(0, 50)) + '\n\n';
    mensagem += 'Procure e liste apenas:\n';
    mensagem += '1. Celulas vazias em colunas que deveriam ter dados\n';
    mensagem += '2. Texto em colunas de numeros\n';
    mensagem += '3. Datas invalidas\n';
    mensagem += '4. Valores duplicados\n';
    mensagem += '5. Numeros muito fora do padrao\n\n';
    mensagem += 'Se nao encontrar problemas, diga que os dados estao ok.\n';
    mensagem += 'IMPORTANTE: Responda em texto simples, sem tabelas nem pipes. Use bullets com >.';
    var options = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: MODELO, max_tokens: 1024, messages: [{ role: 'user', content: mensagem }] }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      var erro = JSON.parse(response.getContentText());
      throw new Error(erro.error ? erro.error.message : 'Erro HTTP ' + response.getResponseCode());
    }
    return JSON.parse(response.getContentText()).content[0].text;
  }
  function gerarRelatorio() {
    var apiKey = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) throw new Error('API key nao configurada');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dr = sheet.getDataRange();
    var dados = dr.getValues();
    var mensagem = 'Analise esta planilha chamada "' + sheet.getName() + '" com ' + dr.getNumRows() + ' linhas e ' +
  dr.getNumColumns() + ' colunas.\n\n';
    mensagem += 'Dados:\n' + JSON.stringify(dados.slice(0, 50)) + '\n\n';
    mensagem += 'Gere um relatorio completo em portugues com:\n';
    mensagem += '1. O que sao esses dados\n';
    mensagem += '2. Totais e soma dos valores numericos\n';
    mensagem += '3. Medias e porcentagens relevantes\n';
    mensagem += '4. Observacoes importantes\n';
    mensagem += '5. Sugestoes de melhoria para a planilha\n\n';
    mensagem += 'IMPORTANTE: Responda em texto simples. Nao use tabelas nem pipes. Use bullets com >.';
    var options = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: MODELO, max_tokens: 2048, messages: [{ role: 'user', content: mensagem }] }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      var erro = JSON.parse(response.getContentText());
      throw new Error(erro.error ? erro.error.message : 'Erro HTTP ' + response.getResponseCode());
    }
    return JSON.parse(response.getContentText()).content[0].text;
  }

  function processarMensagem(historico) {
    // Limita historico a ultimas 10 mensagens para evitar rate limit
    if (historico.length > 6) {
      historico = historico.slice(historico.length - 6);
    }

    var apiKey = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) throw new Error('API key nao configurada');
    var mensagens = historico;
    var acoesRealizadas = [];
    for (var t = 0; t < 10; t++) {
      var resposta = chamarClaude(mensagens, apiKey);
      if (resposta.stop_reason === 'end_turn') {
        var texto = null;
        for (var i = 0; i < resposta.content.length; i++) {
          if (resposta.content[i].type === 'text') { texto = resposta.content[i].text; break; }
        }
        return { texto: texto || 'Concluido.', acoes: acoesRealizadas };
      }
      if (resposta.stop_reason === 'tool_use') {
        mensagens.push({ role: 'assistant', content: resposta.content });
        var resultados = [];
        for (var j = 0; j < resposta.content.length; j++) {
          var bloco = resposta.content[j];
          if (bloco.type !== 'tool_use') continue;
          acoesRealizadas.push(bloco.name);
          try {
            var resultado = executarFerramenta(bloco.name, bloco.input);
            resultados.push({ type: 'tool_result', tool_use_id: bloco.id, content: JSON.stringify(resultado) });
          } catch(err) {
            resultados.push({ type: 'tool_result', tool_use_id: bloco.id, content: 'Erro: ' + err.message, is_error:
  true });
          }
        }
        mensagens.push({ role: 'user', content: resultados });
      }
    }
    return { texto: 'Concluido.', acoes: acoesRealizadas };
  }

  function chamarClaude(mensagens, apiKey) {
    var options = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: MODELO, max_tokens: 4096, system: SYSTEM_PROMPT, tools: FERRAMENTAS, messages:
  mensagens }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      var erro = JSON.parse(response.getContentText());
      throw new Error(erro.error ? erro.error.message : 'Erro HTTP ' + response.getResponseCode());
    }
    return JSON.parse(response.getContentText());
  }
 function salvarBackup() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dr = sheet.getDataRange();
    var backup = {
      nome_aba: sheet.getName(),
      valores: dr.getValues(),
      formatos: dr.getNumberFormats(),
      intervalo: dr.getA1Notation()
    };
    PropertiesService.getUserProperties().setProperty('BACKUP', JSON.stringify(backup));
  }

  function desfazerUltimaAcao() {
    var backupStr = PropertiesService.getUserProperties().getProperty('BACKUP');
    if (!backupStr) throw new Error('Nenhuma acao para desfazer.');
    var backup = JSON.parse(backupStr);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(backup.nome_aba);
    if (!sheet) throw new Error('Aba "' + backup.nome_aba + '" nao encontrada.');
    ss.setActiveSheet(sheet);
    sheet.clearContents();
    var inicio = sheet.getRange(backup.intervalo.split(':')[0]);
    sheet.getRange(inicio.getRow(), inicio.getColumn(), backup.valores.length,
  backup.valores[0].length).setValues(backup.valores);
    return 'Acao desfeita com sucesso!';
  }
  function executarFerramenta(nome, params) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var abaAnterior = PropertiesService.getUserProperties().getProperty('ABA_ANTERIOR');

    if (nome === 'ler_celulas') {
      var r = sheet.getRange(params.intervalo);
      return { intervalo: params.intervalo, valores: r.getValues(), formatos: r.getNumberFormats() };
    }if (nome === 'escrever_valores' || nome === 'aplicar_formula' || nome === 'formatar_celulas' || nome ===
  'ordenar_dados') {
      salvarBackup();
    }
    if (nome === 'escrever_valores') {
      var inicio = sheet.getRange(params.celula_inicio);
      sheet.getRange(inicio.getRow(), inicio.getColumn(), params.valores.length,
  params.valores[0].length).setValues(params.valores);
      return { sucesso: true };
    }
    if (nome === 'aplicar_formula') {
      var cell = sheet.getRange(params.celula);
      cell.setFormula(params.formula);
      SpreadsheetApp.flush();
      return { sucesso: true, resultado: cell.getValue() };
    }
    if (nome === 'formatar_celulas') {
      var r = sheet.getRange(params.intervalo);
      if (params.cor_fundo) r.setBackground(params.cor_fundo);
      if (params.cor_texto) r.setFontColor(params.cor_texto);
      if (params.negrito !== undefined) r.setFontWeight(params.negrito ? 'bold' : 'normal');
      if (params.italico !== undefined) r.setFontStyle(params.italico ? 'italic' : 'normal');
      if (params.alinhamento) r.setHorizontalAlignment(params.alinhamento);
      if (params.tamanho_fonte) r.setFontSize(params.tamanho_fonte);
      if (params.formato_numero) r.setNumberFormat(params.formato_numero);
      return { sucesso: true };
    }
    if (nome === 'info_planilha') {
      var dr = sheet.getDataRange();
      return { nome: sheet.getName(), total_linhas: dr.getNumRows(), total_colunas: dr.getNumColumns(), intervalo_usado:
   dr.getA1Notation(), previa_dados: dr.getValues().slice(0, 8) };
    }
    if (nome === 'criar_grafico') {
      var r = sheet.getRange(params.intervalo);
      var tipos = { 'COLUMN': Charts.ChartType.COLUMN, 'LINE': Charts.ChartType.LINE, 'PIE': Charts.ChartType.PIE,
  'BAR': Charts.ChartType.BAR, 'AREA': Charts.ChartType.AREA };
      var builder = sheet.newChart().setChartType(tipos[params.tipo] ||
  Charts.ChartType.COLUMN).addRange(r).setPosition(3, 6, 0, 0);
      if (params.titulo) builder = builder.setOption('title', params.titulo);
      sheet.insertChart(builder.build());
      return { sucesso: true };
    }
    if (nome === 'ordenar_dados') {
      var r = sheet.getRange(params.intervalo);
      r.sort({ column: r.getColumn() + params.coluna - 1, ascending: params.crescente !== false });
      return { sucesso: true };
    }
    if (nome === 'listar_abas') {
      var sheets = ss.getSheets();
      var nomes = [];
      for (var i = 0; i < sheets.length; i++) { nomes.push(sheets[i].getName()); }
      return { abas: nomes, total: nomes.length };
    }
    if (nome === 'trocar_aba') {
      PropertiesService.getUserProperties().setProperty('ABA_ANTERIOR', sheet.getName());
      var aba = ss.getSheetByName(params.nome_aba);
      if (!aba) throw new Error('Aba "' + params.nome_aba + '" nao encontrada');
      ss.setActiveSheet(aba);
      return { sucesso: true, aba_atual: params.nome_aba };
    }
    if (nome === 'criar_aba') {
      PropertiesService.getUserProperties().setProperty('ABA_ANTERIOR', sheet.getName());
      var novaAba = ss.insertSheet(params.nome_aba);
      ss.setActiveSheet(novaAba);
      return { sucesso: true, aba_criada: params.nome_aba };
    }
    if (nome === 'voltar_aba') {
      if (!abaAnterior) return { aviso: 'Nenhuma aba anterior registrada.' };
      var aba = ss.getSheetByName(abaAnterior);
      if (!aba) throw new Error('Aba anterior "' + abaAnterior + '" nao encontrada');
      ss.setActiveSheet(aba);
      return { sucesso: true, voltou_para: abaAnterior };
    }
    throw new Error('Ferramenta desconhecida: ' + nome);
  }