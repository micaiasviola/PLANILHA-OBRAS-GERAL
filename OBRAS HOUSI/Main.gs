/*************************
 * GATILHOS PRINCIPAIS
 *************************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const S = CONFIG.SHEETS;

  // Menu operacional (uso diário da equipe).
  ui.createMenu("⚙️ Automacao ECQUA")
    .addItem("🌟 Abrir Painel de Gestão (Dashboard)", "abrirDashboardInterface")
    .addSeparator()
    .addItem("➕ Inserir 1 linha no topo (INFO. GERAIS)", "inserirNovaLinhaTopoInformacoesGerais")
    .addItem("🧱 Gerar templates pendentes FASE-OBRA", "gerarTemplatesPendentesFaseObra")
    .addItem("🚀 Gerar/Atualizar PEDIDOS-GERAL (HOUSI)", "sincronizarTodosPedidosHousi")
    .addItem("🚚 Sincronizar envios para FASE-ENTREGA", "sincronizarTodosEnviosParaFaseEntrega")
    .addItem("📑 Atualizar Informes dos Prestadores", "atualizarInformePrestadores")
    .addItem("⏰ Atualizar Indicador de Atrasos (INFO. GERAIS)", "atualizarIndicadorServicosAtrasados")
    .addItem("📋 Atualizar PENDÊNCIAS GERAIS (DASHBOARD)", "atualizarPendenciasGeraisDashboard")
    .addItem("📑 Ordenar Fase-Preliminar (Base INFO. GERAIS)", "ordenarPreliminarIgualInformacoesGerais")
    .addItem("🔄 Atualizar TODAS AS PENDÊNCIAS (Sinc. Global)", "sincronizacaoManualGlobal")
    .addItem("⏯️ Reordenar FASE-OBRA (Manual)", "executarAtualizarFaseObraDiaria")
    .addSeparator()
    .addToUi();

  // Menu técnico (manutenção/admin).
  ui.createMenu("⚙️ Admin ECQUA")
    .addItem("☑️ Configurar coluna de envio (checkbox)", "configurarColunaEnvioFaseEntrega")
    .addItem("⏰ Criar acionador diário central (01:00)", "criarTriggerDiarioCentralizado01h")
    .addItem("🔄 Forçar Atualização desta Aba (A->B)", "atualizarTodaAAbaAtiva")
    .addItem("💰 Recalcular Verba Teto (W -> X)", "recalcularTodasVerbasTetoFaseObra")
    .addItem("📆 Recalcular indicador de cronograma (OBRA E)", "sincronizarTodosIndicadoresCronogramaFaseObra")
    .addItem("🌐 Sincronizar ocorrências abertas (G)", "sincronizarTodasOcorrenciasAbertasParaPreliminar")
    .addSeparator()
    .addItem("📸 Criar Coluna Foto Tomada (PRELIMINAR)", "configurarColunaFotoTomadaPreliminar")
    .addItem("🧱 Criar Coluna Ativador Manual (FASE-OBRA)", "configurarColunaEnviarObraPreliminar")
    .addItem("🚚 Criar Coluna Envio para Entrega (INF. GERAIS)", "configurarColunaEnviarEntregaInfoGerais")
    .addItem("📊 Criar Coluna Status Obra (INF. GERAIS)", "configurarColunaStatusObraInfoGerais")
    .addItem("📆 Criar Colunas Cronograma (FASE-OBRA)", "configurarColunasCronogramaFaseObra")
    .addItem("📅 Criar Coluna Semana do Mês (FASE-OBRA)", "configurarColunaSemanaMesFaseObra")
    .addSeparator()
    .addItem("📆 Recalcular Semanas Cronograma (OBRA)", "sincronizarTodaAbaObraSemanasCronograma")
    .addItem("📅 Recalcular Semanas do Mês (OBRA)", "sincronizarTodaAbaObraSemanaMes")
    .addSeparator()
    .addItem("⏰ Criar acionador diário FASE-OBRA (03:30)", "criarTriggerDiariaAtualizarFaseObra_")
    .addToUi();
  criarMenuPagamentos();
}

/**
 * Event Router para onEdit.
 * Melhora a performance ao evitar verificações redundantes em cada edição.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const S = CONFIG.SHEETS;

  // Mapa de roteamento por aba
  const handlers = {
    [S.OBRA]: handleObraEdit,
    [S.PEDIDOS]: handlePedidosEdit,
    [S.PRELIMINAR]: handlePreliminarEdit,
    [S.INFO_GERAIS]: handleInfoGeraisEdit,
    [S.OCORRENCIAS]: handleOcorrenciasEdit,
    [S.ENTREGA]: handleEntregaEdit_v2
  };

  if (handlers[sheetName]) {
    try {
      handlers[sheetName](e);
    } catch (err) {
      // Suprimir toast visível quando o erro for de lock do documento — registrar em Logger apenas
      try {
        if (err && err.message && /Não foi possível obter lock do documento/.test(err.message)) {
          Logger.log('Lock ocupado (suprimido toast): ' + err.message);
        } else {
          SpreadsheetApp.getActiveSpreadsheet().toast("Erro na automação: " + err.message, "⚠️ Erro", 5);
        }
      } catch (e) {
        // Evita falha ao tentar mostrar toast em contextos sem UI
        Logger.log('Falha ao exibir toast de erro: ' + (e && e.message));
      }
      console.error(err);
    }
  }
}

/**
 * Funções gerais disparadas pelos menus.
 */
function sincronizacaoManualGlobal() {
  SpreadsheetApp.getUi().alert("Iniciando a Sincronização Global. Por favor, aguarde...");
  executarSincronizacaoFinalDoDia();
  recalcularServicosAtrasados_();
  atualizarInformePrestadores();
  atualizarPendenciasGeraisDashboard();
  SpreadsheetApp.getUi().alert("Pronto! Sincronização Global Concluída.");
}

function atualizarTodaAAbaAtiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getActiveSheet();
  const linhaCorte = obterLinhaInicialPorAba(aba.getName());
  processarIntervaloAparaB_(aba, aba.getRange(linhaCorte, 1, aba.getLastRow() - linhaCorte + 1, 1));
}

/**
 * Cria (ou recria) o acionador que executa às 1h da manhã diariamente.
 * Chame esta função UMA VEZ pelo menu Admin ECQUA.
 */
/**
 * Cria (ou recria) o acionador que executa a sincronização completa às 1h da manhã.
 */
function configurarRotinaSincronizacaoMadrugada() {
  Logger.log('Deprecado: configurarRotinaSincronizacaoMadrugada() — use criarTriggerDiarioCentralizado01h().');
  try {
    if (typeof criarTriggerDiarioCentralizado01h === 'function') {
      criarTriggerDiarioCentralizado01h();
    } else {
      Logger.log('Função criarTriggerDiarioCentralizado01h não encontrada.');
    }
  } catch (e) {
    try { SpreadsheetApp.getUi().alert('Erro ao criar trigger central: ' + (e && e.message)); } catch (er) {}
    Logger.log('configurarRotinaSincronizacaoMadrugada erro: ' + (e && e.message));
  }
}

/**
 * Função executada pelo acionador noturno às 1h.
 * Roda a sincronização global + recálculo de atrasos sem exibir alertas de UI.
 */
/**
 * Função principal executada diariamente pelo acionador noturno (01h).
 * Consolida todas as sincronias pesadas e recálculos necessários.
 */
function executarSincronizacaoGlobalMadrugada_() {
  executarComDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Iniciando Rotina Global de Madrugada (01:00)...");

    // 1) Push para Pedidos: Gera novas linhas de serviços na PEDIDOS-GERAL vindas da OBRA
    try { sincronizarTodosPedidosHousi(); } catch(e) { console.error("Erro em sincronizarTodosPedidosHousi: " + e.message); }

    // 2) Sincronização Base: Pull Status Pedidos, Sync InfoGerais <-> Preliminar
    try { executarSincronizacaoFinalDoDia(); } catch(e) { console.error("Erro em executarSincronizacaoFinalDoDia: " + e.message); }

    // 3) Ocorrências: Sincroniza contagem de ocorrências abertas para a Preliminar
    try { sincronizarOcorrenciasAbertasParaPreliminar_(null, false); } catch(e) { console.error("Erro em sincronizarOcorrenciasAbertasParaPreliminar: " + e.message); }

    // 4) Cronograma FASE-OBRA: Recalcula Semanas e Indicadores (E)
    try { sincronizarTodaAbaObraSemanasCronograma(); } catch(e) { console.error("Erro em sincronizarTodaAbaObraSemanasCronograma: " + e.message); }
    try { sincronizarTodosIndicadoresCronogramaFaseObra(); } catch(e) { console.error("Erro em sincronizarTodosIndicadoresCronogramaFaseObra: " + e.message); }

    // 5) Consolidated Reports: Recalcula indicador de atrasos e Status Obra em INFO GERAIS
    try { recalcularServicosAtrasados_(); } catch(e) { console.error("Erro em recalcularServicosAtrasados: " + e.message); }
    try { sincronizarStatusObraGeral_(); } catch(e) { console.error("Erro em sincronizarStatusObraGeral: " + e.message); }

    // 6) Alinhamento visual FASE-PRELIMINAR -> INFO GERAIS
    try { ordenarPreliminarIgualInformacoesGerais(); } catch(e) { console.error("Erro em ordenarPreliminarIgualInformacoesGerais: " + e.message); }

    // 7) Reordenar FASE-OBRA para seguir INFORMAÇÕES GERAIS (novo requisito)
    try { atualizarOrdemFaseObraPorInformacoesGerais_(); } catch(e) { console.error("Erro em atualizarOrdemFaseObraPorInformacoesGerais_: " + e.message); }

    console.log("✅ Rotina Global de Madrugada concluída com sucesso!");
  }, 300000); // 5 minutos de lock para garantir execução completa
}

/**
 * Wrapper de menu para atualizar o indicador de atrasos.
 */
function atualizarIndicadorServicosAtrasados() {
  recalcularServicosAtrasados_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Indicador de atrasos atualizado em INFORMAÇÕES GERAIS!",
    "Automação",
    5
  );
}

/**
 * Wrapper de menu para sincronização de ocorrências abertas.
 */
function sincronizarTodasOcorrenciasAbertasParaPreliminar() {
  sincronizarOcorrenciasAbertasParaPreliminar_(null, true);
}

/**
 * Wrapper de menu para recálculo em lote de verba teto na FASE-OBRA.
 */
function recalcularTodasVerbasTetoFaseObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!obra) return;

  const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  if (C.VERBA_HOUSI <= 0 || C.VERBA_TETO <= 0) {
    ss.toast("Colunas de verba não encontradas na FASE-OBRA.", "⚠️ Automação", 5);
    return;
  }

  const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
  const last = obra.getLastRow();
  if (last < ini) return;

  const numRows = last - ini + 1;
  const maxCol = Math.max(C.VERBA_HOUSI, C.VERBA_TETO);
  const dados = obra.getRange(ini, 1, numRows, maxCol).getValues();
  const saida = [];

  for (let i = 0; i < numRows; i++) {
    const numero = converterParaNumero_(dados[i][C.VERBA_HOUSI - 1]);
    saida.push([numero === null ? "" : numero * 0.9]);
  }

  obra.getRange(ini, C.VERBA_TETO, numRows, 1).setValues(saida);
  ss.toast("✅ Verba teto recalculada para toda a FASE-OBRA.", "Automação", 5);
}

/**
 * Lista os acionadores do projeto e retorna um array com detalhes básicos.
 * Execute esta função no editor do Apps Script para ver os acionadores atuais.
 */
function listarAcionadoresProjeto() {
  const triggers = ScriptApp.getProjectTriggers();
  const detalhes = triggers.map(function(t) {
    let info = {};
    try { info.handler = t.getHandlerFunction(); } catch (e) { info.handler = null; }
    try { info.eventType = t.getEventType ? String(t.getEventType()) : null; } catch (e) { info.eventType = null; }
    try { info.triggerSource = t.getTriggerSource ? String(t.getTriggerSource()) : null; } catch (e) { info.triggerSource = null; }
    try { info.sourceId = t.getTriggerSourceId ? String(t.getTriggerSourceId()) : null; } catch (e) { info.sourceId = null; }
    return info;
  });

  try {
    const ui = SpreadsheetApp.getUi();
    if (detalhes.length === 0) {
      ui.alert('Nenhum acionador encontrado neste projeto.');
    } else {
      let msg = 'Acionadores encontrados:\n';
      detalhes.forEach(function(d, i) {
        msg += (i+1) + '. Handler: ' + (d.handler || '-') + ' | EventType: ' + (d.eventType || '-') + ' | Source: ' + (d.triggerSource || '-') + '\n';
      });
      if (msg.length > 5000) {
        Logger.log(msg);
        ui.alert('Detalhes dos acionadores gravados em Logger. Veja o Log.');
      } else {
        ui.alert(msg);
      }
    }
  } catch (e) {
    Logger.log(JSON.stringify(detalhes, null, 2));
  }

  return detalhes;
}

/**
 * Cria um acionador diário às 01:00 que executa `sincronizarTodosPedidosHousi`.
 * Se já existir um acionador com o mesmo handler, nada é criado.
 */
/** Rotina centralizada que executa as tarefas diárias agrupadas às 01:00. */
function executarRotinaDiariaCentralizada_() {
  executarComDocumentLock_(function() {
    try { executarSincronizacaoGlobalMadrugada_(); } catch (e) { console.error('central: executarSincronizacaoGlobalMadrugada_ erro: ' + (e && e.message)); }
    try { autorunSincronizarStatusPagamentos(); } catch (e) { console.error('central: autorunSincronizarStatusPagamentos erro: ' + (e && e.message)); }
    try { autorunGerarRelatorio(); } catch (e) { console.error('central: autorunGerarRelatorio erro: ' + (e && e.message)); }
    try { atualizarInformePrestadores(); } catch (e) { console.error('central: atualizarInformePrestadores erro: ' + (e && e.message)); }
    try { atualizarPendenciasGeraisDashboard(); } catch (e) { console.error('central: atualizarPendenciasGeraisDashboard erro: ' + (e && e.message)); }
  }, 300000);
}

/** Recria um único acionador diário às 01:00 executando `executarRotinaDiariaCentralizada_`.
 * Remove acionadores antigos relacionados antes de criar o novo.
 */
function criarTriggerDiarioCentralizado01h() {
  const FN = 'executarRotinaDiariaCentralizada_';
  const toRemove = [
    'executarSincronizacaoFinalDoDia',
    'executarSincronizacaoGlobalMadrugada_',
    'executarAtualizarFaseObraDiaria',
    'autorunGerarRelatorio',
    'autorunSincronizarStatusPagamentos',
    'sincronizarTodosPedidosHousi'
  ];

  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    try {
      const h = existing[i].getHandlerFunction && existing[i].getHandlerFunction();
      if (h && (toRemove.indexOf(h) >= 0 || h === FN)) {
        ScriptApp.deleteTrigger(existing[i]);
        Logger.log('Trigger removido para handler: ' + h);
      }
    } catch (e) {}
  }

  ScriptApp.newTrigger(FN).timeBased().everyDays(1).atHour(1).nearMinute(0).create();
  Logger.log('Acionador diário central criado para ' + FN + ' às 01:00.');
  try { SpreadsheetApp.getUi().alert('Acionador diário central criado para 01:00.'); } catch (e) {}
  return true;
}
