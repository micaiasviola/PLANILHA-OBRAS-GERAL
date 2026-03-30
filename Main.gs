/*************************
 * GATILHOS PRINCIPAIS
 *************************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const S = CONFIG.SHEETS;

  // Menu operacional (uso diário da equipe).
  ui.createMenu("⚙️ Automacao ECQUA")
    .addItem("➕ Inserir 1 linha no topo (INFO. GERAIS)", "inserirNovaLinhaTopoInformacoesGerais")
    .addItem("🧱 Gerar templates pendentes FASE-OBRA", "gerarTemplatesPendentesFaseObra")
    .addItem("🚀 Gerar/Atualizar PEDIDOS-GERAL (HOUSI)", "sincronizarTodosPedidosHousi")
    .addItem("🚚 Sincronizar envios para FASE-ENTREGA", "sincronizarTodosEnviosParaFaseEntrega")
    .addItem("⏰ Atualizar Indicador de Atrasos (INFO. GERAIS)", "atualizarIndicadorServicosAtrasados")
    .addItem("🔄 Atualizar TODAS AS PENDÊNCIAS (Sinc. Global)", "sincronizacaoManualGlobal")
    .addToUi();

  // Menu técnico (manutenção/admin).
  ui.createMenu("⚙️ Admin ECQUA")
    .addItem("☑️ Configurar coluna de envio (checkbox)", "configurarColunaEnvioFaseEntrega")
    .addItem("⏰ Criar acionador fechamento diário (23h)", "criarAcionadorSincronizacaoFinalDoDia")
    .addItem("🌙 Criar acionador sincronização completa (01h)", "configurarRotinaSincronizacaoMadrugada")
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
    .addToUi();
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
      SpreadsheetApp.getActiveSpreadsheet().toast("Erro na automação: " + err.message, "⚠️ Erro", 5);
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
  const HANDLER = "executarSincronizacaoGlobalMadrugada_";

  // Remove acionadores anteriores com o mesmo handler para evitar duplicação
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(1)       // 01:00 no fuso do projeto
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Acionador de Sincronização Completa criado!\n" +
    "Toda a planilha (Pedidos, Atrasos, Ocorrências e Informações Gerais)\n" +
    "será atualizada automaticamente todos os dias às 01:00."
  );
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
