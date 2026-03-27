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
    .addItem("🌙 Criar acionador atrasos diários (01h)", "criarAcionadorAtrasosDiarios")
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
    .addSeparator()
    .addItem("📆 Recalcular Semanas Cronograma (OBRA)", "sincronizarTodaAbaObraSemanasCronograma")
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
function criarAcionadorAtrasosDiarios() {
  const HANDLER = "executarAtualizacaoAtrasosNocturna";

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
    "✅ Acionador criado!\n" +
    "O indicador de atrasos será atualizado automaticamente\n" +
    "todos os dias por volta de 01:00 (fuso do projeto)."
  );
}

/**
 * Função executada pelo acionador noturno às 1h.
 * Roda a sincronização global + recálculo de atrasos sem exibir alertas de UI.
 */
function executarAtualizacaoAtrasosNocturna() {
  executarComDocumentLock_(function() {
    // 1) Sincronização completa (Pedidos, Entrega, Preliminar)
    executarSincronizacaoFinalDoDia();

    // 2) Recálculo do indicador de atrasos em INFO GERAIS
    recalcularServicosAtrasados_();
  }, 120000);
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
