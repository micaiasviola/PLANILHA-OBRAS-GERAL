/**
 * Sincronizador diário de status de pagamentos.
 * - `sincronizarStatusPagamentosFromFaseObra(dryRun)`  : verifica FASE-OBRA -> PAGAMENTOS e atualiza status PAGO
 * - `testarSincronizarStatusPagamentosDryRun()`       : wrapper para dry-run (retorna resultado)
 * - `autorunSincronizarStatusPagamentos()`           : entrypoint para trigger
 * - `criarTriggerSincronizarStatusPagamentosDiario()` : cria trigger diário (substitui triggers existentes)
 * - `removerTriggerSincronizarStatusPagamentos()`     : remove triggers criados
 *
 * Uso sugerido:
 * 1) Rodar `testarSincronizarStatusPagamentosDryRun()` e revisar Logger
 * 2) Rodar `criarTriggerSincronizarStatusPagamentosDiario()` para agendar execução diária
 */

function sincronizarStatusPagamentosFromFaseObra(dryRun) {
  dryRun = !!dryRun;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obraName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.OBRA) ? CONFIG.SHEETS.OBRA : 'FASE-OBRA';
  const pagosName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.PAGAMENTOS) ? CONFIG.SHEETS.PAGAMENTOS : 'PAGAMENTOS';

  const obra = ss.getSheetByName(obraName);
  const pagos = ss.getSheetByName(pagosName);
  if (!obra) { Logger.log('Aba não encontrada: %s', obraName); return { updated: 0, reason: 'obra_not_found' }; }
  if (!pagos) { Logger.log('Aba não encontrada: %s', pagosName); return { updated: 0, reason: 'pagos_not_found' }; }

  const iniObra = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(obraName) : 3;
  const iniPagos = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(pagosName) : 3;
  if (obra.getLastRow() < iniObra) { Logger.log('FASE-OBRA sem dados'); return { updated: 0, reason: 'obra_empty' }; }
  if (pagos.getLastRow() < iniPagos) { Logger.log('PAGAMENTOS sem dados'); return { updated: 0, reason: 'pagos_empty' }; }

  // Tentativa de resolver colunas via resolveSheetColumns_ (se disponível)
  const CO = (typeof resolveSheetColumns_ === 'function' && typeof CONFIG !== 'undefined') ? resolveSheetColumns_(obra, (CONFIG.HEADERS_COLS && CONFIG.HEADERS_COLS.OBRA) || null, (CONFIG.COLUMNS && CONFIG.COLUMNS.OBRA) || null) : null;
  const CP = (typeof resolveSheetColumns_ === 'function' && typeof CONFIG !== 'undefined') ? resolveSheetColumns_(pagos, (CONFIG.HEADERS_COLS && CONFIG.HEADERS_COLS.PAGAMENTOS) || null, (CONFIG.COLUMNS && CONFIG.COLUMNS.PAGAMENTOS) || null) : null;

  let chaveColObra = (CO && CO.CHAVE) ? CO.CHAVE : (typeof obterIndiceColunaChavePorAba_ === 'function' ? obterIndiceColunaChavePorAba_(obra) : -1);
  let chaveColPag = (CP && CP.CHAVE) ? CP.CHAVE : (typeof obterIndiceColunaChavePorAba_ === 'function' ? obterIndiceColunaChavePorAba_(pagos) : -1);

  const headerRowObra = iniObra - 1;
  const headerRowPagos = iniPagos - 1;

  function findColLike(sheet, headerRow, substr) {
    try {
      const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] || [];
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i] ? String(headers[i]).toUpperCase() : '';
        if (h.indexOf(substr.toUpperCase()) !== -1) return i + 1;
      }
    } catch (e) { Logger.log('findColLike erro: %s', e && e.message); }
    return -1;
  }

  let statusColObra = (CO && CO.STATUS) ? CO.STATUS : findColLike(obra, headerRowObra, 'STATUS');
  let statusColPag = (CP && CP.STATUS) ? CP.STATUS : findColLike(pagos, headerRowPagos, 'STATUS');

  if (!chaveColObra || chaveColObra <= 0) Logger.log('Aviso: Coluna CHAVE não encontrada em FASE-OBRA (sincronização usará apenas linhas com CHAVE).');
  if (!chaveColPag || chaveColPag <= 0) Logger.log('Aviso: Coluna CHAVE não encontrada em PAGAMENTOS (sincronização por CHAVE não será possível).');
  if (!statusColObra || statusColObra <= 0) { Logger.log('Coluna STATUS não encontrada em FASE-OBRA'); return { updated: 0, reason: 'status_obra_not_found' }; }
  if (!statusColPag || statusColPag <= 0) { Logger.log('Coluna STATUS não encontrada em PAGAMENTOS'); return { updated: 0, reason: 'status_pagos_not_found' }; }

  const numObraRows = obra.getLastRow() - iniObra + 1;
  const numPagosRows = pagos.getLastRow() - iniPagos + 1;

  const obraChaves = (chaveColObra > 0) ? obra.getRange(iniObra, chaveColObra, numObraRows, 1).getValues() : [];
  const obraStatus = obra.getRange(iniObra, statusColObra, numObraRows, 1).getDisplayValues();

  // Map CHAVE -> STATUS (apenas linhas com CHAVE)
  const mapaObra = new Map();
  for (let i = 0; i < numObraRows; i++) {
    const k = (chaveColObra > 0) ? String((obraChaves[i] && obraChaves[i][0]) || '').trim() : '';
    if (!k) continue;
    const s = String((obraStatus[i] && obraStatus[i][0]) || '').trim().toUpperCase();
    mapaObra.set(k, s);
  }

  // Ler PAGAMENTOS (apenas CHAVE e STATUS)
  const pagosChaves = (chaveColPag > 0) ? pagos.getRange(iniPagos, chaveColPag, numPagosRows, 1).getValues() : [];
  const pagosStatus = pagos.getRange(iniPagos, statusColPag, numPagosRows, 1).getDisplayValues();

  const updates = [];
  const changed = [];
  for (let i = 0; i < numPagosRows; i++) {
    const k = (chaveColPag > 0) ? String((pagosChaves[i] && pagosChaves[i][0]) || '').trim() : '';
    const pagoSraw = String((pagosStatus[i] && pagosStatus[i][0]) || '').trim();
    const pagoS = pagoSraw.toUpperCase();
    if (!k) { updates.push([pagoSraw]); continue; }
    const obraS = mapaObra.has(k) ? mapaObra.get(k) : null;
    if (!obraS) { updates.push([pagoSraw]); continue; }
    // Apenas atualiza quando FASE-OBRA estiver marcado como PAGO e PAGAMENTOS ainda não
    if (obraS === 'PAGO' && pagoS !== 'PAGO') {
      updates.push([obraS]);
      changed.push({ row: iniPagos + i, chave: k, from: pagoS, to: obraS });
    } else {
      updates.push([pagoSraw]);
    }
  }

  if (changed.length === 0) {
    Logger.log('sincronizarStatusPagamentosFromFaseObra: nada a atualizar.');
    return { updated: 0, details: [] };
  }

  if (dryRun) {
    Logger.log('Dry-run sincronizarStatusPagamentosFromFaseObra: %s atualizações encontradas: %s', changed.length, JSON.stringify(changed.slice(0, 200), null, 2));
    return { updated: changed.length, sample: changed.slice(0, 200), dryRun: true };
  }

  // Escrita segura com lock quando disponível
  const writer = function() {
    try {
      pagos.getRange(iniPagos, statusColPag, numPagosRows, 1).setValues(updates);
    } catch (e) {
      Logger.log('Erro ao gravar atualizações de status: %s', e && e.message);
      throw e;
    }
  };
  if (typeof executarComDocumentLock_ === 'function') {
    executarComDocumentLock_(writer);
  } else {
    writer();
  }

  Logger.log('sincronizarStatusPagamentosFromFaseObra: gravadas %s atualizações. Exemplos: %s', changed.length, JSON.stringify(changed.slice(0, 200), null, 2));
  try { SpreadsheetApp.getUi().alert('Sincronização de status concluída: ' + changed.length + ' alterações. Veja Logger.'); } catch (e) {}
  return { updated: changed.length, details: changed };
}

function testarSincronizarStatusPagamentosDryRun() {
  return sincronizarStatusPagamentosFromFaseObra(true);
}

function autorunSincronizarStatusPagamentos() {
  try {
    sincronizarStatusPagamentosFromFaseObra(false);
  } catch (e) {
    Logger.log('autorunSincronizarStatusPagamentos erro: %s', e && e.message);
  }
}

function criarTriggerSincronizarStatusPagamentosDiario(hour) {
  Logger.log('Deprecado: criarTriggerSincronizarStatusPagamentosDiario() — use criarTriggerDiarioCentralizado01h().');
  try {
    if (typeof criarTriggerDiarioCentralizado01h === 'function') {
      criarTriggerDiarioCentralizado01h();
      return { ok: true, forwarded: true };
    }
    return { ok: false, reason: 'criarTriggerDiarioCentralizado01h not found' };
  } catch (e) {
    Logger.log('Erro ao encaminhar criação de trigger central: %s', e && e.message);
    return { ok: false, reason: e && e.message };
  }
}

function removerTriggerSincronizarStatusPagamentos() {
  let removed = 0;
  try {
    const existing = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existing.length; i++) {
      const t = existing[i];
      if (t.getHandlerFunction && t.getHandlerFunction() === 'autorunSincronizarStatusPagamentos') {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    }
    Logger.log('Triggers removidos: %s', removed);
    try { SpreadsheetApp.getUi().alert('Triggers removidos: ' + removed); } catch (e) {}
    return { removed: removed };
  } catch (e) {
    Logger.log('Erro ao remover triggers: %s', e && e.message);
    return { removed: removed, reason: e && e.message };
  }
}

/* Trigger helpers para gerar relatório automaticamente */
function autorunGerarRelatorio() {
  try {
    // chamar a função existente em Payments.gs que gera o relatório
    if (typeof gerarRelatorioPagamentos === 'function') {
      // executar sem UI (não dry-run)
      try { gerarRelatorioPagamentos(false); } catch(e) { Logger.log('autorunGerarRelatorio erro (inner): %s', e && e.message); }
    } else {
      Logger.log('autorunGerarRelatorio: função gerarRelatorioPagamentos não encontrada.');
    }
  } catch (e) {
    Logger.log('autorunGerarRelatorio erro: %s', e && e.message);
  }
}

function criarTriggerGerarRelatorioDiario(hour) {
  Logger.log('Deprecado: criarTriggerGerarRelatorioDiario() — use criarTriggerDiarioCentralizado01h().');
  try {
    if (typeof criarTriggerDiarioCentralizado01h === 'function') {
      criarTriggerDiarioCentralizado01h();
      return { ok: true, forwarded: true };
    }
    return { ok: false, reason: 'criarTriggerDiarioCentralizado01h not found' };
  } catch (e) {
    Logger.log('Erro ao encaminhar criação de trigger central: %s', e && e.message);
    return { ok: false, reason: e && e.message };
  }
}

function removerTriggerGerarRelatorio() {
  let removed = 0;
  try {
    const existing = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existing.length; i++) {
      const t = existing[i];
      if (t.getHandlerFunction && t.getHandlerFunction() === 'autorunGerarRelatorio') {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    }
    Logger.log('Triggers de gerar-relatorio removidos: %s', removed);
    try { SpreadsheetApp.getUi().alert('Triggers de gerar-relatório removidos: ' + removed); } catch (e) {}
    return { removed: removed };
  } catch (e) {
    Logger.log('Erro ao remover triggers de gerar-relatorio: %s', e && e.message);
    return { removed: removed, reason: e && e.message };
  }
}

