/**
 * Payments module (PAGAMENTOS) — versão enxuta.
 * Mantém apenas as funções necessárias para gerar e atualizar o relatório de pagamentos.
 */

/** Resolve column mapping with aliases for PAGAMENTOS sheet. Returns 0-based indices or -1 if missing. */
function resolvePaymentsColMap(sh) {
  const data = sh.getDataRange().getValues();
  const headers = (data && data.length) ? data[0] : [];
  const idx = {};
  headers.forEach((h, i) => { idx[String(h).trim()] = i; });
  const pick = (...names) => { for (const n of names) if (typeof idx[n] !== 'undefined') return idx[n]; return -1; };
  return {
    ID: pick('ID','PAYMENT_UUID','PAYMENT_ID','PAY-'),
    CHAVE_SERVICO: pick('CHAVE_SERVICO','CHAVE','CHAVE_SERVICO'),
    EMPREENDIMENTO: pick('EMPREENDIMENTO','EMPREEND','EMPREENDIMENTO'),
    UNID: pick('UNID','UNIDADE'),
    PRESTADOR: pick('PRESTADOR','FORNECEDOR'),
    PARCELA_NUM: pick('PARCELA_NUM','PARCELA'),
    TOTAL_SERVICO: pick('TOTAL_SERVICO','VALOR_TOTAL_SERVICO','VALOR_TOTAL','TOTAL_SERVICO','TOTAL'),
    VALOR: pick('VALOR','VALOR_PARCELA','VALOR_PARCELA','VALOR'),
    DATA_PREVISTA: pick('DATA_PREVISTA','DATA_PREVISTO'),
    DATA_PAGAMENTO: pick('DATA_PAGAMENTO','DATA_PAGO'),
    STATUS: pick('STATUS'),
    FORMA_PAGAMENTO: pick('FORMA_PAGAMENTO','METODO_PAGAMENTO','FORMA_PAGTO'),
    DOCUMENTO_LINK: pick('DOCUMENTO_LINK','NOTAS'),
    OBS: pick('OBS','NOTAS'),
    CREATED_BY: pick('CRIADO_POR','CREATED_BY','CRIADO_POR'),
    CREATED_AT: pick('CRIADO_EM','CREATED_AT','CRIADO_EM'),
    UPDATED_BY: pick('ATUALIZADO_POR','UPDATED_BY','ATUALIZADO_POR'),
    UPDATED_AT: pick('ATUALIZADO_EM','UPDATED_AT','ATUALIZADO_EM')
  };
}

/** Create or return PAGAMENTOS sheet with canonical headers used by sync routines. */
function criarAbaPagamentosSimples() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('PAGAMENTOS');
  if (sh) return sh;
  sh = ss.insertSheet('PAGAMENTOS');
  const headers = [
    'PAYMENT_UUID','CHAVE','EMPREENDIMENTO','UNIDADE','CATEGORIA','SUBCATEGORIA','SERVICO','PRESTADOR',
    'PARCELA_NUM','TOTAL_SERVICO','VALOR','DATA_PREVISTA','DATA_PAGAMENTO','STATUS','FORMA_PAGAMENTO',
    'DOCUMENTO_LINK','OBS','CRIADO_POR','CRIADO_EM','ATUALIZADO_POR','ATUALIZADO_EM','MÊS'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  try { sh.setFrozenRows(1); } catch(e) {}
  return sh;
}

/** Validate sum of parcels for a given service key. Returns { totalService, sumParcels, diff } */
function validarSoma(chave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('PAGAMENTOS');
  if (!sh) throw new Error('Aba PAGAMENTOS não encontrada.');

  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return { totalService: null, sumParcels: 0, diff: null };

  const colMap = resolvePaymentsColMap(sh);

  let sum = 0;
  let totalService = null;
  for (let r = 1; r < data.length; r++) {
    if (colMap.CHAVE_SERVICO >= 0 && String(data[r][colMap.CHAVE_SERVICO]) === String(chave)) {
      const val = (colMap.VALOR >= 0) ? (Number(data[r][colMap.VALOR]) || 0) : 0;
      sum += val;
      if (colMap.TOTAL_SERVICO >= 0 && !totalService && data[r][colMap.TOTAL_SERVICO]) totalService = Number(data[r][colMap.TOTAL_SERVICO]) || null;
    }
  }
  return { totalService: totalService, sumParcels: sum, diff: (totalService !== null ? (totalService - sum) : null) };
}

/** Aggregate payment summary for a service key and update FASE-OBRA summary columns. */
function agregarResumoParaFaseObra(chave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payments = validarSoma(chave);
  const paid = payments.sumParcels || 0;
  const pending = (payments.totalService !== null) ? Math.max(0, payments.totalService - paid) : null;

  const obra = ss.getSheetByName('FASE-OBRA');
  if (!obra) throw new Error('Aba FASE-OBRA não encontrada.');

  try {
    const C = (typeof resolveSheetColumns_ === 'function') ? resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA) : null;
    const chaveCol = (C && C.CHAVE) ? C.CHAVE : null;
    if (!chaveCol) throw new Error('Coluna CHAVE não encontrada na FASE-OBRA.');
    const ini = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA || 'FASE-OBRA') : 3;
    const last = obra.getLastRow();
    if (last < ini) throw new Error('FASE-OBRA não possui dados.');
    const vals = obra.getRange(ini, chaveCol, last - ini + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === String(chave)) {
        const row = ini + i;
        const headerRow = (typeof getHeaderRow === 'function') ? getHeaderRow(obra) : obra.getRange(1,1,1,obra.getLastColumn()).getValues()[0];
        let paidColIdx = headerRow.indexOf('PAID_SUM');
        let pendingColIdx = headerRow.indexOf('PENDING_SUM');
        if (paidColIdx === -1 || pendingColIdx === -1) {
          const lastCol = obra.getLastColumn();
          obra.insertColumnsAfter(lastCol, 2);
          obra.getRange(1, lastCol+1).setValue('PAID_SUM');
          obra.getRange(1, lastCol+2).setValue('PENDING_SUM');
          obra.getRange(row, lastCol+1).setValue(paid);
          obra.getRange(row, lastCol+2).setValue(pending);
        } else {
          obra.getRange(row, paidColIdx+1).setValue(paid);
          obra.getRange(row, pendingColIdx+1).setValue(pending);
        }
        return { paid: paid, pending: pending };
      }
    }
  } catch (e) {
    throw new Error('Erro ao agregar resumo para FASE-OBRA: ' + (e && e.message));
  }

  throw new Error('Serviço com CHAVE não encontrado em FASE-OBRA: ' + chave);
}

/** Server wrapper: runs the sync and returns result. */
function gerarRelatorioPagamentos(dryRun) {
  try {
    // Incluir serviços já marcados como PAGO que ainda não estão no relatório
    const res = sincronizarPagamentosSimplesFromFaseObraFixed(!!dryRun, true);
    return { success: true, result: res };
  } catch (e) {
    console.error('Erro em gerarRelatorioPagamentos:', e);
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

/** Menu wrapper: gera relatório em modo dry-run e mostra resumo ao usuário. */
function gerarRelatorioPagamentosDryRun() {
  const ui = SpreadsheetApp.getUi();
  const res = gerarRelatorioPagamentos(true);
  if (!res || !res.success) {
    ui.alert('Erro ao gerar relatório (dry-run): ' + (res && res.error ? res.error : 'sem detalhes'));
    return res;
  }
  const r = res.result || {};
  ui.alert('Dry-run concluído: ' + (r.imported || 0) + ' lançamentos (amostra registrada no Logger).');
  Logger.log('Relatório pagamentos (dry-run): %s', JSON.stringify(r));
  return r;
}

/** Menu wrapper: gera relatório e grava os lançamentos na aba PAGAMENTOS (confirmável). */
function gerarRelatorioPagamentosRun() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Confirmação', 'Executar relatório e gravar lançamentos na aba PAGAMENTOS? Esta ação é irreversível.', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return { cancelled: true };
  const res = gerarRelatorioPagamentos(false);
  if (!res || !res.success) {
    ui.alert('Erro ao gerar relatório: ' + (res && res.error ? res.error : 'sem detalhes'));
    return res;
  }
  const r = res.result || {};
  ui.alert('Execução concluída: ' + (r.imported || 0) + ' lançamentos importados.');
  Logger.log('Relatório pagamentos (exec): %s', JSON.stringify(r));
  return r;
}

/** Cria um menu enxuto apenas com ações de relatório (opcional). Chamado por onOpen() se existir. */
function criarMenuPagamentos() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('💳 Pagamentos');
    menu.addItem('Gerar Relatório (Dry-run)', 'gerarRelatorioPagamentosDryRun');
    menu.addItem('Gerar Relatório (Executar)', 'gerarRelatorioPagamentosRun');
    menu.addToUi();
  } catch (e) {}
}
