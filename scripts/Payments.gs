/**
 * Payments module (PAGAMENTOS)
 * Scaffolded by Copilot CLI
 * Responsibilities:
 * - Provide functions to create/update payments linked to FASE-OBRA via CHAVE (AY)
 * - Validate parcel sums vs total service value
 * - Aggregate summaries back to FASE-OBRA (PAID_SUM, PENDING_SUM)
 */

/** Opens the payments sheet in the UI (helper). */
function abrirPagamentos() {
  // UI helper: open the sheet named 'PAGAMENTOS'
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('PAGAMENTOS');
    if (!sh) throw new Error('Aba PAGAMENTOS não encontrada.');
    ss.setActiveSheet(sh);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro ao abrir PAGAMENTOS: ' + e.message);
  }
}

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
    CATEGORIA: pick('CATEGORIA'),
    SUBCATEGORIA: pick('SUBCATEGORIA'),
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

/** Quick creation wrapper used by menu. Expects a minimal payload. */
function criarPagamentoRapido(payload) {
  // payload: {chave, prestador, valor, data_prevista, parcela_num}
  return criarPagamento(payload);
}

/** Create a payment record in the PAGAMENTOS sheet.
 * Returns the created payment ID or throws on error.
 */
function criarPagamento(opts) {
  // opts: { CHAVE_SERVICO, PRESTADOR, VALOR, DATA_PREVISTA, PARCELA_NUM, TOTAL_SERVICO, OBS }
  opts = opts || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('PAGAMENTOS');
  if (!sh) {
    throw new Error('Aba PAGAMENTOS não encontrada. Crie o scaffold primeiro.');
  }

  // Basic ID
  const id = 'PAY-' + new Date().getTime();
  const row = [
    id,
    opts.CHAVE_SERVICO || '',
    opts.EMPREENDIMENTO || '',
    opts.UNID || '',
    opts.CATEGORIA || '',
    opts.SUBCATEGORIA || '',
    opts.PRESTADOR || '',
    opts.PARCELA_NUM || 1,
    opts.TOTAL_SERVICO || '',
    opts.VALOR || 0,
    opts.DATA_PREVISTA || '',
    '', // DATA_PAGAMENTO
    opts.STATUS || 'PENDENTE',
    opts.FORMA_PAGAMENTO || '',
    opts.DOCUMENTO_LINK || '',
    opts.OBS || '',
    Session.getActiveUser().getEmail() || '',
    new Date()
  ];

  sh.appendRow(row);
  // return id for reference
  return id;
}

/** Update an existing payment row by ID (partial update allowed). */
function atualizarPagamento(id, changes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('PAGAMENTOS');
  if (!sh) throw new Error('Aba PAGAMENTOS não encontrada.');

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  let foundRow = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id)) { foundRow = r+1; break; }
  }
  if (foundRow === -1) throw new Error('Pagamento ID não encontrado: ' + id);

  // header->col 1-based map
  const headerIndex = {};
  for (let i = 0; i < headers.length; i++) headerIndex[headers[i]] = i+1;

  for (const k in changes) {
    if (headerIndex[k]) {
      sh.getRange(foundRow, headerIndex[k]).setValue(changes[k]);
      continue;
    }
    // try canonical aliases
    const canonical = {
      'UPDATED_BY':'ATUALIZADO_POR','UPDATED_AT':'ATUALIZADO_EM','CREATED_BY':'CRIADO_POR','CREATED_AT':'CRIADO_EM',
      'VALOR':'VALOR','VALOR_PARCELA':'VALOR','CHAVE_SERVICO':'CHAVE_SERVICO','CHAVE':'CHAVE_SERVICO'
    };
    const targetHeader = canonical[k];
    if (targetHeader && headerIndex[targetHeader]) {
      sh.getRange(foundRow, headerIndex[targetHeader]).setValue(changes[k]);
      continue;
    }
    // ignore unknown keys
  }

  // touch updated meta using resolved columns
  const colMap = resolvePaymentsColMap(sh);
  if (colMap.UPDATED_BY >= 0) sh.getRange(foundRow, colMap.UPDATED_BY + 1).setValue(Session.getActiveUser().getEmail() || '');
  if (colMap.UPDATED_AT >= 0) sh.getRange(foundRow, colMap.UPDATED_AT + 1).setValue(new Date());

  return true;
}

/** Validate sum of parcels for a given service key.
 * Returns { totalService, sumParcels, diff }
 */
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

/** Aggregate payment summary for a service key and update FASE-OBRA summary columns.
 * Updates PAID_SUM and PENDING_SUM columns in the corresponding FASE-OBRA row (if found).
 */
function agregarResumoParaFaseObra(chave) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payments = validarSoma(chave);
  const paid = payments.sumParcels || 0;
  const pending = (payments.totalService !== null) ? Math.max(0, payments.totalService - paid) : null;

  // Find FASE-OBRA sheet and CHAVE column
  const obra = ss.getSheetByName('FASE-OBRA');
  if (!obra) throw new Error('Aba FASE-OBRA não encontrada.');

  try {
    const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const chaveCol = C.CHAVE;
    if (!chaveCol) throw new Error('Coluna CHAVE não encontrada na FASE-OBRA.');
    const ini = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA || 'FASE-OBRA');
    const last = obra.getLastRow();
    if (last < ini) throw new Error('FASE-OBRA não possui dados.');
    const vals = obra.getRange(ini, chaveCol, last - ini + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === String(chave)) {
        const row = ini + i;
        const headerRow = obra.getRange(1,1,1,obra.getLastColumn()).getValues()[0];
        const paidColIdx = headerRow.indexOf('PAID_SUM');
        const pendingColIdx = headerRow.indexOf('PENDING_SUM');
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
    throw new Error('Erro ao agregar resumo para FASE-OBRA: ' + e.message);
  }

  throw new Error('Serviço com CHAVE não encontrado em FASE-OBRA: ' + chave);
}


/**
 * Synchronize (import) services from FASE-OBRA into PAGAMENTOS as initial payment rows.
 * - Creates one initial PENDENTE parcela per service found with a non-zero total value
 * - Skips services that already have a payment row (matching CHAVE)
 */
function sincronizarPagamentosDaFaseObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const S = CONFIG.SHEETS || {};
  const obraName = S.OBRA || 'FASE-OBRA';
  const obra = ss.getSheetByName(obraName);
  const paySh = ss.getSheetByName('PAGAMENTOS');
  if (!obra) throw new Error('Aba FASE-OBRA não encontrada.');
  if (!paySh) throw new Error('Aba PAGAMENTOS não encontrada.');

  const ini = obterLinhaInicialPorAba(obraName);
  const lastRow = obra.getLastRow();
  const lastCol = obra.getLastColumn();
  if (lastRow < ini) { SpreadsheetApp.getUi().alert('FASE-OBRA não possui dados.'); return; }

  const headerRow = obra.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h).trim());
  let chaveIdx = headerRow.indexOf('CHAVE');
  if (chaveIdx === -1) chaveIdx = headerRow.indexOf('CHAVE_SERVICO');

  // If still not found, fallback to fixed column mapping CONFIG.COLUMNS.OBRA.CHAVE (1-based index, e.g., 51 for AY)
  if (chaveIdx === -1) {
    if (CONFIG && CONFIG.COLUMNS && CONFIG.COLUMNS.OBRA && CONFIG.COLUMNS.OBRA.CHAVE) {
      const fixed = Number(CONFIG.COLUMNS.OBRA.CHAVE);
      if (!isNaN(fixed) && fixed > 0) {
        chaveIdx = fixed - 1; // zero-based
      }
    }
  }

  if (chaveIdx === -1) throw new Error('Coluna CHAVE não encontrada em FASE-OBRA.');

  const totalCandidates = ['VALOR_TOTAL_SERVICO','VALOR_TOTAL','VALOR','VALOR_SERVICO','TOTAL_SERVICO'];
  let totalIdx = -1; for (let c of totalCandidates) { const i = headerRow.indexOf(c); if (i !== -1) { totalIdx = i; break; } }

  // if totalIdx not found in headers, try common fixed positions or leave as -1
  const obraData = obra.getRange(ini,1,lastRow-ini+1,lastCol).getValues();

  const payHeaders = paySh.getRange(1,1,1,paySh.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  const paymentsData = paySh.getDataRange().getValues();
  const payChaveIdx = payHeaders.indexOf('CHAVE') !== -1 ? payHeaders.indexOf('CHAVE') : payHeaders.indexOf('CHAVE_SERVICO');

  const rowsToAppend = [];
  for (let r = 0; r < obraData.length; r++) {
    const row = obraData[r];
    const chave = row[chaveIdx];
    if (!chave) continue;
    // skip if already exists in payments
    if (payChaveIdx !== -1) {
      let exists = false;
      for (let pr = 1; pr < paymentsData.length; pr++) {
        if (String(paymentsData[pr][payChaveIdx]) === String(chave)) { exists = true; break; }
      }
      if (exists) continue;
    }
    const total = (totalIdx !== -1) ? Number(row[totalIdx]) || 0 : 0;
    if (total === 0) continue;

    // build payments row aligned to payHeaders
    const out = new Array(payHeaders.length).fill('');
    for (let i = 0; i < payHeaders.length; i++) {
      const h = payHeaders[i];
      if (h === 'PAYMENT_UUID' || h === 'PAYMENT_ID' || h === 'ID') out[i] = 'PAY-' + Date.now() + '-' + Math.floor(Math.random()*1000);
      else if (h === 'CHAVE' || h === 'CHAVE_SERVICO') out[i] = chave;
      else if (h === 'EMPREENDIMENTO') {
        const idx = headerRow.indexOf('EMPREENDIMENTO'); if (idx !== -1) out[i] = row[idx];
      } else if (h === 'UNIDADE' || h === 'UNID') { const idx = headerRow.indexOf('UNIDADE'); if (idx !== -1) out[i] = row[idx]; }
      else if (h === 'CATEGORIA') { const idx = headerRow.indexOf('CATEGORIA'); if (idx !== -1) out[i] = row[idx]; }
      else if (h === 'SUBCATEGORIA') { const idx = headerRow.indexOf('SUBCATEGORIA'); if (idx !== -1) out[i] = row[idx]; }
      else if (h === 'SERVICO') { const idx = headerRow.indexOf('SERVICO'); if (idx !== -1) out[i] = row[idx]; }
      else if (h === 'PRESTADOR' || h === 'FORNECEDOR') out[i] = '';
      else if (h === 'PARCELA_NUM' || h === 'PARCELA') out[i] = 1;
      else if (h === 'VALOR' || h === 'VALOR_PARCELA' || h === 'VALOR_PARCELA') out[i] = total;
      else if (h === 'DATA_PAGAMENTO' || h === 'DATA_PAGO' || h === 'DATA_PREVISTA') out[i] = '';
      else if (h === 'STATUS') out[i] = 'PENDENTE';
      else if (h === 'METODO_PAGAMENTO' || h === 'FORMA_PAGAMENTO') out[i] = '';
      else if (h === 'NOTAS' || h === 'OBS' || h === 'DOCUMENTO_LINK') out[i] = '';
      else if (h === 'CRIADO_POR' || h === 'CREATED_BY') out[i] = Session.getActiveUser().getEmail() || '';
      else if (h === 'CRIADO_EM' || h === 'CREATED_AT') out[i] = new Date();
      else if (h === 'ATUALIZADO_POR' || h === 'UPDATED_BY') out[i] = '';
      else if (h === 'ATUALIZADO_EM' || h === 'UPDATED_AT') out[i] = '';
    }
    rowsToAppend.push(out);
  }

  if (rowsToAppend.length === 0) {
    SpreadsheetApp.getUi().alert('Nenhum novo lançamento a importar da FASE-OBRA.');
    return { imported: 0 };
  }

  const startRow = paySh.getLastRow() + 1;
  paySh.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  SpreadsheetApp.getUi().alert('Importados ' + rowsToAppend.length + ' lançamentos para PAGAMENTOS.');
  return { imported: rowsToAppend.length };
}
