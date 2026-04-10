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
  // This is a simple implementation: scan sheet for ID and merge changes
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

  // Map of header->col
  const headerIndex = {};
  for (let i = 0; i < headers.length; i++) headerIndex[headers[i]] = i+1;

  const updates = [];
  for (const k in changes) {
    if (headerIndex[k]) {
      sh.getRange(foundRow, headerIndex[k]).setValue(changes[k]);
    }
  }
  // touch updated meta
  if (headerIndex['UPDATED_BY']) sh.getRange(foundRow, headerIndex['UPDATED_BY']).setValue(Session.getActiveUser().getEmail() || '');
  if (headerIndex['UPDATED_AT']) sh.getRange(foundRow, headerIndex['UPDATED_AT']).setValue(new Date());

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
  const headers = data[0];
  const colMap = {};
  for (let i = 0; i < headers.length; i++) colMap[headers[i]] = i;

  let sum = 0;
  let totalService = null;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][colMap['CHAVE_SERVICO']]) === String(chave)) {
      const val = Number(data[r][colMap['VALOR']]) || 0;
      sum += val;
      if (!totalService && data[r][colMap['TOTAL_SERVICO']]) totalService = Number(data[r][colMap['TOTAL_SERVICO']]) || null;
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

  // Try to detect CHAVE column using helper (if present)
  try {
    const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const chaveCol = C.CHAVE || C.CHAVE; // fallback
    const ini = obterLinhaInicialPorAba('FASE-OBRA');
    const last = obra.getLastRow();
    const vals = obra.getRange(ini, chaveCol, last - ini + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === String(chave)) {
        const row = ini + i;
        // Ensure summary cols exist (PAID_SUM / PENDING_SUM) — choose columns near CHAVE or append
        const paidCol = obra.getRange(1,1,1,obra.getLastColumn()).getValues()[0].indexOf('PAID_SUM')+1;
        const pendingCol = obra.getRange(1,1,1,obra.getLastColumn()).getValues()[0].indexOf('PENDING_SUM')+1;
        if (paidCol <= 0 || pendingCol <= 0) {
          // append two columns at the end
          const lastCol = obra.getLastColumn();
          obra.insertColumnsAfter(lastCol, 2);
          obra.getRange(1, lastCol+1).setValue('PAID_SUM');
          obra.getRange(1, lastCol+2).setValue('PENDING_SUM');
          obra.getRange(row, lastCol+1).setValue(paid);
          obra.getRange(row, lastCol+2).setValue(pending);
        } else {
          obra.getRange(row, paidCol).setValue(paid);
          obra.getRange(row, pendingCol).setValue(pending);
        }
        return { paid: paid, pending: pending };
      }
    }
  } catch (e) {
    throw new Error('Erro ao agregar resumo para FASE-OBRA: ' + e.message);
  }

  throw new Error('Serviço com CHAVE não encontrado em FASE-OBRA: ' + chave);
}
