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

  // generate id
  const id = 'PAY-' + Date.now() + '-' + Math.floor(Math.random()*1000);

  // Try to align to existing headers if present
  const lastCol = sh.getLastColumn();
  const headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h=>String(h).trim()) : [];
  if (headers && headers.length) {
    const out = new Array(headers.length).fill('');
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (h === 'PAYMENT_UUID' || h === 'PAYMENT_ID' || h === 'ID') out[i] = id;
      else if (h === 'CHAVE' || h === 'CHAVE_SERVICO') out[i] = opts.CHAVE_SERVICO || opts.CHAVE || '';
      else if (h === 'EMPREENDIMENTO') out[i] = opts.EMPREENDIMENTO || '';
      else if (h === 'UNIDADE' || h === 'UNID') out[i] = opts.UNID || opts.UNIDADE || '';
      else if (h === 'CATEGORIA') out[i] = opts.CATEGORIA || '';
      else if (h === 'SUBCATEGORIA') out[i] = opts.SUBCATEGORIA || '';
      else if (h === 'SERVICO') out[i] = opts.SERVICO || '';
      else if (h === 'PRESTADOR' || h === 'FORNECEDOR') out[i] = opts.PRESTADOR || '';
      else if (h === 'PARCELA_NUM' || h === 'PARCELA') out[i] = opts.PARCELA_NUM || 1;
      else if (h === 'TOTAL_SERVICO' || h === 'VALOR_TOTAL_SERVICO' || h === 'VALOR_TOTAL') out[i] = opts.TOTAL_SERVICO || opts.TOTAL || '';
      else if (h === 'VALOR' || h === 'VALOR_PARCELA') out[i] = (typeof opts.VALOR !== 'undefined') ? opts.VALOR : (opts.TOTAL_SERVICO || opts.TOTAL || 0);
      else if (h === 'DATA_PREVISTA' || h === 'DATA_PAGAMENTO' || h === 'DATA_PAGO') out[i] = opts.DATA_PREVISTA || opts.DATA_PAGAMENTO || '';
      else if (h === 'STATUS') out[i] = opts.STATUS || 'PENDENTE';
      else if (h === 'FORMA_PAGAMENTO' || h === 'METODO_PAGAMENTO') out[i] = opts.FORMA_PAGAMENTO || '';
      else if (h === 'NOTAS' || h === 'OBS' || h === 'DOCUMENTO_LINK') out[i] = opts.OBS || '';
      else if (h === 'CRIADO_POR' || h === 'CREATED_BY') out[i] = opts.CRIADO_POR || Session.getActiveUser().getEmail() || '';
      else if (h === 'CRIADO_EM' || h === 'CREATED_AT') out[i] = opts.CRIADO_EM || new Date();
      else if (h === 'ATUALIZADO_POR' || h === 'UPDATED_BY') out[i] = '';
      else if (h === 'ATUALIZADO_EM' || h === 'UPDATED_AT') out[i] = '';
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, out.length).setValues([out]);
    return id;
  }

  // Fallback: legacy positional row
  const row = [
    id,
    opts.CHAVE_SERVICO || opts.CHAVE || '',
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

/** Create a small payments menu. Called from onOpen(). */
function criarMenuPagamentos() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💳 Pagamentos')
    .addItem('Abrir PAGAMENTOS', 'abrirPagamentos')
    .addItem('Sincronizar da FASE-OBRA', 'sincronizarPagamentosDaFaseObra')
    .addItem('Importar (planilha manual)', 'sincronizarPagamentosDaPlanilhaManualPrompt')
    .addToUi();
}

/** Prompt wrapper to ask for external sheet ID/URL and dry-run option. */
function sincronizarPagamentosDaPlanilhaManualPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Importar pagamentos (planilha manual)', 'Cole o ID ou URL da planilha manual (deixe vazio para usar a aba ativa desta planilha):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();

  const dry = ui.alert('Dry-run?', 'Executar em modo dry-run (não grava) e apenas mostrar o que seria importado?', ui.ButtonSet.YES_NO) === ui.Button.YES;
  try {
    const res = importarPagamentosDePlanilhaManual(input || null, dry);
    ui.alert('Importação concluída', JSON.stringify(res), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro na importação: ' + e.message);
    console.error(e);
  }
}

/**
 * Import payments from another spreadsheet (manual control).
 * - sheetIdOrUrl: optional. If null, use active sheet in current spreadsheet.
 * - dryRun: if true, do not write; just return summary.
 */
function importarPagamentosDePlanilhaManual(sheetIdOrUrl, dryRun) {
  const ui = SpreadsheetApp.getUi();
  let srcSS;
  if (!sheetIdOrUrl) {
    srcSS = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    // extract id from url or accept id
    const m = String(sheetIdOrUrl).match(/[-\w]{25,}/);
    if (!m) throw new Error('Não foi possível extrair ID da URL/entrada.');
    srcSS = SpreadsheetApp.openById(m[0]);
  }

  const srcSh = srcSS.getSheets()[0];
  const lastColSrc = srcSh.getLastColumn();
  const lastRowSrc = srcSh.getLastRow();
  if (lastRowSrc < 2) return { imported: 0, reason: 'planilha sem dados' };

  const rawHeaders = srcSh.getRange(1,1,1,lastColSrc).getValues()[0].map(h=>String(h||'').trim());
  const norm = (v) => String(v||'').toUpperCase().normalize('NFD').replace(/[ -\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
  const normHeaders = rawHeaders.map(norm);
  const findIdx = (names) => { const want = names.map(norm); for (let i=0;i<normHeaders.length;i++) if (want.indexOf(normHeaders[i])!==-1) return i; return -1; };

  const empIdx = findIdx(['EMPREENDIMENTO','EMP']);
  const prestIdx = findIdx(['PRESTADOR','FORNECEDOR']);
  const dateIdx = findIdx(['DATA','DATA_PAGAMENTO','DATA_PAGO']);
  const statusIdx = findIdx(['STATUS']);
  const valIdx = findIdx(['VALOR','VALOR_PARCELA','VALOR_TOTAL']);
  const chaveIdxSrc = findIdx(['CHAVE','CHAVE_SERVICO']);

  const paySh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PAGAMENTOS');
  if (!paySh) throw new Error('Aba PAGAMENTOS não encontrada na planilha atual.');
  const payHeaders = paySh.getRange(1,1,1,paySh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const paymentsData = paySh.getDataRange().getValues();

  const payChaveIdx = payHeaders.indexOf('CHAVE') !== -1 ? payHeaders.indexOf('CHAVE') : payHeaders.indexOf('CHAVE_SERVICO');

  const toImport = [];
  const dupChecks = {};

  const srcData = srcSh.getRange(2,1,lastRowSrc-1,lastColSrc).getValues();
  for (let r=0;r<srcData.length;r++) {
    const row = srcData[r];
    const emp = empIdx>=0?row[empIdx]:'';
    const prest = prestIdx>=0?row[prestIdx]:'';
    const date = dateIdx>=0?row[dateIdx]:'';
    const status = statusIdx>=0?row[statusIdx]:'';
    const val = valIdx>=0?row[valIdx]:'';
    const chave = chaveIdxSrc>=0?row[chaveIdxSrc]:'';

    // normalize
    const valNum = (typeof val === 'string') ? Number(String(val).replace(/[^0-9,.-]/g,'').replace(',','.')) : Number(val);
    const dateVal = date instanceof Date ? date : (new Date(date));

    // dedupe by chave if present, else by emp+prest+date+val
    let isDup = false;
    if (chave) {
      if (payChaveIdx!==-1) {
        for (let pr=1; pr<paymentsData.length; pr++) {
          if (String(paymentsData[pr][payChaveIdx]) === String(chave)) { isDup = true; break; }
        }
      }
    } else {
      const key = [String(emp).trim().toUpperCase(), String(prest).trim().toUpperCase(), (dateVal?dateVal.toISOString().slice(0,10):''), String(valNum)].join('|');
      if (dupChecks[key]) { isDup = true; }
      dupChecks[key]=true;
      // also scan paymentsData for same tuple
      for (let pr=1; pr<paymentsData.length; pr++) {
        const pEmp = String(paymentsData[pr][payHeaders.indexOf('EMPREENDIMENTO')]||'').trim().toUpperCase();
        const pPrest = String(paymentsData[pr][payHeaders.indexOf('PRESTADOR')]||'').trim().toUpperCase();
        const pDate = paymentsData[pr][payHeaders.indexOf('DATA_PAGAMENTO')] instanceof Date ? paymentsData[pr][payHeaders.indexOf('DATA_PAGAMENTO')].toISOString().slice(0,10) : String(paymentsData[pr][payHeaders.indexOf('DATA_PAGAMENTO')]||'');
        const pVal = Number(paymentsData[pr][payHeaders.indexOf('VALOR')]||0);
        if (pEmp===pEmp && pPrest===pPrest && pDate=== (dateVal?dateVal.toISOString().slice(0,10):'') && pVal===valNum) { isDup=true; break; }
      }
    }

    if (isDup) continue;

    // build out row aligned to payHeaders
    const out = new Array(payHeaders.length).fill('');
    for (let i=0;i<payHeaders.length;i++) {
      const h = payHeaders[i];
      if (h==='PAYMENT_UUID' || h==='PAYMENT_ID' || h==='ID') out[i] = 'PAY-' + Date.now() + '-' + Math.floor(Math.random()*1000);
      else if (h==='CHAVE' || h==='CHAVE_SERVICO') out[i] = chave || '';
      else if (h==='EMPREENDIMENTO') out[i] = emp || '';
      else if (h==='PRESTADOR' || h==='FORNECEDOR') out[i] = prest || '';
      else if (h==='DATA_PAGAMENTO' || h==='DATA_PAGO') out[i] = dateVal instanceof Date && !isNaN(dateVal) ? dateVal : '';
      else if (h==='STATUS') out[i] = status || 'PENDENTE';
      else if (h==='VALOR' || h==='VALOR_PARCELA') out[i] = isNaN(valNum)?'' : valNum;
      else if (h==='CRIADO_POR' || h==='CREATED_BY') out[i] = Session.getActiveUser().getEmail()||'';
      else if (h==='CRIADO_EM' || h==='CREATED_AT') out[i] = new Date();
    }
    toImport.push({out: out, chave: chave});
  }

  if (toImport.length===0) {
    return { imported:0, reason: 'Nenhum novo lançamento a importar.' };
  }

  if (dryRun) {
    return { imported: toImport.length, sample: toImport.slice(0,10).map(x=>x.out) };
  }

  // append
  const startRow = paySh.getLastRow()+1;
  paySh.getRange(startRow,1,toImport.length,toImport[0].out.length).setValues(toImport.map(x=>x.out));

  // aggregate for chaves
  const chaves = Array.from(new Set(toImport.map(x=>x.chave).filter(Boolean)));
  for (let c of chaves) {
    try { agregarResumoParaFaseObra(c); } catch(e) { console.warn('agregarResumo error for',c,e.message); }
  }

  return { imported: toImport.length };
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
  const norm = (v) => String(v || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]/g, '');
  const normalizedHeaders = headerRow.map(norm);
  const findIdxByNames = (names) => {
    const wanted = names.map(norm);
    for (let i = 0; i < normalizedHeaders.length; i++) {
      if (wanted.indexOf(normalizedHeaders[i]) !== -1) return i;
    }
    return -1;
  };
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

  const totalCandidates = [
    'VALOR TOTAL SERVIÇO',
    'VALOR TOTAL SERVICO',
    'VALOR_TOTAL_SERVICO',
    'VALOR TOTAL',
    'TOTAL SERVICO',
    'TOTAL_SERVICO'
  ];
  let totalIdx = findIdxByNames(totalCandidates);
  // Fallback conhecido do layout: coluna AA (27) inicia bloco financeiro
  if (totalIdx === -1 && lastCol >= 27) totalIdx = 26;

  // if totalIdx not found in headers, try common fixed positions or leave as -1
  const obraData = obra.getRange(ini,1,lastRow-ini+1,lastCol).getValues();
  const C = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
  const empIdx = C.EMP ? C.EMP - 1 : findIdxByNames(['EMPREENDIMENTO']);
  const uniIdx = C.UNI ? C.UNI - 1 : findIdxByNames(['UNID', 'UNIDADE']);
  const catIdx = C.CAT ? C.CAT - 1 : findIdxByNames(['CATEGORIA DE SERVIÇO', 'CATEGORIA DE SERVICO', 'CATEGORIA']);
  const subIdx = C.SUB ? C.SUB - 1 : findIdxByNames(['SUB-CATEGORIA DE SERVIÇO', 'SUB-CATEGORIA DE SERVICO', 'SUBCATEGORIA']);
  const servIdx = findIdxByNames(['SERVIÇO', 'SERVICO', 'DESCRIÇÃO', 'DESCRICAO']);

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
      else if (h === 'EMPREENDIMENTO') out[i] = empIdx >= 0 ? row[empIdx] : '';
      else if (h === 'UNIDADE' || h === 'UNID') out[i] = uniIdx >= 0 ? row[uniIdx] : '';
      else if (h === 'CATEGORIA') out[i] = catIdx >= 0 ? row[catIdx] : '';
      else if (h === 'SUBCATEGORIA') out[i] = subIdx >= 0 ? row[subIdx] : '';
      else if (h === 'SERVICO') out[i] = servIdx >= 0 ? row[servIdx] : '';
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
