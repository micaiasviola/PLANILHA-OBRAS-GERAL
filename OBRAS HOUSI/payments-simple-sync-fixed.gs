/**
 * Sincronizador simples e enxuto para geração/atualização do relatório PAGAMENTOS.
 * Mantém apenas a lógica necessária para detectar o 1º pagamento, importar LIBERADO
 * e (quando solicitado) também importar PAGO que ainda não conste no relatório.
 */

function sincronizarPagamentosSimplesFromFaseObraFixed(dryRun, includePaid) {
  dryRun = !!dryRun;
  includePaid = !!includePaid;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obraName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.OBRA) ? CONFIG.SHEETS.OBRA : 'FASE-OBRA';
  const obra = ss.getSheetByName(obraName);
  if (!obra) throw new Error('Aba FASE-OBRA não encontrada: ' + obraName);

  let paySh = ss.getSheetByName('PAGAMENTOS');
  if (!paySh) paySh = criarAbaPagamentosSimples();

  // ensure MÊS header exists
  const payLastCol = paySh.getLastColumn();
  const payHeadersRow = payLastCol ? paySh.getRange(1,1,1,payLastCol).getValues()[0].map(h=>String(h||'').trim()) : [];
  if (payHeadersRow.indexOf('MÊS') === -1) {
    paySh.getRange(1, Math.max(1, payLastCol) + 1).setValue('MÊS');
  }
  const payHeaders = paySh.getRange(1,1,1,paySh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());

  // header normalization helpers
  const payHeaderNorm = payHeaders.map(h => String(h||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,''));
  const findPayContains = (substr) => { for (let i=0;i<payHeaderNorm.length;i++) if (String(payHeaderNorm[i]||'').indexOf(substr)!==-1) return i; return -1; };

  // prefer parcel VALOR (not TOTAL)
  let payValIdx = -1;
  for (let i=0;i<payHeaderNorm.length;i++) {
    if (String(payHeaderNorm[i]).indexOf('VALOR')!==-1 && String(payHeaderNorm[i]).indexOf('TOTAL')===-1) { payValIdx = i; break; }
  }
  if (payValIdx === -1) payValIdx = findPayContains('VALOR');
  const payStatusIdx = findPayContains('STATUS');
  let payChaveIdx = findPayContains('CHAVE');

  // detect columns in FASE-OBRA
  const lastCol = obra.getLastColumn();
  const headerRow = lastCol ? obra.getRange(1,1,1,lastCol).getValues()[0] : [];
  const normalize = txt => String(txt||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
  const normalized = headerRow.map(normalize);

  // detectar índice da CHAVE
  const sourceChaveIdx = normalized.findIndex(h => (h||'').indexOf('CHAVE') !== -1);

  // Procurar primeiro pelo cabeçalho literal pedido pelo usuário.
  const targetPrestadorHeader = 'FORNECEDOR/ PRESTADOR ECQUA EXECUÇÃO';
  const sourcePrestadorIdxExact = headerRow.findIndex(h => String(h||'').trim().toUpperCase() === targetPrestadorHeader.toUpperCase());
  let sourcePrestadorIdx;
  if (sourcePrestadorIdxExact !== -1) {
    sourcePrestadorIdx = sourcePrestadorIdxExact;
  } else {
    // Fallback: procurar por sinônimos no cabeçalho normalizado
    sourcePrestadorIdx = normalized.findIndex(h => {
      const v = (h||'');
      return v.indexOf('PRESTADOR') !== -1 || v.indexOf('FORNECEDOR') !== -1 || v.indexOf('EXECUCA') !== -1 || v.indexOf('ECQUA') !== -1;
    });
  }

  const sourceServicoIdxExact = headerRow.findIndex(h => String(h || '').trim().toUpperCase() === 'CATEGORIA DE SERVIÇO');
  let sourceServicoIdx = sourceServicoIdxExact;
  if (sourceServicoIdx < 0) {
    sourceServicoIdx = normalized.findIndex(h => {
      const v = (h || '');
      return v.indexOf('CATEGORIA') !== -1 && v.indexOf('SERVICO') !== -1;
    });
  }

  // helper: detect if a payment column candidate refers to the 1º pagamento
  const isFirstPaymentCandidate = (pc) => {
    try {
      const candidates = [];
      if (typeof pc.date === 'number' && pc.date >= 0) candidates.push(normalized[pc.date] || '');
      if (typeof pc.status === 'number' && pc.status >= 0) candidates.push(normalized[pc.status] || '');
      if (typeof pc.val === 'number' && pc.val >= 0) candidates.push(normalized[pc.val] || '');
      for (const v of candidates) {
        if (!v) continue;
        if (v.indexOf('PRIMEIRO') !== -1) return true;
        if (v.indexOf('1') !== -1 && !v.match(/1[0-9]/)) return true;
      }
    } catch (e) {}
    return false;
  };

  // detect multiple payment columns (VALOR with nearby DATE/STATUS)
  const paymentCols = [];
  for (let i=0;i<normalized.length;i++) {
    const rawHeader = String(headerRow[i]||'').toUpperCase().trim();
    if (!/\bVALOR\b/.test(rawHeader)) continue;
    let d = -1, s = -1;
    let dPayment = -1, dAny = -1;
    for (let j=Math.max(0,i-6); j<=Math.min(normalized.length-1,i+6); j++) {
      const h = String(normalized[j]||'');
      if (dPayment === -1 && h.indexOf('DATA')!==-1 && (h.indexOf('PAG')!==-1 || h.indexOf('PAGO')!==-1 || h.indexOf('PREV')!==-1 || h.indexOf('PARCELA')!==-1)) dPayment = j;
      if (dAny === -1 && h.indexOf('DATA')!==-1) dAny = j;
      if (s === -1 && h.indexOf('STATUS')!==-1) s = j;
    }
    d = (dPayment !== -1) ? dPayment : dAny;
    paymentCols.push({val:i, date:d, status:s});
  }
  if (paymentCols.length === 0) {
    // fallback: try to find generic columns
    const dateIdx = normalized.findIndex(h => (h||'').indexOf('DATA')!==-1);
    const statusIdx = normalized.findIndex(h => (h||'').indexOf('STATUS')!==-1);
    const valIdx = normalized.findIndex(h => (h||'').indexOf('VALOR')!==-1);
    paymentCols.push({val: valIdx, date: dateIdx, status: statusIdx});
  }

  const ini = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(obraName) : 3;
  const lastRow = obra.getLastRow();
  if (lastRow < ini) return { imported:0, reason: 'FASE-OBRA sem dados' };
  const obraData = obra.getRange(ini,1,lastRow-ini+1,lastCol).getValues();

  // read existing payments for dedupe / upsert
  const existingRowCount = Math.max(0, paySh.getLastRow() - 1);
  let existing = [];
  if (existingRowCount > 0) existing = paySh.getRange(2, 1, existingRowCount, paySh.getLastColumn()).getValues();
  const getCell = (arr, idx) => (Array.isArray(arr) && typeof idx === 'number' && idx >= 0 && idx < arr.length) ? arr[idx] : '';

  // helper: compare two rows (shallow) for exact equality
  const rowsAreEqual_ = (a, b) => {
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    const len = Math.max(a.length, b.length);
    for (let i = 0; i < len; i++) {
      const va = a[i] === undefined || a[i] === null ? '' : String(a[i]);
      const vb = b[i] === undefined || b[i] === null ? '' : String(b[i]);
      if (va !== vb) return false;
    }
    return true;
  };

  // currency / date helpers (defined early so they're available to subsequent logic)
  const _localParseCurrency = function(v){
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return v;
    let s = String(v).trim();
    s = s.replace(/[R$\s]/g, '');
    const comma = s.indexOf(',');
    const dot = s.indexOf('.');
    if (comma !== -1 && dot === -1) { s = s.replace('.', '').replace(',', '.'); } else { s = s.replace(/,/g,''); }
    const n = Number(s); return isNaN(n) ? null : n;
  };
  const parseCurrencyToNumberLocal = (typeof parseCurrencyToNumber === 'function') ? parseCurrencyToNumber : _localParseCurrency;

  const normalizeDateKey = (cell) => {
    if (cell instanceof Date && !isNaN(cell.getTime())) return cell.toISOString().slice(0,10);
    if (typeof cell === 'number' && !isNaN(cell)) {
      try { const d = new Date(Math.round((cell - 25569) * 86400000)); if (!isNaN(d.getTime())) return d.toISOString().slice(0,10); } catch(e) {}
    }
    const s = String(cell || '').trim(); if (!s) return '';
    const m = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      let day = Number(m[1]), month = Number(m[2]) - 1, year = Number(m[3]); if (year < 100) year += 2000;
      const d = new Date(year, month, day); if (!isNaN(d.getTime())) return d.toISOString().slice(0,10);
    }
    const parsed = Date.parse(s); if (!isNaN(parsed)) return (new Date(parsed)).toISOString().slice(0,10);
    return '';
  };

  const parseToDate = (cell) => {
    if (cell instanceof Date && !isNaN(cell.getTime())) return cell;
    if (typeof cell === 'number' && !isNaN(cell)) {
      try { const d = new Date(Math.round((cell - 25569) * 86400000)); if (!isNaN(d.getTime())) return d; } catch(e) {}
    }
    const s = String(cell || '').trim(); if (!s) return null;
    const m = String(s).match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
    if (m) { let day = Number(m[1]), month = Number(m[2]) - 1, year = Number(m[3]); if (year < 100) year += 2000; const d = new Date(year, month, day); if (!isNaN(d.getTime())) return d; }
    const parsed = Date.parse(s); if (!isNaN(parsed)) return new Date(parsed);
    return null;
  };

  // helper: decide se a linha nova deve substituir a existente (preferência por VALOR preenchido e status mais "final")
  const preferRowByValueAndStatus_ = (existingRow, newRow, valIdxHdr, statusIdxHdr) => {
    const get = (r, idx) => (Array.isArray(r) && typeof idx === 'number' && idx >= 0 && idx < r.length) ? r[idx] : '';
    const existingVal = parseCurrencyToNumberLocal(get(existingRow, valIdxHdr));
    const newVal = parseCurrencyToNumberLocal(get(newRow, valIdxHdr));
    const existingHasVal = existingVal !== null && existingVal !== '';
    const newHasVal = newVal !== null && newVal !== '';

    const normalizeStatus = s => String(s || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
    const existingStatus = normalizeStatus(get(existingRow, statusIdxHdr));
    const newStatus = normalizeStatus(get(newRow, statusIdxHdr));

    if (!existingHasVal && newHasVal) return true;
    if (existingHasVal !== newHasVal) return newHasVal;
    if (existingStatus.indexOf('PREVISTO') !== -1 && (newStatus.indexOf('LIBERADO') !== -1 || newStatus.indexOf('PAGO') !== -1)) return true;
    if (existingStatus.indexOf('LIBERADO') !== -1 && newStatus.indexOf('PAGO') !== -1) return true;
    if (existingVal !== null && newVal !== null && existingVal !== newVal) {
      if (newStatus.indexOf('PAGO') !== -1) return true;
      if (existingStatus === '' && newStatus !== '') return true;
    }
    return false;
  };

  // build quick lookup structures for existing rows to allow upserts
  const existingEntries = [];
  const existingByChave = new Map(); // key: 'CHAVE:xxx' -> entry
  const existingByTuple = new Map(); // key: 'EMP|UNI|PREST|DATE' -> [entries]
  if (existing && existing.length) {
    for (let i = 0; i < existing.length; i++) {
      const r = existing[i];
      const chaveExisting = (typeof payChaveIdx !== 'undefined' && payChaveIdx >= 0) ? String(getCell(r, payChaveIdx) || '').trim() : '';
      const payEmpIdx = findPayContains('EMPREEND');
      const payUnidIdx = findPayContains('UNID');
      const payPrestIdx = (findPayContains('PRESTADOR') >= 0) ? findPayContains('PRESTADOR') : findPayContains('FORNECEDOR');
      const payDateIdx = findPayContains('DATA');
      const payValIdxFinal = (typeof payValIdx === 'number' && payValIdx >= 0) ? payValIdx : findPayContains('VALOR');

      const empV = String(getCell(r, payEmpIdx) || '').trim();
      const uniV = String(getCell(r, payUnidIdx) || '').trim();
      const prestV = String(getCell(r, payPrestIdx) || '').trim();
      const dateKey = normalizeDateKey(getCell(r, payDateIdx));
      const valNumExisting = parseCurrencyToNumberLocal(getCell(r, payValIdxFinal));
      const tupleKeyExisting = [empV, uniV, prestV, dateKey].join('|');

      const entry = { chave: chaveExisting, emp: empV, uni: uniV, prest: prestV, dateKey: dateKey, valNum: valNumExisting, rowIndex: i + 2, rowArray: r };
      existingEntries.push(entry);
      if (chaveExisting) existingByChave.set('CHAVE:' + chaveExisting, entry);
      if (!existingByTuple.has(tupleKeyExisting)) existingByTuple.set(tupleKeyExisting, []);
      existingByTuple.get(tupleKeyExisting).push(entry);
    }
  }
  // helpers already defined earlier in the file; avoid redeclaration here

  const existingKeys = new Set(existing.map(r => {
    const chaveExisting = (typeof payChaveIdx !== 'undefined' && payChaveIdx >= 0) ? String(getCell(r, payChaveIdx) || '').trim() : '';
    if (chaveExisting) return 'CHAVE:' + chaveExisting;

    // mapear índices de PAGAMENTOS usando cabeçalhos detectados
    const payEmpIdx = findPayContains('EMPREEND');
    const payUnidIdx = findPayContains('UNID');
    const payPrestIdx = (findPayContains('PRESTADOR') >= 0) ? findPayContains('PRESTADOR') : findPayContains('FORNECEDOR');
    const payDateIdx = findPayContains('DATA');
    const payValIdxFinal = (typeof payValIdx === 'number' && payValIdx >= 0) ? payValIdx : findPayContains('VALOR');

    const empV = String(getCell(r, payEmpIdx) || '').trim();
    const uniV = String(getCell(r, payUnidIdx) || '').trim();
    const prestV = String(getCell(r, payPrestIdx) || '').trim();
    const dateKey = normalizeDateKey(getCell(r, payDateIdx));
    const valNumExisting = parseCurrencyToNumberLocal(getCell(r, payValIdxFinal));
    const valKey = valNumExisting === null ? '' : String(valNumExisting);
    return [empV, uniV, prestV, dateKey, valKey].join('|');
  }));

  const outMap = new Map();
  const updates = []; // {rowIndex, values}
  const months = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

  for (let r=0;r<obraData.length;r++) {
    const row = obraData[r];
    const emp = (typeof row[0] !== 'undefined') ? row[0] : '';
    const uni = (typeof row[1] !== 'undefined') ? row[1] : '';
    const sourceChaveVal = (typeof sourceChaveIdx === 'number' && sourceChaveIdx >= 0) ? row[sourceChaveIdx] : '';
    const sourcePrestadorVal = (typeof sourcePrestadorIdx === 'number' && sourcePrestadorIdx >= 0) ? row[sourcePrestadorIdx] : '';
    const sourceServicoVal = (typeof sourceServicoIdx === 'number' && sourceServicoIdx >= 0) ? row[sourceServicoIdx] : '';

    for (const pc of paymentCols) {
      const valRaw = (pc.val>=0)?row[pc.val]:'';
      const valNum = parseCurrencyToNumberLocal(valRaw);
      let dateValRaw = (pc.date>=0) ? row[pc.date] : '';
      if ((!dateValRaw || String(dateValRaw).trim() === '') ) {
        const anyDateIdx = normalized.findIndex(h => (h||'').indexOf('DATA') !== -1);
        if (anyDateIdx !== -1) dateValRaw = row[anyDateIdx];
      }
      let dateObj = parseToDate(dateValRaw);
      if (!dateObj && pc.val >= 0) {
        for (let jj = Math.max(0, pc.val-6); jj <= Math.min(row.length-1, pc.val+6); jj++) {
          if (jj === pc.val || (pc.date >= 0 && jj === pc.date)) continue;
          const cand = parseToDate(row[jj]); if (cand) { dateObj = cand; dateValRaw = row[jj]; break; }
        }
      }

      const statusRaw = (pc.status>=0)?row[pc.status]:'';
      const statusNorm = String(statusRaw||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');

      if (!isFirstPaymentCandidate(pc)) continue;

      const isLiberado = statusNorm.indexOf('LIBERADO') !== -1;
      const isPago = statusNorm.indexOf('PAGO') !== -1;
      if (!isLiberado && !(includePaid && isPago)) continue;

      const dateKey = dateObj ? dateObj.toISOString().slice(0,10) : normalizeDateKey(dateValRaw);
      const tupleKey = [String(emp||'').trim(),String(uni||'').trim(),String(sourcePrestadorVal||'').trim(),String(sourceServicoVal||'').trim(), dateKey].join('|');
      const chaveCandidate = String(sourceChaveVal || '').trim();
      const key = chaveCandidate ? ('CHAVE:' + chaveCandidate) : tupleKey;
      // build the output row for this candidate (do this before checking for existing matches)
      const rowOut = new Array(payHeaders.length).fill('');
      for (let i=0;i<payHeaders.length;i++) {
        const hn = String(payHeaders[i]||'').toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g,'');
        if (hn.indexOf('PAY')===0 || hn.indexOf('ID')===0) rowOut[i] = 'PAY-' + Date.now() + '-' + Math.floor(Math.random()*1000);
        else if (hn.indexOf('CHAVE')!==-1) rowOut[i] = sourceChaveVal || '';
        else if (hn.indexOf('EMPREEND')!==-1) rowOut[i] = emp||'';
        else if (hn.indexOf('UNID')!==-1) rowOut[i] = uni||'';
        else if (hn.indexOf('PRESTADOR')!==-1 || hn.indexOf('FORNECEDOR')!==-1) rowOut[i] = sourcePrestadorVal || '';
        else if (hn.indexOf('CATEGORIA')!==-1 && hn.indexOf('SERVICO')!==-1) rowOut[i] = sourceServicoVal || '';
        else if (hn.indexOf('SERVICO')!==-1 && hn.indexOf('PRESTADOR')===-1 && hn.indexOf('FORNECEDOR')===-1) rowOut[i] = sourceServicoVal || '';
        else if (hn.indexOf('DATA')!==-1) rowOut[i] = (dateObj instanceof Date && !isNaN(dateObj)) ? dateObj : (dateValRaw || '');
        else if (hn.indexOf('STATUS')!==-1) rowOut[i] = statusRaw||'';
        else if (hn.indexOf('VALOR')!==-1 && hn.indexOf('TOTAL')===-1) rowOut[i] = (valNum===null?'':valNum);
        else if (hn.indexOf('MES')!==-1) {
          if (dateObj instanceof Date && !isNaN(dateObj)) rowOut[i] = (('0'+(dateObj.getMonth()+1)).slice(-2) + '.' + months[dateObj.getMonth()] + '-' + String(dateObj.getFullYear()).slice(-2));
        }
      }

      // --- UPDATES: check if this candidate matches an existing row and should update it ---
      let matchedExisting = false;
      if (chaveCandidate) {
        const exist = existingByChave.get('CHAVE:' + chaveCandidate);
        if (exist) {
          matchedExisting = true;
          if (preferRowByValueAndStatus_(exist.rowArray, rowOut, payValIdx, payStatusIdx)) {
            updates.push({ rowIndex: exist.rowIndex, values: rowOut });
            // reflect change for subsequent comparisons
            exist.rowArray = rowOut;
          }
        }
      } else {
        const arr = existingByTuple.get(tupleKey) || [];
        if (arr.length) {
          for (const exist of arr) {
            if (rowsAreEqual_(exist.rowArray, rowOut)) { matchedExisting = true; break; }
            if (preferRowByValueAndStatus_(exist.rowArray, rowOut, payValIdx, payStatusIdx)) {
              updates.push({ rowIndex: exist.rowIndex, values: rowOut });
              exist.rowArray = rowOut;
              matchedExisting = true;
              break;
            }
          }
        }
      }
      if (matchedExisting) continue;

      const existingEntry = outMap.get(key);
      if (existingEntry) {
        const preferNew = preferRowByValueAndStatus_(existingEntry, rowOut, payValIdx, payStatusIdx);
        if (preferNew) outMap.set(key, rowOut);
      } else {
        outMap.set(key, rowOut);
      }
      existingKeys.add(key);
    }
  }

  const outRows = Array.from(outMap.values());
  if (outRows.length === 0 && updates.length === 0) {
    if (dryRun) return { imported:0, updated:0, reason: 'Nenhum novo lançamento a importar.' };
    return { imported:0, updated:0, reason: 'Nenhum novo lançamento a importar.' };
  }
  if (dryRun) return { imported: outRows.length, updated: updates.length, sample: outRows.slice(0,20) };

  // apply updates first, then inserts
  try {
    // apply updates in-place
    for (let u = 0; u < updates.length; u++) {
      const rowIdx = updates[u].rowIndex;
      const values = updates[u].values;
      try {
        paySh.getRange(rowIdx, 1, 1, values.length).setValues([values]);
      } catch (e) {
        Logger.log('Falha ao atualizar linha %s em PAGAMENTOS: %s', rowIdx, e && e.message);
      }
    }

    // then insert new rows
    if (outRows && outRows.length) {
      const startRowInsert = Math.max(2, paySh.getLastRow()+1);
      if (typeof setValuesPreservandoColunaChave_ === 'function') {
        setValuesPreservandoColunaChave_(paySh, startRowInsert, 1, outRows);
      } else {
        paySh.getRange(startRowInsert,1,outRows.length,outRows[0].length).setValues(outRows);
      }
    }
    Logger.log('sincronizarPagamentosSimplesFromFaseObraFixed: atualizados %s, gravados %s lançamentos em PAGAMENTOS', updates.length || 0, outRows.length);
  } catch (e) {
    Logger.log('Erro ao gravar/atualizar lançamentos em PAGAMENTOS: %s', e && e.message);
    throw e;
  }
  try {
    const payHeaderNormFinal = payHeaders.map(h => String(h||'').toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g,''));
    let formatValIdx = -1;
    for (let i=0;i<payHeaderNormFinal.length;i++) if (payHeaderNormFinal[i].indexOf('VALOR')!==-1 && payHeaderNormFinal[i].indexOf('TOTAL')===-1) { formatValIdx = i; break; }
    if (formatValIdx === -1) for (let i=0;i<payHeaderNormFinal.length;i++) if (payHeaderNormFinal[i].indexOf('VALOR')!==-1) { formatValIdx = i; break; }
    if (formatValIdx !== -1) paySh.getRange(startRow, formatValIdx+1, outRows.length, 1).setNumberFormat('R$ #,##0.00');
  } catch(e) {}

  return { imported: outRows.length };
}
