/**
 * Admin helpers: geração de templates e verificação de CHAVE duplicada.
 * Use: executar `gerarTemplatesEVerificarDuplicatas` no editor Apps Script.
 */
function checarDuplicatasChaveObra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obraName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.OBRA) ? CONFIG.SHEETS.OBRA : 'FASE-OBRA';
  const obra = ss.getSheetByName(obraName);
  if (!obra) { Logger.log('Aba FASE-OBRA não encontrada: ' + obraName); return []; }
  const C = (typeof resolveSheetColumns_ === 'function') ? resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA) : null;
  const ini = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(obraName) : 3;
  const last = obra.getLastRow();
  if (last < ini) { Logger.log('FASE-OBRA sem dados'); return []; }
  let chaveCol = (C && C.CHAVE && C.CHAVE > 0) ? C.CHAVE : obterIndiceColunaChavePorAba_(obra);
  if (!chaveCol || chaveCol <= 0) { Logger.log('Coluna CHAVE não encontrada'); return []; }
  const vals = obra.getRange(ini, chaveCol, last - ini + 1, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const seen = new Map();
  const dup = [];
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (!v) continue;
    if (seen.has(v)) {
      const firstLine = seen.get(v);
      const existing = dup.find(d => d.chave === v);
      if (existing) existing.linhas.push(ini + i);
      else dup.push({ chave: v, linhas: [firstLine, ini + i] });
    } else {
      seen.set(v, ini + i);
    }
  }
  Logger.log('Duplicatas encontradas: %s', JSON.stringify(dup, null, 2));
  try { SpreadsheetApp.getUi().alert('Checagem duplicatas concluída — veja Logger (Apps Script).'); } catch (e) {}
  return dup;
}

function gerarTemplatesEVerificarDuplicatas() {
  // Executa a geração padrão de templates (insere linhas em FASE-OBRA)
  try {
    gerarTemplatesPendentesFaseObra();
  } catch (e) {
    Logger.log('Erro ao executar gerarTemplatesPendentesFaseObra: %s', e && e.message);
  }
  // Em seguida, checa duplicatas de CHAVE e retorna resultado
  const duplicates = checarDuplicatasChaveObra();
  return { generated: true, duplicates: duplicates };
}

/**
 * Gera CHAVE (ID) para linhas da aba FASE-OBRA que estejam sem CHAVE.
 * Se `dryRun` for verdadeiro, apenas retorna quantas e quais seriam geradas.
 */
function gerarChavesFaltantesFaseObra(dryRun) {
  dryRun = !!dryRun;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obraName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.OBRA) ? CONFIG.SHEETS.OBRA : 'FASE-OBRA';
  const obra = ss.getSheetByName(obraName);
  if (!obra) {
    Logger.log('Aba FASE-OBRA não encontrada: %s', obraName);
    return { generated: 0, reason: 'FASE-OBRA não encontrada' };
  }

  // Resolve coluna CHAVE (1-based). Usa resolveSheetColumns_ quando disponível.
  const C = (typeof resolveSheetColumns_ === 'function') ? resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA) : null;
  let chaveCol = (C && C.CHAVE && C.CHAVE > 0) ? C.CHAVE : obterIndiceColunaChavePorAba_(obra);

  // Se não encontrou, cria coluna CHAVE ao final
  if (!chaveCol || chaveCol <= 0) {
    const lastCol = obra.getLastColumn();
    const insertAfter = Math.max(1, lastCol);
    obra.insertColumnAfter(insertAfter);
    chaveCol = insertAfter + 1;
    try { obra.getRange(1, chaveCol).setValue('CHAVE'); } catch (e) {}
    if (typeof limparCacheResolucaoColunas_ === 'function') limparCacheResolucaoColunas_();
  }

  const ini = (typeof obterLinhaInicialPorAba === 'function') ? obterLinhaInicialPorAba(obraName) : 3;
  const lastRow = obra.getLastRow();
  if (lastRow < ini) return { generated: 0, reason: 'FASE-OBRA sem dados' };

  const numRows = lastRow - ini + 1;
  const range = obra.getRange(ini, chaveCol, numRows, 1);
  const vals = range.getValues(); // [[val],[val],...]

  const out = [];
  const generated = [];
  for (let i = 0; i < vals.length; i++) {
    const cur = vals[i][0];
    const curStr = (cur === null || typeof cur === 'undefined') ? '' : String(cur).trim();
    if (!curStr || curStr.indexOf('FO_ROW_') === 0) {
      const newId = gerarUUID_();
      out.push([newId]);
      generated.push({ row: ini + i, chave: newId });
    } else {
      out.push([curStr]);
    }
  }

  if (dryRun) {
    Logger.log('Dry-run gerarChavesFaltantesFaseObra: %s', JSON.stringify({ count: generated.length, sample: generated.slice(0, 50) }));
    return { generated: generated.length, sample: generated.slice(0, 50), dryRun: true };
  }

  // Escreve em bloco com lock
  executarComDocumentLock_(function() {
    try {
      range.setValues(out);
    } catch (e) {
      Logger.log('Erro ao gravar CHAVE na FASE-OBRA: %s', e && e.message);
      throw e;
    }
  });

  try { SpreadsheetApp.getUi().alert('Geração de CHAVE concluída: ' + generated.length + ' chaves geradas. Veja Logger.'); } catch (e) {}
  Logger.log('gerarChavesFaltantesFaseObra: geradas %s chaves (exemplo): %s', generated.length, JSON.stringify(generated.slice(0,50), null, 2));
  return { generated: generated.length, details: generated };
}

