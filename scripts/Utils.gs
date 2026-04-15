/*************************
 * UTILITÁRIOS GERAIS
 *************************/

/**
 * Obtém a linha inicial de dados para cada aba baseada no cabeçalho.
 */
function obterLinhaInicialPorAba(nomeAba) {
  const S = CONFIG.SHEETS;
  if (nomeAba === S.PRELIMINAR) return 4;
  if (nomeAba === S.OBRA) return 3;
  if (nomeAba === S.ENTREGA) return 4;
  if (nomeAba === S.PEDIDOS) return 2;
  if (nomeAba === S.OCORRENCIAS) return 3;
  if (nomeAba === S.INFO_GERAIS) return 4;
  return 2;
}

/**
 * Flag de reentrância: evita deadlock quando uma função com lock
 * chama outra função que também tenta adquirir lock.
 */
let _lockAtivo_ = false;

/**
 * Executa uma função com um lock de documento para evitar concorrência.
 * Reentrante: se já estiver dentro de um lock, executa diretamente.
 */
function executarComDocumentLock_(callback, timeoutMs = 20000) {
  // Se já estamos dentro de um lock, executa direto (reentrância segura)
  if (_lockAtivo_) return callback();

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(timeoutMs)) {
    throw new Error("Não foi possível obter lock do documento. Tente novamente em instantes.");
  }
  _lockAtivo_ = true;
  try {
    return callback();
  } finally {
    _lockAtivo_ = false;
    lock.releaseLock();
  }
}

/**
 * Normaliza um texto para comparação (sem acentos, caixa alta, sem espaços extras).
 */
function textoNormalizadoSemAcento_(valor) {
  return String(valor || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

/**
 * Normaliza datas para comparar apenas o dia (sem hora).
 */
function normalizarDataSomenteDia_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }
  const texto = String(valor || "").trim();
  if (!texto) return null;

  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(texto);
  if (!m) return null;

  const dia = Number(m[1]);
  const mes = Number(m[2]) - 1;
  let ano = Number(m[3]);
  if (ano < 100) ano += 2000;

  const dt = new Date(ano, mes, dia);
  return isNaN(dt.getTime()) ? null : dt;
}

/**
 * Converte valor para número de forma flexível (trata R$, pontos e vírgulas).
 */
function converterParaNumero_(valor) {
  if (typeof valor === "number") return valor;
  let texto = String(valor || "").trim().replace(/\s/g, "").replace(/[R$]/g, "");
  if (!texto) return null;
  if (texto.indexOf(",") >= 0) texto = texto.replace(/\./g, "").replace(",", ".");
  const numero = Number(texto);
  return Number.isFinite(numero) ? numero : null;
}

/**
 * Verifica se um intervalo intercepta uma determinada coluna.
 */
function intervaloInterceptaColuna(range, coluna) {
  const colIni = range.getColumn();
  const colFim = colIni + range.getNumColumns() - 1;
  return coluna >= colIni && coluna <= colFim;
}

/**
 * Busca o índice de uma coluna baseado em uma lista de cabeçalhos possíveis.
 */
function obterColunaPorCabecalho_(aba, headers, linhaBusca = 1) {
  if (!aba || !headers || headers.length === 0) return -1;
  const lastCol = aba.getLastColumn();
  if (lastCol < 1) return -1;

  const normalizados = headers.map(h => textoNormalizadoSemAcento_(h));
  const cabecalhos = aba.getRange(linhaBusca, 1, 1, lastCol).getDisplayValues()[0];

  for (let c = 0; c < cabecalhos.length; c++) {
    const atual = textoNormalizadoSemAcento_(cabecalhos[c]);
    if (atual && normalizados.includes(atual)) return c + 1;
  }
  return -1;
}

/**
 * Busca o índice da coluna em múltiplas linhas de cabeçalho, na ordem informada.
 */
function obterColunaPorCabecalhoEmLinhas_(aba, headers, linhasBusca) {
  if (!Array.isArray(linhasBusca) || linhasBusca.length === 0) return -1;
  for (let i = 0; i < linhasBusca.length; i++) {
    const linha = Number(linhasBusca[i]) || 0;
    if (linha <= 0) continue;
    const col = obterColunaPorCabecalho_(aba, headers, linha);
    if (col > 0) return col;
  }
  return -1;
}

/**
 * Detecta a coluna técnica de chave estável (UUID) pela assinatura "ID-...-...".
 * Usado como fallback resiliente quando o cabeçalho da chave não está disponível.
 */
function detectarColunaChaveUUID_(aba, linhaInicialDados) {
  const lastCol = aba.getLastColumn();
  const lastRow = aba.getLastRow();
  if (lastCol < 1 || lastRow < linhaInicialDados) return -1;

  const numRows = Math.min(60, lastRow - linhaInicialDados + 1);
  const dados = aba.getRange(linhaInicialDados, 1, numRows, lastCol).getDisplayValues();
  const regexUUID = /^ID-[A-Z0-9]+-[A-Z0-9]+$/i;
  const contagemPorColuna = new Array(lastCol).fill(0);

  for (let r = 0; r < dados.length; r++) {
    for (let c = 0; c < lastCol; c++) {
      const valor = String(dados[r][c] || "").trim();
      if (regexUUID.test(valor)) contagemPorColuna[c]++;
    }
  }

  let melhorColuna = -1;
  let melhorPontuacao = 0;
  for (let c = 0; c < contagemPorColuna.length; c++) {
    if (contagemPorColuna[c] > melhorPontuacao) {
      melhorPontuacao = contagemPorColuna[c];
      melhorColuna = c + 1;
    }
  }

  return melhorPontuacao > 0 ? melhorColuna : -1;
}

/**
 * Gera um UUID simples para identificação de linhas.
 */
/**
 * Gera um identificador para a linha.
 * Se passados os 4 campos, gera um "Concatenado" (estilo antigo do usuário).
 * Se não, gera um UUID randômico e estável.
 */
function gerarUUID_() {
  try {
    // Prefer Apps Script UUID generator for strong uniqueness
    if (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function') {
      const full = Utilities.getUuid().toUpperCase().replace(/-/g, '');
      // Keep a short legacy-friendly form: ID-<8hex>-<4hex>
      const part1 = full.slice(0, 8);
      const part2 = full.slice(8, 12);
      return 'ID-' + part1 + '-' + part2;
    }
  } catch (e) {
    // fallthrough to fallback
  }
  // Fallback: original (less robust) approach
  return "ID-" + Math.random().toString(36).substring(2, 9).toUpperCase() + "-" + new Date().getTime().toString(36).toUpperCase();
}

/**
 * Garante que a aba tenha pelo menos o número de colunas especificado.
 */
function garantirColunasAte_(aba, numColunaDestino) {
  const maxCols = aba.getMaxColumns();
  if (maxCols < numColunaDestino) {
    aba.insertColumnsAfter(maxCols, numColunaDestino - maxCols);
  }
}

/**
 * Garante que a aba tenha pelo menos o número de linhas especificado.
 */
function garantirLinhasAte_(aba, numLinhaDestino) {
  const maxRows = aba.getMaxRows();
  if (maxRows < numLinhaDestino) {
    aba.insertRowsAfter(maxRows, numLinhaDestino - maxRows);
  }
}

/**
 * Atalho para obter a aba de backup.
 */
function obterAbaBackup_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const aba = spreadsheet.getSheetByName(CONFIG.SHEETS.BACKUP) 
      || spreadsheet.getSheetByName("BACKUP") 
      || spreadsheet.getSheetByName("Backup");
  
  if (!aba) {
    console.error("Aba de Backup não encontrada! Verifique o nome da aba no arquivo Config.gs.");
  }
  return aba;
}
/**
 * Versão robusta do getLastRow que ignora células com fórmulas vazias ou formatação sem dados.
 */
function obterUltimaLinhaDados_(aba, coluna = 1) {
  const lastRow = aba.getLastRow();
  if (lastRow === 0) return 0;
  const values = aba.getRange(1, coluna, lastRow, 1).getDisplayValues();
  for (let i = lastRow - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== "") return i + 1;
  }
  return 0;
}

// Cache de colunas resolvidas por aba (válido por execução do script)
const _colCache_ = {};

/**
 * Resolve dinamicamente os índices de coluna para uma aba
 * baseado nos nomes de cabeçalho definidos em headerDefs.
 *
 * @param {Sheet} aba - A aba do Sheets.
 * @param {Object} headerDefs - Definições de cabeçalhos, ex: { EMP: ["EMPREENDIMENTO"], CHAVE: 36 }
 * @param {Object} fallbacks  - CONFIG.COLUMNS.XXX, usado quando o cabeçalho não é encontrado.
 * @returns {Object} Mapa { chave: índiceColuna }
 */
function resolveSheetColumns_(aba, headerDefs, fallbacks) {
  const nomeAba = aba.getName();
  if (_colCache_[nomeAba]) return _colCache_[nomeAba];

  const linhaHeader = Math.max(1, obterLinhaInicialPorAba(nomeAba) - 1);
  const linhaInicialDados = Math.max(2, obterLinhaInicialPorAba(nomeAba));
  const linhasBuscaHeader = [linhaHeader, 1, 2, 3].filter((v, i, arr) => arr.indexOf(v) === i);
  const result = {};

  for (const key in headerDefs) {
    const def = headerDefs[key];
    if (typeof def === "number") {
      // Coluna fixa (sem cabeçalho, ex: CHAVE)
      result[key] = def;
    } else {
      // Busca dinâmica por cabeçalho
      const col = obterColunaPorCabecalhoEmLinhas_(aba, def, linhasBuscaHeader);
      
      if (col > 0) {
        result[key] = col;
      } else if (/CHAVE/i.test(String(key))) {
        // Se a chave técnica deslocou de posição, tenta localizar pelo padrão UUID.
        const colUUID = detectarColunaChaveUUID_(aba, linhaInicialDados);
        if (colUUID > 0) {
          result[key] = colUUID;
        } else if (fallbacks && fallbacks[key] != null) {
          result[key] = fallbacks[key];
        } else {
          result[key] = -1;
        }
      } else if (fallbacks && fallbacks[key] != null) {
        result[key] = fallbacks[key];
      } else {
        result[key] = -1;
      }
    }
  }

  result["_source_"] = "HEADERS";
  _colCache_[nomeAba] = result;
  return result;
}

/**
 * Limpa o cache de colunas (útil em testes ou após redesenho da planilha).
 */
function limparCacheResolucaoColunas_() {
  Object.keys(_colCache_).forEach(k => delete _colCache_[k]);
}

/**
 * Obtém o índice da coluna técnica de chave (CHAVE/AY) para a aba informada.
 * Tenta resolver pelos mappings CONFIG.HEADERS_COLS / CONFIG.COLUMNS e, se falhar,
 * usa detecção por padrão UUID como fallback.
 */
function obterIndiceColunaChavePorAba_(aba) {
  if (!aba) return -1;
  const nome = aba.getName();
  // Tenta encontrar a chave pelo mapeamento de sheets em CONFIG.SHEETS
  for (const key in CONFIG.SHEETS) {
    if (CONFIG.SHEETS[key] === nome) {
      const headerDefs = CONFIG.HEADERS_COLS[key];
      const fallbacks = CONFIG.COLUMNS[key];
      if (headerDefs && fallbacks) {
        try {
          const cols = resolveSheetColumns_(aba, headerDefs, fallbacks);
          if (cols && cols.CHAVE && cols.CHAVE > 0) return cols.CHAVE;
        } catch (e) {
          // ignore
        }
      }
    }
  }

  // fallback: detecta pela assinatura UUID nas linhas de dados
  const linhaInicial = Math.max(2, obterLinhaInicialPorAba(nome));
  const detected = detectarColunaChaveUUID_(aba, linhaInicial);
  return detected > 0 ? detected : -1;
}

/**
 * Grava valores em um intervalo garantindo que a coluna técnica de CHAVE (AY) seja preservada.
 * Se o intervalo escrito incluir a coluna CHAVE, os valores atuais dessa coluna serão mantidos.
 * Uso seguro: substitui gravações em lote que poderiam sobrescrever a coluna técnica.
 */
function setValuesPreservandoColunaChave_(aba, startRow, startCol, values) {
  if (!aba || !Array.isArray(values) || values.length === 0) return;
  const numRows = values.length;
  const numCols = values[0].length || 0;
  if (numRows <= 0 || numCols <= 0) return;

  const chaveCol = obterIndiceColunaChavePorAba_(aba);
  if (chaveCol < 0) {
    // sem coluna chave detectada — grava normalmente
    aba.getRange(startRow, startCol, numRows, numCols).setValues(values);
    return;
  }

  // Se a coluna chave não está dentro do intervalo a ser gravado, grava normalmente
  const colFim = startCol + numCols - 1;
  if (chaveCol < startCol || chaveCol > colFim) {
    aba.getRange(startRow, startCol, numRows, numCols).setValues(values);
    return;
  }

  // Precisamos preservar a coluna chave: ler valores existentes e mesclar
  const existing = aba.getRange(startRow, startCol, numRows, numCols).getValues();
  const idx = chaveCol - startCol;
  const merged = values.map((row, r) => {
    const newRow = row.slice();
    // Se a nova linha não tem posição para a chave, garante espaço
    if (newRow.length <= idx) {
      for (let k = newRow.length; k <= idx; k++) newRow[k] = "";
    }
    // Preserva valor existente da coluna chave quando novo valor é vazio/null/undefined
    const candidate = newRow[idx];
    if (candidate === undefined || candidate === null || String(candidate).trim() === "") {
      newRow[idx] = existing[r][idx];
    }
    return newRow;
  });

  aba.getRange(startRow, startCol, numRows, numCols).setValues(merged);
}


/**
 * Reordena as linhas da FASE-OBRA para seguir a ordem definida em INFORMAÇÕES GERAIS.
 * Preserva validações e formatações por linha ao usar sort nativo do Sheets.
 * Linhas sem correspondência em INFORMAÇÕES GERAIS vão para o final,
 * mantendo a ordem original entre elas.
 */
function atualizarOrdemFaseObraPorInformacoesGerais_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(CONFIG.SHEETS.INFO_GERAIS);
  const obra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  if (!info || !obra) return;

  executarComDocumentLock_(function() {
    const C_INFO = resolveSheetColumns_(info, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
    const C_OBRA = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    const iniInfo = obterLinhaInicialPorAba(CONFIG.SHEETS.INFO_GERAIS);
    const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const lastInfo = obterUltimaLinhaDados_(info, C_INFO.EMP);
    const lastObra = obterUltimaLinhaDados_(obra, C_OBRA.EMP);
    if (lastInfo < iniInfo || lastObra < iniObra) return;

    const maxColInfo = Math.max(C_INFO.EMP, C_INFO.UNI);
    const dadosInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();
    const mapaOrdem = new Map();
    for (let i = 0; i < dadosInfo.length; i++) {
      const emp = String(dadosInfo[i][C_INFO.EMP - 1] || "").trim().toUpperCase();
      const uni = String(dadosInfo[i][C_INFO.UNI - 1] || "").trim();
      if (!emp || !uni) continue;
      const chave = emp + "|" + uni;
      if (!mapaOrdem.has(chave)) mapaOrdem.set(chave, i + 1);
    }
    if (mapaOrdem.size === 0) return;

    const lastColObra = obra.getLastColumn();
    const total = lastObra - iniObra + 1;
    const maxColChave = Math.max(C_OBRA.EMP, C_OBRA.UNI);
    const dadosChaveObra = obra.getRange(iniObra, 1, total, maxColChave).getDisplayValues();

    // Criar mapa de agrupamento por unidade e preservar ordem interna
    const rank = [];
    const seqWithin = [];
    const seqGlobal = [];
    const unmatchedBase = 900000;
    const unmatchedMap = new Map();
    let unmatchedCounter = 0;
    const countersPerUnit = {};

    for (let i = 0; i < dadosChaveObra.length; i++) {
      const emp = String(dadosChaveObra[i][C_OBRA.EMP - 1] || "").trim().toUpperCase();
      const uni = String(dadosChaveObra[i][C_OBRA.UNI - 1] || "").trim();
      const chave = (emp && uni) ? (emp + "|" + uni) : "__ROW__" + i;

      let groupRank;
      if (mapaOrdem.has(chave)) {
        groupRank = mapaOrdem.get(chave);
      } else {
        if (!unmatchedMap.has(chave)) {
          unmatchedCounter++;
          unmatchedMap.set(chave, unmatchedCounter);
        }
        groupRank = unmatchedBase + unmatchedMap.get(chave);
      }

      rank.push([groupRank]);

      countersPerUnit[chave] = (countersPerUnit[chave] || 0) + 1;
      seqWithin.push([countersPerUnit[chave]]);
      seqGlobal.push([i + 1]);
    }

    // Inserir 3 colunas auxiliares: rank (grupo), seqWithin (ordem interna), seqGlobal (tie-break)
    const maxCols = obra.getMaxColumns();
    obra.insertColumnsAfter(maxCols, 3);
    const colRank = maxCols + 1;
    const colSeqWithin = maxCols + 2;
    const colSeqGlobal = maxCols + 3;

    obra.getRange(iniObra, colRank, total, 1).setValues(rank);
    obra.getRange(iniObra, colSeqWithin, total, 1).setValues(seqWithin);
    obra.getRange(iniObra, colSeqGlobal, total, 1).setValues(seqGlobal);

    const rangeSort = obra.getRange(iniObra, 1, total, colSeqGlobal);
    rangeSort.sort([
      { column: colRank, ascending: true },
      { column: colSeqWithin, ascending: true },
      { column: colSeqGlobal, ascending: true }
    ]);

    // Remover colunas auxiliares (direita -> esquerda)
    obra.deleteColumn(colSeqGlobal);
    obra.deleteColumn(colSeqWithin);
    obra.deleteColumn(colRank);

    // Garantir que a linha acima dos dados (espacador) esteja vazia
    const spacerRow = iniObra - 1;
    if (spacerRow >= 1) {
      const spacerRange = obra.getRange(spacerRow, 1, 1, obra.getLastColumn());
      const spacerVals = spacerRange.getValues()[0];
      let anyNonEmpty = false;
      for (let j = 0; j < spacerVals.length; j++) {
        if (String(spacerVals[j]).trim() !== '') { anyNonEmpty = true; break; }
      }
      if (anyNonEmpty) spacerRange.clearContent();

      // Revalida subcategorias de serviço nas linhas ordenadas para evitar células "Inválido"
      try {
        const C_OBRA2 = resolveSheetColumns_(obra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
        const colStart = Math.min(C_OBRA2.CAT, C_OBRA2.SUB);
        const colSpan = Math.abs(C_OBRA2.SUB - C_OBRA2.CAT) + 1;
        const intervaloReval = obra.getRange(iniObra, colStart, total, colSpan);
        processarSubcategoriasObra_(obra, intervaloReval, { revalidate: true });
      } catch (err) {
        console.error('Erro ao revalidar subcategorias após ordenação: ' + err);
      }
      }
  });
}

/** Wrapper executável por trigger (registro diário) */
function executarAtualizarFaseObraDiaria() {
  try {
    atualizarOrdemFaseObraPorInformacoesGerais_();
  } catch (err) {
    console.error('Erro ao atualizar ordem FASE-OBRA: ' + err);
  }
}

/** Cria um trigger diário (3:30) para atualizar a ordem da FASE-OBRA */
function criarTriggerDiariaAtualizarFaseObra_() {
  const fn = 'executarAtualizarFaseObraDiaria';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(3).nearMinute(30).create();
  Logger.log('Trigger criada para ' + fn);
}
