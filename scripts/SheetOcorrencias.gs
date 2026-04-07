/*************************
 * MÓDULO: OCORRÊNCIAS
 *************************/

/**
 * Handler disparado pelo router onEdit quando a aba OCORRÊNCIAS é editada.
 */
function handleOcorrenciasEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);

  // 1) Empreendimento -> Unidade (A -> B)
  if (intervaloInterceptaColuna(range, C.EMP)) {
    processarIntervaloAparaB_(sheet, range);
  }

  // 2) Categoria -> Subcategoria (F -> G)
  if (intervaloInterceptaColuna(range, C.CAT)) {
    processarSubcategoriasOcorrencias_(sheet, range);
  }

  // 3) Cálculo de Status Geral (I) baseado nas visitas (J, K, P, Q, V, W)
  const colsVisitas = [C.VISITA_1_STATUS, C.VISITA_1_DATA, C.VISITA_2_STATUS, C.VISITA_2_DATA, C.VISITA_3_STATUS, C.VISITA_3_DATA];
  if (colsVisitas.some(c => intervaloInterceptaColuna(range, c))) {
    sincronizarStatusGeralOcorrenciasPorEdicao_(e);
  }

  // 4) Atualiza Resumo de Ocorrências na FASE-PRELIMINAR
  if (intervaloInterceptaColuna(range, C.EMP) || intervaloInterceptaColuna(range, C.UNI) || intervaloInterceptaColuna(range, C.STATUS_GERAL)) {
    sincronizarOcorrenciasAbertasParaPreliminarPorEdicaoOcorrencias_(e);
  }
}

/**
 * Calcula o Status Geral da Ocorrência.
 */
function atualizarStatusGeralOcorrenciasPorIntervalo_(ocorrencias, primeiraLinha, numLinhas) {
  const C = resolveSheetColumns_(ocorrencias, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);
  const maxCol = Math.max(
    C.VISITA_1_STATUS,
    C.VISITA_1_DATA,
    C.VISITA_2_STATUS,
    C.VISITA_2_DATA,
    C.VISITA_3_STATUS,
    C.VISITA_3_DATA,
    C.STATUS_GERAL
  );
  const dados = ocorrencias.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();
  const saida = [];

  for (let i = 0; i < dados.length; i++) {
    const row = dados[i];
    const s1 = row[C.VISITA_1_STATUS - 1];
    const d1 = row[C.VISITA_1_DATA - 1];
    const s2 = row[C.VISITA_2_STATUS - 1];
    const d2 = row[C.VISITA_2_DATA - 1];
    const s3 = row[C.VISITA_3_STATUS - 1];
    const d3 = row[C.VISITA_3_DATA - 1];

    saida.push([calcularStatusGeralOcorrencia_(s1, d1, s2, d2, s3, d3)]);
  }

  ocorrencias.getRange(primeiraLinha, C.STATUS_GERAL, numLinhas, 1).setValues(saida);
}

/**
 * Atualiza dropdowns de subcategoria baseados na categoria na aba Ocorrências.
 */
function processarSubcategoriasOcorrencias_(abaOco, intervalo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCad = ss.getSheetByName(CONFIG.SHEETS.CADASTROS);
  if (!abaCad) return;

  const C = resolveSheetColumns_(abaOco, CONFIG.HEADERS_COLS.OCORRENCIAS, CONFIG.COLUMNS.OCORRENCIAS);
  const linhaIni = obterLinhaInicialPorAba(CONFIG.SHEETS.OCORRENCIAS);
  const rowStart = Math.max(intervalo.getRow(), linhaIni);
  const numRows = intervalo.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  // Carrega categorias/subcategorias da aba CADASTROS
  const CAD_FIRST_ROW = 6;
  const CAD_FIRST_COL = 2;
  const CAD_NUM_COLS = 2;
  const CAD_NUM_ROWS = Math.max(0, abaCad.getLastRow() - (CAD_FIRST_ROW - 1));
  const dadosCad = CAD_NUM_ROWS > 0 ? abaCad.getRange(CAD_FIRST_ROW, CAD_FIRST_COL, CAD_NUM_ROWS, CAD_NUM_COLS).getValues() : [];

  for (let i = 0; i < numRows; i++) {
    const row = rowStart + i;
    const cat = String(abaOco.getRange(row, C.CAT).getValue()).trim();
    const celulaSub = abaOco.getRange(row, C.SUB);

    if (!cat) {
      celulaSub.clearDataValidations().clearContent();
      continue;
    }

    const subcategorias = dadosCad
      .filter(r => String(r[0]).trim() === cat && String(r[1]).trim() !== "")
      .map(r => String(r[1]).trim());

    if (subcategorias.length === 0) {
      celulaSub.clearDataValidations().clearContent();
      continue;
    }

    const regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(subcategorias, true)
      .setAllowInvalid(true)
      .build();

    celulaSub.setDataValidation(regra);
  }
}
