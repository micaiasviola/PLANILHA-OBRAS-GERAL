/*************************
 * MÓDULO: FASE-ENTREGA
 *************************/

/**
 * Handler disparado pelo router onEdit quando a aba FASE-ENTREGA é editada. (v2)
 */
function handleEntregaEdit_v2(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const C = resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);

  // 1) Empreendimento -> Unidade (A -> B)
  if (intervaloInterceptaColuna(range, C.EMP)) {
    processarIntervaloAparaB_(sheet, range);
    sincronizarRespPreliminarParaEntrega_(e); // Busca RESP OPR da preliminar
  }

  // 2) Sincronização de Status Geral (I) baseado nas vistorias
  const colsVistoriaFinal = [C.STATUS_VISTORIA_1, C.DATA_VISTORIA_1, C.STATUS_REV_1, C.DATA_REV_1, C.STATUS_REV_2, C.DATA_REV_2];
  if (colsVistoriaFinal.some(c => intervaloInterceptaColuna(range, c))) {
    sincronizarStatusGeralVistoriaFinalPorEdicao_(e);
  }
}

/**
 * Calcula o Status Geral da Vistoria Final de Entrega em lote para um intervalo.
 */
function atualizarStatusGeralVistoriaFinalPorIntervalo_(entrega, primeiraLinha, numLinhas) {
  const C = resolveSheetColumns_(entrega, CONFIG.HEADERS_COLS.ENTREGA, CONFIG.COLUMNS.ENTREGA);
  
  // Pegamos a largura necessária baseada nas colunas de vistoria até REV_2
  const maxCol = Math.max(C.STATUS_VISTORIA_1, C.DATA_VISTORIA_1, C.STATUS_REV_1, C.DATA_REV_1, C.STATUS_REV_2, C.DATA_REV_2);
  const dados = entrega.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();
  const saida = [];

  for (let i = 0; i < dados.length; i++) {
    const row = dados[i];
    const s1 = row[C.STATUS_VISTORIA_1 - 1]; 
    const d1 = row[C.DATA_VISTORIA_1 - 1];   
    const sr1 = row[C.STATUS_REV_1 - 1];     
    const dr1 = row[C.DATA_REV_1 - 1];       
    const sr2 = row[C.STATUS_REV_2 - 1];     
    const dr2 = row[C.DATA_REV_2 - 1];       

    saida.push([calcularStatusGeralVistoriaFinal_(s1, d1, sr1, dr1, sr2, dr2)]);
  }

  entrega.getRange(primeiraLinha, C.STATUS_GERAL, numLinhas, 1).setValues(saida);
}
