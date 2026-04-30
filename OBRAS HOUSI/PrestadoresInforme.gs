/**
 * Gera a aba INFORME PRESTADORES com 3 blocos verticais:
 * N1, N2 e N3.
 * Cada bloco corresponde a um período de medição e exibe:
 * prestador, empreendimento, unidade e serviço realizado.
 * Registros antigos vão para a seção LEGADOS / ANTIGOS no rodapé de cada bloco.
 */
function atualizarInformePrestadores() {
  return executarComDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = ss.getSheetByName('PAGAMENTOS');
    if (!source) {
      throw new Error('Aba PAGAMENTOS não encontrada.');
    }

    const targetName = (CONFIG && CONFIG.SHEETS && CONFIG.SHEETS.INFORME_PRESTADORES) ? CONFIG.SHEETS.INFORME_PRESTADORES : 'INFORME PRESTADORES';
    let target = ss.getSheetByName(targetName);
    if (!target) {
      target = ss.insertSheet(targetName);
    }

    const map = resolvePaymentsColMap(source);
    const lastRow = source.getLastRow();
    const lastCol = source.getLastColumn();
    const data = lastRow > 1 ? source.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    const grupos = {
      N1: { label: '26 ao 05', pagamento: 'dia 10', active: [], legacy: [] },
      N2: { label: '06 ao 15', pagamento: 'dia 20', active: [], legacy: [] },
      N3: { label: '16 ao 25', pagamento: 'ultimo dia util do mes', active: [], legacy: [] }
    };

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = String((map.STATUS >= 0) ? row[map.STATUS] : '').trim();
      if (textoNormalizadoSemAcento_(status).indexOf('LIBERADO') === -1) continue;

      const prestador = String((map.PRESTADOR >= 0) ? row[map.PRESTADOR] : '').trim() || 'Sem prestador';
      const empreendimento = String((map.EMPREENDIMENTO >= 0) ? row[map.EMPREENDIMENTO] : '').trim() || 'Sem empreendimento';
      const unidade = String((map.UNID >= 0) ? row[map.UNID] : '').trim() || 'Sem unidade';
      const servico = obterServicoInformePrestadores_(row, map);
      const valor = obterValorInformePrestadores_(row, map);

      const baseDate = obterDataBaseInformePrestadores_(row, map);
      if (!baseDate) continue;

      const periodo = classificarPeriodoMedicao_(baseDate);
      const pagamentoPrevisto = calcularPagamentoPrevisto_(baseDate);
      if (!pagamentoPrevisto) continue;

      const notaKey = mapearNotaPorPeriodo_(periodo.rotulo);
      if (!notaKey || !grupos[notaKey]) continue;

      const pagamentoKey = new Date(pagamentoPrevisto.getFullYear(), pagamentoPrevisto.getMonth(), pagamentoPrevisto.getDate()).getTime();
      const registro = {
        prestador: prestador,
        empreendimento: empreendimento,
        unidade: unidade,
        servico: servico,
        valor: valor,
        periodo: periodo.rotulo,
        pagamento: pagamentoPrevisto,
        pagamentoKey: pagamentoKey,
        chave: [String(row[map.ID] || row[map.CHAVE_SERVICO] || ''), prestador, empreendimento, unidade, servico, notaKey].join('|')
      };

      if (registro.pagamentoKey >= hojeRefAtual_()) {
        grupos[notaKey].active.push(registro);
      } else {
        grupos[notaKey].legacy.push(registro);
      }
    }

    const blocoN1 = montarBlocoInformePrestadores_(grupos.N1, 'N1');
    const blocoN2 = montarBlocoInformePrestadores_(grupos.N2, 'N2');
    const blocoN3 = montarBlocoInformePrestadores_(grupos.N3, 'N3');
    const matriz = mesclarBlocosInformePrestadores_(blocoN1, blocoN2, blocoN3);

    target.clearContents();
    target.clearFormats();
    try { target.getRange(1, 1, target.getMaxRows(), 9).breakApart(); } catch (e) {}

    target.getRange(1, 1, 1, 12).mergeAcross();
    target.getRange(1, 1).setValue('INFORME PRESTADORES');

    target.getRange(2, 1, 1, 12).setValues([[
      'N1 - 26 ao 05 | paga dia 10', '', '', '',
      'N2 - 06 ao 15 | paga dia 20', '', '', '',
      'N3 - 16 ao 25 | ultimo dia util', '', '', ''
    ]]);
    target.getRange(2, 1, 1, 4).mergeAcross();
    target.getRange(2, 5, 1, 4).mergeAcross();
    target.getRange(2, 9, 1, 4).mergeAcross();

    target.getRange(3, 1, 1, 12).setValues([[
      'PRESTADOR', 'EMPREENDIMENTO / UNIDADE', 'SERVIÇO', 'VALOR',
      'PRESTADOR', 'EMPREENDIMENTO / UNIDADE', 'SERVIÇO', 'VALOR',
      'PRESTADOR', 'EMPREENDIMENTO / UNIDADE', 'SERVIÇO', 'VALOR'
    ]]);

    if (matriz.length > 0) {
      target.getRange(4, 1, matriz.length, 12).setValues(matriz);
    }

    target.setFrozenRows(3);
    target.setRowHeight(1, 30);
    target.setRowHeight(2, 24);
    target.setRowHeight(3, 24);
    target.getRange(1, 1, 1, 12).setBackground('#16324f').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    target.getRange(2, 1, 1, 12).setBackground('#2f6f9f').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    target.getRange(3, 1, 1, 12).setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');

    if (matriz.length > 0) {
      aplicarEstiloInformePrestadores_(target, 4, matriz.length);
    }

    target.autoResizeColumns(1, 12);
    for (let c = 1; c <= 12; c++) {
      target.setColumnWidth(c, c % 4 === 1 ? 150 : (c % 4 === 2 ? 185 : (c % 4 === 3 ? 135 : 95)));
    }

    // --- Construir tabela auxiliar de serviços (da FASE-OBRA) a partir da coluna O (15)
    try {
      const colStart = 15; // coluna O
      const hdrs = [
        'EMPREENDIMENTO', 'UNID', 'CATEGORIA DE SERVIÇO', 'FORNECEDOR / PRESTADOR ECQUA',
        'DATA INÍCIO REAL EXECUÇÃO', 'DATA FIM REAL EXECUÇÃO', 'MÊS DO SERVIÇO', 'STATUS PGTO', 'VALOR PGTO'
      ];

      // obter dados da aba FASE-OBRA (para enriquecer pelo CHAVE)
      const obraSh = ss.getSheetByName(CONFIG.SHEETS.OBRA);
      let mapaFasePorChave = {};
      if (obraSh) {
        const C_OBRA = resolveSheetColumns_(obraSh, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
        const raw = obraSh.getDataRange().getValues();
        for (let r = 1; r < raw.length; r++) {
          const row = raw[r];
          const chave = (C_OBRA && C_OBRA.CHAVE) ? String(row[C_OBRA.CHAVE - 1] || '').trim() : '';
          if (!chave) continue;
          mapaFasePorChave[chave] = { row: row, cols: C_OBRA };
        }
      }

      // Preparar linhas de saída: buscar em PAGAMENTOS os registros LIBERADO ou PAGO
      const out = [];
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const status = String((map.STATUS >= 0) ? row[map.STATUS] : '').trim().toUpperCase();
        if (!(status.indexOf('LIBERADO') !== -1 || status.indexOf('PAGO') !== -1)) continue;

        const chaveServico = String((map.CHAVE_SERVICO >= 0) ? row[map.CHAVE_SERVICO] : (map.ID >= 0 ? row[map.ID] : '')).trim();
        const pagamentoStatus = String((map.STATUS >= 0) ? row[map.STATUS] : '').trim();
        const pagamentoValorRaw = (map.VALOR >= 0) ? row[map.VALOR] : '';
        const pagamentoValor = converterParaNumero_(pagamentoValorRaw) || 0;

        let emp = String((map.EMPREENDIMENTO >= 0) ? row[map.EMPREENDIMENTO] : '').trim();
        let uni = String((map.UNID >= 0) ? row[map.UNID] : '').trim();
        let categoria = '';
        let fornecedor = String((map.PRESTADOR >= 0) ? row[map.PRESTADOR] : '').trim();
        let dataInicio = null;
        let dataFim = null;

        if (chaveServico && mapaFasePorChave[chaveServico]) {
          const info = mapaFasePorChave[chaveServico];
          const r = info.row;
          const C = info.cols;
          if (C && C.EMP) emp = String(r[C.EMP - 1] || emp).trim();
          if (C && C.UNI) uni = String(r[C.UNI - 1] || uni).trim();
          if (C && C.CAT) categoria = String(r[C.CAT - 1] || '').trim();
          if (C && C.FORNECEDOR) fornecedor = String(r[C.FORNECEDOR - 1] || fornecedor).trim();
          // tentar campos de data (fallbacks comuns)
          if (C && C.DATA_INICIO_PLANEJADO) dataInicio = parseDateFlexivel_(r[C.DATA_INICIO_PLANEJADO - 1]);
          if (C && C.DATA_PRAZO) dataFim = parseDateFlexivel_(r[C.DATA_PRAZO - 1]);
        }

        // mês do serviço (formato: MM. MMM)
        const mesLabel = (function(d) {
          const dt = d instanceof Date ? d : null;
          if (!dt) return '';
          const n = dt.getMonth() + 1;
          const names = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
          return (n < 10 ? '0' + n : '' + n) + '. ' + names[dt.getMonth()];
        })(dataInicio || obterDataBaseInformePrestadores_(row, map));

        out.push([
          emp || '',
          uni || '',
          categoria || '',
          fornecedor || '',
          dataInicio instanceof Date ? Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          dataFim instanceof Date ? Utilities.formatDate(dataFim, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          mesLabel || '',
          pagamentoStatus || '',
          pagamentoValor ? ('R$ ' + pagamentoValor.toFixed(2).replace('.', ',')) : ''
        ]);
      }

      // Escrever cabeçalho e linhas na planilha (coluna O = 15)
      if (out.length > 0) {
        // cabeçalho na mesma linha 3 para alinhamento visual
        target.getRange(3, colStart, 1, hdrs.length).setValues([hdrs]);
        target.getRange(4, colStart, out.length, hdrs.length).setValues(out);
        // ajustar larguras
        for (let ci = 0; ci < hdrs.length; ci++) target.setColumnWidth(colStart + ci, ci === 1 ? 120 : 150);
      }
    } catch (err) {
      console.error('Erro ao montar tabela auxiliar de FASE-OBRA: ' + err);
    }

    ss.toast('Informe de prestadores atualizado.', 'Concluído', 4);
    return { total: matriz.length };
  });
}

function obterDataBaseInformePrestadores_(row, map) {
  const indicesPreferenciais = [map.DATA_PREVISTA, map.DATA_PAGAMENTO, map.CREATED_AT, map.UPDATED_AT];
  for (let i = 0; i < indicesPreferenciais.length; i++) {
    const idx = indicesPreferenciais[i];
    if (typeof idx === 'number' && idx >= 0) {
      const dt = parseDateFlexivel_(row[idx]);
      if (dt) return dt;
    }
  }

  return null;
}

function obterServicoInformePrestadores_(row, map) {
  const candidatos = [];
  if (typeof map.SERVICO === 'number' && map.SERVICO >= 0) candidatos.push(row[map.SERVICO]);
  if (typeof map.CATEGORIA === 'number' && map.CATEGORIA >= 0) candidatos.push(row[map.CATEGORIA]);
  if (typeof map.SUBCATEGORIA === 'number' && map.SUBCATEGORIA >= 0) candidatos.push(row[map.SUBCATEGORIA]);
  if (typeof map.OBS === 'number' && map.OBS >= 0) candidatos.push(row[map.OBS]);

  const texto = candidatos.map(function(v) { return String(v || '').trim(); }).filter(function(v) { return v; }).join(' / ');
  return texto || 'Serviço não informado';
}

function obterValorInformePrestadores_(row, map) {
  if (typeof map.VALOR !== 'number' || map.VALOR < 0) return 0;
  
  const valor = row[map.VALOR];
  if (typeof valor === 'number') return Math.max(0, valor);
  
  // Se for texto (ex: "R$ 140,00"), extrai o número
  const txt = String(valor || '').trim();
  if (!txt) return 0;
  
  // Remove "R$", espaços e trata separador decimal (vírgula → ponto)
  const num = txt.replace(/[R$\s]/g, '').replace(',', '.');
  const result = parseFloat(num);
  return Number.isFinite(result) ? Math.max(0, result) : 0;
}

function mapearNotaPorPeriodo_(periodo) {
  const texto = textoNormalizadoSemAcento_(periodo);
  if (texto.indexOf('26') !== -1 || texto.indexOf('05') !== -1) return 'N1';
  if (texto.indexOf('06') !== -1 || texto.indexOf('15') !== -1) return 'N2';
  if (texto.indexOf('16') !== -1 || texto.indexOf('25') !== -1) return 'N3';
  return '';
}

function hojeRefAtual_() {
  const hoje = new Date();
  return new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
}

function classificarPeriodoMedicao_(dataMedicao) {
  if (!(dataMedicao instanceof Date) || isNaN(dataMedicao.getTime())) {
    return { rotulo: 'SEM DATA' };
  }

  const dia = dataMedicao.getDate();
  if (dia >= 26 || dia <= 5) return { rotulo: '26-05' };
  if (dia >= 6 && dia <= 15) return { rotulo: '06-15' };
  return { rotulo: '16-25' };
}

function calcularPagamentoPrevisto_(dataMedicao) {
  if (!(dataMedicao instanceof Date) || isNaN(dataMedicao.getTime())) return null;

  const ano = dataMedicao.getFullYear();
  const mes = dataMedicao.getMonth();
  const dia = dataMedicao.getDate();

  if (dia >= 26) return new Date(ano, mes + 1, 10);
  if (dia <= 5) return new Date(ano, mes, 10);
  if (dia >= 6 && dia <= 15) return new Date(ano, mes, 20);
  return ultimoDiaUtilDoMes_(ano, mes);
}

function ultimoDiaUtilDoMes_(ano, mes) {
  const d = new Date(ano, mes + 1, 0);
  while (d.getDay() === 0 || d.getDay() === 6) {
    d.setDate(d.getDate() - 1);
  }
  return d;
}

function parseDateFlexivel_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  if (typeof valor === 'number' && isFinite(valor)) {
    try {
      const d = new Date(Math.round((valor - 25569) * 86400000));
      if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    } catch (e) {}
  }

  const txt = String(valor || '').trim();
  if (!txt) return null;

  const m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    let dia = Number(m[1]);
    let mes = Number(m[2]) - 1;
    let ano = Number(m[3]);
    if (ano < 100) ano += 2000;
    const d = new Date(ano, mes, dia);
    if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  const parsed = Date.parse(txt);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}

function montarBlocoInformePrestadores_(grupo, titulo) {
  const saida = [];
  const ativos = (grupo.active || []).slice().sort(function(a, b) {
    if (a.pagamentoKey !== b.pagamentoKey) return a.pagamentoKey - b.pagamentoKey;
    return String(a.prestador).localeCompare(String(b.prestador));
  });
  const legados = (grupo.legacy || []).slice().sort(function(a, b) {
    if (a.pagamentoKey !== b.pagamentoKey) return a.pagamentoKey - b.pagamentoKey;
    return String(a.prestador).localeCompare(String(b.prestador));
  });

  saida.push([titulo, '', '', '']);
  if (grupo.label) {
    saida.push(['PERÍODO: ' + grupo.label, 'PAGAMENTO: ' + grupo.pagamento, '', '']);
  }

  if (ativos.length > 0) {
    saida.push(['ATIVOS', '', '', '']);
  }

  for (let i = 0; i < ativos.length; i++) {
    saida.push([
      ativos[i].prestador,
      ativos[i].empreendimento + ' / ' + ativos[i].unidade,
      ativos[i].servico,
      'R$ ' + ativos[i].valor.toFixed(2).replace('.', ',')
    ]);
  }

  if (legados.length > 0) {
    saida.push(['LEGADOS / ANTIGOS', '', '', '']);
    for (let i = 0; i < legados.length; i++) {
      saida.push([
        legados[i].prestador,
        legados[i].empreendimento + ' / ' + legados[i].unidade,
        legados[i].servico,
        'R$ ' + legados[i].valor.toFixed(2).replace('.', ',')
      ]);
    }
  }

  return saida;
}

function mesclarBlocosInformePrestadores_(bloco1, bloco2, bloco3) {
  const max = Math.max(bloco1.length, bloco2.length, bloco3.length);
  const out = [];
  for (let i = 0; i < max; i++) {
    const r1 = bloco1[i] || ['', '', '', ''];
    const r2 = bloco2[i] || ['', '', '', ''];
    const r3 = bloco3[i] || ['', '', '', ''];
    out.push([r1[0], r1[1], r1[2], r1[3], r2[0], r2[1], r2[2], r2[3], r3[0], r3[1], r3[2], r3[3]]);
  }
  return out;
}

function aplicarEstiloInformePrestadores_(target, startRow, rowCount) {
  const blocks = [1, 5, 9];
  for (let i = 0; i < blocks.length; i++) {
    const col = blocks[i];
    try {
      target.getRange(startRow, col, rowCount, 4).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
      target.getRange(startRow, col, rowCount, 4).setVerticalAlignment('middle');
      // Alinhar valores (coluna VALOR = col+3) à direita
      target.getRange(startRow, col + 3, rowCount, 1).setHorizontalAlignment('right');
      target.getRange(startRow, col + 1, rowCount, 1).setIndent(1);
      target.getRange(startRow, col + 2, rowCount, 1).setIndent(1);
      for (let r = 0; r < rowCount; r++) {
        const rowIndex = startRow + r;
        const label = String(target.getRange(rowIndex, col).getValue() || '').trim().toUpperCase();
        if (label.indexOf('PERÍODO:') !== -1 || label.indexOf('PERIODO:') !== -1) {
          target.getRange(rowIndex, col, 1, 4).setBackground('#f8fafc').setFontWeight('bold');
        }
        if (label === 'ATIVOS') {
          target.getRange(rowIndex, col, 1, 4).setBackground('#dcfce7').setFontWeight('bold');
        } else if (label.indexOf('LEGADOS') !== -1) {
          target.getRange(rowIndex, col, 1, 4).setBackground('#fee2e2').setFontWeight('bold');
        }
      }
    } catch (e) {}
  }
}