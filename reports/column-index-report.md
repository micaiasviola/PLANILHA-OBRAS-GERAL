# Column Index Report

Generated: 2026-04-08T21:06:01.639Z

Found 126 potential usages:

- scripts\Dashboard.gs:44 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosInfo = abaInfo.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();`

- scripts\Dashboard.gs:83 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPed = abaPedidos.getRange(iniPed, 1, lastPed - iniPed + 1, maxColPed).getDisplayValues();`

- scripts\Dashboard.gs:112 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosOco = abaOcor.getRange(iniOco, 1, lastOco - iniOco + 1, maxColOco).getDisplayValues();`

- scripts\Main.gs:93 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `processarIntervaloAparaB_(aba, aba.getRange(linhaCorte, 1, aba.getLastRow() - linhaCorte + 1, 1));`

- scripts\Main.gs:204 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = obra.getRange(ini, 1, numRows, maxCol).getValues();`

- scripts\SheetEntrega.gs:34 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = entrega.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();`

- scripts\SheetInfoGerais.gs:67 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosEmpUni = info.getRange(primeiraLinha, 1, numLinhas, maxColInfo).getDisplayValues();`

- scripts\SheetInfoGerais.gs:114 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const origem = info.getRange(linhaMolde, 1, 1, maxCols);`

- scripts\SheetInfoGerais.gs:115 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const alvo = info.getRange(linhaIni, 1, qtd, maxCols);`

- scripts\SheetInfoGerais.gs:136 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `info.getRange(linhaIni, 1, qtd, colFimDados).clearContent();`

- scripts\SheetInfoGerais.gs:141 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `processarIntervaloAparaB_(info, info.getRange(linhaIni, 1, qtd, 1));`

- scripts\SheetInfoGerais.gs:179 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = abaObra.getRange(iniObra, 1, lastObra - iniObra + 1, maxColObra)`

- scripts\SheetInfoGerais.gs:208 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosInfo = abaInfo.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo)`

- scripts\SheetInfoGerais.gs:232 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `console.log("Indicador de atrasos atualizado: " + novosResumos.filter(r => r[0]).length + " unidade(s) com atraso.");`

- scripts\SheetObra.gs:93 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `.filter(r => String(r[0]).trim() === cat && String(r[1]).trim() !== "")`

- scripts\SheetObra.gs:94 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `.map(r => String(r[1]).trim());`

- scripts\SheetObra.gs:114 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `celulaSub.setValue(subcategorias[0]);`

- scripts\SheetObra.gs:173 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const valsObra = obra.getRange(obraIni, 1, obraLast - obraIni + 1, Math.max(C_OBRA.EMP, C_OBRA.UNI)).getDisplayValues();`

- scripts\SheetObra.gs:223 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPre = pre.getRange(preIni, 1, preLast - preIni + 1, lastColPre).getValues();`

- scripts\SheetObra.gs:254 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `linhaNova[C_OBRA.CAT - 1] = t[0];`

- scripts\SheetObra.gs:255 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `linhaNova[C_OBRA.SUB - 1] = t[1];`

- scripts\SheetObra.gs:256 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_2
  
    Code: `linhaNova[C_OBRA.ATRELADO - 1] = t[2];`

- scripts\SheetObra.gs:274 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeInsert = obra.getRange(iniLivre, 1, arrayLoteCompleto.length, numMaxColsLinha);`

- scripts\SheetObra.gs:300 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const vals = obra.getRange(linha, 1, 1, Math.max(C.CHAVE, C.ATRELADO)).getValues()[0];`

- scripts\SheetObra.gs:300 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const vals = obra.getRange(linha, 1, 1, Math.max(C.CHAVE, C.ATRELADO)).getValues()[0];`

- scripts\SheetObra.gs:319 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const dadosPed = abaPedidos.getRange(linhaPed, C_PED.STATUS, 1, C_PED.FORNECEDOR - C_PED.STATUS + 1).getValues()[0];`

- scripts\SheetObra.gs:320 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `obra.getRange(linha, C.STATUS).setValue(dadosPed[0]);`

- scripts\SheetObra.gs:321 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `obra.getRange(linha, C.FORNECEDOR).setValue(dadosPed[1]);`

- scripts\SheetObra.gs:431 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = sheetObra.getRange(rowStart, 1, numRows, maxCol).getDisplayValues();`

- scripts\SheetOcorrencias.gs:49 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = ocorrencias.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();`

- scripts\SheetOcorrencias.gs:99 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `.filter(r => String(r[0]).trim() === cat && String(r[1]).trim() !== "")`

- scripts\SheetOcorrencias.gs:100 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `.map(r => String(r[1]).trim());`

- scripts\SheetPedidos.gs:106 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const nome = textoNormalizadoSemAcento_(f[0]);`

- scripts\SheetPedidos.gs:139 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPedidos = sheet.getRange(rowStartAdjusted, 1, numRowsAdjusted, maxColPed).getValues();`

- scripts\SheetPedidos.gs:157 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = abaObra.getRange(iniObra, 1, numRowsObra, maxColObra).getValues();`

- scripts\SheetPreliminar.gs:42 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = pre.getRange(primeiraLinha, 1, numLinhas, maxCol).getValues();`

- scripts\SheetPreliminar.gs:86 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosEmpUni = pre.getRange(primeiraLinha, 1, numLinhas, Math.max(C.EMP, C.UNI)).getDisplayValues();`

- scripts\SheetPreliminar.gs:87 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const controles = pre.getRange(primeiraLinha, 1, numLinhas, ultimaColControle).getDisplayValues();`

- scripts\SheetPreliminar.gs:89 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const cabecalhosChecklist = pre.getRange(linhaHeader, C.CHECKLIST_INI, 1, C.CHECKLIST_FIM - C.CHECKLIST_INI + 1).getDisplayValues()[0];`

- scripts\SheetPreliminar.gs:263 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = pre.getRange(primeiraLinha, 1, numLinhas, maxCol).getDisplayValues();`

- scripts\SyncLogic.gs:22 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = sheetObra.getRange(rowStart, 1, numRows, C_OBRA.CHAVE).getValues();`

- scripts\SyncLogic.gs:30 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `dadosPedAll = abaPedidos.getRange(iniPed, 1, lastPed - iniPed + 1, C_PED.CHAVE).getValues();`

- scripts\SyncLogic.gs:110 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `abaPedidos.getRange(iniPed, 1, dadosPedAll.length, C_PED.CHAVE).setValues(dadosPedAll);`

- scripts\SyncLogic.gs:117 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `abaPedidos.getRange(linhaDest, 1, 1, C_PED.CHAVE).setValues([dados]);`

- scripts\SyncLogic.gs:156 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = sheetObra.getRange(linhaIniObra, 1, numLinhasObra, C_OBRA.CHAVE).getValues();`

- scripts\SyncLogic.gs:163 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `dadosPed = abaPedidos.getRange(linhaIniPed, 1, numRowsPed, C_PED.CHAVE).getValues();`

- scripts\SyncLogic.gs:239 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `abaPedidos.getRange(linhaIniPed, 1, rowsParaLimpar, C_PED.CHAVE).clearContent();`

- scripts\SyncLogic.gs:251 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `abaPedidos.getRange(linhaIniPed, 1, listaFinalPedidos.length, C_PED.CHAVE).setValues(listaFinalPedidos);`

- scripts\SyncLogic.gs:277 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = sheetObra.getRange(rowStart, 1, numRows, maxColObra).getValues();`

- scripts\SyncLogic.gs:295 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPed = abaPedidos.getRange(iniPed, 1, numRowsPed, maxColPed).getValues();`

- scripts\SyncLogic.gs:331 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const chave = String(dados[i][0]).trim();`

- scripts\SyncLogic.gs:350 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `if (String(chaves[i][0]).trim() === String(chave).trim()) return ini + i;`

- scripts\SyncLogic.gs:367 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `if (String(chaves[i][0]).trim() === String(chave).trim()) return ini + i;`

- scripts\SyncLogic.gs:407 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const empOriginal = String(valsEmp[i][0]).trim();`

- scripts\SyncLogic.gs:418 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `.filter(r => textoNormalizadoSemAcento_(r[0]) === empBusca)`

- scripts\SyncLogic.gs:419 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `.map(r => String(r[1]).trim())`

- scripts\SyncLogic.gs:461 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const valsA = pre.getRange(linhaIni, 1, last - linhaIni + 1, 1).getDisplayValues();`

- scripts\SyncLogic.gs:463 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `if (CONFIG.STATUS.MARCADOR_BASE.test(valsA[i][0])) {`

- scripts\SyncLogic.gs:497 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_9
  
    Code: `range: abaPedidos.getRange(obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS), 1, abaPedidos.getLastRow(), 9),`

- scripts\SyncLogic.gs:522 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const registrosInfo = info.getRange(linhaInicialInfo, 1, lastInfo - linhaInicialInfo + 1, maxColInfo).getValues();`

- scripts\SyncLogic.gs:542 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `? pre.getRange(linhaHeaderPre, 1, 1, lastColPre).getDisplayValues()[0]`

- scripts\SyncLogic.gs:542 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `? pre.getRange(linhaHeaderPre, 1, 1, lastColPre).getDisplayValues()[0]`

- scripts\SyncLogic.gs:559 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_2
  
    Code: `const dadosPre = pre.getRange(linhaInicialPre, 1, lastPre - linhaInicialPre + 1, 2).getDisplayValues();`

- scripts\SyncLogic.gs:602 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const origem = pre.getRange(linhaMolde, 1, 1, maxCols);`

- scripts\SyncLogic.gs:603 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const alvo = pre.getRange(linhaInicialPre, 1, qtd, maxCols);`

- scripts\SyncLogic.gs:658 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const minRow = linhas[0];`

- scripts\SyncLogic.gs:674 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `pre.getRange(minRow, 2, maxRow - minRow + 1, 1).clearDataValidations();`

- scripts\SyncLogic.gs:693 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const valoresA = abaPedidos.getRange(linhaInicial, 1, limiteBusca - linhaInicial + 1, 1).getDisplayValues();`

- scripts\SyncLogic.gs:717 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const valsA = aba.getRange(linhaInicial, 1, last - linhaInicial + 1, 1).getDisplayValues();`

- scripts\SyncLogic.gs:752 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = info.getRange(linhaIniInfo, 1, numLinhas, numCols).getDisplayValues();`

- scripts\SyncLogic.gs:835 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const emp = String(existentes[i][0]).trim().toUpperCase();`

- scripts\SyncLogic.gs:836 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const uni = String(existentes[i][1]).trim();`

- scripts\SyncLogic.gs:843 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const emp = String(listaEmpUni[i][0]).trim();`

- scripts\SyncLogic.gs:844 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const uni = String(listaEmpUni[i][1]).trim();`

- scripts\SyncLogic.gs:894 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const c = String(chavesObra[i][0]).trim();`

- scripts\SyncLogic.gs:911 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const chavePed = String(chavesPedidos[i][0]).trim();`

- scripts\SyncLogic.gs:972 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_2
  
    Code: `const dados = sheet.getRange(rowStart, 1, numRows, 2).getValues();`

- scripts\SyncLogic.gs:1001 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosOco = abaOco.getRange(iniOco, 1, lastOco - iniOco + 1, Math.max(C_OCO.EMP, C_OCO.UNI, C_OCO.STATUS_GERAL)).getValues();`

- scripts\SyncLogic.gs:1024 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangePre = abaPre.getRange(iniPre, 1, numRowsPre, maxColPre);`

- scripts\SyncLogic.gs:1071 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPre = pre.getRange(iniPre, 1, lastPre - iniPre + 1, maxColPre).getValues();`

- scripts\SyncLogic.gs:1100 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, Math.max(C_INFO.EMP, C_INFO.UNI));`

- scripts\SyncLogic.gs:1162 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosEdicao = abaPedidos.getRange(rowStart, 1, numRows, abaPedidos.getLastColumn()).getValues();`

- scripts\SyncLogic.gs:1184 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeObra = abaObra.getRange(iniObra, 1, lastObra - iniObra + 1, C_OBRA.CHAVE);`

- scripts\SyncLogic.gs:1223 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_9
  
    Code: `range: abaPedidos.getRange(obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS), 1, abaPedidos.getLastRow(), 9),`

- scripts\SyncLogic.gs:1255 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosEntrega = entrega.getRange(rowStart, 1, numRows, Math.max(C_ENTREGA.EMP, C_ENTREGA.UNI)).getValues();`

- scripts\SyncLogic.gs:1263 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPre = pre.getRange(preIni, 1, preLast - preIni + 1, Math.max(C_PRE.EMP, C_PRE.UNI, C_PRE.RESP_OPR, C_PRE.RESP_ADM)).getValues();`

- scripts\SyncLogic.gs:1313 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `range: entrega.getRange(ini, 1, last - ini + 1, 1), // simula alcance englobando coluna A (EMP)`

- scripts\SyncLogic.gs:1420 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = abaPre.getRange(ini, 1, last - ini + 1, Math.max(C.EMP, C.UNI, C.DATA_LOTE)).getValues();`

- scripts\SyncLogic.gs:1461 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados   = sheetObra.getRange(rowStart, 1, numRows, numCols).getValues();`

- scripts\SyncLogic.gs:1528 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosEnt = entrega.getRange(iniEnt, 1, lastEnt - iniEnt + 1, Math.max(C_ENT.UNI, C_ENT.STATUS_GERAL || 0)).getValues();`

- scripts\SyncLogic.gs:1543 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, Math.max(C_INFO.EMP, C_INFO.UNI));`

- scripts\SyncLogic.gs:1576 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `range: obra.getRange(ini, 1, last - ini + 1, 1),`

- scripts\SyncLogic.gs:1670 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `{ range: obra.getRange(ini, 1, last - ini + 1, 1), source: ss },`

- scripts\SyncLogic.gs:1703 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const fakeEvent = { range: info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, 1) };`

- scripts\SyncLogic.gs:1732 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosInfo = info.getRange(rowStart, 1, numRows, maxColInfo).getDisplayValues();`

- scripts\SyncLogic.gs:1757 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosObra = obra.getRange(iniObra, 1, numObraRows, maxColObra).getValues();`

- scripts\SyncLogic.gs:1824 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = sheetObra.getRange(rowStart, 1, numRows, maxCol).getValues();`

- scripts\SyncLogic.gs:1850 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `range: obra.getRange(ini, 1, last - ini + 1, 1),`

- scripts\SyncLogic.gs:1882 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosInfo = abaInfo.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();`

- scripts\SyncLogic.gs:1905 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPre = abaPre.getRange(iniPre, 1, lastPre - iniPre + 1, maxColPre).getDisplayValues();`

- scripts\SyncLogic.gs:1914 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_999999
  
    Code: `etiquetas.push([999999]); // Marcador vai pro fundo absoluto`

- scripts\SyncLogic.gs:1923 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_99999
  
    Code: `etiquetas.push([99999]); // Itens não listados, brancos ou lixos caem pro fundo`

- scripts\SyncLogic.gs:1936 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeTotal = abaPre.getRange(iniPre, 1, etiquetas.length, colTemp);`

- scripts\Tests.gs:79 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `sheetObra.getRange(novaLinhaObra, 1, 1, dadosTesteObra[0].length).setValues(dadosTesteObra);`

- scripts\Tests.gs:83 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `range: sheetObra.getRange(novaLinhaObra, 1, 1, C_OBRA.ATRELADO),`

- scripts\Tests.gs:94 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPedidos = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed - iniPed + 1), C_PED.EMP).getDisplayValues();`

- scripts\Tests.gs:116 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPedidos2Before = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed2Before - iniPed + 1), C_PED.EMP).getDisplayValues();`

- scripts\Tests.gs:125 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `sheetObra.getRange(novaLinhaObra2, 1, 1, dadosTesteObra2[0].length).setValues(dadosTesteObra2);`

- scripts\Tests.gs:128 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `range: sheetObra.getRange(novaLinhaObra2, 1, 1, C_OBRA.ATRELADO),`

- scripts\Tests.gs:136 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosPedidos2After = abaPedidos.getRange(iniPed, 1, Math.max(1, lastPed2After - iniPed + 1), C_PED.EMP).getDisplayValues();`

- scripts\Tests.gs:364 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const status = dados[i][1];`

- scripts\Utils.gs:71 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dia = Number(m[1]);`

- scripts\Utils.gs:72 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_2
  
    Code: `const mes = Number(m[2]) - 1;`

- scripts\Utils.gs:73 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_3
  
    Code: `let ano = Number(m[3]);`

- scripts\Utils.gs:110 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const cabecalhos = aba.getRange(linhaBusca, 1, 1, lastCol).getDisplayValues()[0];`

- scripts\Utils.gs:110 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const cabecalhos = aba.getRange(linhaBusca, 1, 1, lastCol).getDisplayValues()[0];`

- scripts\Utils.gs:143 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dados = aba.getRange(linhaInicialDados, 1, numRows, lastCol).getDisplayValues();`

- scripts\Utils.gs:220 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `if (String(values[i][0]).trim() !== "") return i + 1;`

- scripts\Utils.gs:385 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosInfo = info.getRange(iniInfo, 1, lastInfo - iniInfo + 1, maxColInfo).getDisplayValues();`

- scripts\Utils.gs:399 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const dadosChaveObra = obra.getRange(iniObra, 1, total, maxColChave).getDisplayValues();`

- scripts\Utils.gs:444 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const rangeSort = obra.getRange(iniObra, 1, total, colSeqGlobal);`

- scripts\Utils.gs:459 — pattern: getRange_numeric — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `const spacerRange = obra.getRange(spacerRow, 1, 1, obra.getLastColumn());`

- scripts\Utils.gs:460 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `const spacerVals = spacerRange.getValues()[0];`

- tests.js:230 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_0
  
    Code: `suite4.assertEqual(dados[0][0], "EMPREENDIMENTO", "Cabeçalho correto");`

- tests.js:231 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `suite4.assertEqual(dados[1][0], "MODERN BUNTANTÃ", "Primeiro EMP correto");`

- tests.js:239 — pattern: array_index — suggestion: SUGGESTED_HEADER_COL_1
  
    Code: `suite4.assertEqual(dados[1][1], "11999999999", "Contato do primeiro fornecedor");`



Suggested next steps:
- For each occurrence, inspect and replace numeric index with resolveSheetColumns_(sheet, CONFIG.HEADERS_COLS.<SHEET>, CONFIG.COLUMNS.<SHEET>).
- Add a descriptive key to CONFIG.HEADERS_COLS for the header text you will use.
