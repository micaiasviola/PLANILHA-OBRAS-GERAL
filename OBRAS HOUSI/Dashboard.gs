/**
 * Popula a aba DASHBOARD (Pendencias Gerais) a partir de:
 * - INFORMAÇÕES GERAIS: EMPREENDIMENTO, UNID e DATA LOTE (para dias de obra)
 * - FASE-OBRA: agregacoes de servicos e pagamentos
 */
function atualizarPendenciasGeraisDashboard() {
	executarComDocumentLock_(function() {
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const S = CONFIG.SHEETS;

		const abaInfo = ss.getSheetByName(S.INFO_GERAIS);
		const abaObra = ss.getSheetByName(S.OBRA);
		const abaDash = ss.getSheetByName(S.DASHBOARD);
		if (!abaInfo || !abaObra || !abaDash) {
			throw new Error("Abas obrigatorias nao encontradas: INFORMAÇÕES GERAIS, FASE-OBRA e/ou DASHBOARD.");
		}

		const C_INFO = resolveSheetColumns_(abaInfo, CONFIG.HEADERS_COLS.INFO_GERAIS, CONFIG.COLUMNS.INFO_GERAIS);
		const C_OBRA = resolveSheetColumns_(abaObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
		const C_DASH = resolveSheetColumns_(abaDash, CONFIG.HEADERS_COLS.DASHBOARD, CONFIG.COLUMNS.DASHBOARD);

		const linhaIniInfo = obterLinhaInicialPorAba(S.INFO_GERAIS);
		const linhaIniObra = obterLinhaInicialPorAba(S.OBRA);
		const linhaIniDash = obterLinhaInicialPorAba(S.DASHBOARD);

		// 1) Base de unidades unicas vindas de INFORMAÇÕES GERAIS (preserva ordem de aparicao)
		const unidades = carregarUnidadesInfoGerais_(abaInfo, C_INFO, linhaIniInfo);

		// 2) Agregacoes por unidade vindas de FASE-OBRA
		const agregadosObra = agregarDadosObraPorUnidade_(abaObra, C_OBRA, linhaIniObra);

		// 3) Limpa conteudo anterior no bloco A:H
		const lastDash = abaDash.getLastRow();
		if (lastDash >= linhaIniDash) {
			abaDash.getRange(linhaIniDash, 1, lastDash - linhaIniDash + 1, 8).clearContent();
		}

		if (unidades.length === 0) {
			ss.toast("Dashboard atualizado: nenhuma unidade encontrada em INFORMAÇÕES GERAIS.", "Concluido", 4);
			return;
		}

		garantirLinhasAte_(abaDash, linhaIniDash + unidades.length - 1);

		const maxColDash = Math.max(
			C_DASH.EMP || 0,
			C_DASH.UNI || 0,
			C_DASH.SEMANA_CRONOGRAMA || 0,
			C_DASH.SERVICOS_CONCLUIDOS || 0,
			C_DASH.SERVICOS_PENDENTES || 0,
			C_DASH.VERBA_UTILIZADA || 0,
			C_DASH.PGTOS_PENDENTES || 0,
			C_DASH.ALERTA || 0,
			8
		);

		const saida = [];
		for (let i = 0; i < unidades.length; i++) {
			const reg = unidades[i];
			const agg = agregadosObra.get(reg.chave) || {
				servicosConcluidos: 0,
				servicosPendentes: 0,
				verbaUtilizada: 0,
				pgtosPendentes: 0
			};
			const diasObra = calcularDiasCorridosDesdeDataLote_(reg.dataLote);

			const linha = new Array(maxColDash).fill("");
			if (C_DASH.EMP > 0) linha[C_DASH.EMP - 1] = reg.emp;
			if (C_DASH.UNI > 0) linha[C_DASH.UNI - 1] = reg.uni;
			if (C_DASH.SEMANA_CRONOGRAMA > 0) linha[C_DASH.SEMANA_CRONOGRAMA - 1] = diasObra;
			if (C_DASH.SERVICOS_CONCLUIDOS > 0) linha[C_DASH.SERVICOS_CONCLUIDOS - 1] = valorOuVazioSeZero_(agg.servicosConcluidos);
			if (C_DASH.SERVICOS_PENDENTES > 0) linha[C_DASH.SERVICOS_PENDENTES - 1] = valorOuVazioSeZero_(agg.servicosPendentes);
			if (C_DASH.VERBA_UTILIZADA > 0) linha[C_DASH.VERBA_UTILIZADA - 1] = valorOuVazioSeZero_(round2_(agg.verbaUtilizada));
			if (C_DASH.PGTOS_PENDENTES > 0) linha[C_DASH.PGTOS_PENDENTES - 1] = valorOuVazioSeZero_(round2_(agg.pgtosPendentes));
			if (C_DASH.ALERTA > 0) {
				linha[C_DASH.ALERTA - 1] = (agg.servicosPendentes > 0 || agg.pgtosPendentes > 0) ? "ALERTA" : "";
			}

			saida.push(linha);
		}

		abaDash.getRange(linhaIniDash, 1, saida.length, maxColDash).setValues(saida);
		ss.toast("Dashboard atualizado com " + saida.length + " unidade(s).", "Sucesso", 4);
	});
}

/**
 * Retorna lista de unidades unicas da aba INFORMAÇÕES GERAIS.
 */
function carregarUnidadesInfoGerais_(abaInfo, C_INFO, linhaIniInfo) {
	const lastInfo = obterUltimaLinhaDados_(abaInfo, C_INFO.EMP || 1);
	if (lastInfo < linhaIniInfo) return [];

	const maxCol = Math.max(C_INFO.EMP || 0, C_INFO.UNI || 0, C_INFO.DATA_LOTE || 0, 2);
	const dados = abaInfo.getRange(linhaIniInfo, 1, lastInfo - linhaIniInfo + 1, maxCol).getValues();

	const vistos = new Set();
	const idxByChave = new Map();
	const lista = [];
	for (let i = 0; i < dados.length; i++) {
		const emp = String(dados[i][(C_INFO.EMP || 1) - 1] || "").trim();
		const uni = String(dados[i][(C_INFO.UNI || 2) - 1] || "").trim();
		if (!emp || !uni) continue;
		const dataLote = C_INFO.DATA_LOTE > 0 ? dados[i][C_INFO.DATA_LOTE - 1] : "";

		const chave = textoNormalizadoSemAcento_(emp) + "|" + textoNormalizadoSemAcento_(uni);
		if (vistos.has(chave)) {
			const idx = idxByChave.get(chave);
			if (idx != null && idx >= 0) {
				const atual = lista[idx].dataLote;
				if (!normalizarDataSomenteDia_(atual) && normalizarDataSomenteDia_(dataLote)) {
					lista[idx].dataLote = dataLote;
				}
			}
			continue;
		}

		vistos.add(chave);
		idxByChave.set(chave, lista.length);
		lista.push({ emp: emp, uni: uni, chave: chave, dataLote: dataLote });
	}
	return lista;
}

/**
 * Agrega os dados da FASE-OBRA por unidade (EMP|UNI).
 */
function agregarDadosObraPorUnidade_(abaObra, C_OBRA, linhaIniObra) {
	const lastObra = obterUltimaLinhaDados_(abaObra, C_OBRA.EMP || 1);
	const mapa = new Map();
	if (lastObra < linhaIniObra) return mapa;

	const linhasBuscaHeader = [Math.max(1, linhaIniObra - 1), 1, 2, 3];
	const colStatusAprov = obterColunaPorCabecalhoEmLinhas_(
		abaObra,
		(CONFIG.HEADERS && CONFIG.HEADERS.OBRA_STATUS_APROVACAO) || ["STATUS APROVACAO SERVICO EXECUTADO", "STATUS APROVAÇÃO SERVIÇO EXECUTADO"],
		linhasBuscaHeader
	) || C_OBRA.STATUS || -1;

	const pagamentos = resolverColunasPagamentosObra_(abaObra, linhasBuscaHeader);

	const maxCol = Math.max(
		C_OBRA.EMP || 0,
		C_OBRA.UNI || 0,
		C_OBRA.TIPO || 0,
		C_OBRA.CAT || 0,
		C_OBRA.SUB || 0,
		colStatusAprov || 0,
		pagamentos.maxCol || 0,
		2
	);

	const dados = abaObra.getRange(linhaIniObra, 1, lastObra - linhaIniObra + 1, maxCol).getDisplayValues();

	for (let i = 0; i < dados.length; i++) {
		const row = dados[i];
		const emp = String(row[(C_OBRA.EMP || 1) - 1] || "").trim();
		const uni = String(row[(C_OBRA.UNI || 2) - 1] || "").trim();
		if (!emp || !uni) continue;

		// Evita contar linhas vazias sem servico descrito.
		const tipo = C_OBRA.TIPO > 0 ? String(row[C_OBRA.TIPO - 1] || "").trim() : "";
		const cat = C_OBRA.CAT > 0 ? String(row[C_OBRA.CAT - 1] || "").trim() : "";
		const sub = C_OBRA.SUB > 0 ? String(row[C_OBRA.SUB - 1] || "").trim() : "";
		if (!tipo && !cat && !sub) continue;

		const chave = textoNormalizadoSemAcento_(emp) + "|" + textoNormalizadoSemAcento_(uni);
		if (!mapa.has(chave)) {
			mapa.set(chave, {
				servicosConcluidos: 0,
				servicosPendentes: 0,
				verbaUtilizada: 0,
				pgtosPendentes: 0
			});
		}
		const agg = mapa.get(chave);

		// D/E) Contagem por status de aprovacao do servico.
		const statusAprovRaw = colStatusAprov > 0 ? row[colStatusAprov - 1] : "";
		const statusAprovNorm = textoNormalizadoSemAcento_(statusAprovRaw);
		const ehAprovado100 = /100/.test(statusAprovNorm) && /APROVAD/.test(statusAprovNorm);
		const ehCancelado = /CANCELAD/.test(statusAprovNorm);

		if (ehAprovado100) {
			agg.servicosConcluidos++;
		}
		if (!ehAprovado100 && !ehCancelado) {
			agg.servicosPendentes++;
		}

		// F/G) Soma pagamentos por parcelas 1..5
		for (let p = 0; p < pagamentos.pares.length; p++) {
			const par = pagamentos.pares[p];
			if (!par.colStatus || !par.colValor) continue;

			const statusPag = textoNormalizadoSemAcento_(row[par.colStatus - 1]);
			if (!statusPag) continue;

			const valor = converterParaNumero_(row[par.colValor - 1]);
			const valorNum = (valor == null ? 0 : valor);

			if (/\bPAGO\b/.test(statusPag)) {
				agg.verbaUtilizada += valorNum;
			} else {
				agg.pgtosPendentes += valorNum;
			}
		}
	}
	return mapa;
}

/**
 * Resolve colunas de pagamentos na FASE-OBRA (STATUS n PAG / VALOR LIB n PAG, n=1..5).
 */
function resolverColunasPagamentosObra_(abaObra, linhasBuscaHeader) {
	const pares = [];
	let maxCol = 0;

	for (let n = 1; n <= 5; n++) {
		const colStatus = obterColunaPorCabecalhoEmLinhas_(abaObra, aliasesStatusPagamento_(n), linhasBuscaHeader);
		const colValor = obterColunaPorCabecalhoEmLinhas_(abaObra, aliasesValorPagamento_(n), linhasBuscaHeader);
		if (colStatus > maxCol) maxCol = colStatus;
		if (colValor > maxCol) maxCol = colValor;
		pares.push({ colStatus: colStatus, colValor: colValor });
	}

	return { pares: pares, maxCol: maxCol };
}

function aliasesStatusPagamento_(n) {
	return [
		"STATUS " + n + "º PAG",
		"STATUS " + n + "° PAG",
		"STATUS " + n + "O PAG",
		"STATUS " + n + " PAG",
		"STATUS " + n + "º PGTO",
		"STATUS " + n + " PGTO",
		"STATUS " + n + " PAGAMENTO"
	];
}

function aliasesValorPagamento_(n) {
	return [
		"VALOR LIB " + n + "º PAG",
		"VALOR LIB " + n + "° PAG",
		"VALOR LIB " + n + "O PAG",
		"VALOR LIB " + n + " PAG",
		"VALOR LIBERADO " + n + "º PAG",
		"VALOR LIBERADO " + n + " PAG",
		"VALOR " + n + "º PAG",
		"VALOR " + n + " PAG"
	];
}

function extrairNumeroFlex_(valor) {
	const n = converterParaNumero_(valor);
	if (n != null) return n;

	const txt = String(valor || "").trim();
	if (!txt) return null;
	const m = txt.match(/-?\d+[\.,]?\d*/);
	if (!m) return null;
	return converterParaNumero_(m[0]);
}

function round2_(n) {
	const v = Number(n || 0);
	return Math.round(v * 100) / 100;
}

function valorOuVazioSeZero_(valor) {
	const n = converterParaNumero_(valor);
	if (n != null) {
		return n === 0 ? "" : valor;
	}
	return valor;
}

/**
 * Calcula dias corridos de obra com base na DATA LOTE da INFORMAÇÕES GERAIS.
 * Retorna vazio para data invalida ou futura.
 */
function calcularDiasCorridosDesdeDataLote_(dataLote) {
	const lote = normalizarDataSomenteDia_(dataLote);
	if (!lote) return "";

	const hoje = normalizarDataSomenteDia_(new Date());
	const diffMs = hoje.getTime() - lote.getTime();
	if (diffMs < 0) return "";

	const diffDias = Math.floor(diffMs / 86400000);
	return diffDias;
}

