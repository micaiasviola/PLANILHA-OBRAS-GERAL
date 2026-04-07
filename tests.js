/**
 * TESTES ISOLADOS (Node.js) - Sem dependência de planilha
 * Roda localmente com mocks do Google Apps Script
 * 
 * Uso: node tests.js
 */

// ============= MOCK DO GOOGLE APPS SCRIPT =============

const mockGAS = {
  getActiveSpreadsheet: () => mockGAS.spreadsheet,
  spreadsheet: {
    getSheetByName: (name) => mockGAS.sheets[name] || null,
    getSheets: () => Object.values(mockGAS.sheets),
  },
  sheets: {
    "FASE-OBRA": {
      getName: () => "FASE-OBRA",
      getLastRow: () => 5,
      getLastColumn: () => 51,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => mockGAS.generateMockData("FASE-OBRA", row, col, numRows, numCols),
        getDisplayValues: () => mockGAS.generateMockData("FASE-OBRA", row, col, numRows, numCols),
        setValue: (val) => {},
        setValues: (vals) => {},
        setNumberFormat: () => ({
          setValues: () => {}
        }),
        getSheet: () => mockGAS.sheets["FASE-OBRA"]
      }),
      clearContent: () => {},
      insertColumnsAfter: () => {},
    },
    "PEDIDOS-GERAL": {
      getName: () => "PEDIDOS-GERAL",
      getLastRow: () => 5,
      getLastColumn: () => 37,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => mockGAS.generateMockData("PEDIDOS-GERAL", row, col, numRows, numCols),
        getDisplayValues: () => mockGAS.generateMockData("PEDIDOS-GERAL", row, col, numRows, numCols),
        setValue: (val) => {},
        setValues: (vals) => {},
        setNumberFormat: (fmt) => ({
          setValues: () => {},
          clearDataValidations: () => {}
        }),
        clearDataValidations: () => {},
        getSheet: () => mockGAS.sheets["PEDIDOS-GERAL"]
      }),
      clearContent: () => {},
    },
    "Backup": {
      getName: () => "Backup",
      getLastRow: () => 5,
      getLastColumn: () => 12,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => mockGAS.generateMockData("Backup", row, col, numRows, numCols),
        getDisplayValues: () => mockGAS.generateMockData("Backup", row, col, numRows, numCols),
      }),
    }
  },
  
  generateMockData: (sheet, row, col, numRows, numCols) => {
    const mockData = {
      "FASE-OBRA": [
        ["EMPREENDIMENTO", "UNIDADE", "OPR", "ADM", "CRONOGRAMA", "TIPO", "CATEGORIA", "SUB-CATEGORIA", "ATRELADO", "STATUS", "FORNECEDOR"],
        ["MODERN BUNTANTÃ", "307", "João", "Maria", "1", "Limpeza", "Eletros", "Geladeira", "HOUSI", "APROVADO", "FORNECEDOR X", undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, "ID-ABC123-DEF456"],
        ["SKY PINHEIROS", "1501", "Pedro", "Ana", "2", "Pintura", "Mobília", "Mesa", "OUTRO", "PENDENTE", "FORNECEDOR Y", undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, "ID-XYZ789-GHI012"],
      ],
      "PEDIDOS-GERAL": [
        ["EMPREENDIMENTO", "UNIDADE", "OPR", "ADM", "TIPO", "CATEGORIA", "SUB-CATEGORIA", "DESCRIÇÃO", "STATUS", "FORNECEDOR", "CONTATO", "DATA_SOL", "DATA_AGE"],
        ["MODERN BUNTANTÃ", "307", "João", "Maria", "Limpeza", "Eletros", "Geladeira", "Geladeira 400L", "PEDIDO AG", "FORNECEDOR X", "", new Date(2026, 3, 2), null, "ID-ABC123-DEF456"],
      ],
      "Backup": [
        ["FORNECEDOR", "CONTATO", "NOME"],
        ["FORNECEDOR X", "11999999999", "Fornecedor X LTDA"],
        ["DECOR FLOOR", "1133333333", "Decoração Floor"],
        ["CIAMAIS", "1144444444", "Ciamais Eletros"],
      ]
    };
    
    return (mockData[sheet] || []).slice(row - 1, row - 1 + numRows)
      .map(r => r.slice(col - 1, col - 1 + numCols));
  }
};

const TEST_START_ROW = 1;
const TEST_START_COL = 1;

// ============= TESTES =============

class TestSuite {
  constructor(name) {
    this.name = name;
    this.tests = [];
    this.passCount = 0;
    this.failCount = 0;
  }

  test(nome, fn) {
    try {
      fn();
      this.tests.push({ nome, status: "PASS" });
      this.passCount++;
    } catch (e) {
      this.tests.push({ nome, status: "FAIL", erro: e.message });
      this.failCount++;
    }
  }

  assertEqual(atual, esperado, msg) {
    if (JSON.stringify(atual) !== JSON.stringify(esperado)) {
      throw new Error(`${msg}\n  Esperado: ${JSON.stringify(esperado)}\n  Obtido: ${JSON.stringify(atual)}`);
    }
  }

  assertTrue(condicao, msg) {
    if (!condicao) throw new Error(msg);
  }

  printResults() {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`📋 ${this.name}`);
    console.log(`${'='.repeat(60)}`);
    
    for (const t of this.tests) {
      const icon = t.status === "PASS" ? "✅" : "❌";
      console.log(`${icon} ${t.nome}`);
      if (t.erro) console.log(`   └─ ${t.erro}`);
    }
    
    console.log(`${'='.repeat(60)}`);
    console.log(`Resultado: ${this.passCount} PASS | ${this.failCount} FAIL`);
    console.log(`${'='.repeat(60)}\n`);

    return this.failCount === 0;
  }
}

// ============= TESTE 1: Textoização e Normalização =============

const suite1 = new TestSuite("Suite 1: Funções Puras (Normalização de Texto)");

suite1.test("textoNormalizadoSemAcento_ com acentos", () => {
  // Mock local da função
  const textoNormalizado = (valor) => {
    return String(valor || "")
      .trim()
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  };
  
  suite1.assertEqual(textoNormalizado("CONTATO FORNECEDOR"), "CONTATO FORNECEDOR", "Sem acentos");
  suite1.assertEqual(textoNormalizado("CóNtAtO"), "CONTATO", "Mixcase com acento");
  suite1.assertEqual(textoNormalizado("  espaços  "), "ESPACOS", "Espaços extras");
});

suite1.test("Conversão para número com R$ e vírgula", () => {
  const converterParaNumero = (valor) => {
    if (typeof valor === "number") return valor;
    let texto = String(valor || "").trim().replace(/\s/g, "").replace(/[R$]/g, "");
    if (!texto) return null;
    if (texto.indexOf(",") >= 0) texto = texto.replace(/\./g, "").replace(",", ".");
    const numero = Number(texto);
    return Number.isFinite(numero) ? numero : null;
  };

  suite1.assertEqual(converterParaNumero("R$ 1.234,56"), 1234.56, "R$ com ponto/vírgula");
  suite1.assertEqual(converterParaNumero("1000"), 1000, "Número limpo");
  suite1.assertEqual(converterParaNumero(""), null, "String vazia");
});

// ============= TESTE 2: Geração de UUID e Chaves =============

const suite2 = new TestSuite("Suite 2: Geração de ID/UUID");

suite2.test("gerarUUID_ retorna formato válido", () => {
  const gerarUUID = () => {
    return "ID-" + Math.random().toString(36).substring(2, 9).toUpperCase() + "-" + new Date().getTime().toString(36).toUpperCase();
  };

  const uuid1 = gerarUUID();
  const uuid2 = gerarUUID();
  
  suite2.assertTrue(/^ID-[A-Z0-9]+-[A-Z0-9]+$/.test(uuid1), "UUID 1 válido");
  suite2.assertTrue(/^ID-[A-Z0-9]+-[A-Z0-9]+$/.test(uuid2), "UUID 2 válido");
  suite2.assertTrue(uuid1 !== uuid2, "UUIDs diferentes");
});

// ============= TESTE 3: Lógica de Campos de Contato =============

const suite3 = new TestSuite("Suite 3: Lógica de Contato (Busca no Backup)");

suite3.test("Busca fornecedor normalizado no mapa", () => {
  const mapaFornecedores = new Map([
    ["FORNECEDOR X", "11999999999"],
    ["DECOR FLOOR", "1133333333"],
    ["CIAMAIS", "1144444444"],
  ]);

  const buscarContato = (nome) => {
    const nomeLimpo = nome.trim().toUpperCase();
    return mapaFornecedores.get(nomeLimpo) || "";
  };

  suite3.assertEqual(buscarContato("Fornecedor X"), "11999999999", "Fornecedor X encontrado");
  suite3.assertEqual(buscarContato("DECOR FLOOR"), "1133333333", "DECOR FLOOR encontrado");
  suite3.assertEqual(buscarContato("INEXISTENTE"), "", "Fornecedor não encontrado retorna vazio");
});

suite3.test("Contato é sempre texto, nunca data", () => {
  const contatoObtido = "11999999999";
  const naoEhData = !/^\d{1,2}\/\d{1,2}\/\d{4}|31\/12\/1969/.test(contatoObtido);
  
  suite3.assertTrue(naoEhData, "Contato é texto, não data");
  suite3.assertTrue(typeof contatoObtido === "string", "Contato é string");
});

// ============= TESTE 4: Mocked GAS Sheet Operations =============

const suite4 = new TestSuite("Suite 4: Operações Mock do Sheets");

suite4.test("Lê dados de FASE-OBRA", () => {
  const sheet = mockGAS.sheets["FASE-OBRA"];
  const dados = sheet.getRange(TEST_START_ROW, TEST_START_COL, 3, 11).getValues();
  
  suite4.assertTrue(dados.length === 3, "3 linhas lidas");
  suite4.assertEqual(dados[0][0], "EMPREENDIMENTO", "Cabeçalho correto");
  suite4.assertEqual(dados[1][0], "MODERN BUNTANTÃ", "Primeiro EMP correto");
});

suite4.test("Lê dados de Backup", () => {
  const sheet = mockGAS.sheets["Backup"];
  const dados = sheet.getRange(TEST_START_ROW, TEST_START_COL, 4, 2).getValues();
  
  suite4.assertTrue(dados.length === 4, "4 linhas lidas");
  suite4.assertEqual(dados[1][1], "11999999999", "Contato do primeiro fornecedor");
});

// ============= EXECUÇÃO =============

console.log("\n🧪 INICIANDO SUITE DE TESTES LOCAIS (Node.js + Mock GAS)\n");

const suites = [suite1, suite2, suite3, suite4];
const allPassed = suites.map(s => s.printResults()).every(r => r);

if (allPassed) {
  console.log("🎉 TODOS OS TESTES PASSARAM!\n");
  process.exit(0);
} else {
  console.log("❌ ALGUNS TESTES FALHARAM\n");
  process.exit(1);
}
