# 📚 Guia para Agentes: Alterações em Testes

> **Documento para manutenção e extensão da suite de testes automatizados**
> Leia antes de fazer qualquer alteração em Tests.gs ou tests.js

---

## 🎯 Visão Geral Rápida

### Arquitetura de 2 Camadas

```
┌─────────────────────────────────────┐
│ TESTES DE INTEGRAÇÃO (Tests.gs)     │
│ Roda: Google Apps Script            │
│ Função raiz: executarTodosTestes()  │
│ Saída: Aba TEST_DATA (colorida)     │
└─────────────────────────────────────┘

┌─────────────────────────────────────┐
│ TESTES ISOLADOS (tests.js)          │
│ Roda: Node.js local                 │
│ Função raiz: (4 suites automáticas) │
│ Saída: Console colorido + exit code │
└─────────────────────────────────────┘
```

### Quando usar cada um

| Situação | Usar | Velocidade | Precisa Planilha |
|----------|------|-----------|------------------|
| Desenvolvendo lógica pura (normalização, UUID) | `tests.js` | ~1s | ❌ Não |
| Testando sincronizações, fluxos | `Tests.gs` | ~20s | ✅ Sim |
| Validação pré-deploy | Ambos | ~25s | ✅ Sim |

---

## 🔍 Estrutura Atual

### Tests.gs (Google Apps Script)

**Função raiz:** `executarTodosTestes()`
```
├─ testarSincronizacaoPedidosHousi_()      [Teste 1]
├─ testarBuscaContatoFornecedor_()         [Teste 2]
├─ testarMapeamentoDinamicoColunas_()      [Teste 3]
└─ gerarRelatorioTestes_()                 [Formatação + Saída]
```

**Estrutura de um teste:**
```javascript
function testarMeuCenario_(abaTest, ss) {
  const resultados = [];
  const meuSheet = ss.getSheetByName("MINHA_ABA");
  
  if (!meuSheet) {
    resultados.push({ 
      nome: "Meu Teste", 
      status: "SKIP", 
      motivo: "Aba não encontrada" 
    });
    return resultados;
  }

  try {
    // Seu teste aqui
    const resultado = true; // boolean
    
    resultados.push({
      nome: "Descrição do Teste",
      status: resultado ? "PASS" : "FAIL",
      detalhes: "Contexto ou valor retornado"
    });
    
  } catch (e) {
    resultados.push({
      nome: "Descrição do Teste",
      status: "ERROR",
      motivo: e.message
    });
  }

  return resultados;
}
```

### tests.js (Node.js)

**Estrutura de suite:**
```javascript
const suiteX = new TestSuite("Suite X: Meu Conceito");

suiteX.test("Descrição do teste 1", () => {
  const resultado = meuAlgoritmo(input);
  suiteX.assertEqual(resultado, esperado, "Mensagem de erro");
});

suiteX.test("Descrição do teste 2", () => {
  suiteX.assertTrue(condicao, "Condição falhou");
});
```

---

## ➕ Como Adicionar um Novo Teste

### Cenário 1: Teste de Lógica Pura (Node.js)

Exemplo: Teste de cálculo de cronograma

```javascript
// 1. Criar nova suite em tests.js
const suite5 = new TestSuite("Suite 5: Cálculo de Cronograma");

// 2. Adicionar testes
suite5.test("Calcula semana corretamente (data positiva)", () => {
  const dataLote = new Date(2026, 2, 27);      // 27/03/2026
  const dataPlanejada = new Date(2026, 3, 3); // 03/04/2026 (7 dias depois)
  
  const semana = Math.ceil((dataPlanejada - dataLote) / (7 * 24 * 60 * 60 * 1000));
  
  suite5.assertEqual(semana, 1, "Primeira semana");
});

suite5.test("Calcula semana corretamente (data negativa)", () => {
  const dataLote = new Date(2026, 2, 27);
  const dataPlanejada = new Date(2026, 2, 20); // Antes do lote
  
  const semana = Math.ceil((dataPlanejada - dataLote) / (7 * 24 * 60 * 60 * 1000));
  
  suite5.assertEqual(semana, -1, "Semana negativa");
});

// 3. Registrar suite no final (linha ~160)
const suites = [suite1, suite2, suite3, suite4, suite5];
```

**Então raie:** `npm run test:local`

### Cenário 2: Teste de Sincronização (Apps Script)

Exemplo: Teste de sincronização tipo Y

```javascript
// 1. Adicionar função em Tests.gs (após linha 250)
function testarMinhaNovaFeature_(abaTest, ss) {
  const resultados = [];
  const meuSheet = ss.getSheetByName("MINHA_ABA");
  
  if (!meuSheet) {
    resultados.push({ 
      nome: "Minha Nova Feature", 
      status: "SKIP", 
      motivo: "Aba não encontrada" 
    });
    return resultados;
  }

  try {
    // Setup: Criar dados de teste
    const C = resolveSheetColumns_(meuSheet, CONFIG.HEADERS_COLS.MINHA_ABA, CONFIG.COLUMNS.MINHA_ABA);
    const linhaIni = obterLinhaInicialPorAba("MINHA_ABA");
    
    // Ação: Executar função que quer testar
    minhaFuncao_({ range: meuSheet.getRange(linhaIni, 1, 1, 10), source: ss });
    
    // Verificação: Checar resultado
    const resultado = meuSheet.getRange(linhaIni, C.MINHA_COLUNA).getDisplayValue();
    const passou = resultado === "ESPERADO";
    
    resultados.push({
      nome: "Teste Minha Nova Feature: Caso 1",
      status: passou ? "PASS" : "FAIL",
      detalhes: passou ? "OK" : `Esperado: ESPERADO, Obtido: ${resultado}`
    });
    
  } catch (e) {
    resultados.push({
      nome: "Minha Nova Feature",
      status: "ERROR",
      motivo: e.message
    });
  }

  return resultados;
}

// 2. Registrar no executarTodosTestes() (linha ~27)
resultados.push(...testarMinhaNovaFeature_(abaTest, ss));

// 3. Executar no Apps Script
```

**Então raie:** Extensões → Apps Script → ▶ Executar

---

## 📝 Convenções Obrigatórias

### Nomes de Funções

```javascript
testar[CONCEITO]_()          // Sempre com underscore final
testarSincronizador()        // ❌ Errado
testarSincronizador_()       // ✅ Certo
```

### Estrutura de Resultado

Cada teste deve retornar array com objetos:

```javascript
{
  nome: "String descritiva do teste",      // ✅ Obrigatório
  status: "PASS|FAIL|ERROR|SKIP",          // ✅ Obrigatório
  detalhes: "Resultado ou motivo",         // ⚠️  Recomendado
  motivo: "Por que pulou"                  // ⚠️  Se status === SKIP
}
```

### Status Meanings

| Status | Quando usar | Exemplo |
|--------|------------|---------|
| `PASS` | Teste passou ✅ | `resultado === esperado` |
| `FAIL` | Teste falhou ❌ | `resultado !== esperado` |
| `ERROR` | Exceção lançada | `throw new Error()` captado em `catch` |
| `SKIP` | Não pôde testar | Aba não existe, dados incompletos |

---

## 🔗 Dependências Entre Testes

### Cuidado: Não quebre testes existentes!

**Ordem de execução (Tests.gs):**
```javascript
1. testarSincronizacaoPedidosHousi_()
2. testarBuscaContatoFornecedor_()
3. testarMapeamentoDinamicoColunas_()
```

**Impacto cruzado:**
- ❌ Não modifique estrutura de dados das suites anteriores
- ❌ Não delete colunas do CONFIG.HEADERS_COLS
- ✅ Você pode adicionar novos testes ao final
- ✅ Você pode refinar lógica DENTRO de um teste

---

## 🧪 Checklist: Antes de Submeter Alteração

- [ ] Teste local passou? `npm run test:local` retornou 0?
- [ ] Teste em Sheets passou? Aba TEST_DATA está 100% verde?
- [ ] Adicionei comentários explicando meu novo teste?
- [ ] Meu teste é isolado e não quebra testes existentes?
- [ ] Nomeei funções com `_()` no final?
- [ ] Retorno sempre arrays com objetos `{ nome, status, detalhes }`?
- [ ] Tratei exceções em `try/catch`?
- [ ] Atualizei este documento se mudei a arquitetura?

---

## 🐛 Troubleshooting para Agentes

### Problema: "TypeError: console.clear is not a function"

**Causa:** Google Apps Script não tem `console.clear()`  
**Solução:** Remove linha com `console.clear()`, use apenas `console.log()`

```javascript
// ❌ Errado
console.clear();

// ✅ Certo
console.log("Iniciando testes...");
```

### Problema: Teste retorna SKIP para todas as abas

**Causa:** Abas não existem ou estão com nome errado  
**Solução:** Verifica CONFIG.SHEETS

```javascript
// Verifica nomes corretos em Config.gs:
console.log(CONFIG.SHEETS.OBRA);      // "FASE-OBRA"?
console.log(CONFIG.SHEETS.PEDIDOS);   // "PEDIDOS-GERAL"?
```

### Problema: Teste passa localmente mas falha em Sheets

**Causa:** Mock em `tests.js` não replica comportamento real  
**Solução:** Teste sempre em Sheets também com `executarTodosTestes()`

```javascript
// Faça ambos:
npm run test:local        # Local rápido
# Depois execute Tests.gs no Apps Script  # Validação real
```

---

## 📊 Cobertura Atual vs Necessária

```
✅ IMPLEMENTADO:
  ├─ Sincronização Obra→Pedidos (60%)
  ├─ Busca Contato Fornecedor (80%)
  └─ Mapeamento Dinâmico (50%)

❌ TODO:
  ├─ Cálculo Cronograma/Semana (0%)
  ├─ Sincronização Bidirecional (0%)
  ├─ Limpeza de Órfãos (0%)
  └─ Validações gerais (0%)
```

**Próximo agente:** Considere adicionar testes para **Cronograma** (alto impacto, baixa complexidade)

---

## 📞 Contexto da Correção Original

### Por que Tests.gs foi criado?

A função `buscarContatoFornecedor_()` em [SheetPedidos.gs](SheetPedidos.gs#L37) tinha um bug:
- ❌ Retornava datas aleatórias (31/12/1969) em vez de telefone
- ❌ Fallback cego podia puxar coluna errada
- ❌ Nenhum teste verificava esta lógica

**Solução implementada:**
1. ✅ Corrigir lógica em SheetPedidos.gs (buscar em múltiplas linhas de header)
2. ✅ Forçar formato de texto na coluna CONTATO
3. ✅ Criar teste que valida isso NÃO acontece novamente

**Este documento existe para:** Garantir que futuros agentes entendam a cadeia de causa-efeito e não reintroduzam o bug.

---

## 🚀 Exemplo Completo: Adicionar Novo Teste

### Tarefa: Testar validação de EMPREENDIMENTO vazio

**Passo 1:** Crie função em Tests.gs

```javascript
function testarValidacaoEmpreendimentoVazio_(abaTest, ss) {
  const resultados = [];
  const sheetObra = ss.getSheetByName(CONFIG.SHEETS.OBRA);
  
  if (!sheetObra) {
    resultados.push({ 
      nome: "Validação Empreendimento Vazio", 
      status: "SKIP", 
      motivo: "Aba FASE-OBRA não encontrada" 
    });
    return resultados;
  }

  try {
    const iniObra = obterLinhaInicialPorAba(CONFIG.SHEETS.OBRA);
    const C_OBRA = resolveSheetColumns_(sheetObra, CONFIG.HEADERS_COLS.OBRA, CONFIG.COLUMNS.OBRA);
    
    // Tenta criar linha com EMP vazio
    const novaLinha = iniObra + 99;
    const dadosTest = [
      ["", "1501", "João", "Maria", "", "Teste", "Cat", "Sub", "HOUSI", "PEDIDO AG", "FORN", null, null]
    ];
    
    sheetObra.getRange(novaLinha, 1, 1, C_OBRA.ATRELADO).setValues(dadosTest);
    
    // Valida: não sincroniza se EMP está vazio
    sincronizarPedidosHousiPorEdicao_({
      range: sheetObra.getRange(novaLinha, 1, 1, C_OBRA.ATRELADO),
      source: ss
    });
    
    const abaPedidos = ss.getSheetByName(CONFIG.SHEETS.PEDIDOS);
    const iniPed = obterLinhaInicialPorAba(CONFIG.SHEETS.PEDIDOS);
    const C_PED = resolveSheetColumns_(abaPedidos, CONFIG.HEADERS_COLS.PEDIDOS, CONFIG.COLUMNS.PEDIDOS);
    const linhasVazias = abaPedidos.getRange(iniPed, 1, 10, 1).getDisplayValues()
      .filter(r => String(r[0]).trim() === "").length;
    
    resultados.push({
      nome: "Validação Empreendimento Vazio: Não sincroniza",
      status: linhasVazias === 0 ? "PASS" : "FAIL",
      detalhes: linhasVazias === 0 ? "OK - Não sincronizou" : `Erro: Criou ${linhasVazias} linhas vazias`
    });
    
  } catch (e) {
    resultados.push({
      nome: "Validação Empreendimento Vazio",
      status: "ERROR",
      motivo: e.message
    });
  }

  return resultados;
}
```

**Passo 2:** Registre em `executarTodosTestes()`

```javascript
// Após linha ~31, adicione:
resultados.push(...testarValidacaoEmpreendimentoVazio_(abaTest, ss));
```

**Passo 3:** Raie e valide

```bash
# Abra Google Apps Script
# Extensões → Apps Script
# Execute executarTodosTestes()
# Verifique aba TEST_DATA
```

---

## 📚 Referências

- [QUICK_START.md](QUICK_START.md) — Como executar testes
- [TESTES.md](TESTES.md) — Documentação detalhada
- [SheetPedidos.gs](SheetPedidos.gs#L37) — Função que motivou os testes
- [tests.js](tests.js) — Suite local Node.js
- [Config.gs](Config.gs#L1) — Definições de nomes de colunas

---

**Última atualização:** 2 de abril de 2026  
**Status:** ✅ Ativo — Leia antes de qualquer alteração  
**Mantido por:** Agentes de IA + Engenheiros
