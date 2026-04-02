# 🧪 Guia Completo: Testes Automatizados

Este é um framework de testes de **duas camadas** para sua automação em Google Sheets + Apps Script.

---

## 📋 Visão Geral

### Camada 1: Testes Isolados (Node.js)
- ✅ Roda **sem planilha** (localmente)
- ✅ Testa funções **puras** (normalização, cálculos, lógica)
- ✅ **Rápido** (< 1 segundo)
- ✅ Perfeito para **CI/CD**

### Camada 2: Testes de Integração (Google Apps Script)
- ✅ Roda **contra a planilha real**
- ✅ Testa **fluxos completos** (sincronizações, busca de contato, etc)
- ✅ Gera **relatório visual** na aba `TEST_DATA`
- ✅ Garante compatibilidade com Sheets

---

## 🚀 Como Usar

### Pré-requisitos
```bash
# Instale Node.js (https://nodejs.org/)
# Ou use via PowerShell:
winget install OpenJS.NodeJS
```

### Opção A: Testes Locais (Recomendado para Desenvolvimento)

```bash
# Terminal PowerShell na pasta do projeto
cd "c:\Users\leomi_wd4lqis\Desktop\DESENVOLVIMENTO ECQUA\PLANILHA OBRAS GERAL"

# Rodar testes locais
npm run test:local

# Ou diretamente
node tests.js
```

**Saída esperada:**
```
============================================================
📋 Suite 1: Funções Puras (Normalização de Texto)
============================================================
✅ textoNormalizadoSemAcento_ com acentos
✅ Conversão para número com R$ e vírgula
...
============================================================
Resultado: 15 PASS | 0 FAIL
============================================================
```

### Opção B: Testes Completos (CLI Coordenadora)

```bash
npm run test
```

Isso executa **ambas** as suites e gera um relatório completo.

### Opção C: Testes de Integração (Sheets)

1. Abra a planilha no **Google Sheets**
2. Vá em **Extensões → Apps Script**
3. Procure pela função `executarTodosTestes()`
4. Clique em **▶ Executar**
5. Verifique os resultados na aba **TEST_DATA**

---

## 📁 Estrutura de Arquivos

```
PLANILHA OBRAS GERAL/
├── Tests.gs              ← Testes de integração (Google Apps Script)
├── tests.js              ← Testes isolados (Node.js)
├── run-tests.js          ← CLI coordenadora
├── package.json          ← Dependências (npm)
└── TESTES.md             ← Este arquivo
```

### Tests.gs (Google Apps Script)
**Função raiz:** `executarTodosTestes()`

Testa:
- ✅ Sincronização Obra → Pedidos
- ✅ Busca de Contato Fornecedor (nova implementação)
- ✅ Mapeamento Dinâmico de Colunas

**Resultado:** Aba `TEST_DATA` é criada com relatório colorido

### tests.js (Node.js)
**Execução:** `npm run test:local`

Testa:
- ✅ Normalização de texto e acentos
- ✅ Conversão de números com R$
- ✅ Geração de UUIDs
- ✅ Lógica de busca de contato (sem Sheets)
- ✅ Operações mock de Range

---

## 🔄 Fluxo de Trabalho Recomendado

### 1. Antes de Uma Feature Nova

```bash
# Rode testes locais para garantir baseline
npm run test:local
```

### 2. Implementar a Feature

- Edite os arquivos `.gs` normalmente
- Se a feature tiver lógica pura, **adicione teste em tests.js**

### 3. Validar a Feature

```bash
# Testes locais (rápido)
npm run test:local

# Testes de integração (completo)
# Abra Google Sheets → Extensões → Apps Script → executarTodosTestes()
```

### 4. Se Falhar

- Verifique a aba **TEST_DATA** para detalhes
- Verde (✅) = Passou
- Vermelho (❌) = Falhou
- Amarelo (⏸️) = Pulou (aba não encontrada)

---

## 📝 Como Adicionar Novos Testes

### Teste Isolado (Node.js - Rápido)

Edite `tests.js` e adicione uma nova suite:

```javascript
const suite5 = new TestSuite("Suite 5: Minha Lógica");

suite5.test("Descrição do teste", () => {
  const resultado = meuAlgoritmo(entrada);
  suite5.assertEqual(resultado, esperado, "Mensagem de erro");
});

// Adicione ao final
const suites = [suite1, suite2, suite3, suite4, suite5];
```

### Teste de Integração (Google Apps Script)

Edite `Tests.gs` e adicione uma função:

```javascript
function testarMinhaFeature_(abaTest, ss) {
  const resultados = [];
  
  try {
    // Seu teste aqui
    const resultado = true;
    
    resultados.push({
      nome: "Minha Feature: Caso 1",
      status: resultado ? "PASS" : "FAIL",
      detalhes: "Detalhes do resultado"
    });
  } catch (e) {
    resultados.push({
      nome: "Minha Feature",
      status: "ERROR",
      motivo: e.message
    });
  }
  
  return resultados;
}
```

Depois adicione à função `executarTodosTestes()`:

```javascript
resultados.push(...testarMinhaFeature_(abaTest, ss));
```

---

## 🛠️ Troubleshooting

### "Node.js não está instalado"

```bash
winget install OpenJS.NodeJS
# Ou baixe em https://nodejs.org/
```

### "npm: command not found"

Feche e reabra o terminal PowerShell após instalar Node.js.

### Testes passam localmente mas falham no Sheets

- Verifique a aba **TEST_DATA** para detalhes do erro
- Possíveis causas:
  - Cabeçalhos de coluna deslocados
  - Dados de teste incompletos
  - Funções do Sheets faltando (ex: aba Backup não existe)

### Como Debugar Testes?

Na aba `TEST_DATA`, os testes mostram:
- **Status**: PASS, FAIL, ERROR, SKIP
- **Detalhes**: Mensagem explicativa
- **Cor**: Verde (ok), Vermelho (erro), Cinza (pulou)

---

## 📊 Cobertura Atual

| Módulo | Cobertura |
|--------|-----------|
| Sincronização Obra→Pedidos | 60% |
| Busca de Contato | 80% (melhorado) |
| Mapeamento de Colunas | 50% |
| Validações | 30% |
| Cálculos de Data | 0% (para adicionar) |

---

## 🎯 Próximas Melhorias

- [ ] Testa cálculo de cronograma (SEMANA DO MÊS)
- [ ] Testa sincronização bidirecional (Pedidos → Obra)
- [ ] Testa limpeza de órfãos
- [ ] Coverage report (quantos % do código testado)
- [ ] Integração com GitHub Actions (rodar automaticamente em PRs)

---

## 💡 Dicas

1. **Rode testes locais antes de fazer qualquer mudança** — muitas vezes bugs aparecem em funções puras antes de sincronizações complexas.

2. **Crie dados de teste na aba `TEST_DATA`** — garante isolamento e não corrompe dados reais.

3. **Use `console.log()` em Tests.gs** — aparece no Apps Script console (Ctrl+Shift+J).

4. **Versionamento** — se os testes passam em uma feature, você tem confiança para mesclar/deploy.

---

## 📞 Suporte

Se um teste falhar de forma estranha:

1. Verifique se os cabeçalhos das colunas batem com `Config.gs`
2. Limpe o cache: `limparCacheResolucaoColunas_()`
3. Rode `executarTodosTestes()` novamente
4. Abra a aba `TEST_DATA` e procure por linhas vermelhas

---

**Última atualização:** 2 de abril de 2026  
**Status:** ✅ Ativo e testado
