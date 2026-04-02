# ⚡ Quick Start: Testes Automatizados

## 1️⃣ Seu Primeiro Teste Local (30 segundos)

```powershell
cd "c:\Users\leomi_wd4lqis\Desktop\DESENVOLVIMENTO ECQUA\PLANILHA OBRAS GERAL"
node tests.js
```

**Resultado esperado:** ✅ TODOS OS TESTES PASSARAM!

---

## 2️⃣ Testar Sua Planilha (3 minutos)

1. Abra a planilha no **Google Sheets**
2. Clique em **Extensões → Apps Script**
3. Procure : `executarTodosTestes`
4. Clique no botão **▶ Executar**
5. Aguarde 10-20 segundos
6. Verifique a aba **TEST_DATA** (nova aba será criada)

---

## 3️⃣ Adicionar um Novo Teste

### Se for lógica pura (sem Sheets):

Edite `tests.js`, encontre o final:

```javascript
const suite5 = new TestSuite("Minha Suite");
suite5.test("Meu teste", () => {
  suite5.assertEqual(resultado, esperado, "mensagem");
});
```

### Se for fluxo completo (com Sheets):

Edite `Tests.gs`, procure `executarTodosTestes()`:

```javascript
resultados.push(...testarMinhaFeature_(abaTest, ss));
```

---

## 🎯 Checklist: Antes de fazer Deploy

- [ ] `npm run test:local` passou?
- [ ] `executarTodosTestes()` passou (aba TEST_DATA verde)?
- [ ] Testei manualmente o cenário crítico?
- [ ] Cabeçalhos de coluna não mudaram?

---

## 📞 Estrutura de Arquivos

```
Tests.gs             ← Testes no Sheets (rodar no Apps Script)
tests.js             ← Testes locais (rodar no terminal: npm run test:local)
run-tests.js         ← CLI (coordena ambos)
package.json         ← Dependências (npm)
TESTES.md            ← Documentação completa
QUICK_START.md       ← Este arquivo
```

---

**Próximo passo:** `npm run test:local` agora mesmo! 🚀
