# Regras de Negócio e Arquitetura - Planilha de Gestão de Obras ECQUA/HOUSI

Este documento serve como a **Fonte da Verdade** (Contexto Mestre) para qualquer inteligência artificial ou desenvolvedor humano que vá iterar sob o código desta automação no Google Apps Script.

---

## 📌 1. Objetivo Principal do Sistema
O sistema (desenvolvido em Google Sheets + Google Apps Script) tem como objetivo gerenciar a esteira completa ("pipeline") de obras de implantação/reforma de apartamentos.
Ele permite acompanhar desde a vistoria inicial e liberação de chaves, passando por cronograma de execuções de serviços, logística de entrega de pedidos (mobília, eletros, etc.), controle de atrasos gerais e, por fim, a vistoria final de devolução da unidade.

Tudo é organizado por **Empreendimento (EMP)** e **Unidade (UNI)**.

---

## 🏛️ 2. Arquitetura Modular (GAS)
O monólito antigo `script.js` foi abolido. O projeto utiliza uma arquitetura reativa, separada por escopo funcional (módulos separadas em arquivos `.gs` individuais).

*   **`Config.gs`**: O "Cérebro" estático. Contém os nomes exatos das abas, os REGEX de status, e o dicionário de variáveis com o NOME DOS CABEÇALHOS (para resolução dinâmica) e seus índices de fallback.
*   **`Utils.gs`**: Funções genéricas puras e auxiliares (Locks, Triggers, Tratamento de Datas, UUID, Resolvedor Dinâmico de Colunas `resolveSheetColumns_`).
*   **`Main.gs`**: Ponto de entrada. Declara o `onOpen` (menus) e o "Event Router" `onEdit`, que intercepta a edição do usuário e delega rapidamente para o handler da aba correspondente (com tratamento hierárquico de erro).
*   **`Sheet[Aba].gs`**: Ex.: `SheetObra.gs`, `SheetPedidos.gs`. Cada aba possui seu próprio tratador (Ex: `handleObraEdit(e)`), responsável por regras visuais rápidas ou preparo antes de jogar dados para outras abas.
*   **`SyncLogic.gs`**: Funções de sincronização "Cruzadas", responsáveis por processar transferências demoradas entre duas ou mais abas diferentes (ex: Orfanatos, sincronização Obra <-> Pedidos).

---

## 🔄 3. Funcionamento Geral e o Novo Resolvedor de Colunas (Dinâmico)
### 3.1. Mapeamento por Cabeçalho (Fim das Colunas Fixas)
O sistema foi reescrito para sobreviver à alteração visual promovida por supervisores.
O código **não amarra** mais, por exemplo, o índice numérico (ex: "Coluna E") para ler uma data. 

*   O script chama a função `resolveSheetColumns_(sheet, configHeadersObj, configIndicesFallback)`.
*   Esta função procura (nas primeiras 3 linhas da tabela) **exatamente** o texto do cabeçalho mapeado em `Config.gs` (A propriedade `HEADERS_COLS`).
*   Se o script encontrar a coluna "EMPREENDIMENTO" na Coluna K, ele passará a usar a Coluna K dinamicamente para aquela execução. *(Isso aplica-se apenas aos cabeçalhos previamente mapeados no `Config.gs`)*.
*   A leitura de colunas possui **cache temporário (`_colCache_`)**, ou seja, só varre o cabeçalho a 1ª vez durante uma sessão de execução de Script (evitando dezenas de chamadas lentas à API).

---

## 🛠️ 4. Fluxos de Trabalho (A Esteira de Abas)

### 4.1. INFORMAÇÕES GERAIS ("Dashboard Operacional")
*   **Papel**: Listar TUDO de forma macro. Onde o usuário cria Novas Linhas no Topo (via Menu Automático) para introduzir uma nova unidade no pipeline.
*   **Sincronização**: Os resumos de PENDÊNCIAS, ATRASOS, OCORRÊNCIAS aparecem aqui automaticamente, calculados por triggers Noturnos (23:00 e 01:00 AM) baseados no andamento da `FASE-OBRA`.
*   **Status da Unidade**: Possui uma coluna de **"STATUS OBRA"** (ATIVA/FINALIZADA) calculada automaticamente com base no status da `FASE-ENTREGA`. Unidades FINALIZADAS são ignoradas em filtros de obras ativas.
*   **Dropdowns Mestre**: A lista de unidades por empreendimento é derivada da "Aba Base Backup" com validação `DataValidation`.
    *   Em **INFORMAÇÕES GERAIS** o dropdown é aplicado como base operacional principal.
    *   Em **OCORRÊNCIAS** também há aplicação automática de Unidade por Empreendimento no fluxo de edição (A -> B).
    *   Nas demais abas, prevalece o modelo de texto para reduzir custo de recálculo e travamentos.

### 4.2. FASE-PRELIMINAR (Vistoria Inicial)
*   **Papel**: Checklist de check-in (vistoria preliminar) e pendências construtivas. Onde são avaliadas as condições do imóvel e definidos o *Responsável Operacional (OPR)* e *Administrativo (ADM)*.
*   **Gatilho Mestre de Obras**: A unidade fica qualificada para iniciar obras através do ativador manual na coluna **"FASE-OBRA"** (Dropdown SIM/NÃO).
*   **Sincronizações**: Quando a linha atualiza suas Ocorrências (vindos da aba Ocorrências), envia as quantidades e pendências de volta para a Informações Gerais. As colunas G e H também disparam fluxos automáticos.

### 4.3. FASE-OBRA (A Espinha Dorsal)
*   **Papel**: Controle granulado (serviço por serviço).
*   **Geração de Templates**: O usuário clica no Botão Superior: `"Gerar templates pendentes FASE-OBRA"`. O Script rastreia em massa todas as unidades que possuem o ativador manual **"FASE-OBRA"** como **SIM** (e que já não estejam na Obra). E então ele gera em LOTE (na memóra O(1)) 30 linhas de Serviços Padrões por unidade nova (Limpeza, Elétrica, Mobília, Eletros, Mármores...).
*   **Sincronização com Pedidos (Housi)**:
    *   Sempre que um serviço tiver a Coluna `CATEGORIA` preenchida e O Fornecedor Atrelado (`ATRELADO`) for igual a `"HOUSI"`, **essa linha é espelhada instantaneamente para a aba PEDIDOS-GERAL**.
    *   Caso os dados mudem, a sincronização é bidirecional de alguns campos logísticos.
*   **Gestão de Cronograma (Semana)**:
    *   Existe uma coluna **"SEMANA CRONOGRAMA"** que calcula automaticamente em qual semana o serviço está agendado.
    *   **Cálculo**: `Math.ceil((Data Planejada - Data Lote) / 7)`.
    *   Exemplo: Se o lote é 27/03, o dia 03/04 (7º dia) é 1ª semana. O dia 09/04 (13º dia) é 2ª semana.
*   **Gatilho Fase Entrega**: O usuário pode marcar Ativo um Dropdown customizado `"ENVIAR P/ ENTREGA"`. Ao fazer isso massivamente, o Automator ("Sincronizar Envios para Fase-Entrega" via UI) copiará a Unidade para a aba final.
*   **IDs Estáveis UUID (A Magia)**: Ao criar serviços ou transferi-los prara Pedidos, o script gera silenciosamente uma **CHAVE ESTATICA UUID (Coluna Especial Técnica AY)**. É por essa chave garantida (e não mais por nome do Empreendimento + Serviço) que o script sabe quem é quem na tabela cruzada. Nunca delete as Chaves da extremidade final das colunas!.

### 4.4. PEDIDOS-GERAL (Logística e Suprimentos)
*   **Papel**: Visão focada do departamento de compras/recebimentos. Basicamente os serviços com `ATRELADO="HOUSI"` listados.
*   **Bidirecionalidade**:
    *   Se na aba Pedidos o usuário editar o **Status do Fornecedor / Observação do Pedido**, ou alterar a **Data Agendada ADM**, o script fará uma travessia O(1) pelo UUID (Chave) e substituirá a informação de volta na `FASE-OBRA`.
    *   Se um serviço for Aprovado/Removido na `FASE-OBRA` ou tiver Empreendimento alterado, o script conciliará essa mudança nos Pedidos, removendo "Órfãos" ou movendo o ID.

### 4.5. OCORRÊNCIAS (Pós-Obra ou Pontual)
*   **Papel**: Chamados específicos e assistências técnicas.
*   **Comportamento**: Avalia três fases sucessivas de visitas (Vistoria 1, Revistoria 2, Revistoria 3...) e calcula uma fórmula lógica complexa no Back-End que dita o **STATUS GERAL (Aberto/Fechado/Cancelado)**. Se houver ocorrências *"NÃO CONCLUÍDAS* para um condomínio/unidade", esse número "sobe" para a coluna de Resumo lá na FASE-PRELIMINAR e INFO GERAIS.

### 4.6. FASE-ENTREGA (Check-out)
*   **Papel**: Vistoria final e revisão antes do fechamento total.
*   **Calculo de Vistorias**: Possui diversas colunas de Revisão e Vistoria parecidas com ocorrências, cujo status Geral Consolida a média do "Melhor Pior Caso". Se tem algo negado, prende o Status Geral.

### 4.7. DASHBOARD (Pendências Gerais)
*   **Papel**: Consolidação gerencial por Unidade (EMP + UNID), com base em INFORMAÇÕES GERAIS e FASE-OBRA, preparando evolução futura para painel.
*   **Estrutura de colunas**:
    *   A `EMPREENDIMENTO`
    *   B `UNID`
    *   C `DIAS DE OBRA`
    *   D `SERVIÇOS CONCLUIDOS`
    *   E `SERVIÇOS PENDENTES`
    *   F `VERBA UTILIZADA`
    *   G `PGTOS PENDENTES`
    *   H `ALERTA`
*   **Fonte dos dados**:
    *   A/B vêm de **INFORMAÇÕES GERAIS** (unidades únicas, preservando ordem).
    *   C é calculada por **DATA LOTE** da INFORMAÇÕES GERAIS (sem depender da coluna de semana da FASE-OBRA).
    *   D/E/F/G vêm de agregações da **FASE-OBRA** por unidade.
*   **Regras de cálculo**:
    *   `DIAS DE OBRA` = quantidade de dias corridos desde a DATA LOTE até hoje (datas futuras/invalidas ficam vazias).
    *   `SERVIÇOS CONCLUIDOS` = contagem de serviços com status de aprovação 100% aprovado.
    *   `SERVIÇOS PENDENTES` = contagem de serviços exceto 100% aprovado e cancelado.
    *   `VERBA UTILIZADA` = soma dos valores de parcelas cujo STATUS n PAG = `PAGO` (n de 1 a 5).
    *   `PGTOS PENDENTES` = soma dos valores de parcelas com STATUS n PAG preenchido e diferente de `PAGO`.
    *   `ALERTA` = exibido quando houver serviço pendente ou pagamento pendente na unidade.
*   **Higiene visual**: valores numéricos iguais a zero não são escritos (célula fica vazia).
*   **Execução**:
    *   Função principal: `atualizarPendenciasGeraisDashboard()` em `Dashboard.gs`.
    *   Disponível no menu Automação (`📋 Atualizar PENDÊNCIAS GERAIS (DASHBOARD)`).
    *   Incluída no acionador diário centralizado das 01:00.
    *   Incluída em `sincronizacaoManualGlobal()`.

---

## ⚡ 5. Regras de Ouro e Dicas para o Backend App Script
1.  **Manipule Dados em Lotes Arrays In-Memory (Sempre!)**: O Google Sheets é lento ao ler propriedades (`Range.getDisplayValues()`) usando loops. Tudo no código já refatorado foi colocado em Lote (`Set`, `Map`) no Javascript. Se for inserir templates novos: Crie as 10 mil arrays virtuais local e dê um único tiro `.setValues()`.
2.  **Validações Causam Erros `setValues` em lote**: Ás vezes na manipulação de colunas, se você passar um Array num .setValues() para dentro de uma coluna que contém validações (Dropdown `DataValidation`), a requisição inteira será cancelada pelo servidor. Por isso os módulos usam `range.clearDataValidations()` antes de escrever o Array denso, e no final reaplicam as Validações.
3.  **Locks Redundantes**: Scripts que interagem com tabelas longas (como Sincronizações Noturnas) exigem uso criterioso da ferramenta utilitária `executarComDocumentLock_(callback)` de `Utils.gs`.
4.  **Expansão**: Para adicionar uma nova coluna ao mapeamento dinâmico e proteger o código disso, basta:
    1. Ir na aba desejada do sheets e dar um Título Exato ao Cabeçalho.
    2. Ir em `Config.gs` -> `CONFIG.HEADERS_COLS`.
    3. Adicionar uma nova linha com a Chave (Nome Backend) Mapeada ao Valor (CABEÇALHO FRONTEND). E ajustar as constantes default de FALLBACK (para criar compatibilidade regressiva).
    4. Passar a usar a propriedade de `resolveSheetColumns_(sheet, ..)` na função específica.
5.  **Dashboards de saída**: em tabelas de consolidação (ex.: DASHBOARD), preferir saída "limpa" (não imprimir `0` quando não agrega valor visual) e derivar indicadores temporais de data-base oficial (`DATA LOTE`) para evitar divergência entre abas.

---
**Fim de Documento.**
(Conserve este documento na raiz virtual do projeto para manter coerência semântica perante o desenvolvimento de novas features.)
## Atualizações do script (FASE-OBRA reorder, DASHBOARD e CI)

- Implementada função atualizarOrdemFaseObraPorInformacoesGerais_() em scripts/Utils.gs para reordenar a aba 'FASE-OBRA' seguindo a ordem de 'INFORMAÇÕES GERAIS'. Linhas sem correspondência são movidas ao final.
- Adicionada executarAtualizarFaseObraDiaria() e criarTriggerDiariaAtualizarFaseObra_() para criação de trigger diário (03:30) — nota: trigger precisa ser ativada no editor do Apps Script.
- Rotina executarSincronizacaoGlobalMadrugada_ (scripts/Main.gs) foi atualizada para chamar a reordenação durante a sincronização noturna das 01:00.
- Preservação da coluna técnica 'CHAVE' (coluna AY) ao regravar dados via setValuesPreservandoColunaChave_.
- Ajustes em scripts/Config.gs: novos HEADERS_COLS e fallbacks para suportar resolução dinâmica de colunas sem depender de índices numéricos.
- Novo módulo `Dashboard.gs` com função `atualizarPendenciasGeraisDashboard()` para popular a aba `DASHBOARD` (Pendências Gerais).
- Inclusão da aba `DASHBOARD` em `CONFIG.SHEETS`, `CONFIG.COLUMNS.DASHBOARD` e `CONFIG.HEADERS_COLS.DASHBOARD`.
- Ajuste do vínculo Empreendimento -> Unidade para também aplicar dropdown em `OCORRÊNCIAS` (além de `INFORMAÇÕES GERAIS`).
- Menu `⚙️ Automacao ECQUA` atualizado com `📋 Atualizar PENDÊNCIAS GERAIS (DASHBOARD)`.
- `sincronizacaoManualGlobal()` passou a executar `atualizarPendenciasGeraisDashboard()`.
- Rotina central diária `executarRotinaDiariaCentralizada_()` passou a executar `atualizarPendenciasGeraisDashboard()`.
- Coluna C do DASHBOARD renomeada para `DIAS DE OBRA`.
- Regra de C no DASHBOARD alterada para cálculo de dias corridos desde a `DATA LOTE` da unidade.
- Regra de saída no DASHBOARD: não imprimir valores `0` nas colunas numéricas de consolidação.
- CI: Agent Guard workflow (.github/workflows/agent-guard.yml) agora fornece check run 'agent-checks'. Proteção de branch deve exigir esse contexto para desobstruir merges.


Detalhes e passo a passo para ativar o **Acionador Diário Centralizado (01:00)**

- Função que cria o gatilho: `criarTriggerDiarioCentralizado01h()`
- Função executada pelo acionador: `executarRotinaDiariaCentralizada_()`
- Horário padrão: 01:00 (fuso do projeto). O acionador é `time-driven` e roda diariamente às 01:00.

- O que o acionador faz (resumo operacional):
    - Executa em sequência, com lock e tolerância a falhas, as rotinas importantes de sincronização e relatório:
        1. `executarSincronizacaoGlobalMadrugada_()` — sincronização completa (push/pull entre FASE-OBRA, PEDIDOS, PRELIMINAR e INFO_GERAIS).
        2. `autorunSincronizarStatusPagamentos()` — atualiza STATUS em `PAGAMENTOS` a partir de `FASE-OBRA` (marcas PAGO).
        3. `autorunGerarRelatorio()` — gera/atualiza o relatório de `PAGAMENTOS` (chama `gerarRelatorioPagamentos`).
        4. `atualizarPendenciasGeraisDashboard()` — atualiza a consolidação da aba `DASHBOARD`.
    - Cada item é protegido por `try/catch`; se uma subtarefa falhar, a próxima ainda será executada.
    - A execução é envolvida em `executarComDocumentLock_` para reduzir condições de corrida.

- Comportamento de limpeza de triggers:
    - Ao criar o acionador central, a função `criarTriggerDiarioCentralizado01h()` remove automaticamente triggers antigos relacionados para evitar duplicação. Handlers removidos incluem (lista não-exaustiva):
        - `executarSincronizacaoFinalDoDia`
        - `executarSincronizacaoGlobalMadrugada_`
        - `executarAtualizarFaseObraDiaria`
        - `autorunGerarRelatorio`
        - `autorunSincronizarStatusPagamentos`
        - `sincronizarTodosPedidosHousi`

- Garantias e observações operacionais:
    - Escritas em lote preservam a coluna técnica `CHAVE` usando `setValuesPreservandoColunaChave_`.
    - O lock reduz concorrência, porém recomenda-se não agendar outras tarefas pesadas no mesmo minuto (01:00) para evitar disputas de recursos.
    - Limite de tempo: se a rotina exceder o tempo máximo de execução do Apps Script (situação rara, dependendo do volume de dados), migre subtarefas pesadas para jobs menores ou escalone horários.

- Como ativar manualmente (passo a passo):
    1. Abra a planilha no Google Sheets → Extensões → Apps Script.
    2. No editor do Apps Script, selecione a função `criarTriggerDiarioCentralizado01h` e clique em Executar (Run).
    3. Conceda as permissões OAuth solicitadas (se necessário).
    4. No editor, abra o painel "Triggers" e confirme que existe um trigger `Time-driven` configurado para rodar diariamente às 01:00, com handler `executarRotinaDiariaCentralizada_`.

- Testes pós-ativação recomendados:
    - Execute `executarRotinaDiariaCentralizada_()` manualmente e verifique os logs para cada subtarefa.
    - Chame `listarAcionadoresProjeto()` para confirmar que o acionador central está ativo e que triggers antigos foram removidos.
    - Valide que relatórios, sincronizações e atualizações de status ocorreram conforme esperado e que a coluna `CHAVE` foi preservada.

Automatização (opcional)
- Para criar triggers programaticamente (CI), usar a Apps Script API com credenciais de serviço ou usar `clasp` com uma conta que tenha permissões. Esse fluxo requer deploy e permissões administrativas e não é executado automaticamente ao fazer merge.

Ações necessárias após merge
- Com a branch mesclada, execute `criarTriggerDiarioCentralizado01h()` no editor do Apps Script com a conta de deploy para ativar o acionador central.
- Testar em uma cópia de staging: fazer backup, executar `executarRotinaDiariaCentralizada_()` manualmente e verificar os logs; confirmar que `PEDIDOS-GERAL` e `PAGAMENTOS` foram atualizados conforme esperado.
- Conferir logs e validar que UUIDs (coluna AY) foram preservadas; verificar se triggers antigos foram removidos (use `listarAcionadoresProjeto()`).

Notas sobre CI/Proteção de branch
- O check run do Agent Guard é chamado "agent-checks". Garanta que a regra de proteção do branch principal (main) exige esse contexto para liberar merges.
- Se houver mismatch entre o nome esperado pela proteção e o check run real, ajustar na UI do GitHub: Settings → Branches → Edit rule → Required status checks.


