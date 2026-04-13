# AGENT-LOGS

Este arquivo serve como um canal de comunicação simples entre agentes (bots) que modificam o repositório e este assistente humano/CLI. O objetivo é registrar, em append, ações relevantes realizadas por agentes automáticos para que possam ser monitoradas aqui.

Formato recomendado — duas linhas por evento (legível + JSON estruturado):

- 2026-04-08T18:45:00Z | AgentName | ACTION | resumo curto | commit: <sha> | pr: <#>
- AGENT-LOG-JSON: {"timestamp":"2026-04-08T18:45:00Z","agent":"AgentName","action":"ACTION","summary":"resumo curto","commit":"<sha>","pr":<#>,"repo":"owner/repo","run_id":null,"workflow":"agent-log","files":[],"tags":["auto"],"severity":"info","notes":"opcional"}

Observações:
- A primeira linha é para leitura humana rápida; a segunda linha (prefixada por `AGENT-LOG-JSON:`) é JSON válido que permite parsing por outros agentes.
- Campos mínimos recomendados no JSON: timestamp, agent, action, summary, commit. Campos opcionais úteis: pr, repo, run_id, workflow, files, tags, severity, notes.

Exemplos:
- 2026-04-08T18:45:00Z | Copilot-Local | COMMIT_PUSH | adicionou menu e trigger | commit: 996c8f0 | pr: 11
- AGENT-LOG-JSON: {"timestamp":"2026-04-08T18:45:00Z","agent":"Copilot-Local","action":"COMMIT_PUSH","summary":"adicionou menu e trigger","commit":"996c8f0","pr":11,"repo":"micaiasviola/PLANILHA-OBRAS-GERAL","workflow":"agent-log","files":["scripts/CompatibilityWrappers.gs"],"tags":["menu","fix"],"severity":"info"}

Instruções para append seguro (BASH):
- Criar conteúdo em variável e gravar em arquivo temporário para evitar writes parciais:
  TIMESTAMP=$(date -u --iso-8601=seconds)
  HUMAN_LINE="- ${TIMESTAMP} | ${AGENT_NAME:-AppAgent} | ${ACTION:-ACTION} | ${SUMMARY:-descrição} | commit: ${GIT_COMMIT:-} | pr: ${PR_NUMBER:-}"
  JSON_LINE="AGENT-LOG-JSON: {\"timestamp\":\"${TIMESTAMP}\",\"agent\":\"${AGENT_NAME:-AppAgent}\",\"action\":\"${ACTION:-ACTION}\",\"summary\":\"${SUMMARY:-descrição}\",\"commit\":\"${GIT_COMMIT:-}\",\"pr\":${PR_NUMBER:-null}}"
  printf "%s\n%s\n" "$HUMAN_LINE" "$JSON_LINE" >> AGENT-LOGS.md
  git add AGENT-LOGS.md && git commit -m "agent-log: ${TIMESTAMP} ${ACTION}" && git push origin $(git rev-parse --abbrev-ref HEAD)

Instruções para append seguro (PowerShell):
  $ts = (Get-Date).ToUniversalTime().ToString("s") + "Z"
  $human = "- $ts | $env:AGENT_NAME | $env:ACTION | $env:SUMMARY | commit: $env:GIT_COMMIT | pr: $env:PR_NUMBER"
  $json = "AGENT-LOG-JSON: {\"timestamp\":\"$ts\",\"agent\":\"$env:AGENT_NAME\",\"action\":\"$env:ACTION\",\"summary\":\"$env:SUMMARY\",\"commit\":\"$env:GIT_COMMIT\",\"pr\":$($env:PR_NUMBER -ne '' ? $env:PR_NUMBER : 'null')}"
  "$human" | Out-File -FilePath .\AGENT-LOGS.md -Encoding utf8 -Append
  "$json"  | Out-File -FilePath .\AGENT-LOGS.md -Encoding utf8 -Append
  git add .\AGENT-LOGS.md; git commit -m "agent-log: $ts $env:ACTION"; git push origin (git rev-parse --abbrev-ref HEAD)

Recomendações de comportamento para o agente que escreve:
- Sempre anexar (append) — nunca sobrescrever o arquivo.
- Incluir commit SHA (git rev-parse --short HEAD) e, se aplicável, run_id / workflow.
- Evitar churn: agrupar eventos quando apropriado.
- Validar que o JSON é bem formado antes de gravar (ex.: jq -c . >/dev/null).
- Mensagens de commit padronizadas: começar com `agent-log:`.

Alertas e monitoramento (para este assistente):
- Este assistente fará parsing das linhas AGENT-LOG-JSON para acionar análises automatizadas.
- Workflows devem atualizar AGENT-LOGS.md ao final de operações críticas; preferir ACTIONS_PAT para commits diretos quando seguro.

---


---

Obs: este arquivo foi criado automaticamente para suportar monitoramento entre agentes. Mantê-lo curto e legível.
- 2026-04-08T19:10:08.091Z | Copilot-Local | REVALIDATE_POST_SORT | Revalidação de subcategorias após ordenação | commit: 0577002 
- 2026-04-08T20:12:07.017Z | Copilot-Local | REMOVE_LIMPO_OCORRENCIAS | Remove 'LIMPO' em ocorrências automatizadas | commit: c9cee9d 
- 2026-04-10T14:25:28Z | agent-log-retrofit | PR-MERGED | feat(sheet): usar RESP_ADM resolvido por cabeçalho | commit: 9fdd39b | pr: #13
- 2026-04-10T14:57:49Z | agent-log-retrofit | PR-MERGED | feat(sync): resolver largura fakeEvent (SyncLogic) | commit: 3bb73bc | pr: #14
- 2026-04-10T15:37:32Z | agent-log-retrofit | PR-MERGED | refactor(sync): usar larguras calculadas via maxCols (SyncLogic) | commit: 55c78a0 | pr: #15
- 2026-04-10T15:39:29Z | agent-log-retrofit | PR-MERGED | chore(agent-logs): add workflow and retrofit logs for PRs #13,#14 | commit: 0bcd112 | pr: #16
- 2026-04-10T15:46:43Z | agent-log-retrofit | PR-MERGED | fix(workflow): corrigir sintaxe agent-log-on-merge.yml | commit: eba857b | pr: #17
- 2026-04-10T15:52:05Z | agent-log-retrofit | PR-MERGED | refactor(obra): usar maxCols para evitar indices fixos | commit: 148229e | pr: #18
2026-04-10T18:01:31Z | PR-MERGER(micaiasviola) | PR-MERGED | test(agent-logs): trigger append workflow | commit: 0b59060f98c7a4a9b37ca21d2ab66fdd6309bb20 | pr: #31
2026-04-10T18:01:31Z | PR-MERGER(micaiasviola) | PR-MERGED | test(agent-logs): trigger append workflow | commit: 0b59060f98c7a4a9b37ca21d2ab66fdd6309bb20 | pr: #31
2026-04-10T18:04:15Z | PUSH_TO_MAIN | Merge pull request #32 from micaiasviola/agent-log/entry-24256832424-20260410180132 | commit: 91446d685f393615708a003a16c9bcbbfa06389f | pusher: micaiasviola
2026-04-10T18:04:16Z | PR-MERGER(micaiasviola) | PR-MERGED | Append AGENT log entry | commit: 91446d685f393615708a003a16c9bcbbfa06389f | pr: #32
2026-04-10T18:07:48Z | PUSH_TO_MAIN | ci/test: trigger agent-log workflows | commit: 359d7236227ebfe03ac1b4723d3b91a5e7147c8e | pusher: micaiasviola
2026-04-10T18:47:37Z | PUSH_TO_MAIN | feat(wrapper): adicionar compatibilidade para funções de menu legadas | commit: b477780c1e96edea6ceb030058a9468fa09866a5 | pusher: micaiasviola
2026-04-10T21:30:42Z | PUSH_TO_MAIN | chore(agent-logs): structured JSON log lines for agent logs | commit: f471cd2a3472e9642cf05ede8e24630733d971af | pusher: micaiasviola
2026-04-10T22:29:07Z | PR-MERGER(micaiasviola) | PR-MERGED | feat(payments): módulo PAGAMENTOS + sincronização inicial | commit: e3e9864cd003ffa2b1a0d9eb08e89c9d464232a2 | pr: #33
2026-04-10T22:29:09Z | PUSH_TO_MAIN | Merge pull request #33 from micaiasviola/feature/payments | commit: e3e9864cd003ffa2b1a0d9eb08e89c9d464232a2 | pusher: micaiasviola
2026-04-13T14:02:00Z | PUSH_TO_MAIN | feat(payments): add importer for manual sheet and payments menu | commit: 486b1b3395dbe243ea91fa54898997518816624b | pusher: micaiasviola
