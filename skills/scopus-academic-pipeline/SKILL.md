---
name: scopus-academic-pipeline
description: "Pipeline multiagente para planejar, escrever, revisar e consolidar textos acadêmicos com RAG local em coleções Scopus/PDF. Use quando o usuário invocar `$scopus-academic-pipeline`, quiser estruturar artigo a partir de código/experimento, planejar referencial teórico/metodologia, gerar texto acadêmico com citações ABNT, auditar evidências e produzir versão final em Markdown."
---

# Scopus Academic Pipeline

Use esta skill para conduzir um workflow acadêmico baseado em evidências recuperadas do RAG local.

Esta skill substitui o fluxo antigo pesado de descoberta/canonicalização externa quando o usuário já possui coleções locais Scopus/PDF confiáveis.

## Casos de uso

Use esta skill quando o usuário pedir para:

- planejar estrutura de artigo;
- transformar experimento/código/notebook em artigo;
- definir seções, referencial teórico, metodologia, resultados e discussão;
- escrever texto acadêmico com base em RAG local;
- revisar aderência entre citações e claims;
- gerar versão final em Markdown;
- auditar fontes recuperadas de coleções Scopus/PDF.

## Entradas mínimas

Para escrita com RAG:

- `collection`: nome base da coleção local;
- `writing_task`: tarefa de escrita;
- `topic_or_outline`: tema, hipótese, objetivo, experimento ou esboço.

Para planejamento de artigo:

- descrição do experimento, código, notebook, resultados ou ideia;
- objetivo do artigo ou problema de pesquisa, se houver;
- área-alvo ou evento/periódico, se houver.

Não há coleção padrão.

Se a coleção estiver ausente e a tarefa exigir RAG, execute:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections
```

Liste as coleções disponíveis e peça ao usuário escolher.

## Modelo de coleções

Cada base costuma ter duas coleções:

- `<collection>`: metadados Scopus, abstracts, keywords e registros bibliográficos.
- `<collection>_pdf`: chunks dos PDFs/texto integral.

Use papéis distintos:

- `<collection>` identifica estudos candidatos e metadados para referência.
- `<collection>_pdf` valida e aprofunda claims que exigem suporte textual mais forte.

## Princípios

- Queries de busca sempre em inglês.
- Saída textual em português do Brasil, salvo pedido contrário.
- Não usar web como evidência.
- Não usar memória/modelo como evidência científica.
- Não inventar referências.
- Não citar fonte apenas por proximidade temática genérica.
- Priorizar fonte `core` para claims conceituais.
- Usar fonte `applied` apenas para exemplos/aplicações/domínios específicos.
- Rejeitar fonte `adjacent` ou `reject`.
- Manter texto cauteloso quando a evidência for parcial.

## Agentes

O pipeline usa cinco papéis:

1. `article_planner`
2. `rag_writer`
3. `evidence_reviewer`
4. `rag_reviser`
5. `final_postcheck`

Leia os arquivos em `agents/` antes de executar cada papel.

## Modos de execução

### Modo A — Planejamento de artigo

Use quando o usuário começa pelo experimento/código/notebook e precisa transformar isso em estrutura de artigo.

Fluxo:

1. `article_planner`
2. opcionalmente `rag_writer` para criar seções iniciais
3. opcionalmente `evidence_reviewer`
4. opcionalmente `rag_reviser`
5. `final_postcheck`

Artefatos:

```text
.ai/handoff/scopus_academic_pipeline_<tag>/
├── article_plan.md
├── article_plan.json
├── theoretical_framework_queries.md
├── methodology_blueprint.md
└── gaps_and_next_steps.md
```

### Modo B — Escrita com RAG

Use quando o usuário já quer gerar texto.

Fluxo:

1. `rag_writer`
2. `evidence_reviewer`
3. `rag_reviser`
4. `final_postcheck`

Artefatos:

```text
.ai/handoff/scopus_academic_pipeline_<tag>/
├── draft.md
├── evidence_map.json
├── evidence_review.json
├── final.md
├── postcheck.md
└── search_log_compact.json
```

### Modo C — Revisão/auditoria de texto existente

Use quando o usuário fornece texto já escrito.

Fluxo:

1. `evidence_reviewer`
2. `rag_reviser`
3. `final_postcheck`

Artefatos:

```text
.ai/handoff/scopus_academic_pipeline_<tag>/
├── input_text.md
├── evidence_review.json
├── final.md
└── postcheck.md
```

## Estratégia de recuperação

Default:

- 3 a 5 queries em inglês por tarefa/evidence need;
- `--top-k 5` em `<collection>`;
- buscar `<collection>` primeiro;
- buscar `<collection>_pdf` quando a evidência for fraca, genérica, insuficiente ou quando o claim exigir validação textual;
- `--top-k 5` em `<collection>_pdf`;
- deduplicar por DOI, depois título normalizado + ano + primeiro autor;
- armazenar somente notas compactas de evidência.

Para claims centrais ou evidência fraca:

- até 6 queries;
- `--top-k 8`;
- incluir uma query literal;
- incluir uma query com termos específicos do domínio.

Modo diagnóstico:

- até 8 queries;
- `--top-k 10`;
- buscar em ambas as coleções;
- gerar matriz de evidência.

## Contrato de busca

Use apenas:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections

python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py search --collection <collection> --query "<english query>" --original-query "<intent or claim>" --top-k 5
```

## Evidence classification gate

Classifique cada fonte candidata antes de citar:

- `core`: sustenta diretamente conceito, método, mecanismo ou claim.
- `applied`: aplica conceito relevante em domínio específico.
- `adjacent`: tangencia o tema, mas não sustenta o claim.
- `reject`: não sustenta o claim.

Regras:

- Claims conceituais precisam preferencialmente de fontes `core`.
- Fontes `applied` servem para exemplos, aplicações e domínio específico.
- Não usar `adjacent` para sustentar conceito geral.
- Nunca citar `reject`.
- Se só houver `adjacent`/`reject`, refazer busca com queries mais literais e específicas.
- Se ainda faltar evidência, suavizar ou declarar lacuna.

## Saída padrão

Para escrita final, retornar somente:

1. texto final com citações ABNT autor-data;
2. exatamente duas linhas em branco;
3. título `Referências bibliográficas`;
4. referências ABNT-like apenas das obras citadas.

## Compatibilidade com fluxo antigo

Papéis removidos ou rebaixados:

- `doi_hunter`: não é etapa padrão. Usar apenas em fluxo externo de descoberta bibliográfica.
- `citation_validator`: substituído por `evidence_reviewer` para coleções locais Scopus/PDF.
- `source_agents`: não são etapa padrão se a coleção local já contém fontes válidas.
- `zotero_ingest_router`: opcional, não obrigatório.
- `zotero_semantic_refresher`: opcional, não obrigatório.
- `dedup_router_agent`: substituído por deduplicação leve por DOI/título.
- `consolidator`: substituído por `final_postcheck` + artefatos compactos.

## Antes de executar

Leia:

- `references/workflow.md`
- `agents/article_planner.md` quando houver planejamento de artigo;
- `agents/rag_writer.md` quando houver escrita;
- `agents/evidence_reviewer.md` quando houver auditoria de evidência;
- `agents/rag_reviser.md` quando houver revisão final;
- `agents/final_postcheck.md` antes de entregar resultado final.
