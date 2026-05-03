# Workflow

## Objetivo

Produzir planejamento, texto acadêmico e revisão de evidências a partir de uma base local Scopus/PDF, sem depender de discovery externo ou canonicalização obrigatória em Zotero.

## Modo A — Planejamento de artigo

Use quando o usuário começa pelo experimento, código, notebook ou resultado e precisa descobrir como transformar isso em artigo.

Etapas:

1. Ler a descrição do experimento/código/resultados.
2. Identificar contribuição potencial.
3. Propor problema de pesquisa.
4. Propor objetivo geral e objetivos específicos.
5. Propor estrutura do artigo.
6. Mapear referenciais teóricos candidatos.
7. Mapear metodologia.
8. Mapear resultados esperados.
9. Mapear lacunas que exigem RAG/fonte.
10. Gerar `article_plan.md` e `article_plan.json`.

## Modo B — Escrita com RAG

Etapas:

1. Decompor tarefa em evidence needs.
2. Gerar queries em inglês.
3. Buscar em `<collection>`.
4. Buscar em `<collection>_pdf` quando necessário.
5. Deduplicar.
6. Classificar evidência.
7. Escrever draft.
8. Gerar `draft.md` e `evidence_map.json`.
9. Revisar evidência.
10. Gerar `evidence_review.json`.
11. Aplicar melhorias.
12. Gerar `final.md`.
13. Rodar postcheck.

## Modo C — Auditoria/Revisão

Etapas:

1. Ler texto fornecido.
2. Dividir em claims.
3. Verificar citações existentes.
4. Buscar fontes adicionais apenas se necessário.
5. Classificar evidência.
6. Reposicionar citações.
7. Suavizar/remover claims fracos.
8. Gerar versão final e postcheck.

## Artefatos mínimos

Planejamento:

- `article_plan.md`
- `article_plan.json`
- `theoretical_framework_queries.md`
- `methodology_blueprint.md`
- `gaps_and_next_steps.md`

Escrita:

- `draft.md`
- `evidence_map.json`
- `evidence_review.json`
- `final.md`
- `postcheck.md`
- `search_log_compact.json`

## Gates

### Gate 1 — Escopo

Nunca processar artigo inteiro de uma vez se o usuário pediu algo muito amplo. Quebrar em:

- plano geral;
- seção;
- subseção;
- grupo de parágrafos.

### Gate 2 — Evidência

Antes de citar, classificar fonte como:

- `core`
- `applied`
- `adjacent`
- `reject`

### Gate 3 — Citação

Toda citação precisa sustentar o claim próximo.

### Gate 4 — Referência

Toda referência deve estar citada e toda citação deve ter referência.

### Gate 5 — Prontidão

Se houver fonte fraca usada em claim central, marcar como pendência ou fazer segunda rodada de busca.
