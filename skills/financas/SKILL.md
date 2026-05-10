---
name: financas
description: Consumir e automatizar a API local de finanças em D:\Financas. Use para toda mensagem ou áudio que solicite "atualizar finanças" ou "usar a skill finanças para ..."; use também quando o agente precisar consultar Swagger/OpenAPI, listar ou criar contas, inserir receitas/despesas/PIX, extrair comprovante PIX, importar extratos/faturas, categorizar, mesclar categorias, reconciliar lançamentos, resolver deduplicação ou consultar dashboards e relatórios financeiros pela API Django/DRF local.
---

# Finanças

Use esta skill para interagir com a API local de controle financeiro. Quando o pedido vier por mensagem ou áudio com "atualizar finanças", primeiro descubra o tipo de atualização: importar arquivos, lançar despesa/receita/PIX, categorizar, reconciliar ou consultar relatório.

## Base

- API local preferida quando o agente estiver dentro de container Docker: `http://host.docker.internal:8000/api`
- API local alternativa quando o agente estiver rodando direto no host: `http://localhost:8000/api`
- API local alternativa no host: `http://127.0.0.1:8000/api`
- Swagger: `http://host.docker.internal:8000/api/docs/`
- OpenAPI JSON vivo: `http://host.docker.internal:8000/api/schema/?format=json`
- Contrato fonte interno: `openapi.yaml` no diretório desta skill
- Referência operacional: `references/api.md`

Não há autenticação. Envie JSON com `Content-Type: application/json`, exceto endpoints de upload, que usam `multipart/form-data`.

## Regra Contract-First

1. Se o servidor estiver acessível, leia `GET /api/schema/?format=json` antes de escolher endpoint, parâmetros ou formato de payload.
2. Use `openapi.yaml` no diretório desta skill como contrato offline quando o servidor estiver fora do ar.
3. Use `references/api.md` como fallback e guia resumido quando precisar de fluxo operacional.
4. Trate todos os paths do schema como relativos ao servidor `/api`; exemplo: path `/accounts/` vira `http://host.docker.internal:8000/api/accounts/` quando estiver dentro de container Docker.
5. Se houver divergência entre esta skill, o contrato offline e o schema vivo, siga o schema vivo e registre a divergência no resultado.

## Fluxos

1. Para descobrir IDs, comece por `GET /accounts/`, `GET /categories/` e, se necessário, `GET /transactions/`.
2. Para comprovantes PIX, use primeiro `POST /transactions/manual-pix/extract/` com o PDF. Não invente `descricao`; use o campo `lancamento` informado pelo usuário ao criar o PIX.
3. Para inserir PIX revisado pelo usuário, use `POST /transactions/manual-pix/`; use `manual-pix-upload` somente quando o objetivo explícito for importar o comprovante como documento de origem.
4. Para despesa simples, use `POST /transactions/manual-expense/`; para receita, use `POST /transactions/manual-income/` as opções de meio são apenas `cartao|pix` normalize a informação antes de inserir o registro.
5. Para importação de arquivos, use `POST /imports/bank-statements/` para extratos com `conta_id` e `POST /imports/card-statements/` para faturas/cartões. Ambos aceitam múltiplos arquivos.
6. Para conciliação, use `GET /reconciliation/uncategorized/`, depois `PATCH /reconciliation/transactions/{id}/`.
7. Para deduplicação, use `GET /reconciliation/deduplication/` e resolva com `POST /reconciliation/deduplication/{id}/resolve/`.
8. Para relatórios, use `GET /reports/dashboard/?month=YYYY-MM&since=YYYY-MM-DD` e `GET /reports/monthly/?month=YYYY-MM`.

## Cuidados

- Valores monetários vão como string decimal com duas casas, por exemplo `"20.00"`.
- Datas usam `YYYY-MM-DD`; meses usam `YYYY-MM`; horas usam `HH:MM:SS`.
- `banco` aceita apenas `itau`, `nubank`, `bradesco`.
- `titular` aceita apenas `danielle`, `alexandre`.
- `meio_pagamento` em despesas aceita `pix` ou `cartao`; em receitas aceita `pix`, `cartao` ou `outro`.
- Uploads são idempotentes por hash de arquivo; resposta 200 pode indicar que nada novo foi inserido.
- Status `200` em lançamentos manuais pode significar duplicado; confira o campo `duplicado`.
- Em uploads, campo `arquivos` aceita múltiplos arquivos PDF, CSV ou OFX.
- Em extratos de conta, `conta_id` é obrigatório; em faturas/cartões, `vencimento` é opcional.
- Preserve masking e não exponha agência, conta, documentos pessoais ou chaves PIX completas em mensagens ao usuário.

## Referência Completa

Leia `references/api.md` quando precisar de parâmetros obrigatórios, valores válidos, exemplos de payload, formatos de resposta ou códigos HTTP por endpoint.
