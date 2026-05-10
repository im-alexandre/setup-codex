# Referência da API Local de Finanças

Base URL preferida: `http://localhost:8000/api`

Base URL alternativa: `http://127.0.0.1:8000/api`

Sem autenticação. Use JSON nos endpoints comuns e `multipart/form-data` nos uploads. O Swagger completo fica em `/api/docs/` e o contrato OpenAPI vivo em `/api/schema/?format=json`.

## Lógica de Uso para OpenClaw

1. Ao receber "atualizar finanças" por mensagem ou áudio, classifique a intenção em uma destas ações: importar arquivo, lançar despesa, lançar receita, lançar PIX, extrair comprovante PIX, categorizar, reconciliar, deduplicar ou consultar relatório.
2. Antes de montar payloads, leia `GET http://localhost:8000/api/schema/?format=json` se o servidor estiver ativo. O schema vivo é a fonte de verdade.
3. Resolva IDs necessários antes de escrever: use `/accounts/` para contas e `/categories/` para categorias.
4. Para arquivos, use multipart e preserve o nome `arquivos` para uploads múltiplos ou `arquivo` para comprovante PIX único.
5. Depois de inserções/importações, confira `duplicado`, `lancamentos_criados`, `pendencias` e listas de conciliação/deduplicação.
6. Se a fala do usuário não trouxer categoria, conta, descrição ou arquivo suficiente, consulte listas disponíveis e peça apenas o dado faltante.

## Convenções

- Dinheiro: string decimal com duas casas, por exemplo `"20.00"` ou `"-20.00"`.
- Data: `YYYY-MM-DD`.
- Mês: `YYYY-MM`.
- Hora: `HH:MM:SS`.
- Date-time: ISO 8601, por exemplo `2026-05-07T19:23:03-03:00`.
- Banco: `itau`, `nubank`, `bradesco`.
- Titular: `danielle`, `alexandre`.
- Meio de pagamento: despesas aceitam `pix` ou `cartao`; receitas aceitam `pix`, `cartao` ou `outro`.
- IDs são inteiros positivos.

## Sistema

### GET `/schema/`
Retorna o contrato OpenAPI completo em JSON.

### GET `/docs/`
Retorna o Swagger UI.

## Contas

### GET `/accounts/`
Lista contas bancárias ativas.

Resposta: array de:
```json
{
  "id": 1,
  "nome": "Itaú Alexandre",
  "agencia": "3830",
  "numero_conta": "14196-0",
  "banco": "itau",
  "titular": "alexandre",
  "ativa": true
}
```

### POST `/accounts/`
Cria conta bancária.

Obrigatórios: `nome`, `banco`, `titular`.
Opcionais: `agencia`, `numero_conta`.

```json
{
  "nome": "Nubank Danielle",
  "agencia": "",
  "numero_conta": "",
  "banco": "nubank",
  "titular": "danielle"
}
```

Status: `201` criado, `400` inválido.

### GET/PATCH/DELETE `/accounts/{id}/`
Consulta, altera parcialmente ou desativa conta. `DELETE` é soft delete e retorna `204`.

## Documentos

### GET `/documents/`
Lista documentos de origem importados. Campos: `id`, `tipo`, `caminho`, `importado_em`, `conta`, `vencimento`, `metadados`.

### GET `/documents/{id}/`
Detalha documento de origem.

## Lançamentos

### GET `/transactions/`
Lista lançamentos por data decrescente.

Resposta:
```json
{
  "id": 440,
  "data": "2026-03-27",
  "hora_aproximada": null,
  "descricao": "PIX QRS GIANT S BAR27 03",
  "valor": "-20.00",
  "categoria": 5,
  "meio_pagamento": "",
  "documento_origem": 13,
  "pendente_revisao": false,
  "metadados": {}
}
```

### POST `/transactions/manual-expense/`
Cria despesa manual simples. Retorna `201` quando cria e `200` quando já existia.

Obrigatórios: `valor`, `data`, `meio_pagamento`, `categoria_id`, `descricao`.
Opcionais: `hora_aproximada`, `id_externo`.

```json
{
  "valor": "20.00",
  "data": "2026-03-27",
  "hora_aproximada": "20:18:57",
  "meio_pagamento": "pix",
  "categoria_id": 5,
  "descricao": "PIX QRS GIANT S BAR",
  "id_externo": "opcional"
}
```

### POST `/transactions/manual-income/`
Cria receita manual. Mesmo contrato de despesa, mas `meio_pagamento` aceita `pix`, `cartao` ou `outro`.

```json
{
  "valor": "1500.00",
  "data": "2026-05-09",
  "hora_aproximada": "09:30:00",
  "meio_pagamento": "pix",
  "categoria_id": 12,
  "descricao": "Receita de projeto"
}
```

### POST `/transactions/manual-pix/`
Cria PIX manual. Não infere descrição automaticamente.

Obrigatórios: `valor`, `data_pix`, `hora_pix`, `lancamento`, `destinatario`, `categoria_id`.
Opcionais: `conta_id`, `emitido_em`, `id_externo`.

```json
{
  "valor": "20.00",
  "data_pix": "2026-03-27",
  "hora_pix": "20:18:57",
  "lancamento": "PIX QRS GIANT S BAR",
  "destinatario": "giant s bar",
  "categoria_id": 5,
  "conta_id": 1,
  "emitido_em": "2026-05-07T19:23:03-03:00"
}
```

Resposta manual:
```json
{
  "id": 501,
  "data": "2026-03-27",
  "hora_aproximada": "20:18:57",
  "descricao": "PIX QRS GIANT S BAR",
  "valor": "-20.00",
  "categoria": 5,
  "meio_pagamento": "pix",
  "documento_origem": 22,
  "duplicado": false
}
```

### POST `/transactions/manual-pix/extract/`
Alias: `/transactions/manual-pix-preview/`.

Multipart obrigatório: `arquivo` com PDF do comprovante.

Resposta:
```json
{
  "valor": "-20.00",
  "data_pix": "2026-03-27",
  "hora_pix": "20:18:57",
  "banco": "itau",
  "titular": "alexandre",
  "destinatario": "giant s bar",
  "descricao": "",
  "campos_ausentes": [],
  "conta_id": 1,
  "conta_nome": "Itaú Alexandre"
}
```

`campos_ausentes` pode conter `valor`, `data_pix`, `hora_pix`. `descricao` deve permanecer vazia.

Fluxo recomendado para áudio/mensagem com comprovante PIX:

1. Envie o PDF para `/transactions/manual-pix/extract/`.
2. Mostre os campos extraídos e destaque `campos_ausentes`.
3. Peça ou use uma descrição explícita do usuário para `lancamento`.
4. Use `POST /transactions/manual-pix/` para criar o lançamento revisado.
5. Confira `duplicado`; se `true`, informe que o registro não foi recriado.

### POST `/transactions/manual-pix-upload/`
Multipart: `arquivo` obrigatório, `categoria_id` opcional. Importa o comprovante como documento de origem e é idempotente por hash.

## Categorias

### GET `/categories/`
Lista categorias ativas.

### POST `/categories/`
Cria categoria. Obrigatório: `nome`. Opcional: `descricao`.

### PATCH `/categories/{id}/`
Edita `nome`, `descricao` e/ou `ativa`.

### POST `/categories/{id}/merge/`
Mescla categoria da URL em outra categoria ativa.

```json
{
  "categoria_destino_id": 8
}
```

Resposta:
```json
{
  "categoria_origem_id": 5,
  "categoria_destino_id": 8,
  "lancamentos_atualizados": 12,
  "sugestoes_atualizadas": 3
}
```

### PUT `/transactions/{id}/category/`
Define categoria de lançamento.

```json
{
  "categoria_id": 5,
  "aplicar_no_periodo_atual": true
}
```

### GET `/suggestions/`
Lista sugestões de categoria.

### POST `/suggestions/{id}/decision/`
Decide sugestão. Body: `{ "aceita": true }`. Retorna `204`.

## Importação

### POST `/imports/run`
Sem body, importa fontes configuradas. Com body, importa caminhos locais explícitos.

```json
{
  "files": [
    { "source_type": "ofx", "path": "D:\\Financas\\dados_contas\\extratos\\arquivo.ofx" }
  ]
}
```

`source_type`: `statement`, `ofx`, `card`, `pix`.

### GET `/imports/pending`
Retorna `{ "pendencias": 0 }`.

### POST `/imports/upload/`
Upload genérico multipart.

Campos:
- `tipo`: `extrato` ou `cartao`.
- `conta_id`: obrigatório quando `tipo=extrato`.
- `vencimento`: opcional para cartão, `YYYY-MM-DD`.
- `arquivos`: um ou mais arquivos PDF, CSV ou OFX.

Resposta:
```json
{
  "documentos_processados": 2,
  "lancamentos_criados": 10,
  "pendencias": 1
}
```

### POST `/imports/bank-statements/`
Alias para extratos. Multipart: `conta_id` obrigatório, `arquivos` múltiplos.

### POST `/imports/card-statements/`
Alias para faturas/cartões. Multipart: `arquivos` múltiplos, `vencimento` opcional.

Após qualquer importação, consulte:

- `/imports/pending` para contagem geral de pendências.
- `/reconciliation/uncategorized/` para lançamentos sem categoria.
- `/reconciliation/deduplication/` para registros a deduplicar.

## Conciliação e Deduplicação

### GET `/reconciliation/uncategorized/`
Lista lançamentos sem categoria, excluindo registros ainda em `a_deduplicar`.

### PATCH `/reconciliation/transactions/{id}/`
Edita lançamento durante conciliação.

```json
{
  "descricao": "Padaria Central",
  "categoria_id": 7
}
```

### GET `/reconciliation/deduplication/`
Lista pares para comparação humana.

```json
{
  "id": 3,
  "motivo": "Data, valor e destinatario coincidem, mas a hora ou a chave forte diferem.",
  "existente": { "id": 440 },
  "novo": { "id": 501 }
}
```

### POST `/reconciliation/deduplication/{id}/resolve/`
Resolve pelo formato do app:

```json
{ "action": "same" }
```

Valores: `same` remove o novo registro como duplicado; `different` mantém o novo registro sem categoria.

### GET `/deduplication/`
Lista pendências no formato backend, com `lancamento_existente` e `lancamento_novo`.

### POST `/deduplication/{id}/decision/`
Resolve pelo formato backend:

```json
{ "decisao": "mesmo_registro" }
```

Valores: `mesmo_registro`, `registro_diferente`.

## Relatórios

### GET `/reports/dashboard/`
Query params:
- `month`: opcional, `YYYY-MM`.
- `since`: opcional, `YYYY-MM-DD`.

Campos principais: `saldo_atual_contas`, `total_cartoes_mes_atual`, `total_despesas_mes_atual`, `total_receitas_mes_atual`, `receitas_por_categoria`, `despesas_por_categoria`.

### GET `/reports/monthly/`
Query obrigatório: `month=YYYY-MM`.

Resposta: `referencia`, `total_receitas`, `total_despesas`, `saldo`, `pendencias_revisao`, `totais_categoria`, `resumo_diario`.

### GET `/reports/`
Lista relatórios periódicos salvos.
