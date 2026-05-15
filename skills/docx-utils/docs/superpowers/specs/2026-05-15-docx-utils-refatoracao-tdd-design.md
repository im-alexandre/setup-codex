# Refatoracao TDD do docx-utils

Data: 2026-05-15

## Contexto

A skill `docx-utils` tem um CLI .NET/Open XML concentrado principalmente em `src/DocxOpenXmlTools/Program.cs`. O arquivo atua como roteador, parser, executor de comandos, manipulador de DOCX, validador de planos, emissor de relatorios e deposito de modelos auxiliares. A refatoracao deve reduzir esse acoplamento sem alterar o contrato externo.

A suite atual fica em `src/DocxOpenXmlTools.Tests/DocxOpenXmlToolsCliTests.cs`. Ela nao cobre apenas happy path: ja existem testes para contratos invalidos, comando desconhecido, formatos de comentarios, autoria automatica e explicita, `replace-table`, `replace-blocks`, `create-docx` e `create-article`. Mesmo assim, muitos comandos do `switch` principal ainda nao tem caracterizacao dedicada antes da extracao.

Na exploracao inicial, a suite estava com baseline vermelho em `ReplaceTable_by_ordinal_preserves_table_and_cell_styles`: o teste esperava o estilo `TabelaOriginal`, mas observou `tabelauerj`. Esse caso deve ser classificado na auditoria inicial antes de qualquer correcao ou extracao.

## Objetivos

- Preservar exatamente o contrato publico do executavel `docx-utils`.
- Aplicar TDD first com red/green/refactor.
- Revisar e fortalecer testes antes de refatorar.
- Transformar `Program.cs` em um entrypoint mais fino, com dispatch e delegacao para modulos de dominio.
- Dividir a implementacao por arquivos-alvo claros.
- Usar o agente especifico `dotnet-docx-maintainer` para execucao pesada.
- Validar cada pacote antes de passar para um agente final validador/agregador.

## Fora de Escopo

- Alterar nomes de comandos, argumentos, codigos de saida ou formatos de saida.
- Reescrever a arquitetura inteira em um unico passo.
- Trocar dependencias NuGet sem necessidade direta.
- Converter a refatoracao em mudanca funcional.
- Corrigir lacunas de baixo risco que nao bloqueiem a primeira extracao segura.

## Contrato Que Deve Permanecer Identico

O comportamento observavel deve permanecer igual para:

- comandos disponiveis e comandos desconhecidos;
- argumentos e opcoes obrigatorias;
- codigos de saida;
- stdout e stderr relevantes;
- formatos `json`, `markdown`, `raw`, `auto` e tabela de terminal quando aplicavel;
- mutacoes em DOCX;
- relatorios gerados;
- validacao Open XML;
- contratos declarativos em `references/plan-contracts.md` e `references/plan-contracts.json`;
- autoria automatica da thread principal e autoria explicita de subagents.

## Primeira Etapa: Auditoria Dos Testes

A primeira tarefa executavel e uma auditoria da cobertura existente. Ela deve gerar um artefato versionado em:

`docs/refactor/docx-utils-test-coverage-audit.md`

A matriz deve conter pelo menos:

| comando | testes atuais | cobre happy path | cobre erro | cobre contrato externo | lacuna | decisao |
| --- | --- | --- | --- | --- | --- | --- |

Cada comando no `switch` do CLI deve aparecer na matriz. A auditoria deve classificar os testes como:

- happy path;
- erro de argumento ou opcao obrigatoria;
- arquivo ou plano inexistente;
- contrato JSON invalido;
- comando desconhecido ou formato invalido;
- efeito real em DOCX;
- stdout, stderr e codigo de saida;
- teste fragil ou acoplado demais ao detalhe interno.

O teste vermelho de `replace-table` deve ser explicado dentro dessa matriz. A decisao pode ser corrigir expectativa, corrigir producao ou dividir o teste para proteger contratos diferentes, mas isso nao deve acontecer antes da classificacao.

## Fluxo TDD

Cada extracao deve seguir:

1. Red ou caracterizacao: criar ou ajustar teste do dominio antes de mover codigo.
2. Green inicial: provar o comportamento ainda com o codigo antigo quando o teste for de caracterizacao.
3. Refactor: extrair uma unidade pequena para novo arquivo ou classe.
4. Green final: rodar o teste alvo e a suite relevante.
5. Contrato: comparar comportamento observavel quando houver stdout, stderr, codigo de saida, DOCX ou relatorio.

Nenhuma extracao deve entrar se alterar contrato externo sem aprovacao explicita.

## Arquitetura Alvo

O executavel continua tendo a mesma superficie publica. Internamente, `Program.cs` deve caminhar para:

- inicializacao de encoding;
- parse inicial de comando;
- tratamento de `help`, ausencia de argumentos e arquivo DOCX inexistente;
- dispatch para comandos;
- funcoes minimas compartilhadas que ainda nao tenham destino claro.

As responsabilidades extraidas devem ir para arquivos de dominio. A auditoria decide a ordem final, mas a direcao esperada e:

- `Cli/CommandDispatcher.cs`;
- `Cli/CliOptions.cs`;
- `PlanContracts/PlanContractCommands.cs`, se a divisao do suporte atual trouxer clareza;
- `Comments/CommentCommands.cs`;
- `Comments/CommentModels.cs`;
- `Blocks/BlockMutationCommands.cs`;
- `Tables/TableCommands.cs`;
- `Figures/FigureCommands.cs`;
- `Math/FormulaCommands.cs`;
- `Styles/StyleCommands.cs`;
- `Layout/LayoutRepairCommands.cs`.

Dominios com alto risco, como formulas, MathML/OMML, estilos, layout e ABNT, devem vir depois dos dominios com melhor caracterizacao.

## Plano Por Arquivos-Alvo

Depois da auditoria, o main thread deve transformar a matriz em pacotes por arquivo-alvo. Cada pacote deve declarar:

- arquivo novo a criar;
- trecho ou dominio que sai de `Program.cs`;
- testes de caracterizacao obrigatorios;
- comando de validacao alvo;
- dependencias com outros pacotes;
- risco de conflito;
- criterio de pronto.

Pacotes que editem a mesma regiao de `Program.cs` nao devem rodar em paralelo. Eles entram em uma fila de integracao.

## Uso Do Agente dotnet-docx-maintainer

O trabalho pesado deve ser delegado ao `dotnet-docx-maintainer` quando o `agent_type` estiver disponivel na sessao. Se a sessao ainda nao reconhecer o tipo diretamente, o fallback e usar um `worker` com as instrucoes completas de:

- `C:\Users\imale\.codex\skills\docx-utils\SKILL.md`;
- `C:\Users\imale\.codex\skills\docx-utils\agents\mantenedor-dotnet-docx.md`;
- `C:\Users\imale\.codex\agents\dotnet-docx-maintainer.toml`.

Cada agente de implementacao deve:

- trabalhar em um pacote fechado;
- escrever ou ajustar testes antes de extrair;
- rodar os testes do seu escopo;
- devolver arquivos alterados, comandos executados, resultados e riscos.

Depois dos agentes de implementacao, um agente final `dotnet-docx-maintainer` deve atuar como validador/agregador. Ele deve revisar os diffs integrados, rodar a suite completa, conferir contrato externo e listar conflitos ou correcoes objetivas.

O main thread fica responsavel por orquestrar, integrar em ordem, resolver conflitos e decidir se o pacote passa para a proxima etapa.

## Validacao

Validacoes esperadas durante a execucao:

- `dotnet restore src\DocxOpenXmlTools.Tests\DocxOpenXmlTools.Tests.csproj`;
- `dotnet build src\DocxOpenXmlTools.Tests\DocxOpenXmlTools.Tests.csproj --configuration Release --no-restore`;
- `dotnet test src\DocxOpenXmlTools.Tests\DocxOpenXmlTools.Tests.csproj --configuration Release --no-build --no-restore`;
- testes focados por filtro quando um pacote for pequeno;
- validacao Open XML quando o pacote mutar DOCX;
- validacao pelo binario publicado quando a mudanca afetar uso operacional.

## Criterios De Aceite

- Auditoria de testes criada e versionada.
- Baseline vermelho atual classificado antes de qualquer refatoracao.
- Pacotes de implementacao definidos por arquivos-alvo.
- Cada pacote executado com red/green/refactor.
- Cada agente valida suas modificacoes antes de entregar.
- Agente validador/agregador roda a suite relevante completa.
- `Program.cs` diminui por extracoes pequenas e rastreaveis.
- Contrato externo permanece identico.

