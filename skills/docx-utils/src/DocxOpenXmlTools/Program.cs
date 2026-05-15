using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using DocxOpenXmlTools.Cli;
using DocxOpenXmlTools.PlanContracts;
using DocxOpenXmlTools.Mutation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;
using W16Cex = DocumentFormat.OpenXml.Office2021.Word.CommentsExt;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using M = DocumentFormat.OpenXml.Math;

Console.OutputEncoding = Encoding.UTF8;

if (args.Length == 1 && CliOptions.IsHelpArgument(args[0]))
{
    PrintUsage();
    return 0;
}

if (args.Length == 0)
{
    PrintUsage();
    return 1;
}

var command = args[0].Trim().ToLowerInvariant();

if (command is "help" or "-h" or "--help" or "/?")
{
    PrintUsage();
    return 0;
}

if (command is "plan-contracts" or "plan-contract")
{
    return PlanContractCommands.PrintPlanContracts(args.Skip(1).ToArray());
}

if (command is "create-article")
{
    return ArticleDocxBuilderBridge.Run(args.Skip(1).ToArray());
}

if (command is "create-docx")
{
    if (args.Length < 2)
    {
        PrintUsage();
        return 1;
    }

    return CreateDocxCommand(args.Skip(1).ToArray());
}

if (command is "validate-plan")
{
    if (args.Length < 2)
    {
        PrintUsage();
        return 1;
    }

    return PlanContractCommands.ValidatePlan(args.Skip(1).ToArray());
}

if (args.Length < 2)
{
    PrintUsage();
    return 1;
}

var docxPath = Path.GetFullPath(args[1]);

if (!File.Exists(docxPath))
{
    Console.Error.WriteLine($"DOCX not found: {docxPath}");
    return 2;
}

var commandOptions = CliOptions.Parse(args.Skip(2).ToArray());

return command switch
{
    "paragraphs" => ListParagraphs(docxPath, commandOptions),
    "paragraph-detail" => ParagraphDetail(docxPath, commandOptions),
    "structure-audit" => StructureAudit(docxPath, commandOptions),
    "layout-audit" => LayoutAudit(docxPath, commandOptions),
    "equations-audit" => EquationsAudit(docxPath, commandOptions),
    "math-audit" => MathAudit(docxPath, commandOptions),
    "math-text-audit" => MathTextAudit(docxPath, commandOptions),
    "linear-equation-plan-preview" => LinearEquationPlanPreview(docxPath, commandOptions),
    "revisions" => ListRevisions(docxPath, commandOptions),
    "comments" => ListComments(docxPath, commandOptions),
    "comment-anchors" => ListCommentAnchors(docxPath, commandOptions),
    "next-author" => NextAuthor(docxPath),
    "validate" => Validate(docxPath),
    "export-used-styles" => ExportUsedStyles(docxPath, commandOptions),
    "ensure-canonical-styles" => EnsureCanonicalStylesCommand(docxPath, commandOptions),
    "sync-styles-from-docx" => SyncStylesFromDocxCommand(docxPath, commandOptions),
    "insert-tracked" => InsertTracked(docxPath, commandOptions),
    "insert-blocks" => InsertBlocks(docxPath, commandOptions),
    "replace-blocks" => ReplaceBlocks(docxPath, commandOptions),
    "edit-paragraphs" => EditParagraphs(docxPath, commandOptions),
    "rewrite-equation-blocks" => RewriteEquationBlocks(docxPath, commandOptions),
    "format-equation-paragraphs" => FormatEquationParagraphs(docxPath, commandOptions),
    "style-running-text" => StyleRunningText(docxPath, commandOptions),
    "ensure-style-fonts" => EnsureStyleFonts(docxPath, commandOptions),
    "enable-update-fields-on-open" => EnableUpdateFieldsOnOpen(docxPath, commandOptions),
    "disable-update-fields-on-open" => DisableUpdateFieldsOnOpen(docxPath, commandOptions),
    "normalize-figure-indent" => NormalizeFigureIndent(docxPath, commandOptions),
    "apply-table-design-style" => ApplyTableDesignStyle(docxPath, commandOptions),
    "replace-table" => ReplaceTable(docxPath, commandOptions),
    "replace-figures-from-plan" => ReplaceFiguresFromPlan(docxPath, commandOptions),
    "replace-formulas-with-linear-equations" => ReplaceFormulasWithLinearOfficeMath(docxPath, commandOptions),
    "replace-formulas-with-mathml-omml" => ReplaceFormulasWithMathMlOfficeMath(docxPath, commandOptions),
    "convert-text-formulas-to-omath" => ConvertTextFormulasToOfficeMath(docxPath, commandOptions),
    "repair-article-abnt-layout" => RepairArticleAbntLayout(docxPath, commandOptions),
    "format-abnt-reference-titles" => FormatAbntReferenceTitles(docxPath, commandOptions),
    "apply-crossrefs" => ApplyCrossrefs(docxPath, commandOptions),
    "add-bookmarks" => AddBookmarks(docxPath, commandOptions),
    "rewrite-ref-fields" => RewriteRefFields(docxPath, commandOptions),
    "repair-style-captions" => RepairStyleCaptions(docxPath, commandOptions),
    "repair-layout-pendencies" => RepairLayoutPendencies(docxPath, commandOptions),
    "repair-ref-number-only" => RepairRefNumberOnly(docxPath, commandOptions),
    "insert-figures" => InsertFigures(docxPath, commandOptions),
    "insert-comments" => InsertComments(docxPath, commandOptions),
    "reanchor-comments" => ReanchorComments(docxPath, commandOptions),
    "answer-comments" => AnswerComments(docxPath, commandOptions),
    "reply-comments" => ReplyComments(docxPath, commandOptions),
    "remove-comments" => RemoveComments(docxPath, commandOptions),
    "append-paragraphs" => AppendParagraphs(docxPath, commandOptions),
    "accept-revisions" => AcceptRevisions(docxPath, commandOptions),
    _ => UnknownCommand(command)
};

static int AcceptRevisions(string docxPath, IReadOnlyDictionary<string, string> options)
{
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("--lock is required");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils accept-revisions {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("No main document part.");
    var insAccepted = 0;
    var delAccepted = 0;

    void AcceptInPart(OpenXmlPart? part)
    {
        if (part?.RootElement is null) return;
        var root = part.RootElement;
        var insElems = root.Descendants<InsertedRun>().ToList();
        foreach (var ins in insElems)
        {
            var parent = ins.Parent;
            if (parent is null) continue;
            foreach (var child in ins.ChildElements.ToList())
            {
                child.Remove();
                ins.InsertBeforeSelf(child);
            }
            ins.Remove();
            insAccepted++;
        }

        var delElems = root.Descendants<DeletedRun>().ToList();
        foreach (var del in delElems)
        {
            del.Remove();
            delAccepted++;
        }

        var rPrIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunPropertiesChange>().ToList();
        foreach (var c in rPrIns) c.Remove();

        var paraIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.ParagraphPropertiesChange>().ToList();
        foreach (var c in paraIns) c.Remove();

        var sectIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionPropertiesChange>().ToList();
        foreach (var c in sectIns) c.Remove();

        var tableIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.TablePropertiesChange>().ToList();
        foreach (var c in tableIns) c.Remove();

        var trIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRowPropertiesChange>().ToList();
        foreach (var c in trIns) c.Remove();

        var tcIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCellPropertiesChange>().ToList();
        foreach (var c in tcIns) c.Remove();

        var tblPrExIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.TablePropertyExceptionsChange>().ToList();
        foreach (var c in tblPrExIns) c.Remove();

        // Generic w:ins markers inside rPr/pPr (paragraph mark insertion) â€” these are not InsertedRun; they live as direct child of rPr/pPr
        var paragraphMarkIns = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Inserted>().ToList();
        foreach (var c in paragraphMarkIns) c.Remove();

        var paragraphMarkDel = root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Deleted>().ToList();
        foreach (var c in paragraphMarkDel) c.Remove();

        var delTexts = root.Descendants<DeletedText>().ToList();
        foreach (var dt in delTexts) dt.Remove();

        part.RootElement?.Save();
    }

    AcceptInPart(mainPart);
    foreach (var hp in mainPart.HeaderParts) AcceptInPart(hp);
    foreach (var fp in mainPart.FooterParts) AcceptInPart(fp);
    if (mainPart.FootnotesPart != null) AcceptInPart(mainPart.FootnotesPart);
    if (mainPart.EndnotesPart != null) AcceptInPart(mainPart.EndnotesPart);

    // Optionally disable TrackRevisions when --disable-track is passed
    if (options.TryGetValue("disable-track", out var disable) && string.Equals(disable, "true", StringComparison.OrdinalIgnoreCase))
    {
        var settings = mainPart.DocumentSettingsPart?.Settings;
        if (settings != null)
        {
            var trackRev = settings.Elements<DocumentFormat.OpenXml.Wordprocessing.TrackRevisions>().FirstOrDefault();
            trackRev?.Remove();
            settings.Save();
        }
    }

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var sb = new StringBuilder();
        sb.AppendLine("# Accept Revisions Report");
        sb.AppendLine();
        sb.AppendLine($"- DOCX: `{docxPath}`");
        sb.AppendLine($"- Insertions accepted: {insAccepted}");
        sb.AppendLine($"- Deletions accepted: {delAccepted}");
        File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
    }

    Console.WriteLine($"INSERTIONS_ACCEPTED {insAccepted}");
    Console.WriteLine($"DELETIONS_ACCEPTED {delAccepted}");
    return 0;
}

static int AppendParagraphs(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("--plan is required");
        return 4;
    }
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("--lock is required");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<BlockInsertionPlan>(File.ReadAllText(planPath), new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils append-paragraphs {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var metadata = new RevisionMetadata(author, DateTime.UtcNow);
    var body = doc.MainDocumentPart!.Document.Body!;

    foreach (var spec in plan.Blocks)
    {
        var paragraphs = GetParagraphs(doc);
        var presenceParagraphs = GetAllParagraphEntries(doc);
        if (!string.IsNullOrWhiteSpace(spec.UniqueText) && ContentAlreadyPresent(presenceParagraphs, spec.UniqueText))
        {
            skipped.Add($"{spec.Id}: unique text already present");
            continue;
        }

        if (BlockTableAlreadyPresent(doc, spec))
        {
            skipped.Add($"{spec.Id}: equivalent table already present");
            continue;
        }

        // Find last paragraph in body to use as a style template and insertion point
        var lastP = body.Elements<Paragraph>().LastOrDefault() ?? throw new InvalidOperationException("No paragraphs in body to anchor append.");
        var template = lastP;
        OpenXmlElement insertionPoint = lastP;

        foreach (var item in spec.Items)
        {
            OpenXmlElement block = item.Kind.Equals("table", StringComparison.OrdinalIgnoreCase)
                ? CreateTrackedTable(item, metadata)
                : CreateInsertedParagraphWithStyle(template, item.Text ?? "", item.StyleId, metadata);
            insertionPoint.InsertAfterSelf(block);
            insertionPoint = block;
        }
        applied.Add(spec.Id);
    }

    SaveMainDocumentWithValidationRepair(doc, applied);

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var sb = new StringBuilder();
        sb.AppendLine("# Append Paragraphs Report");
        sb.AppendLine();
        sb.AppendLine($"- DOCX: `{docxPath}`");
        sb.AppendLine($"- Plano: `{planPath}`");
        sb.AppendLine($"- Autor: `{author}`");
        sb.AppendLine($"- Aplicados: {applied.Count}");
        sb.AppendLine($"- Ignorados: {skipped.Count}");
        foreach (var item in applied) sb.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) sb.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int UnknownCommand(string command)
{
    Console.Error.WriteLine($"Unknown command: {command}");
    PrintUsage();
    return 3;
}

static int NextAuthor(string docxPath)
{
    var author = MutationAuthorResolver.Resolve(docxPath, new Dictionary<string, string>());
    Console.WriteLine(author);
    return 0;
}

static void PrintUsage()
{
    Console.WriteLine("""
Uso:
  docx-utils <comando> <docx> [opcoes]
  docx-utils create-article <article_spec.json> <output.docx> [author] [--lock <lockfile>] [--template <template.docx>] [--sbpo] [--blind]
  docx-utils create-docx <output.docx> [--plan <json>]
  docx-utils validate-plan <comando> --plan <json>
  docx-utils plan-contracts [comando] [--format markdown|json]
  docx-utils help | --help | -h | /?   Mostra esta ajuda.

Notas gerais:
  - Comandos que alteram DOCX exigem --lock <lockfile> para escrita exclusiva.
  - Em mutacoes, --author e opcional para a thread principal: se omitido, o utilitario escolhe automaticamente o proximo autor livre no DOCX.
  - Subagents devem passar --author com o nome atribuido ao subagent.
  - Planos JSON devem ser arquivos no formato esperado pelo comando; use --report md para gerar um relatorio auditavel quando disponivel.
  - A listagem externa de autores foi removida; a leitura de autores existe apenas como logica interna do autor automatico.
  - `next-author <docx>` imprime o proximo autor livre sem mutar o documento.
  - `create-article` delega ao comportamento exato do binario `ArticleDocxBuilder`.
  - `create-docx` cria um DOCX vazio quando chamado sem `--plan` e usa um plano JSON quando informado.

Planos de blocos e tabelas:
  - plan-contracts [comando] [--format markdown|json] expõe os contratos operacionais de planos sem ler a implementacao.
  - validate-plan <comando> --plan <json> valida o contrato JSON de create-docx, insert-blocks, replace-blocks ou replace-table antes de mutar o DOCX.
  - insert-blocks, replace-blocks e replace-table seguem os contratos publicados em references/plan-contracts.md e references/plan-contracts.json.
  - replace-blocks remove o intervalo entre `afterPrefix` e `beforePrefix` e insere parágrafos/tabelas declarativamente no lugar.
  - Veja references/plan-contracts.md e references/plan-contracts.json para os contratos minimos e exemplos JSON completos.

Criação de documentos:
  - create-article <article_spec.json> <output.docx> [author] [--lock <lockfile>] [--template <template.docx>] [--sbpo] [--blind]
    Mantém exatamente o comportamento do binário ArticleDocxBuilder.
    Exemplo: docx-utils create-article artigo.json artigo.docx Ultron

  - create-docx <output.docx> [--plan <json>]
    Cria um DOCX do zero. Sem --plan, gera um arquivo vazio; com plano, renderiza title, paragraphs, subtitles, sections e references.
    Exemplo vazio: docx-utils create-docx novo.docx
    Exemplo com plano: docx-utils create-docx novo.docx --plan documento.json

Inspecao e auditoria:
  paragraphs <docx> [--start N] [--count N] [--contains TEXT] [--all true|false]
    Lista paragrafos com indice e texto. Serve para achar ancoras antes de montar planos JSON.
    Exemplo: docx-utils paragraphs tese.docx --contains "Resultados" --count 20

  paragraph-detail <docx> --index N [--all true|false]
    Mostra detalhes de um paragrafo especifico, incluindo estilo e numeracao efetiva.
    Exemplo: docx-utils paragraph-detail tese.docx --index 42

  structure-audit <docx> [--out json]
    Audita estrutura geral: paragrafos, tabelas, figuras e equacoes.
    Exemplo: docx-utils structure-audit tese.docx --out estrutura.json

  layout-audit <docx> [--out json] [--report md]
    Audita layout de tabelas e figuras, incluindo estilos, larguras, alinhamento, fonte e fonte/legenda.
    Exemplo: docx-utils layout-audit tese.docx --out layout.json --report layout.md

  equations-audit <docx> [--out json]
    Audita equacoes, campos SEQ e numeracao de equacoes de dissertacao.
    Exemplo: docx-utils equations-audit tese.docx --out equacoes.json

  math-audit <docx> [--out json]
    Lista blocos Office Math e candidatos matematicos relevantes.
    Exemplo: docx-utils math-audit tese.docx --out math.json

  math-text-audit <docx> [--out json]
    Procura formulas textuais que podem precisar virar Office Math.
    Exemplo: docx-utils math-text-audit tese.docx --out formulas_texto.json

  linear-equation-plan-preview <docx> --plan json --out html
    Gera uma previa HTML de plano de conversao de equacoes lineares antes de mutar o DOCX.
    Exemplo: docx-utils linear-equation-plan-preview tese.docx --plan equacoes.json --out previa.html

  revisions <docx> [--author TEXT]
    Lista revisoes rastreadas, opcionalmente filtradas por autor.
    Exemplo: docx-utils revisions tese.docx --author Ultron

  comments <docx> [--author TEXT] [--format auto|table|json|markdown|raw]
    Lista comentarios e respostas. Sem --format, detecta o ambiente: tabela no CLI e Markdown no app.
    Use --format json para dados estruturados ou --format raw para recuperar a saida textual legada.
    Exemplo: docx-utils comments tese.docx --author Brainiac

  comment-anchors <docx>
    Lista paragrafos que possuem marcadores de comentario e os IDs associados.
    Exemplo: docx-utils comment-anchors tese.docx

  next-author <docx>
    Mostra, sem mutar o DOCX, o proximo nome livre da lista automatica de autores.
    Exemplo: docx-utils next-author tese.docx

  validate <docx>
    Valida o pacote Open XML e informa TrackRevisions, campos e erros acionaveis.
    Exemplo: docx-utils validate tese.docx

Estilos e formatacao:
  export-used-styles <docx> [--out dir]
    Exporta estilos usados no DOCX para inspecao/canonizacao.
    Exemplo: docx-utils export-used-styles tese.docx --out estilos_usados

  ensure-canonical-styles <docx> [--author NAME] --lock <lockfile> [--source dir] [--report md]
    Copia estilos canonicos para o DOCX e normaliza estilo de tabela academica.
    Exemplo: docx-utils ensure-canonical-styles tese.docx --lock tese.lock --source references/estilos --report estilos.md

  sync-styles-from-docx <target.docx> --source-docx <source.docx> [--author NAME] --lock <lockfile> [--report md]
    Sincroniza estilos do DOCX fonte para o DOCX alvo.
    Exemplo: docx-utils sync-styles-from-docx tese.docx --source-docx base.docx --lock tese.lock --report sync.md

  style-running-text <docx> [--author NAME] --lock <lockfile> [--report md]
    Aplica estilo/fonte/tamanho de texto corrido em paragrafos elegiveis, preservando estilos protegidos.
    Exemplo: docx-utils style-running-text tese.docx --lock tese.lock --report texto_corrido.md

  ensure-style-fonts <docx> [--author NAME] --lock <lockfile> [--font NAME] [--report md]
    Garante fonte dos estilos relevantes.
    Exemplo: docx-utils ensure-style-fonts tese.docx --lock tese.lock --font "Times New Roman"

  format-equation-paragraphs <docx> [--author NAME] --lock <lockfile> [--style-id ID] [--seq-name NAME] [--report md]
    Aplica estilo e ajustes de paragrafo em equacoes numeradas.
    Exemplo: docx-utils format-equation-paragraphs tese.docx --lock tese.lock --style-id Equacao --report equacoes.md

  normalize-figure-indent <docx> [--author NAME] --lock <lockfile> [--report md]
    Remove recuo indevido e centraliza paragrafos de figura.
    Exemplo: docx-utils normalize-figure-indent tese.docx --lock tese.lock --report figuras.md

  apply-table-design-style <docx> [--author NAME] --lock <lockfile> --style-id ID [--style-name NAME] [--report md]
    Aplica estilo de tabela, como tabelauerj, em tabelas do documento.
    Exemplo: docx-utils apply-table-design-style tese.docx --lock tese.lock --style-id tabelauerj --style-name tabela_uerj

  replace-table <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Substitui uma tabela existente preservando propriedades da tabela e estilos das celulas conforme o plano.
    Seleciona a tabela por ordinal, bloco, texto da primeira celula, ou por prefixos dos paragrafos vizinhos.
    Exemplo minimo de JSON:
      {
        "tables": [
          {
            "id": "tabela-1",
            "ordinal": 2,
            "rows": [["Novo A1", "Novo A2"], ["Novo B1", "Novo B2"]]
          }
        ]
      }
    Exemplo: docx-utils replace-table tese.docx --plan tabela.json --lock tese.lock --report tabela.md

  replace-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Substitui o intervalo de blocos entre `afterPrefix` e `beforePrefix` e insere os blocos declarados no lugar.
    A skill cria as linhas, celulas e tabelas OpenXML; nao monte `w:tr`/`w:tc` manualmente em PowerShell.
    Exemplo de substituicao por tabela:
      {
        "blocks": [
          {
            "id": "intervalo-1",
            "afterPrefix": "Introducao",
            "beforePrefix": "Conclusao",
            "items": [
              {
                "kind": "table",
                "tableStyleId": "tabelauerj",
                "cellStyleId": "dados",
                "rows": [["A1", "A2"], ["B1", "B2"]]
              }
            ]
          }
        ]
      }
    Exemplo: docx-utils replace-blocks tese.docx --plan blocos.json --lock tese.lock --report blocos.md

  validate-plan <comando> --plan <json>
    Valida o contrato JSON de `create-docx`, `insert-blocks`, `replace-blocks` ou `replace-table` sem mutar o DOCX.
    Exemplo: docx-utils validate-plan insert-blocks --plan blocos.json

  plan-contracts [comando] [--format markdown|json]
    Retorna os contratos operacionais em Markdown ou JSON, com fonte em references/plan-contracts.json.
    Exemplo: docx-utils plan-contracts replace-table --format json

  enable-update-fields-on-open <docx> [--author NAME] --lock <lockfile> [--report md]
    Configura o Word para atualizar campos ao abrir.
    Exemplo: docx-utils enable-update-fields-on-open tese.docx --lock tese.lock

  disable-update-fields-on-open <docx> [--author NAME] --lock <lockfile> [--report md]
    Remove a configuracao de atualizar campos ao abrir.
    Exemplo: docx-utils disable-update-fields-on-open tese.docx --lock tese.lock

Edicao com revisoes rastreadas:
  insert-tracked <docx> --plan <json> [--author NAME] [--lock <lockfile>] [--report md]
    Insere paragrafos com revisao rastreada entre ancoras after/before.
    Exemplo: docx-utils insert-tracked tese.docx --plan insercoes.json --lock tese.lock --report insercoes.md

  insert-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Insere blocos de paragrafo/tabela com revisoes entre afterPrefix e beforePrefix.
    BlockSpec aceita id, afterPrefix, beforePrefix, uniqueText, styleSource e items[].
    BlockItemSpec usa kind=paragraph com text/styleId, ou kind=table com rows/tableStyleId/cellStyleId.
    Exemplo minimo de JSON:
      {
        "blocks": [
          {
            "id": "bloco-1",
            "afterPrefix": "Texto antes",
            "beforePrefix": "Texto depois",
            "items": [
              { "kind": "paragraph", "text": "Paragrafo inserido", "styleId": "CorpoTexto" },
              {
                "kind": "table",
                "tableStyleId": "tabelauerj",
                "cellStyleId": "dados",
                "rows": [["A1", "A2"], ["B1", "B2"]]
              }
            ]
          }
        ]
      }
    Exemplo: docx-utils insert-blocks tese.docx --plan blocos.json --lock tese.lock --report blocos.md

  append-paragraphs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Acrescenta blocos ao final do documento com revisoes rastreadas.
    Exemplo: docx-utils append-paragraphs tese.docx --plan apendice.json --lock tese.lock

  edit-paragraphs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Altera texto e/ou estilo de paragrafos localizados por prefixo.
    Exemplo: docx-utils edit-paragraphs tese.docx --plan edicoes.json --lock tese.lock --report edicoes.md

Figuras, formulas e referencias:
  insert-figures <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Insere figuras com legenda e fonte a partir de plano JSON.
    Exemplo: docx-utils insert-figures tese.docx --plan figuras.json --lock tese.lock --report figuras.md

  replace-figures-from-plan <docx> --plan json [--author NAME] --lock <lockfile> [--report md]
    Substitui imagens existentes mantendo a estrutura indicada pelo plano.
    Exemplo: docx-utils replace-figures-from-plan tese.docx --plan substituir_figuras.json --lock tese.lock

  rewrite-equation-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Reescreve blocos de equacao por ancoras textuais.
    Exemplo: docx-utils rewrite-equation-blocks tese.docx --plan eq_rewrite.json --lock tese.lock

  replace-formulas-with-linear-equations <docx> --plan json [--author NAME] --lock <lockfile> [--keep-linear true|false] [--report md]
    Converte formulas textuais em equacoes Office Math a partir de LaTeX linear.
    Exemplo: docx-utils replace-formulas-with-linear-equations tese.docx --plan formulas.json --lock tese.lock --report formulas.md

  replace-formulas-with-mathml-omml <docx> --plan json [--author NAME] --lock <lockfile> [--xsl MML2OMML.XSL] [--report md]
    Converte MathML para OMML/Office Math usando XSL.
    Exemplo: docx-utils replace-formulas-with-mathml-omml tese.docx --plan mathml.json --lock tese.lock --xsl MML2OMML.XSL

  convert-text-formulas-to-omath <docx> --plan json [--author NAME] --lock <lockfile> [--report md]
    Alias legado: encaminha para o fluxo padronizado de LaTeX linear.
    Exemplo: docx-utils convert-text-formulas-to-omath tese.docx --plan formulas.json --lock tese.lock

  apply-crossrefs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Numera legendas, adiciona bookmarks e converte chamadas em referencias cruzadas.
    Exemplo: docx-utils apply-crossrefs tese.docx --plan crossrefs.json --lock tese.lock --report crossrefs.md

  add-bookmarks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Adiciona bookmarks em paragrafos alvo.
    Exemplo: docx-utils add-bookmarks tese.docx --plan bookmarks.json --lock tese.lock

  rewrite-ref-fields <docx> [--author NAME] --lock <lockfile> --bookmark-prefixes CSV --template TEXT [--report md]
    Reescreve campos REF existentes com base em prefixos de bookmark.
    Exemplo: docx-utils rewrite-ref-fields tese.docx --lock tese.lock --bookmark-prefixes fig_,tab_ --template "REF {0} \h"

Comentarios:
  insert-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Insere comentarios ancorados por prefixo/contains.
    Exemplo: docx-utils insert-comments tese.docx --plan comentarios.json --lock tese.lock --report comentarios.md

  reanchor-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Move marcadores de comentarios existentes para novos paragrafos.
    Exemplo: docx-utils reanchor-comments tese.docx --plan reancorar.json --lock tese.lock

  answer-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Acrescenta resposta textual dentro do comentario existente.
    Exemplo: docx-utils answer-comments tese.docx --plan respostas.json --lock tese.lock

  reply-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Cria replies modernas vinculadas a comentarios pai.
    Exemplo: docx-utils reply-comments tese.docx --plan replies.json --lock tese.lock

  remove-comments <docx> --ids CSV|all [--author NAME] --lock <lockfile> [--report md]
    Remove comentarios por ID ou todos os comentarios.
    Exemplo: docx-utils remove-comments tese.docx --ids 3,4 --lock tese.lock

Reparos e ajustes academicos:
  repair-article-abnt-layout <docx> [--author NAME] --lock <lockfile> [--report md]
    Repara layout ABNT de artigo, incluindo estilos e elementos estruturais previstos pelo utilitario.
    Exemplo: docx-utils repair-article-abnt-layout artigo.docx --lock artigo.lock --report abnt.md

  format-abnt-reference-titles <docx> [--author NAME] --lock <lockfile> [--target publication|article|both] [--emphasis italic|bold] [--report md]
    Aplica destaque ABNT em titulos de referencias conforme alvo e enfase.
    Exemplo: docx-utils format-abnt-reference-titles tese.docx --lock tese.lock --target both --emphasis bold

  repair-style-captions <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]
    Repara estilos de legendas conforme plano de captions/crossrefs.
    Exemplo: docx-utils repair-style-captions tese.docx --plan captions.json --lock tese.lock

  repair-layout-pendencies <docx> [--author NAME] --lock <lockfile> [--report md]
    Corrige pendencias de layout conhecidas em figuras, tabelas e fontes.
    Exemplo: docx-utils repair-layout-pendencies tese.docx --lock tese.lock --report pendencias.md

  repair-ref-number-only <docx> [--author NAME] --lock <lockfile> [--report md]
    Repara referencias cruzadas que devem exibir apenas numero.
    Exemplo: docx-utils repair-ref-number-only tese.docx --lock tese.lock --report refs.md

Finalizacao:
  accept-revisions <docx> --lock <lockfile> [--disable-track true|false] [--report md]
    Aceita insercoes/delecoes rastreadas e opcionalmente desativa TrackRevisions.
    Exemplo: docx-utils accept-revisions tese.docx --lock tese.lock --disable-track true --report aceite.md
""");
}

static void ApplyCommentInitials(Comment comment)
{
    comment.Initials = null;
}

static int ListParagraphs(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var includeAll = CliOptions.IsTrue(options, "all");
    var paragraphs = includeAll ? GetAllParagraphEntries(doc) : GetParagraphs(doc);

    var contains = options.TryGetValue("contains", out var containsValue)
        ? Normalize(containsValue)
        : null;
    var start = options.TryGetValue("start", out var startValue) ? int.Parse(startValue, CultureInfo.InvariantCulture) : 0;
    var count = options.TryGetValue("count", out var countValue) ? int.Parse(countValue, CultureInfo.InvariantCulture) : 20;

    IEnumerable<ParagraphEntry> selected = paragraphs;
    if (!string.IsNullOrWhiteSpace(contains))
    {
        selected = selected.Where(p => Normalize(p.Text).Contains(contains, StringComparison.Ordinal));
    }
    else
    {
        selected = selected.Where(p => p.Index >= start).Take(count);
    }

    foreach (var entry in selected)
    {
        Console.WriteLine($"{(includeAll ? "AP" : "P")}[{entry.Index}] {entry.Text}");
    }

    return 0;
}

static int ParagraphDetail(string docxPath, IReadOnlyDictionary<string, string> options)
{
    if (!options.TryGetValue("index", out var indexValue))
    {
        Console.Error.WriteLine("Missing required option: --index");
        return 4;
    }

    var targetIndex = int.Parse(indexValue, CultureInfo.InvariantCulture);
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var includeAll = CliOptions.IsTrue(options, "all");
    var paragraphs = includeAll ? GetAllParagraphEntries(doc) : GetParagraphs(doc);
    var entry = paragraphs.FirstOrDefault(p => p.Index == targetIndex);
    if (entry is null)
    {
        Console.Error.WriteLine($"Paragraph index not found: {targetIndex}");
        return 5;
    }

    var p = entry.Paragraph;
    var props = p.ParagraphProperties;
    var numbering = props?.NumberingProperties;
    var styleId = ParagraphStyleId(p);
    var style = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
        .Elements<Style>()
        .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
    var styleNumbering = style?.StyleParagraphProperties?.NumberingProperties;
    var numberingId = numbering?.NumberingId?.Val is null
        ? ""
        : numbering.NumberingId.Val.Value.ToString(CultureInfo.InvariantCulture);
    var numberingLevel = numbering?.NumberingLevelReference?.Val is null
        ? ""
        : numbering.NumberingLevelReference.Val.Value.ToString(CultureInfo.InvariantCulture);
    var styleNumberingId = styleNumbering?.NumberingId?.Val is null
        ? ""
        : styleNumbering.NumberingId.Val.Value.ToString(CultureInfo.InvariantCulture);
    var styleNumberingLevel = styleNumbering?.NumberingLevelReference?.Val is null
        ? ""
        : styleNumbering.NumberingLevelReference.Val.Value.ToString(CultureInfo.InvariantCulture);
    var effectiveNumberingId = !string.IsNullOrWhiteSpace(numberingId) ? numberingId : styleNumberingId;
    var effectiveNumberingLevel = !string.IsNullOrWhiteSpace(numberingLevel) ? numberingLevel : styleNumberingLevel;
    var numberingFormat = ResolveNumberingFormat(doc, effectiveNumberingId, effectiveNumberingLevel);
    var fields = ExtractFields(p);

    Console.WriteLine($"{(includeAll ? "AP" : "P")}[{entry.Index}] {entry.Text}");
    Console.WriteLine($"StyleId: {styleId}");
    Console.WriteLine($"StyleName: {style?.StyleName?.Val?.Value ?? ""}");
    Console.WriteLine($"StyleBasedOn: {style?.BasedOn?.Val?.Value ?? ""}");
    Console.WriteLine($"InTable: {p.Ancestors<TableCell>().Any()}");
    Console.WriteLine($"HasDrawing: {HasDrawing(p)}");
    Console.WriteLine($"HasMath: {HasMath(p)}");
    Console.WriteLine($"NumberingId: {numberingId}");
    Console.WriteLine($"NumberingLevel: {numberingLevel}");
    Console.WriteLine($"StyleNumberingId: {styleNumberingId}");
    Console.WriteLine($"StyleNumberingLevel: {styleNumberingLevel}");
    Console.WriteLine($"EffectiveNumberingId: {effectiveNumberingId}");
    Console.WriteLine($"EffectiveNumberingLevel: {effectiveNumberingLevel}");
    Console.WriteLine($"EffectiveNumberingFormat: {numberingFormat}");
    Console.WriteLine($"FieldCount: {fields.Count}");
    foreach (var field in fields)
    {
        Console.WriteLine($"Field: instr=\"{field.Instruction}\" result=\"{field.ResultText}\"");
    }

    return 0;
}

static int MathTextAudit(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var paragraphs = GetAllParagraphEntries(doc);
    var findings = paragraphs
        .Select(AnalyzeMathParagraph)
        .Where(entry => entry.OfficeMathCount > 0 || entry.TextMathCandidates.Count > 0)
        .ToList();

    var result = new MathTextAuditResult(
        docxPath,
        DateTime.UtcNow,
        paragraphs.Count,
        findings.Count,
        findings.Count(entry => entry.OfficeMathCount > 0),
        findings.Count(entry => entry.TextMathCandidates.Count > 0),
        findings.Sum(entry => entry.OfficeMathCount),
        findings);

    var payload = JsonSerializer.Serialize(result, CliOptions.JsonOptionsIndented());
    if (options.TryGetValue("out", out var outPath) && !string.IsNullOrWhiteSpace(outPath))
    {
        var fullOutPath = Path.GetFullPath(outPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullOutPath) ?? ".");
        File.WriteAllText(fullOutPath, payload, Encoding.UTF8);
        Console.WriteLine($"Wrote math text audit: {fullOutPath}");
    }
    else
    {
        Console.WriteLine(payload);
    }

    return 0;
}

static MathTextAuditEntry AnalyzeMathParagraph(ParagraphEntry entry)
{
    var officeMathCount = entry.Paragraph.Descendants<M.OfficeMath>().Count();
    var textOutsideMath = ParagraphTextOutsideOfficeMath(entry.Paragraph);
    var candidates = FindTextMathCandidates(textOutsideMath);
    return new MathTextAuditEntry(
        entry.Index,
        entry.Text,
        entry.Paragraph.Ancestors<TableCell>().Any() ? "TableCell" : "Body",
        officeMathCount,
        candidates);
}

static string ParagraphTextOutsideOfficeMath(Paragraph paragraph)
{
    var builder = new StringBuilder();
    foreach (var text in paragraph.Descendants<Text>())
    {
        if (text.Ancestors().Any(a => a.LocalName is "oMath" or "oMathPara"))
        {
            continue;
        }

        builder.Append(text.Text);
    }

    return builder.ToString();
}

static IReadOnlyList<MathTextCandidate> FindTextMathCandidates(string text)
{
    if (string.IsNullOrWhiteSpace(text))
    {
        return [];
    }

    var patterns = new (string Reason, Regex Regex)[]
    {
        ("greek-letter", new Regex(@"[Î‘-Î©Î±-Ï‰Ï-Ï–]", RegexOptions.CultureInvariant)),
        ("subsup-notation", new Regex(@"\b[\p{L}][\p{L}\d]*\s*[_^]\s*[\p{L}\d]+\b", RegexOptions.CultureInvariant)),
        ("operator-expression", new Regex(@"(?<!R\$)\b[\p{L}\d\)\]]+\s*(?:=|â‰¤|â‰¥|â‰ˆ|<|>|Â±|âˆ’|\+|/|\*)\s*[\p{L}\d\(\[]+", RegexOptions.CultureInvariant)),
        ("function-like", new Regex(@"\b[\p{L}][\p{L}\d]*\([^\)]+\)", RegexOptions.CultureInvariant)),
        ("risk-metric-suffix", new Regex(@"\b(?:VaR|CVaR|TVaR|EVaR|ES|RMSE|MAE|MAPE|R2|RÂ²)\s*\d+(?:[.,]\d+)?\b", RegexOptions.CultureInvariant))
    };

    var candidates = new List<MathTextCandidate>();
    var seen = new HashSet<string>(StringComparer.Ordinal);
    foreach (var (reason, regex) in patterns)
    {
        foreach (Match match in regex.Matches(text))
        {
            var value = match.Value.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            var key = $"{reason}|{value}";
            if (!seen.Add(key))
            {
                continue;
            }

            candidates.Add(new MathTextCandidate(reason, value));
        }
    }

    return candidates;
}

static int StructureAudit(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var blocks = body.Elements<OpenXmlElement>().ToList();
    var paragraphEntries = blocks
        .Select((element, blockIndex) => new { element, blockIndex })
        .Where(x => x.element is Paragraph)
        .Select((x, paragraphIndex) => new ParagraphAuditEntry(
            paragraphIndex,
            x.blockIndex,
            ParagraphText((Paragraph)x.element),
            ParagraphStyleId((Paragraph)x.element),
            HasDrawing((Paragraph)x.element),
            HasMath((Paragraph)x.element)))
        .ToList();

    var paragraphIndexByBlock = paragraphEntries.ToDictionary(p => p.BlockIndex, p => p.Index);
    var tables = new List<TableAuditEntry>();
    var figures = new List<FigureAuditEntry>();
    var equations = new List<EquationAuditEntry>();

    var tableCounter = 0;
    foreach (var item in blocks.Select((element, blockIndex) => new { element, blockIndex }))
    {
        if (item.element is Table table)
        {
            tables.Add(BuildTableAudit(++tableCounter, item.blockIndex, table, paragraphEntries, blocks, paragraphIndexByBlock));
        }

        if (item.element is Paragraph paragraph)
        {
            foreach (var drawing in paragraph.Descendants<Drawing>())
            {
                figures.Add(BuildFigureAudit(figures.Count + 1, item.blockIndex, paragraph, drawing, paragraphEntries, paragraphIndexByBlock));
            }

            if (ContainsMath(paragraph))
            {
                equations.Add(BuildEquationAudit(equations.Count + 1, item.blockIndex, paragraph, paragraphEntries, paragraphIndexByBlock));
            }
        }
    }

    var audit = new StructureAuditReport(
        docxPath,
        DateTime.UtcNow,
        paragraphEntries.Count,
        tables,
        figures,
        equations);

    var json = JsonSerializer.Serialize(audit, new JsonSerializerOptions { WriteIndented = true });
    if (options.TryGetValue("out", out var outPathValue))
    {
        var outPath = Path.GetFullPath(outPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
        File.WriteAllText(outPath, json, Encoding.UTF8);
        Console.WriteLine($"Wrote structure audit: {outPath}");
    }
    else
    {
        Console.WriteLine(json);
    }

    return 0;
}

static int LayoutAudit(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var blocks = body.Elements<OpenXmlElement>().ToList();
    var paragraphEntries = blocks
        .Select((element, blockIndex) => new { element, blockIndex })
        .Where(x => x.element is Paragraph)
        .Select((x, paragraphIndex) => new ParagraphAuditEntry(
            paragraphIndex,
            x.blockIndex,
            ParagraphText((Paragraph)x.element),
            ParagraphStyleId((Paragraph)x.element),
            HasDrawing((Paragraph)x.element),
            HasMath((Paragraph)x.element)))
        .ToList();

    var paragraphIndexByBlock = paragraphEntries.ToDictionary(p => p.BlockIndex, p => p.Index);
    var tables = new List<TableLayoutEntry>();
    var figures = new List<FigureLayoutEntry>();

    foreach (var item in blocks.Select((element, blockIndex) => new { element, blockIndex }))
    {
        if (item.element is Table table)
        {
            tables.Add(BuildTableLayoutEntry(doc, tables.Count + 1, item.blockIndex, table, paragraphEntries));
        }

        if (item.element is Paragraph paragraph)
        {
            foreach (var drawing in paragraph.Descendants<Drawing>())
            {
                figures.Add(BuildFigureLayoutEntry(figures.Count + 1, item.blockIndex, paragraph, drawing, paragraphEntries, paragraphIndexByBlock));
            }
        }
    }

    var report = new LayoutAuditReport(
        docxPath,
        DateTime.UtcNow,
        tables.Count,
        tables.Count(t => t.IsAcademicTable),
        figures.Count,
        figures.Count(f => f.IsAcademicFigure),
        tables,
        figures);

    var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
    if (options.TryGetValue("out", out var outPathValue))
    {
        var outPath = Path.GetFullPath(outPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
        File.WriteAllText(outPath, json, Encoding.UTF8);
        Console.WriteLine($"Wrote layout audit: {outPath}");
    }
    else
    {
        Console.WriteLine(json);
    }

    if (options.TryGetValue("report", out var reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildLayoutAuditMarkdown(report), Encoding.UTF8);
        Console.WriteLine($"Wrote layout report: {reportPath}");
    }

    return 0;
}

static int MathAudit(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var entries = body.Descendants<Paragraph>()
        .Select((paragraph, paragraphIndex) => new { paragraph, paragraphIndex })
        .SelectMany(item => item.paragraph.Descendants<M.OfficeMath>()
            .Select((math, mathIndex) =>
            {
                var serialized = SerializeOfficeMathForAudit(math);
                return new MathAuditEntry(
                    item.paragraphIndex,
                    mathIndex,
                    ParagraphText(item.paragraph),
                    serialized,
                    LooksLikeLinearLatex(serialized));
            }))
        .ToList();

    var json = JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true });
    if (options.TryGetValue("out", out var outPathValue))
    {
        var outPath = Path.GetFullPath(outPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
        File.WriteAllText(outPath, json, Encoding.UTF8);
        Console.WriteLine($"Wrote math audit: {outPath}");
    }
    else
    {
        Console.WriteLine(json);
    }

    return 0;
}

static int LinearEquationPlanPreview(string docxPath, IReadOnlyDictionary<string, string> options)
{
    if (!options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue)
        || !options.TryGetValue("out", out var outPathValue) || string.IsNullOrWhiteSpace(outPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --out");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    var outPath = Path.GetFullPath(outPathValue);
    var plan = JsonSerializer.Deserialize<FormulaConversionPlan>(File.ReadAllText(planPath, Encoding.UTF8), CliOptions.JsonOptions()) ?? new FormulaConversionPlan();
    Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");

    var builder = new StringBuilder();
    builder.AppendLine("<!doctype html>");
    builder.AppendLine("<html lang=\"pt-BR\"><head><meta charset=\"utf-8\">");
    builder.AppendLine("<title>Equation plan preview</title>");
    builder.AppendLine("<script>window.MathJax={tex:{inlineMath:[[\"\\\\(\",\"\\\\)\"]],displayMath:[[\"\\\\[\",\"\\\\]\"]]}};</script>");
    builder.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js\"></script>");
    builder.AppendLine("<style>body{font-family:Arial,sans-serif;margin:24px;line-height:1.4}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:8px;vertical-align:top}code{white-space:pre-wrap}.math{font-size:120%}</style>");
    builder.AppendLine("</head><body>");
    builder.AppendLine("<h1>Equation plan preview</h1>");
    builder.AppendLine($"<p>DOCX: <code>{WebUtility.HtmlEncode(docxPath)}</code></p>");
    builder.AppendLine($"<p>Plan: <code>{WebUtility.HtmlEncode(planPath)}</code></p>");
    builder.AppendLine("<table><thead><tr><th>#</th><th>Anchor text</th><th>LaTeX</th><th>Preview</th></tr></thead><tbody>");
    var index = 1;
    foreach (var formula in plan.Formulas)
    {
        var latex = formula.Latex ?? "";
        builder.AppendLine("<tr>");
        builder.AppendLine($"<td>{index++}</td>");
        builder.AppendLine($"<td>{WebUtility.HtmlEncode(formula.Text)}</td>");
        builder.AppendLine($"<td><code>{WebUtility.HtmlEncode(latex)}</code></td>");
        builder.AppendLine($"<td class=\"math\">\\({WebUtility.HtmlEncode(latex)}\\)</td>");
        builder.AppendLine("</tr>");
    }
    builder.AppendLine("</tbody></table></body></html>");

    File.WriteAllText(outPath, builder.ToString(), Encoding.UTF8);
    Console.WriteLine($"Wrote equation preview: {outPath}");
    return 0;
}

static bool LooksLikeLinearLatex(string text)
{
    var normalized = Regex.Replace(text, "\\s+", " ").Trim();
    return normalized.Contains('\\', StringComparison.Ordinal)
        || Regex.IsMatch(normalized, @"[_^{}]", RegexOptions.CultureInvariant)
        || Regex.IsMatch(normalized, @"\\(frac|sum|hat|alpha|lambda|tau|in|ge|cdot|arg|min)", RegexOptions.CultureInvariant)
        || Regex.IsMatch(normalized, @"^[A-Za-z]+\(.*\)$", RegexOptions.CultureInvariant);
}

static int EquationsAudit(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var blocks = body.Elements<OpenXmlElement>().ToList();
    var paragraphEntries = blocks
        .Select((element, blockIndex) => new { element, blockIndex })
        .Where(x => x.element is Paragraph)
        .Select((x, paragraphIndex) => new ParagraphAuditEntry(
            paragraphIndex,
            x.blockIndex,
            ParagraphText((Paragraph)x.element),
            ParagraphStyleId((Paragraph)x.element),
            HasDrawing((Paragraph)x.element),
            HasMath((Paragraph)x.element)))
        .ToList();

    var paragraphIndexByBlock = paragraphEntries.ToDictionary(p => p.BlockIndex, p => p.Index);
    var mathParagraphs = new List<MathParagraphAuditEntry>();
    var allSeqFields = new List<SeqFieldAuditEntry>();
    var dissertationEquations = new List<NumberedDissertationEquationEntry>();

    foreach (var item in blocks.Select((element, blockIndex) => new { element, blockIndex }))
    {
        if (item.element is not Paragraph paragraph)
        {
            continue;
        }

        var paragraphIndex = paragraphIndexByBlock.TryGetValue(item.blockIndex, out var pIndex) ? pIndex : -1;
        var paragraphText = ParagraphText(paragraph);
        var styleId = ParagraphStyleId(paragraph);
        var hasMath = ContainsMath(paragraph);
        var seqFields = ExtractSeqFields(paragraph).ToList();

        foreach (var seq in seqFields)
        {
            allSeqFields.Add(new SeqFieldAuditEntry(
                allSeqFields.Count + 1,
                item.blockIndex,
                paragraphIndex,
                styleId,
                paragraphText,
                seq.Instruction,
                seq.SequenceName,
                seq.ResultText,
                seq.ResultNumber,
                IsCaptionSeq(seq.SequenceName, paragraphText, styleId)));
        }

        if (!hasMath)
        {
            continue;
        }

        var mathEndOrder = LastMathOrder(paragraph);
        var qualifyingSeq = seqFields
            .Where(seq => seq.BeginOrder > mathEndOrder)
            .Where(seq => seq.ResultNumber is not null)
            .Where(seq => !IsCaptionSeq(seq.SequenceName, paragraphText, styleId))
            .Where(seq => HasParenthesizedSeqResult(paragraphText, seq.ResultNumber!.Value))
            .OrderBy(seq => seq.BeginOrder)
            .FirstOrDefault();

        var isDissertationEquation = qualifyingSeq is not null;
        mathParagraphs.Add(new MathParagraphAuditEntry(
            mathParagraphs.Count + 1,
            item.blockIndex,
            paragraphIndex,
            styleId,
            paragraphText,
            seqFields.Select(seq => new SeqFieldSummary(seq.Instruction, seq.SequenceName, seq.ResultText, seq.ResultNumber)).ToList(),
            isDissertationEquation,
            qualifyingSeq?.ResultNumber));

        if (isDissertationEquation)
        {
            var previous = PreviousParagraph(paragraphEntries, item.blockIndex);
            var next = NextParagraph(paragraphEntries, item.blockIndex);
            dissertationEquations.Add(new NumberedDissertationEquationEntry(
                dissertationEquations.Count + 1,
                item.blockIndex,
                paragraphIndex,
                styleId,
                qualifyingSeq!.ResultNumber!.Value,
                MathOnlyText(paragraph),
                paragraphText,
                qualifyingSeq.Instruction,
                previous,
                next));
        }
    }

    var audit = new EquationSeqAuditReport(
        docxPath,
        DateTime.UtcNow,
        mathParagraphs.Count,
        allSeqFields.Count,
        allSeqFields.Count(s => s.IsCaptionSeq),
        dissertationEquations.Count,
        dissertationEquations.Select(e => e.Number).Distinct().Count(),
        mathParagraphs,
        allSeqFields,
        dissertationEquations);

    var json = JsonSerializer.Serialize(audit, new JsonSerializerOptions { WriteIndented = true });
    if (options.TryGetValue("out", out var outPathValue))
    {
        var outPath = Path.GetFullPath(outPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
        File.WriteAllText(outPath, json, Encoding.UTF8);
        Console.WriteLine($"Wrote equations audit: {outPath}");
    }
    else
    {
        Console.WriteLine(json);
    }

    return 0;
}

static int Validate(string docxPath)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart;
    var trackRevisions = settingsPart?.Settings?.Elements<TrackRevisions>().Any() == true;
    var paragraphs = GetParagraphs(doc);
    var insertedRuns = doc.MainDocumentPart?.Document.Descendants<InsertedRun>().Count() ?? 0;
    var deletedRuns = doc.MainDocumentPart?.Document.Descendants<DeletedRun>().Count() ?? 0;

    Console.WriteLine($"Paragraphs: {paragraphs.Count}");
    Console.WriteLine($"TrackRevisions: {trackRevisions}");
    Console.WriteLine($"InsertedRuns: {insertedRuns}");
    Console.WriteLine($"DeletedRuns: {deletedRuns}");
    var validationErrors = new OpenXmlValidator().Validate(doc).ToList();
    var toleratedErrors = validationErrors.Where(IsToleratedWordTableMarkupValidationError).ToList();
    var actionableErrors = validationErrors.Where(e => !IsToleratedWordTableMarkupValidationError(e)).ToList();
    Console.WriteLine($"OpenXmlValidationErrors: {validationErrors.Count}");
    Console.WriteLine($"OpenXmlValidationErrorsTolerated: {toleratedErrors.Count}");
    Console.WriteLine($"OpenXmlValidationErrorsActionable: {actionableErrors.Count}");
    foreach (var error in actionableErrors.Take(500))
    {
        Console.WriteLine($"  {error.Path?.XPath}: {error.Description}");
    }
    if (actionableErrors.Count == 0 && toleratedErrors.Count > 0)
    {
        Console.WriteLine("  only tolerated Word table markup errors detected (w:tblLook / w:cnfStyle reintroduced by Word save)");
    }
    return 0;
}

static bool IsToleratedWordTableMarkupValidationError(ValidationErrorInfo error)
{
    var path = error.Path?.XPath ?? "";
    var description = error.Description ?? "";
    var isTableLookPath = path.Contains("/w:tblPr[1]/w:tblLook[1]", StringComparison.Ordinal);
    var isCnfStylePath = path.Contains("/w:trPr[1]/w:cnfStyle[1]", StringComparison.Ordinal);
    if (!isTableLookPath && !isCnfStylePath)
    {
        return false;
    }

    if (!description.Contains("attribute is not declared", StringComparison.OrdinalIgnoreCase))
    {
        return false;
    }

    return true;
}

static TableAuditEntry BuildTableAudit(
    int ordinal,
    int blockIndex,
    Table table,
    IReadOnlyList<ParagraphAuditEntry> paragraphs,
    IReadOnlyList<OpenXmlElement> blocks,
    IReadOnlyDictionary<int, int> paragraphIndexByBlock)
{
    var previous = PreviousParagraph(paragraphs, blockIndex);
    var next = NextParagraph(paragraphs, blockIndex);
    var next2 = NextParagraph(paragraphs, blockIndex, 2);
    var rows = table.Elements<TableRow>().ToList();
    var firstRow = rows.FirstOrDefault();
    var columnCount = firstRow?.Elements<TableCell>().Count() ?? 0;
    var tableProperties = table.GetFirstChild<TableProperties>();
    var styleId = tableProperties?.TableStyle?.Val?.Value ?? "";
    var width = tableProperties?.TableWidth?.Width?.Value ?? "";
    var borders = tableProperties?.TableBorders is not null;
    var titleCandidate = previous?.Text ?? "";
    var sourceCandidate = new[] { next?.Text ?? "", next2?.Text ?? "" }.FirstOrDefault(IsSourceLine) ?? "";

    return new TableAuditEntry(
        ordinal,
        blockIndex,
        rows.Count,
        columnCount,
        styleId,
        width,
        borders,
        previous,
        next,
        next2,
        LooksLikeTableTitle(titleCandidate),
        IsSourceLine(sourceCandidate),
        sourceCandidate,
        FirstNonEmptyTableText(table));
}

static FigureAuditEntry BuildFigureAudit(
    int ordinal,
    int blockIndex,
    Paragraph paragraph,
    Drawing drawing,
    IReadOnlyList<ParagraphAuditEntry> paragraphs,
    IReadOnlyDictionary<int, int> paragraphIndexByBlock)
{
    var paragraphIndex = paragraphIndexByBlock.TryGetValue(blockIndex, out var pIndex) ? pIndex : -1;
    var previous = PreviousParagraph(paragraphs, blockIndex);
    var next = NextParagraph(paragraphs, blockIndex);
    var next2 = NextParagraph(paragraphs, blockIndex, 2);
    var titleCandidate = previous?.Text ?? "";
    var sourceCandidate = new[] { next?.Text ?? "", next2?.Text ?? "" }.FirstOrDefault(IsSourceLine) ?? "";
    var docProperties = drawing.Descendants<DocProperties>().FirstOrDefault();

    return new FigureAuditEntry(
        ordinal,
        blockIndex,
        paragraphIndex,
        ParagraphText(paragraph),
        docProperties?.Name?.Value ?? "",
        docProperties?.Description?.Value ?? "",
        previous,
        next,
        next2,
        LooksLikeFigureTitle(titleCandidate),
        IsSourceLine(sourceCandidate),
        sourceCandidate);
}

static TableLayoutEntry BuildTableLayoutEntry(
    WordprocessingDocument doc,
    int ordinal,
    int blockIndex,
    Table table,
    IReadOnlyList<ParagraphAuditEntry> paragraphs)
{
    var previous = PreviousParagraph(paragraphs, blockIndex);
    var next = NextParagraph(paragraphs, blockIndex);
    var next2 = NextParagraph(paragraphs, blockIndex, 2);
    var rows = table.Elements<TableRow>().ToList();
    var cells = table.Descendants<TableCell>().ToList();
    var tableProperties = table.GetFirstChild<TableProperties>();
    var borders = tableProperties?.TableBorders;
    var styleId = tableProperties?.TableStyle?.Val?.Value ?? "";
    var style = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
        .Elements<Style>()
        .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
    var styleBorders = style?.StyleTableProperties?.TableBorders;
    var effectiveBorders = borders ?? styleBorders;
    var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Select(c => c.Width?.Value ?? "").ToList() ?? [];
    var cellParagraphs = cells.SelectMany(c => c.Elements<Paragraph>()).ToList();
    var cellStyles = cellParagraphs
        .Select(ParagraphStyleId)
        .GroupBy(s => string.IsNullOrWhiteSpace(s) ? "(sem estilo)" : s)
        .OrderByDescending(g => g.Count())
        .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
    var directFontSizes = cellParagraphs
        .SelectMany(p => p.Descendants<Run>())
        .Select(r => r.RunProperties?.FontSize?.Val?.Value ?? "")
        .Where(v => !string.IsNullOrWhiteSpace(v))
        .GroupBy(v => v)
        .OrderByDescending(g => g.Count())
        .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
    var justificationCounts = cellParagraphs
        .Select(p => p.ParagraphProperties?.Justification?.Val?.Value.ToString() ?? "(herdado/vazio)")
        .GroupBy(v => v)
        .OrderByDescending(g => g.Count())
        .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
    var cellVerticalAlignments = cells
        .Select(c => c.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value.ToString() ?? "(herdado/vazio)")
        .GroupBy(v => v)
        .OrderByDescending(g => g.Count())
        .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
    var hasMergedCells = cells.Any(c => c.TableCellProperties?.GridSpan is not null || c.TableCellProperties?.VerticalMerge is not null);
    var firstCellText = FirstNonEmptyTableText(table);
    var sourceCandidate = new[] { next?.Text ?? "", next2?.Text ?? "" }.FirstOrDefault(IsSourceLine) ?? "";
    var issues = new List<string>();
    var isAcademic = previous?.StyleId.Equals("Tabela", StringComparison.OrdinalIgnoreCase) == true
        || LooksLikeTableTitle(previous?.Text ?? "");

    if (isAcademic)
    {
        if (previous?.StyleId.Equals("Tabela", StringComparison.OrdinalIgnoreCase) != true)
        {
            issues.Add("titulo acima nao usa estilo Tabela");
        }

        if (!IsSourceLine(sourceCandidate))
        {
            issues.Add("fonte abaixo nao detectada nos dois paragrafos seguintes");
        }

        if (string.IsNullOrWhiteSpace(tableProperties?.TableWidth?.Width?.Value))
        {
            issues.Add("largura da tabela nao declarada explicitamente");
        }

        if (effectiveBorders is null)
        {
            issues.Add("bordas da tabela nao declaradas no nivel da tabela nem no estilo aplicado");
        }
        else if (HasVerticalBorders(effectiveBorders))
        {
            issues.Add("ha bordas verticais declaradas; IBGE geralmente favorece linhas horizontais e evita fechamento lateral em tabelas estatisticas");
        }

        if (cellStyles.ContainsKey("(sem estilo)"))
        {
            issues.Add("ha paragrafos em celulas sem estilo de paragrafo explicito");
        }
    }

    return new TableLayoutEntry(
        ordinal,
        blockIndex,
        isAcademic,
        rows.Count,
        rows.FirstOrDefault()?.Elements<TableCell>().Count() ?? 0,
        cells.Count,
        styleId,
        tableProperties?.TableWidth?.Type?.Value.ToString() ?? "",
        tableProperties?.TableWidth?.Width?.Value ?? "",
        tableProperties?.TableJustification?.Val?.Value.ToString() ?? "",
        tableProperties?.TableLayout?.Type?.Value.ToString() ?? "",
        BorderSummary(effectiveBorders),
        HasVerticalBorders(effectiveBorders),
        gridColumns,
        hasMergedCells,
        cellStyles,
        directFontSizes,
        justificationCounts,
        cellVerticalAlignments,
        previous,
        next,
        next2,
        sourceCandidate,
        firstCellText,
        issues);
}

static FigureLayoutEntry BuildFigureLayoutEntry(
    int ordinal,
    int blockIndex,
    Paragraph paragraph,
    Drawing drawing,
    IReadOnlyList<ParagraphAuditEntry> paragraphs,
    IReadOnlyDictionary<int, int> paragraphIndexByBlock)
{
    var paragraphIndex = paragraphIndexByBlock.TryGetValue(blockIndex, out var pIndex) ? pIndex : -1;
    var next = NextParagraph(paragraphs, blockIndex);
    var next2 = NextParagraph(paragraphs, blockIndex, 2);
    var sourceCandidate = new[] { next?.Text ?? "", next2?.Text ?? "" }.FirstOrDefault(IsSourceLine) ?? "";
    var docProperties = drawing.Descendants<DocProperties>().FirstOrDefault();
    var inline = drawing.Inline;
    var anchor = drawing.Anchor;
    var extent = inline?.Extent ?? anchor?.Extent;
    var widthCm = extent?.Cx is null ? 0 : EmuToCm(extent.Cx.Value);
    var heightCm = extent?.Cy is null ? 0 : EmuToCm(extent.Cy.Value);
    var wrap = anchor?.GetFirstChild<WrapNone>() is not null ? "none"
        : anchor?.GetFirstChild<WrapSquare>() is not null ? "square"
        : anchor?.GetFirstChild<WrapTight>() is not null ? "tight"
        : anchor?.GetFirstChild<WrapTopBottom>() is not null ? "topBottom"
        : anchor?.GetFirstChild<WrapThrough>() is not null ? "through"
        : inline is not null ? "inline"
        : "";
    var hasOutline = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().Any(o => !o.Elements<DocumentFormat.OpenXml.Drawing.NoFill>().Any());
    var isAcademic = string.Equals(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, "Figura", StringComparison.OrdinalIgnoreCase);
    var issues = new List<string>();
    if (isAcademic)
    {
        if (!string.Equals(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, "Figura", StringComparison.OrdinalIgnoreCase))
        {
            issues.Add("figura nao esta ancorada em paragrafo de estilo Figura");
        }

        if (!IsSourceLine(sourceCandidate))
        {
            issues.Add("fonte abaixo nao detectada nos dois paragrafos seguintes");
        }

        if (widthCm > 16.2)
        {
            issues.Add("largura estimada excede area util comum de pagina A4 com margens de 3 cm");
        }

        if (anchor is not null && wrap is not "topBottom" and not "none")
        {
            issues.Add($"figura flutuante com wrap={wrap}; para dissertaÃ§Ã£o costuma ser mais seguro usar inline ou quebra superior/inferior");
        }

        if (hasOutline)
        {
            issues.Add("ha contorno/linha no desenho; normas nao exigem moldura em figuras, usar apenas quando fizer parte da imagem/quadro");
        }
    }

    return new FigureLayoutEntry(
        ordinal,
        blockIndex,
        paragraphIndex,
        isAcademic,
        ParagraphText(paragraph),
        ParagraphStyleId(paragraph),
        docProperties?.Name?.Value ?? "",
        docProperties?.Description?.Value ?? "",
        inline is not null ? "inline" : anchor is not null ? "anchor" : "",
        wrap,
        widthCm,
        heightCm,
        hasOutline,
        anchor?.SimplePosition is not null,
        anchor?.RelativeHeight?.Value ?? 0,
        sourceCandidate,
        next,
        next2,
        issues);
}

static EquationAuditEntry BuildEquationAudit(
    int ordinal,
    int blockIndex,
    Paragraph paragraph,
    IReadOnlyList<ParagraphAuditEntry> paragraphs,
    IReadOnlyDictionary<int, int> paragraphIndexByBlock)
{
    var paragraphIndex = paragraphIndexByBlock.TryGetValue(blockIndex, out var pIndex) ? pIndex : -1;
    var text = ParagraphText(paragraph);
    var previous = PreviousParagraph(paragraphs, blockIndex);
    var next = NextParagraph(paragraphs, blockIndex);
    var hasNumber = Regex.IsMatch(text, @"\(\s*\d+\s*\)\s*$");

    return new EquationAuditEntry(
        ordinal,
        blockIndex,
        paragraphIndex,
        text,
        MathOnlyText(paragraph),
        ParagraphStyleId(paragraph),
        hasNumber,
        previous,
        next);
}

static int ListRevisions(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var paragraphs = GetParagraphs(doc);
    var authorFilter = options.TryGetValue("author", out var authorValue)
        ? Normalize(authorValue)
        : null;

    foreach (var entry in paragraphs)
    {
        var revisions = entry.Paragraph
            .Descendants<OpenXmlElement>()
            .Where(e => e is InsertedRun or DeletedRun)
            .Select(e => new
            {
                Kind = e is InsertedRun ? "ins" : "del",
                Author = GetRevisionAuthor(e),
                Date = GetRevisionDate(e),
                Text = ElementText(e)
            })
            .Where(r => !string.IsNullOrWhiteSpace(r.Text))
            .Where(r => string.IsNullOrEmpty(authorFilter) || Normalize(r.Author).Contains(authorFilter, StringComparison.Ordinal))
            .ToList();

        foreach (var revision in revisions)
        {
            Console.WriteLine($"P[{entry.Index}] {revision.Kind} author=\"{revision.Author}\" date=\"{revision.Date}\" text=\"{revision.Text}\"");
            Console.WriteLine($"    paragraph=\"{entry.Text}\"");
        }
    }

    return 0;
}

static int ListComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var commentsPart = doc.MainDocumentPart?.WordprocessingCommentsPart;
    if (commentsPart?.Comments is null)
    {
        return 0;
    }

    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var paragraphs = body.Elements<Paragraph>()
        .Select((paragraph, index) => new ParagraphEntry(paragraph, index, ParagraphAuditDisplayText(paragraph)))
        .Where(entry => !string.IsNullOrWhiteSpace(entry.Text))
        .ToList();
    var authorFilter = options.TryGetValue("author", out var authorValue)
        ? Normalize(authorValue)
        : null;

    var anchorTextById = new Dictionary<string, string>();
    foreach (var paragraph in paragraphs)
    {
        foreach (var rangeStart in paragraph.Paragraph.Descendants<CommentRangeStart>())
        {
            if (rangeStart.Id?.Value is not { } id)
            {
                continue;
            }

            anchorTextById[id] = paragraph.Text;
        }
    }

    var commentExByParaId = doc.MainDocumentPart?.WordprocessingCommentsExPart?.CommentsEx?
        .Elements<W15.CommentEx>()
        .Where(x => x.ParaId?.Value is not null)
        .ToDictionary(x => x.ParaId!.Value!, StringComparer.Ordinal)
        ?? new Dictionary<string, W15.CommentEx>(StringComparer.Ordinal);
    var commentIdByParaId = commentsPart.Comments.Elements<Comment>()
        .Select(comment => new
        {
            CommentId = comment.Id?.Value,
            ParaId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.CommentId) && !string.IsNullOrWhiteSpace(x.ParaId))
        .ToDictionary(x => x.ParaId!, x => x.CommentId!, StringComparer.Ordinal);

    var entries = new List<CommentListEntry>();
    foreach (var comment in commentsPart.Comments.Elements<Comment>())
    {
        var author = comment.Author?.Value ?? "";
        if (!string.IsNullOrEmpty(authorFilter) && !Normalize(author).Contains(authorFilter, StringComparison.Ordinal))
        {
            continue;
        }

        var id = comment.Id?.Value ?? "";
        anchorTextById.TryGetValue(id, out var anchorText);
        var paraId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
        string? parentCommentId = null;
        if (!string.IsNullOrWhiteSpace(paraId)
            && commentExByParaId.TryGetValue(paraId, out var commentEx)
            && commentEx.ParaIdParent?.Value is { } parentParaId
            && commentIdByParaId.TryGetValue(parentParaId, out var parentId))
        {
            parentCommentId = parentId;
        }

        var parentSuffix = string.IsNullOrWhiteSpace(parentCommentId)
            ? ""
            : $" parentCommentId=\"{parentCommentId}\"";
        entries.Add(new CommentListEntry(
            id,
            author,
            comment.Date?.Value.ToString(CultureInfo.CurrentCulture) ?? "",
            parentCommentId,
            ElementText(comment),
            anchorText ?? ""));
    }

    var formatValue = options.TryGetValue("format", out var requestedFormat)
        ? requestedFormat
        : "auto";
    var format = ResolveCommentsFormat(formatValue);

    if (format is "json")
    {
        WriteCommentsJson(entries);
        return 0;
    }

    if (format is "raw" or "legacy" or "text")
    {
        foreach (var entry in entries)
        {
            var parentSuffix = string.IsNullOrWhiteSpace(entry.ParentCommentId)
                ? ""
                : $" parentCommentId=\"{entry.ParentCommentId}\"";
            Console.WriteLine($"comment id=\"{entry.Id}\" author=\"{entry.Author}\" date=\"{entry.Date}\"{parentSuffix} text=\"{entry.Text}\"");
            if (!string.IsNullOrWhiteSpace(entry.AnchorText))
            {
                Console.WriteLine($"    anchor=\"{entry.AnchorText}\"");
            }
        }

        return 0;
    }

    if (format is "table" or "console" or "console-table")
    {
        WriteTabularOutput(BuildCommentsOutputTable(entries, includeIndex: true), "table");
        return 0;
    }

    if (format is not "md" and not "markdown")
    {
        Console.Error.WriteLine($"Unsupported comments format: {formatValue}");
        return 4;
    }

    WriteTabularOutput(BuildCommentsOutputTable(entries, includeIndex: false), "markdown");
    return 0;
}

static string ResolveDefaultCommentsFormat() =>
    DetectCodexSurface() switch
    {
        "cli" => "table",
        "app" => "markdown",
        _ => "markdown"
    };

static string ResolveCommentsFormat(string? formatValue)
{
    var format = Normalize(formatValue ?? "auto");
    return format is "" or "auto"
        ? ResolveDefaultCommentsFormat()
        : format;
}

static string DetectCodexSurface()
{
    if (string.Equals(Environment.GetEnvironmentVariable("CODEX_MANAGED_BY_NPM"), "1", StringComparison.Ordinal))
    {
        return "cli";
    }

    return "app";
}

static int ListCommentAnchors(string docxPath, IReadOnlyDictionary<string, string> options)
{
    using var stream = OpenSharedRead(docxPath);
    using var doc = WordprocessingDocument.Open(stream, false);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var paragraphs = body.Elements<Paragraph>()
        .Select((paragraph, index) => new ParagraphEntry(paragraph, index, ParagraphAuditDisplayText(paragraph)))
        .Where(entry => !string.IsNullOrWhiteSpace(entry.Text))
        .ToList();

    foreach (var paragraph in paragraphs)
    {
        var ids = paragraph.Paragraph.Descendants<CommentRangeStart>()
            .Select(c => c.Id?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        if (ids.Count == 0)
        {
            continue;
        }

        Console.WriteLine($"P[{paragraph.Index}] ids=\"{string.Join(",", ids)}\" text=\"{paragraph.Text}\"");
    }

    return 0;
}

static int InsertTracked(string docxPath, IReadOnlyDictionary<string, string> options)
{
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var json = File.ReadAllText(planPath);
    var plan = JsonSerializer.Deserialize<List<InsertionSpec>>(json, new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null || plan.Count == 0)
    {
        Console.Error.WriteLine("Plan is empty.");
        return 6;
    }

    var author = MutationAuthorResolver.Resolve(docxPath, options);

    FileStream? lockStream = null;
    string? lockPath = null;
    if (options.TryGetValue("lock", out var lockPathValue) && !string.IsNullOrWhiteSpace(lockPathValue))
    {
        lockPath = Path.GetFullPath(lockPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
        try
        {
            lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            lockStream.SetLength(0);
            lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils insert-tracked {DateTime.UtcNow:O}\n"));
            lockStream.Flush();
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"Unable to acquire exclusive lock: {ex.Message}");
            return 9;
        }
    }

    using (lockStream)
    {
        WordprocessingDocument doc;
        try
        {
            doc = WordprocessingDocument.Open(docxPath, true);
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
            Console.Error.WriteLine("Close the document in Word or any other process holding the file, then retry.");
            return 8;
        }

        var applied = new List<string>();
        var skipped = new List<string>();
        using (doc)
        {
            EnsureTrackRevisions(doc);

            var metadata = new RevisionMetadata(author, DateTime.UtcNow);
            foreach (var spec in plan)
            {
                var paragraphs = GetParagraphs(doc);
                if (ContentAlreadyPresent(paragraphs, spec.Content))
                {
                    Console.WriteLine($"SKIP {spec.Id}: content already present.");
                    skipped.Add(spec.Id);
                    continue;
                }

                var after = FindUniqueParagraph(paragraphs, spec.AfterPrefix, $"{spec.Id}/after");
                var before = FindUniqueParagraph(paragraphs, spec.BeforePrefix, $"{spec.Id}/before");
                if (after.Index >= before.Index)
                {
                    Console.Error.WriteLine($"Anchor order mismatch for {spec.Id}: afterIndex={after.Index}, beforeIndex={before.Index}");
                    return 7;
                }

                var template = spec.StyleSource.Equals("before", StringComparison.OrdinalIgnoreCase)
                    ? before.Paragraph
                    : after.Paragraph;

                var paragraph = CreateInsertedParagraph(template, spec.Content, metadata);
                after.Paragraph.InsertAfterSelf(paragraph);
                applied.Add(spec.Id);
                Console.WriteLine($"APPLY {spec.Id}: inserted after P[{after.Index}] before P[{before.Index}]");
            }

            SaveMainDocumentWithValidationRepair(doc, applied);
        }

        if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
        {
            var reportPath = Path.GetFullPath(reportPathValue);
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
            var builder = new StringBuilder();
            builder.AppendLine("# Insert Tracked Report");
            builder.AppendLine();
            builder.AppendLine($"- DOCX: `{docxPath}`");
            builder.AppendLine($"- Plano: `{planPath}`");
            builder.AppendLine($"- Autor: `{author}`");
            builder.AppendLine($"- Lock: `{lockPath ?? "not requested"}`");
            builder.AppendLine($"- Aplicados: {applied.Count}");
            builder.AppendLine($"- Ignorados: {skipped.Count}");
            if (applied.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("## Applied");
                foreach (var id in applied)
                {
                    builder.AppendLine($"- `{id}`");
                }
            }
            if (skipped.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("## Skipped");
                foreach (var id in skipped)
                {
                    builder.AppendLine($"- `{id}`");
                }
            }
            File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
        }
    }
    return 0;
}

static int InsertBlocks(string docxPath, IReadOnlyDictionary<string, string> options)
{
    return ApplyBlockMutationPlan(docxPath, options, "insert-blocks", replaceBetweenAnchors: false);
}

static int ReplaceBlocks(string docxPath, IReadOnlyDictionary<string, string> options)
{
    return ApplyBlockMutationPlan(docxPath, options, "replace-blocks", replaceBetweenAnchors: true);
}

static int ApplyBlockMutationPlan(
    string docxPath,
    IReadOnlyDictionary<string, string> options,
    string commandName,
    bool replaceBetweenAnchors)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var planJson = File.ReadAllText(planPath, Encoding.UTF8);
    var validation = PlanContractSupport.ValidateInsertBlocksPlan(planJson);
    if (!validation.IsValid)
    {
        PrintPlanValidationErrors(validation.Errors);
        return 6;
    }

    var plan = JsonSerializer.Deserialize<BlockInsertionPlan>(planJson, new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read block insertion plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils {commandName} {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var metadata = new RevisionMetadata(author, DateTime.UtcNow);

    foreach (var spec in plan.Blocks)
    {
        var paragraphs = GetParagraphs(doc);
        var presenceParagraphs = GetAllParagraphEntries(doc);
        if (!replaceBetweenAnchors && !string.IsNullOrWhiteSpace(spec.UniqueText) && ContentAlreadyPresent(presenceParagraphs, spec.UniqueText))
        {
            skipped.Add($"{spec.Id}: unique text already present");
            continue;
        }

        if (!replaceBetweenAnchors && BlockTableAlreadyPresent(doc, spec))
        {
            skipped.Add($"{spec.Id}: equivalent table already present");
            continue;
        }

        var after = FindUniqueParagraph(paragraphs, spec.AfterPrefix, $"{spec.Id}/after");
        var before = FindUniqueParagraph(paragraphs, spec.BeforePrefix, $"{spec.Id}/before");
        if (after.Index >= before.Index)
        {
            Console.Error.WriteLine($"Anchor order mismatch for {spec.Id}: afterIndex={after.Index}, beforeIndex={before.Index}");
            return 7;
        }

        var template = spec.StyleSource.Equals("before", StringComparison.OrdinalIgnoreCase)
            ? before.Paragraph
            : after.Paragraph;

        if (replaceBetweenAnchors)
        {
            RemoveElementsBetweenAnchors(after.Paragraph, before.Paragraph);
        }

        OpenXmlElement insertionPoint = after.Paragraph;
        foreach (var item in spec.Items)
        {
            OpenXmlElement block = item.Kind.Equals("table", StringComparison.OrdinalIgnoreCase)
                ? CreateTrackedTable(item, metadata)
                : CreateInsertedParagraphWithStyle(template, item.Text ?? item.Latex ?? "", item.StyleId, metadata);
            insertionPoint.InsertAfterSelf(block);
            insertionPoint = block;
        }
        applied.Add(spec.Id);
    }

    SaveMainDocumentWithValidationRepair(doc, applied);

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine(replaceBetweenAnchors ? "# Replace Blocks Report" : "# Insert Blocks Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int EditParagraphs(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<ParagraphEditPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read paragraph edit plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils edit-paragraphs {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var metadata = new RevisionMetadata(author, DateTime.UtcNow);

    foreach (var edit in plan.Edits)
    {
        var paragraphs = GetAllParagraphs(doc)
            .Where(p => NormalizedStartsWith(ParagraphText(p), edit.ParagraphPrefix))
            .ToList();

        if (edit.Occurrence is int occurrence && occurrence > 0)
        {
            if (paragraphs.Count < occurrence)
            {
                skipped.Add($"{edit.Id}: paragraph anchor count={paragraphs.Count}");
                continue;
            }

            paragraphs = [paragraphs[occurrence - 1]];
        }

        if (paragraphs.Count != 1)
        {
            skipped.Add($"{edit.Id}: paragraph anchor count={paragraphs.Count}");
            continue;
        }

        var paragraph = paragraphs[0];
        if (!string.IsNullOrWhiteSpace(edit.StyleId))
        {
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            var current = paragraph.ParagraphProperties.ParagraphStyleId;
            if (current is null)
            {
                paragraph.ParagraphProperties.PrependChild(new ParagraphStyleId { Val = edit.StyleId });
            }
            else
            {
                current.Val = edit.StyleId;
            }
        }

        if (edit.ReplacementText is not null)
        {
            var originalText = ParagraphText(paragraph);
            ReplaceParagraphTracked(paragraph, originalText, [CrossrefRunPart.TextPart(edit.ReplacementText)], metadata);
        }

        applied.Add(edit.Id);
    }

    SaveMainDocumentWithValidationRepair(doc, applied);

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Edit Paragraphs Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int StyleRunningText(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var targetStyleId = options.TryGetValue("style", out var styleValue) && !string.IsNullOrWhiteSpace(styleValue)
        ? styleValue
        : "Normal";
    var targetFont = options.TryGetValue("font", out var fontValue) && !string.IsNullOrWhiteSpace(fontValue)
        ? fontValue
        : "Times New Roman";
    var targetHalfPoints = options.TryGetValue("size-half-points", out var halfPointsValue) && !string.IsNullOrWhiteSpace(halfPointsValue)
        ? halfPointsValue
        : "24";

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportEntries = new List<RunningTextStyleReportEntry>();
    var changed = 0;
    var unchanged = 0;
    var skipped = 0;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils style-running-text {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var metadata = new RevisionMetadata(author, DateTime.UtcNow);
    var paragraphs = GetParagraphs(doc);
    var inReferences = false;
    var inDeclarations = false;

    foreach (var entry in paragraphs)
    {
        var text = entry.Text;
        if (IsReferencesHeading(text))
        {
            inReferences = true;
        }

        if (IsDeclarationsHeading(text))
        {
            inDeclarations = true;
        }

        var eligibility = entry.Index == 0
            ? new RunningTextEligibilityResult(false, "document title")
            : RunningTextEligibility(entry.Paragraph, text, inReferences, inDeclarations);
        if (!eligibility.IsEligible)
        {
            var reverted = RevertOwnFormattingRevisions(entry.Paragraph, author);
            skipped++;
            reportEntries.Add(new RunningTextStyleReportEntry(entry.Index, reverted ? "REVERT-SKIP" : "SKIP", eligibility.Reason, text));
            continue;
        }

        var paragraphChanged = ApplyNormalTimesStyle(entry.Paragraph, targetStyleId, targetFont, targetHalfPoints, metadata);
        if (paragraphChanged)
        {
            changed++;
            reportEntries.Add(new RunningTextStyleReportEntry(entry.Index, "CHANGE", "running text", text));
        }
        else
        {
            unchanged++;
            reportEntries.Add(new RunningTextStyleReportEntry(entry.Index, "UNCHANGED", "already matched", text));
        }
    }

    NormalizeFormattingRevisionMarkup(doc);
    SaveMainDocumentWithValidationRepair(doc);

    var validationErrors = new OpenXmlValidator()
        .Validate(doc)
        .Take(20)
        .Select(e => e.Description)
        .ToList();
    var trackRevisions = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.Elements<TrackRevisions>().Any() == true;

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Style Running Text Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Estilo alvo: `{targetStyleId}`");
        builder.AppendLine($"- Fonte alvo: `{targetFont}`");
        builder.AppendLine($"- Tamanho alvo: `{targetHalfPoints}` half-points (12 pt)");
        builder.AppendLine($"- Paragrafos alterados: {changed}");
        builder.AppendLine($"- Paragrafos elegiveis ja conformes: {unchanged}");
        builder.AppendLine($"- Paragrafos ignorados: {skipped}");
        builder.AppendLine($"- TrackRevisions ativo: {trackRevisions}");
        builder.AppendLine($"- Erros Open XML detectados: {validationErrors.Count}");
        foreach (var error in validationErrors)
        {
            builder.AppendLine($"  - {error}");
        }
        builder.AppendLine();
        builder.AppendLine("## Detalhe");
        foreach (var item in reportEntries)
        {
            builder.AppendLine($"- {item.Status} P[{item.Index}] {item.Reason}: {TruncateForReport(item.Text)}");
        }
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    Console.WriteLine($"CHANGED {changed}");
    Console.WriteLine($"UNCHANGED {unchanged}");
    Console.WriteLine($"SKIPPED {skipped}");
    Console.WriteLine($"TRACK_REVISIONS {trackRevisions}");
    Console.WriteLine($"VALIDATION_ERRORS {validationErrors.Count}");
    return validationErrors.Count == 0 ? 0 : 8;
}

static IReadOnlyList<string> RepairValidationBeforeSave(WordprocessingDocument doc)
{
    var changes = new List<string>();
    EnsureTrackRevisions(doc);

    var nextId = 1;
    foreach (var inserted in doc.MainDocumentPart!.Document.Descendants<InsertedRun>())
    {
        inserted.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var deleted in doc.MainDocumentPart!.Document.Descendants<DeletedRun>())
    {
        deleted.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<RunPropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<ParagraphPropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<SectionPropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<TablePropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<TableRowPropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<TableCellPropertiesChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    foreach (var change in doc.MainDocumentPart!.Document.Descendants<TablePropertyExceptionsChange>())
    {
        change.Id = nextId.ToString(CultureInfo.InvariantCulture);
        nextId++;
    }
    changes.Add($"revision_ids_reassigned={nextId - 1}");

    var tableLookCount = 0;
    foreach (var tableLook in doc.MainDocumentPart!.Document.Descendants<TableLook>().ToList())
    {
        tableLook.Remove();
        tableLookCount++;
    }
    changes.Add($"table_look_removed={tableLookCount}");

    var cnfCount = 0;
    foreach (var cnf in doc.MainDocumentPart!.Document.Descendants<ConditionalFormatStyle>().ToList())
    {
        cnf.Remove();
        cnfCount++;
    }
    changes.Add($"conditional_format_style_removed={cnfCount}");

    var ligatureCount = 0;
    if (doc.MainDocumentPart?.StyleDefinitionsPart is { } stylesPart)
    {
        ligatureCount = RemoveWord2010LigaturesFromStylesPart(stylesPart);
    }
    changes.Add($"w14_ligatures_removed={ligatureCount}");

    return changes;
}

static void SaveMainDocumentWithValidationRepair(WordprocessingDocument doc, List<string>? applied = null)
{
    _ = RepairValidationBeforeSave(doc);

    var mainPart = doc.MainDocumentPart
        ?? throw new InvalidOperationException("DOCX sem MainDocumentPart.");
    var mainDocument = mainPart.Document
        ?? throw new InvalidOperationException("DOCX sem documento principal.");
    mainDocument.Save();
}

static int ApplyCrossrefs(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CrossrefPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read crossref plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils apply-crossrefs {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        EnsureUpdateFieldsOnOpen(doc);
        var metadata = new RevisionMetadata(author, DateTime.UtcNow);
        var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
        var bookmarkId = NextBookmarkId(doc);

        foreach (var caption in plan.Captions)
        {
            var paragraphs = GetAllParagraphs(doc);
            var target = FindCaptionParagraph(paragraphs, caption);
            if (target is null)
            {
                skipped.Add($"{caption.Id}: caption anchor not found or ambiguous ({caption.AnchorText})");
                continue;
            }

            if (BookmarkExists(doc, caption.Bookmark))
            {
                skipped.Add($"{caption.Id}: bookmark already exists ({caption.Bookmark})");
                continue;
            }

            var currentText = ParagraphText(target);
            if (Regex.IsMatch(currentText, $"^{Regex.Escape(caption.Label)}\\s+\\d+", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                skipped.Add($"{caption.Id}: caption already appears numbered");
                continue;
            }

            InsertCaptionPrefix(target, caption, metadata, ref bookmarkId);
            applied.Add($"{caption.Id}: caption numbered/bookmarked as {caption.Label} {caption.Number}");
        }

        var bookmarkNames = doc.MainDocumentPart!.Document.Descendants<BookmarkStart>()
            .Select(b => b.Name?.Value ?? "")
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .ToHashSet(StringComparer.Ordinal);

        foreach (var replacement in plan.Replacements)
        {
            var paragraphs = GetAllParagraphs(doc);
            var matches = paragraphs
                .Where(p => NormalizedStartsWith(ParagraphText(p), replacement.ParagraphPrefix))
                .ToList();
            if (matches.Count != 1)
            {
                skipped.Add($"{replacement.Id}: paragraph anchor count={matches.Count}");
                continue;
            }

            var paragraph = matches[0];
            var originalText = ParagraphText(paragraph);
            var planned = BuildPlannedParagraph(originalText, replacement, bookmarkNames, skipped);
            if (planned is null)
            {
                continue;
            }

            ReplaceParagraphTracked(paragraph, originalText, planned, metadata);
            applied.Add($"{replacement.Id}: converted {replacement.Replacements.Count} call(s)");
        }

        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildCrossrefReport(docxPath, plan, lockPath, author, applied, skipped), Encoding.UTF8);
    }

    foreach (var item in applied)
    {
        Console.WriteLine($"APPLY {item}");
    }

    foreach (var item in skipped)
    {
        Console.WriteLine($"SKIP {item}");
    }

    return 0;
}

static int RepairStyleCaptions(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CrossrefPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });
    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read crossref plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils repair-style-captions {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        EnsureUpdateFieldsOnOpen(doc);
        var bookmarkId = NextBookmarkId(doc);

        foreach (var caption in plan.Captions)
        {
            var paragraphs = GetAllParagraphs(doc);
            var target = FindCaptionParagraphAllowingManualPrefix(paragraphs, caption);
            if (target is null)
            {
                skipped.Add($"{caption.Id}: caption anchor not found or ambiguous ({caption.AnchorText})");
                continue;
            }

            var removedPrefix = RemoveLeadingManualCaptionPrefix(target, caption);
            RemoveBookmarkByName(doc, caption.Bookmark);
            AddParagraphBookmark(target, caption.Bookmark, ref bookmarkId);

            var numbering = EffectiveNumberingSummary(doc, target);
            if (string.IsNullOrWhiteSpace(numbering))
            {
                skipped.Add($"{caption.Id}: style numbering not found after repair");
                continue;
            }

            applied.Add($"{caption.Id}: style-numbered caption; removedManualPrefix={removedPrefix}; bookmark={caption.Bookmark}; {numbering}");
        }

        var bookmarkNames = doc.MainDocumentPart!.Document.Descendants<BookmarkStart>()
            .Select(b => b.Name?.Value ?? "")
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Where(name => name.StartsWith("xref_fig_", StringComparison.Ordinal) || name.StartsWith("xref_tab_", StringComparison.Ordinal))
            .ToHashSet(StringComparer.Ordinal);
        var refFieldsUpdated = 0;
        foreach (var fieldCode in doc.MainDocumentPart!.Document.Descendants<FieldCode>())
        {
            var instruction = fieldCode.Text ?? "";
            foreach (var bookmark in bookmarkNames)
            {
                if (!Regex.IsMatch(instruction, $"\\bREF\\s+{Regex.Escape(bookmark)}\\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    continue;
                }

                if (Regex.IsMatch(instruction, "(^|\\s)\\\\r(\\s|$)", RegexOptions.CultureInvariant))
                {
                    continue;
                }

                fieldCode.Text = instruction.Contains("\\h", StringComparison.Ordinal)
                    ? Regex.Replace(instruction, "\\s*\\\\h\\s*", " \\r \\h ", RegexOptions.CultureInvariant)
                    : instruction.TrimEnd() + " \\r ";
                refFieldsUpdated++;
            }
        }

        applied.Add($"updated REF fields with paragraph-number switch: {refFieldsUpdated}");
        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildStyleCaptionRepairReport(docxPath, plan, lockPath, author, applied, skipped), Encoding.UTF8);
    }

    foreach (var item in applied)
    {
        Console.WriteLine($"APPLY {item}");
    }

    foreach (var item in skipped)
    {
        Console.WriteLine($"SKIP {item}");
    }

    return 0;
}

static int AddBookmarks(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<BookmarkPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });
    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read bookmark plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils add-bookmarks {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var bookmarkId = NextBookmarkId(doc);
    foreach (var entry in plan.Bookmarks)
    {
        var paragraphs = GetAllParagraphs(doc);
        var matches = paragraphs
            .Where(p => NormalizedStartsWith(ParagraphText(p), entry.ParagraphPrefix))
            .ToList();
        if (entry.Occurrence is int occurrence && occurrence > 0)
        {
            matches = matches.Skip(occurrence - 1).Take(1).ToList();
        }

        if (matches.Count != 1)
        {
            skipped.Add($"{entry.Id}: paragraph anchor count={matches.Count}");
            continue;
        }

        if (BookmarkExists(doc, entry.Bookmark))
        {
            skipped.Add($"{entry.Id}: bookmark already exists ({entry.Bookmark})");
            continue;
        }

        AddParagraphBookmark(matches[0], entry.Bookmark, ref bookmarkId);
        applied.Add($"{entry.Id}: bookmark={entry.Bookmark}");
    }

    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Add Bookmarks Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        foreach (var item in applied) builder.AppendLine($"- APPLY {item}");
        foreach (var item in skipped) builder.AppendLine($"- SKIP {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int RewriteRefFields(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    if (!options.TryGetValue("bookmark-prefixes", out var prefixesValue) || string.IsNullOrWhiteSpace(prefixesValue))
    {
        Console.Error.WriteLine("Missing required option: --bookmark-prefixes");
        return 4;
    }

    if (!options.TryGetValue("template", out var template) || string.IsNullOrWhiteSpace(template))
    {
        Console.Error.WriteLine("Missing required option: --template");
        return 4;
    }

    var prefixes = prefixesValue
        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Where(p => !string.IsNullOrWhiteSpace(p))
        .ToArray();
    if (prefixes.Length == 0)
    {
        Console.Error.WriteLine("No bookmark prefixes provided.");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils rewrite-ref-fields {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    EnsureUpdateFieldsOnOpen(doc);
    var updated = 0;
    foreach (var fieldCode in doc.MainDocumentPart!.Document.Descendants<FieldCode>())
    {
        var instruction = fieldCode.Text ?? "";
        if (!TryRewriteRefInstruction(instruction, prefixes, template, out var newInstruction))
        {
            continue;
        }
        fieldCode.Text = newInstruction;
        updated++;
    }

    foreach (var simpleField in doc.MainDocumentPart!.Document.Descendants<SimpleField>())
    {
        var instruction = simpleField.Instruction?.Value ?? "";
        if (!TryRewriteRefInstruction(instruction, prefixes, template, out var newInstruction))
        {
            continue;
        }

        simpleField.Instruction = newInstruction;
        updated++;
    }

    applied.Add($"ref_fields_rewritten={updated}");
    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Rewrite REF Fields Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Bookmark prefixes: `{string.Join(", ", prefixes)}`");
        builder.AppendLine($"- Template: `{template}`");
        foreach (var item in applied) builder.AppendLine($"- APPLY {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static string NormalizeFieldInstruction(string? instruction) =>
    Regex.Replace(instruction ?? "", "\\s+", " ").Trim();

static bool TryRewriteRefInstruction(string instruction, string[] prefixes, string template, out string newInstruction)
{
    newInstruction = instruction;
    var match = Regex.Match(instruction ?? "", @"\bREF\s+([A-Za-z0-9_]+)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    if (!match.Success)
    {
        return false;
    }

    var bookmark = match.Groups[1].Value;
    if (!prefixes.Any(prefix => bookmark.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
    {
        return false;
    }

    newInstruction = template.Replace("{bookmark}", bookmark, StringComparison.Ordinal);
    return !string.Equals(NormalizeFieldInstruction(instruction), NormalizeFieldInstruction(newInstruction), StringComparison.Ordinal);
}

static int RepairLayoutPendencies(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils repair-layout-pendencies {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        var metadata = new RevisionMetadata(author, DateTime.UtcNow);
        var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
        var blocks = body.Elements<OpenXmlElement>().ToList();

        foreach (var title in new[]
        {
            "Strings Utilizadas na Pesquisa Base Scopus",
            "Strings Utilizadas na Pesquisa Base Web of Science"
        })
        {
            var table = FindTableAfterCaption(blocks, title);
            if (table is null)
            {
                skipped.Add($"table styles: caption not found or table missing after `{title}`");
                continue;
            }

            var changed = ApplyParagraphStyleInsideTable(table, "dados");
            applied.Add($"table `{title}`: applied paragraph style `dados` to {changed} paragraph(s)");
        }

        foreach (var captionPrefix in new[]
        {
            "SuavizaÃ§Ã£o parcial do histÃ³rico mensal do PLD",
            "SequÃªncia histÃ³rica de clusters no perÃ­odo de treino"
        })
        {
            var paragraph = body.Elements<Paragraph>()
                .Where(p => string.Equals(ParagraphStyleId(p), "Figura", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault(p => Normalize(ParagraphText(p)).StartsWith(Normalize(captionPrefix), StringComparison.Ordinal)
                    && Normalize(ParagraphText(p)).Contains("fonte: autor (2026)", StringComparison.Ordinal));
            if (paragraph is null)
            {
                skipped.Add($"figure source split: caption not found `{captionPrefix}`");
                continue;
            }

            var removed = RemoveSourceTextFromCaption(paragraph, "Fonte: Autor (2026)", metadata);
            if (!removed)
            {
                skipped.Add($"figure source split: source text not removed `{captionPrefix}`");
                continue;
            }

            var sourceParagraph = CreateTrackedSourceParagraph("Fonte: Autor (2026)", "legenda", metadata);
            paragraph.InsertAfterSelf(sourceParagraph);
            applied.Add($"figure `{captionPrefix}`: split source into its own paragraph");
        }

        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildLayoutRepairReport(docxPath, lockPath, author, applied, skipped), Encoding.UTF8);
    }

    foreach (var item in applied)
    {
        Console.WriteLine($"APPLY {item}");
    }

    foreach (var item in skipped)
    {
        Console.WriteLine($"SKIP {item}");
    }

    return 0;
}

static int RepairArticleAbntLayout(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils repair-article-abnt-layout {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        RemoveUpdateFieldsOnOpen(doc);
        applied.Add("settings: removed `w:updateFields` to avoid Word prompt for fields that may reference external files");
        var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

        var fixedFigures = RepairArticleFigures(body);
        applied.AddRange(fixedFigures.Select(x => $"figure: {x}"));

        var tableTitles = new[]
        {
            "SÃ­ntese da base de PLD utilizada",
            "ConfiguraÃ§Ã£o do experimento de geraÃ§Ã£o de cenÃ¡rios",
            "MÃ©tricas de validaÃ§Ã£o dos cenÃ¡rios de PLD em 2025",
            "CritÃ©rios de aceite registrados no experimento"
        };
        foreach (var title in tableTitles)
        {
        var result = RepairArticleTablePair(body, title);
        if (result.applied)
        {
            applied.Add($"table `{title}`: {result.reason}");
            }
            else
            {
                skipped.Add($"table `{title}`: {result.reason}");
            }
        }

        var tableCount = 0;
        foreach (var table in body.Elements<Table>().ToList())
        {
            tableCount++;
            var rebuilt = RebuildAcademicTable(table);
            table.InsertAfterSelf(rebuilt);
            table.Remove();
        }
        applied.Add($"tables: applied `tabelauerj`, paragraph style `dados`, centered width 100% and autofit to {tableCount} remaining table(s)");

        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var text = ParagraphText(paragraph);
            if (text.StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase))
            {
                SetParagraphStyle(paragraph, "legenda");
            }
        }
        applied.Add("sources: applied style `legenda` to all source paragraphs below figures/tables");

        applied.AddRange(NormalizeFigureIndentInDocument(doc, body));

        var reordered = ReorderFontanaBeforeGneiting(body);
        if (reordered)
        {
            applied.Add("references: moved FONTANA before GNEITING to restore alphabetical order");
        }

        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildArticleAbntLayoutReport(docxPath, lockPath, author, applied, skipped), Encoding.UTF8);
    }

    foreach (var item in applied)
    {
        Console.WriteLine($"APPLY {item}");
    }

    foreach (var item in skipped)
    {
        Console.WriteLine($"SKIP {item}");
    }

    return skipped.Count == 0 ? 0 : 9;
}

static int FormatAbntReferenceTitles(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var emphasis = options.TryGetValue("emphasis", out var emphasisValue) && !string.IsNullOrWhiteSpace(emphasisValue)
        ? emphasisValue.Trim().ToLowerInvariant()
        : "italic";
    if (emphasis is not "italic" and not "bold")
    {
        Console.Error.WriteLine("--emphasis must be italic or bold");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils format-abnt-reference-titles {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using (var doc = WordprocessingDocument.Open(docxPath, true))
    {
        EnsureTrackRevisions(doc);
        var metadata = new RevisionMetadata(author, DateTime.UtcNow);
        var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

        var referenceSpecs = BuildArticleMarkovReferenceTitleSpecs();
        var referenceParagraphs = ParagraphsAfterReferencesHeading(body);
        foreach (var spec in referenceSpecs)
        {
            var paragraph = referenceParagraphs
                .FirstOrDefault(p => ParagraphText(p).StartsWith(spec.ReferencePrefix, StringComparison.OrdinalIgnoreCase));
            if (paragraph is null)
            {
                skipped.Add($"{spec.Id}: reference paragraph not found");
                continue;
            }

            if (!ParagraphText(paragraph).Contains(spec.TitleToEmphasize, StringComparison.Ordinal))
            {
                skipped.Add($"{spec.Id}: title span not found `{spec.TitleToEmphasize}`");
                continue;
            }

            var changedRuns = ApplyRunEmphasisToSpan(paragraph, spec.TitleToEmphasize, emphasis, metadata);
            if (changedRuns == 0)
            {
                skipped.Add($"{spec.Id}: no run changed for `{spec.TitleToEmphasize}`");
                continue;
            }

            applied.Add($"{spec.Id}: emphasized `{spec.TitleToEmphasize}` ({changedRuns} run(s))");
        }

        EnableUpdateFieldsOnOpenSetting(doc);
        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildAbntReferenceTitleReport(docxPath, lockPath, author, emphasis, applied, skipped), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static IReadOnlyList<ReferenceTitleSpec> BuildArticleMarkovReferenceTitleSpecs() =>
[
    new("chicco_2012", "CHICCO,", "Energy"),
    new("cleveland_1979", "CLEVELAND,", "Journal of the American Statistical Association"),
    new("diaconis_2006", "DIACONIS,", "The Annals of Statistics"),
    new("gneiting_raftery_2007", "GNEITING,", "Journal of the American Statistical Association"),
    new("fontana_2023", "FONTANA,", "Bernoulli"),
    new("gontijo_2021", "GONTIJO,", "E3S Web of Conferences"),
    new("halilcevic_2011", "HALILCEVIC,", "INTERNATIONAL CONFERENCE ON THE EUROPEAN ENERGY MARKET"),
    new("hong_fan_2016", "HONG,", "International Journal of Forecasting"),
    new("kohonen_1990", "KOHONEN,", "Proceedings of the IEEE"),
    new("lago_2021", "LAGO,", "Applied Energy"),
    new("lauro_2023", "LAURO,", "Energies"),
    new("leonel_2021", "LEONEL,", "Energies"),
    new("lu_2022", "LU,", "Applied Energy"),
    new("marcjasz_2023", "MARCJASZ,", "Energy Economics"),
    new("nowotarski_2018", "NOWOTARSKI,", "Renewable and Sustainable Energy Reviews"),
    new("rasanen_2010", "RASANEN,", "Applied Energy"),
    new("santos_2021", "SANTOS,", "Energies"),
    new("sesia_candes_2020", "SESIA,", "Stat"),
    new("weron_2014", "WERON,", "International Journal of Forecasting"),
    new("zambelli_2011", "ZAMBELLI,", "Sba: Controle & AutomaÃ§Ã£o Sociedade Brasileira de Automatica")
];

static IReadOnlyList<Paragraph> ParagraphsAfterReferencesHeading(Body body)
{
    var result = new List<Paragraph>();
    var inReferences = false;
    foreach (var paragraph in body.Elements<Paragraph>())
    {
        var text = ParagraphText(paragraph);
        if (IsReferencesHeading(text))
        {
            inReferences = true;
            continue;
        }

        if (!inReferences)
        {
            continue;
        }

        if (!string.IsNullOrWhiteSpace(text))
        {
            result.Add(paragraph);
        }
    }

    return result;
}

static int ApplyRunEmphasisToSpan(Paragraph paragraph, string spanText, string emphasis, RevisionMetadata metadata)
{
    if (!TrySplitRunsForAnchor(paragraph, spanText, out var firstRun, out var lastRun) || firstRun is null || lastRun is null)
    {
        return 0;
    }

    var runs = paragraph.Elements<Run>().ToList();
    var start = runs.IndexOf(firstRun);
    var end = runs.IndexOf(lastRun);
    if (start < 0 || end < start)
    {
        return 0;
    }

    var changed = 0;
    foreach (var run in runs.Skip(start).Take(end - start + 1))
    {
        if (string.IsNullOrEmpty(string.Concat(run.Elements<Text>().Select(t => t.Text))))
        {
            continue;
        }

        run.RunProperties ??= new RunProperties();
        if (emphasis == "italic")
        {
            if (run.RunProperties.Italic is not null)
            {
                changed++;
                continue;
            }
            AddRunPropertiesChange(run.RunProperties, metadata);
            run.RunProperties.Italic = new Italic();
        }
        else
        {
            if (run.RunProperties.Bold is not null)
            {
                changed++;
                continue;
            }
            AddRunPropertiesChange(run.RunProperties, metadata);
            run.RunProperties.Bold = new Bold();
        }
        MoveRunPropertiesChangeToEnd(run.RunProperties);
        changed++;
    }

    return changed;
}

static void EnableUpdateFieldsOnOpenSetting(WordprocessingDocument doc)
{
    var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart ?? doc.MainDocumentPart?.AddNewPart<DocumentSettingsPart>();
    if (settingsPart is null)
    {
        throw new InvalidOperationException("Unable to access document settings.");
    }

    settingsPart.Settings ??= new Settings();
    foreach (var existing in settingsPart.Settings.Elements<UpdateFieldsOnOpen>().ToList())
    {
        existing.Remove();
    }
    settingsPart.Settings.AddChild(new UpdateFieldsOnOpen { Val = true }, true);
    settingsPart.Settings.Save();
}

static string BuildAbntReferenceTitleReport(
    string docxPath,
    string lockPath,
    string author,
    string emphasis,
    IReadOnlyList<string> applied,
    IReadOnlyList<string> skipped)
{
    var builder = new StringBuilder();
    builder.AppendLine("# Formatacao ABNT das referencias");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{docxPath}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine($"- Autor da etapa: `{author}`");
    builder.AppendLine("- Utilitario/codigo usado: `docx-utils`, comando `format-abnt-reference-titles` (.NET + Open XML)");
    builder.AppendLine($"- Enfase aplicada: `{emphasis}`");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine("- Regra: aplicar destaque tipografico uniforme aos titulos das publicacoes/obras-fonte nas referencias, mantendo os titulos dos artigos em texto romano.");
    builder.AppendLine("- `w:updateFields` foi preservado/reativado para atualizacao de campos ao abrir no Word.");
    builder.AppendLine();
    builder.AppendLine("## Aplicado");
    foreach (var item in applied) builder.AppendLine($"- {item}");
    builder.AppendLine();
    builder.AppendLine("## Nao aplicado / revisar");
    if (skipped.Count == 0)
    {
        builder.AppendLine("- Nenhum item.");
    }
    else
    {
        foreach (var item in skipped) builder.AppendLine($"- {item}");
    }
    return builder.ToString();
}

static int RewriteEquationBlocks(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue)
        || !options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<EquationRewritePlan>(
        File.ReadAllText(planPath, Encoding.UTF8),
        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read equation rewrite plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils rewrite-equation-blocks {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var metadata = new RevisionMetadata(author, DateTime.UtcNow);

    foreach (var edit in plan.Edits)
    {
        var paragraphs = GetAllParagraphs(doc)
            .Where(p => NormalizedStartsWith(ParagraphText(p), edit.ParagraphPrefix))
            .ToList();
        if (paragraphs.Count != 1)
        {
            skipped.Add($"{edit.Id}: paragraph anchor count={paragraphs.Count}");
            continue;
        }

        var paragraph = paragraphs[0];
        if (edit.ReplacementText is not null)
        {
            var originalText = ParagraphText(paragraph);
            ReplaceParagraphTracked(paragraph, originalText, [CrossrefRunPart.TextPart(edit.ReplacementText)], metadata);
        }

        OpenXmlElement insertionPoint = paragraph;
        foreach (var item in edit.Items)
        {
            OpenXmlElement block = item.Kind.Equals("equation", StringComparison.OrdinalIgnoreCase)
                ? CreateInsertedEquationParagraph(paragraph, item.Latex ?? "", item.StyleId ?? "equao", metadata)
                : CreateInsertedParagraphWithStyle(paragraph, item.Text ?? "", item.StyleId, metadata);

            insertionPoint.InsertAfterSelf(block);
            insertionPoint = block;
        }

        applied.Add(edit.Id);
    }

    doc.MainDocumentPart!.Document.Save();

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Rewrite Equation Blocks Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static int EnsureStyleFonts(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var targetFont = options.TryGetValue("font", out var fontValue) && !string.IsNullOrWhiteSpace(fontValue)
        ? fontValue
        : "Times New Roman";
    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils ensure-style-fonts {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
    if (stylesPart?.Styles is null)
    {
        Console.Error.WriteLine("Document has no style definitions part.");
        return 6;
    }

    var styles = stylesPart.Styles;
    styles.DocDefaults ??= new DocDefaults();
    var runDefaults = styles.DocDefaults.GetFirstChild<RunPropertiesDefault>();
    if (runDefaults is null)
    {
        runDefaults = new RunPropertiesDefault();
        styles.DocDefaults.PrependChild(runDefaults);
    }

    var defaultRunProps = runDefaults.GetFirstChild<RunPropertiesBaseStyle>();
    if (defaultRunProps is null)
    {
        defaultRunProps = new RunPropertiesBaseStyle();
        runDefaults.Append(defaultRunProps);
    }

    SetRunFontsOnly(defaultRunProps, targetFont);
    applied.Add("docDefaults: run font set");

    var changedStyles = 0;
    var skippedTableStyles = 0;
    foreach (var style in styles.Elements<Style>())
    {
        if (style.Type?.Value == StyleValues.Table)
        {
            style.StyleRunProperties?.Remove();
            skippedTableStyles++;
            continue;
        }

        style.StyleRunProperties ??= new StyleRunProperties();
        SetRunFontsOnly(style.StyleRunProperties, targetFont);
        changedStyles++;
    }

    applied.Add($"styles: run font set on {changedStyles} style(s)");
    applied.Add($"table styles: skipped {skippedTableStyles} table style(s) to preserve Open XML validity");
    styles.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Ensure Style Fonts Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Fonte aplicada: `{targetFont}`");
        foreach (var item in applied) builder.AppendLine($"- {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int FormatEquationParagraphs(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var styleId = options.TryGetValue("style-id", out var styleIdValue) && !string.IsNullOrWhiteSpace(styleIdValue)
        ? styleIdValue
        : "equao";
    var seqName = options.TryGetValue("seq-name", out var seqNameValue) && !string.IsNullOrWhiteSpace(seqNameValue)
        ? seqNameValue
        : "Eq";
    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");

    var applied = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils format-equation-paragraphs {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var paragraphs = body.Descendants<Paragraph>()
        .Where(p => string.Equals(ParagraphStyleId(p), styleId, StringComparison.OrdinalIgnoreCase))
        .ToList();

    var equationNumber = 1;
    foreach (var paragraph in paragraphs)
    {
        var mathParagraphs = paragraph.Descendants()
            .Where(e => e.LocalName == "oMathPara")
            .Select(e => (OpenXmlElement)e)
            .ToList();
        var mathElements = mathParagraphs.Count > 0
            ? mathParagraphs
            : paragraph.Descendants<M.OfficeMath>()
                .Where(m => !m.Ancestors().Any(a => a.LocalName == "oMathPara"))
                .Cast<OpenXmlElement>()
                .ToList();
        if (mathElements.Count == 0)
        {
            continue;
        }

        var mathCopies = mathElements
            .Select(m => m.CloneNode(true))
            .ToList();

        var properties = paragraph.ParagraphProperties?.CloneNode(true) as ParagraphProperties ?? new ParagraphProperties();
        EnsureEquationTabs(properties, paragraph);

        paragraph.RemoveAllChildren();
        paragraph.Append(properties);

        foreach (var math in mathCopies)
        {
            paragraph.Append(math);
        }

        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(CreateTextRun("("));
        foreach (var run in CreateComplexFieldRuns($" SEQ {seqName} \\* ARABIC ", equationNumber.ToString(CultureInfo.InvariantCulture)))
        {
            paragraph.Append(run);
        }
        paragraph.Append(CreateTextRun(")"));

        applied.Add($"formatted equation paragraph #{equationNumber}");
        equationNumber++;
    }

    doc.MainDocumentPart!.Document.Save();

    if (options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue))
    {
        var reportPath = Path.GetFullPath(reportPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Format Equation Paragraphs Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Estilo alvo: `{styleId}`");
        builder.AppendLine($"- Campo SEQ: `{seqName}`");
        builder.AppendLine();
        foreach (var item in applied) builder.AppendLine($"- APPLY {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int NormalizeFigureIndent(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;
    var applied = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils normalize-figure-indent {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    applied.AddRange(NormalizeFigureIndentInDocument(doc, body));
    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Normalize Figure Indent Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine("- Regra: zerar recuo esquerdo, direito, primeira linha e deslocamento em legendas, imagens e fontes de figuras.");
        foreach (var item in applied) builder.AppendLine($"- {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int ExportUsedStyles(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var outDir = Path.GetFullPath(options.TryGetValue("out", out var outValue) && !string.IsNullOrWhiteSpace(outValue)
        ? outValue
        : GetDefaultCanonicalStylesDirectory());
    Directory.CreateDirectory(outDir);

    using var doc = WordprocessingDocument.Open(docxPath, false);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart not found.");
    var stylesPart = mainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("StyleDefinitionsPart not found.");

    var usedParagraphStyles = GetAllParagraphs(doc)
        .Select(ParagraphStyleId)
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
        .ToList();
    var usedCharacterStyles = mainPart.Document.Descendants<RunStyle>()
        .Select(s => s.Val?.Value ?? "")
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
        .ToList();
    var usedTableStyles = mainPart.Document.Descendants<TableStyle>()
        .Select(s => s.Val?.Value ?? "")
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
        .ToList();
    AddIfPresentOrRequired(stylesPart.Styles!, usedParagraphStyles, "Normal");
    AddIfPresentOrRequired(stylesPart.Styles!, usedParagraphStyles, "dados");
    AddIfPresentOrRequired(stylesPart.Styles!, usedTableStyles, "tabelauerj");
    var usedNumberingIds = mainPart.Document.Descendants<NumberingId>()
        .Select(n => n.Val?.Value.ToString(CultureInfo.InvariantCulture) ?? "")
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(s => int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : int.MaxValue)
        .ToList();

    var stylesXml = stylesPart.Styles?.OuterXml ?? "";
    File.WriteAllText(Path.Combine(outDir, "styles.xml"), stylesXml, Encoding.UTF8);

    var numberingXml = mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "";
    if (!string.IsNullOrWhiteSpace(numberingXml))
    {
        File.WriteAllText(Path.Combine(outDir, "numbering.xml"), numberingXml, Encoding.UTF8);
    }

    var manifest = new
    {
        source = docxPath,
        exportedUtc = DateTime.UtcNow,
        paragraphStyles = usedParagraphStyles,
        characterStyles = usedCharacterStyles,
        tableStyles = usedTableStyles,
        numberingIds = usedNumberingIds
    };
    var json = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
    File.WriteAllText(Path.Combine(outDir, "used-styles-manifest.json"), json, Encoding.UTF8);

    Console.WriteLine($"OUT {outDir}");
    Console.WriteLine($"PARAGRAPH_STYLES {usedParagraphStyles.Count}");
    Console.WriteLine($"CHARACTER_STYLES {usedCharacterStyles.Count}");
    Console.WriteLine($"TABLE_STYLES {usedTableStyles.Count}");
    Console.WriteLine($"NUMBERING_IDS {usedNumberingIds.Count}");
    return 0;
}

static void AddIfPresentOrRequired(Styles styles, List<string> styleIds, string styleId)
{
    if (!styles.Elements<Style>().Any(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase)))
    {
        return;
    }

    if (!styleIds.Any(s => string.Equals(s, styleId, StringComparison.OrdinalIgnoreCase)))
    {
        styleIds.Add(styleId);
        styleIds.Sort(StringComparer.OrdinalIgnoreCase);
    }
}

static int EnsureCanonicalStylesCommand(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var sourceDir = Path.GetFullPath(options.TryGetValue("source", out var sourceValue) && !string.IsNullOrWhiteSpace(sourceValue)
        ? sourceValue
        : GetDefaultCanonicalStylesDirectory());
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils ensure-canonical-styles {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnableTrackRevisionsOnly(doc);
    var applied = EnsureCanonicalStyles(doc, sourceDir);
    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Ensure Canonical Styles Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Fonte canonica: `{sourceDir}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        foreach (var item in applied) builder.AppendLine($"- {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int SyncStylesFromDocxCommand(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("source-docx", out var sourceDocxValue) || string.IsNullOrWhiteSpace(sourceDocxValue))
    {
        Console.Error.WriteLine("Missing required option: --source-docx");
        return 4;
    }

    var sourceDocxPath = Path.GetFullPath(sourceDocxValue);
    if (!File.Exists(sourceDocxPath))
    {
        Console.Error.WriteLine($"Source DOCX not found: {sourceDocxPath}");
        return 5;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils sync-styles-from-docx {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var targetDoc = WordprocessingDocument.Open(docxPath, true);
    using var sourceDoc = WordprocessingDocument.Open(sourceDocxPath, false);

    EnableTrackRevisionsOnly(targetDoc);
    var applied = SyncStylesFromSourceDocx(targetDoc, sourceDoc);
    targetDoc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Sync Styles From DOCX Report");
        builder.AppendLine();
        builder.AppendLine($"- Target DOCX: `{docxPath}`");
        builder.AppendLine($"- Source DOCX: `{sourceDocxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        foreach (var item in applied) builder.AppendLine($"- {item}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int ApplyTableDesignStyle(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    if (!options.TryGetValue("style-id", out var styleId) || string.IsNullOrWhiteSpace(styleId))
    {
        Console.Error.WriteLine("Missing required option: --style-id");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;
    var styleName = options.TryGetValue("style-name", out var styleNameValue) && !string.IsNullOrWhiteSpace(styleNameValue)
        ? styleNameValue
        : null;
    var applied = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils apply-table-design-style {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    RemoveUpdateFieldsOnOpen(doc);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    var tableCount = 0;
    foreach (var table in body.Descendants<Table>().ToList())
    {
        if (EnsureTableStyle(table, styleId))
        {
            tableCount++;
        }
    }
    applied.Add($"tables: ensured table design style `{styleId}` on {tableCount} table(s) needing change");
    if (styleName is not null)
    {
        EnsureTableStyleDisplayName(doc, styleId, styleName);
        applied.Add($"styles: ensured table style `{styleId}` has display name `{styleName}`");
    }

    doc.MainDocumentPart!.Document.Save();
    WriteAppliedReport(reportPath, "Apply Table Design Style Report", docxPath, author, lockPath, applied, []);
    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    return 0;
}

static int ReplaceTable(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue)
        || !options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;
    var planJson = File.ReadAllText(planPath, Encoding.UTF8);
    var validation = PlanContractSupport.ValidateReplaceTablePlan(planJson);
    if (!validation.IsValid)
    {
        PrintPlanValidationErrors(validation.Errors);
        return 6;
    }

    var plan = JsonSerializer.Deserialize<ReplaceTablePlan>(planJson, CliOptions.JsonOptions()) ?? new ReplaceTablePlan();

    var lockDirectory = Path.GetDirectoryName(lockPath) ?? ".";
    Directory.CreateDirectory(lockDirectory);
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils replace-table {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    foreach (var spec in plan.Tables)
    {
        var selection = SelectTableForReplacement(body, spec);
        if (!selection.IsSuccess)
        {
            skipped.Add($"{spec.Id}: {selection.Message}");
            continue;
        }

        var table = selection.Table!;
        if (!string.IsNullOrWhiteSpace(spec.TableStyleId))
        {
            EnsureTableStyle(table, spec.TableStyleId);
        }

        ReplaceTableRows(table, spec, selection.ColumnWidths, selection.EffectiveCellStyleId);
        applied.Add($"{spec.Id}: table ordinal {selection.Ordinal} block {selection.BlockIndex} replaced with {spec.Rows.Count} row(s)");
    }

    SaveMainDocumentWithValidationRepair(doc, applied);
    WriteAppliedReport(reportPath, "Replace Table Report", docxPath, author, lockPath, applied, skipped);
    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static int CreateDocxCommand(string[] args)
{
    if (args.Length < 1)
    {
        PrintUsage();
        return 1;
    }

    var outputPath = Path.GetFullPath(args[0]);
    var options = CliOptions.Parse(args.Skip(1).ToArray());
    return CreateDocxSupport.CreateDocx(outputPath, options);
}

static void PrintPlanValidationErrors(IReadOnlyList<string> errors)
{
    foreach (var error in errors)
    {
        Console.Error.WriteLine(error);
    }
}

static int ReplaceFiguresFromPlan(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue)
        || !options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    var planPath = Path.GetFullPath(planPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) ? Path.GetFullPath(reportPathValue) : null;
    var plan = JsonSerializer.Deserialize<FigureReplacementPlan>(File.ReadAllText(planPath, Encoding.UTF8), CliOptions.JsonOptions()) ?? new FigureReplacementPlan();
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils replace-figures-from-plan {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    RemoveUpdateFieldsOnOpen(doc);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    foreach (var fix in plan.Figures)
    {
        var result = ReplaceFigureImageAndCaption(doc, body, fix);
        if (result.applied)
        {
            applied.Add(result.message);
        }
        else
        {
            skipped.Add(result.message);
        }
    }

    applied.AddRange(NormalizeFigureIndentInDocument(doc, body));
    doc.MainDocumentPart!.Document.Save();
    WriteAppliedReport(reportPath, "Replace Figures From Plan Report", docxPath, author, lockPath, applied, skipped);
    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static int ConvertTextFormulasToOfficeMath(string docxPath, IReadOnlyDictionary<string, string> options)
{
    Console.Error.WriteLine("convert-text-formulas-to-omath is deprecated; using the standardized linear LaTeX equation path. Plans must include formulas[].latex.");
    return ReplaceFormulasWithLinearOfficeMath(docxPath, options, "convert-text-formulas-to-omath");
}

static int ReplaceFormulasWithLinearOfficeMath(string docxPath, IReadOnlyDictionary<string, string> options, string commandName = "replace-formulas-with-linear-equations")
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue)
        || !options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    var planPath = Path.GetFullPath(planPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) ? Path.GetFullPath(reportPathValue) : null;
    var keepLinear = options.TryGetValue("keep-linear", out var keepLinearValue)
        && bool.TryParse(keepLinearValue, out var keepLinearParsed)
        && keepLinearParsed;
    var plan = JsonSerializer.Deserialize<FormulaConversionPlan>(File.ReadAllText(planPath, Encoding.UTF8), CliOptions.JsonOptions()) ?? new FormulaConversionPlan();
    var formulas = plan.Formulas
        .Select(f => new FormulaSpec(f.Text, f.Latex, f.MathMl, f.Display, f.Occurrence))
        .Where(f => !string.IsNullOrWhiteSpace(f.RequiredText))
        .ToList();
    var missingLatex = formulas.Where(f => string.IsNullOrWhiteSpace(f.Latex)).Select(f => f.RequiredText).ToList();
    if (missingLatex.Count > 0)
    {
        Console.Error.WriteLine("Every formula plan item must include `latex`; plain text must not be used as the equation payload.");
        foreach (var item in missingLatex) Console.Error.WriteLine($"Missing latex for: {item}");
        return 4;
    }

    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils {commandName} {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    RemoveUpdateFieldsOnOpen(doc);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    var formulaResult = ConvertFormulaTextsToOfficeMath(body, formulas, CreateLinearOfficeMath);
    applied.Add($"math: inserted {formulaResult.ConvertedCount} formula occurrence(s) as linear LaTeX Office Math in {formulaResult.ParagraphCount} paragraph(s)");
    if (!keepLinear)
    {
        var promotion = PromoteLinearOfficeMathToProfessional(body, formulas);
        applied.Add($"math: promoted {promotion.PromotedCount} linear Office Math occurrence(s) to professional OMML using .NET/OpenXML parser");
        foreach (var item in promotion.Warnings)
        {
            skipped.Add(item);
        }
    }
    else
    {
        applied.Add("math: kept linear Office Math because --keep-linear=true was requested");
    }
    foreach (var item in formulaResult.Unmatched) skipped.Add($"math pattern not found: {item}");
    doc.MainDocumentPart!.Document.Save();

    WriteAppliedReport(reportPath, "Replace Formulas With Linear Word Equations Report", docxPath, author, lockPath, applied, skipped);
    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static int ReplaceFormulasWithMathMlOfficeMath(string docxPath, IReadOnlyDictionary<string, string> options, string commandName = "replace-formulas-with-mathml-omml")
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue)
        || !options.TryGetValue("plan", out var planPathValue) || string.IsNullOrWhiteSpace(planPathValue))
    {
        Console.Error.WriteLine("Missing required options: --plan and --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    var planPath = Path.GetFullPath(planPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) ? Path.GetFullPath(reportPathValue) : null;
    var plan = JsonSerializer.Deserialize<FormulaConversionPlan>(File.ReadAllText(planPath, Encoding.UTF8), CliOptions.JsonOptions()) ?? new FormulaConversionPlan();
    var formulas = plan.Formulas
        .Select(f => new FormulaSpec(f.Text, f.Latex, f.MathMl, f.Display, f.Occurrence))
        .Where(f => !string.IsNullOrWhiteSpace(f.RequiredText))
        .ToList();
    var missingMathMl = formulas.Where(f => string.IsNullOrWhiteSpace(f.MathMl)).Select(f => f.RequiredText).ToList();
    if (missingMathMl.Count > 0)
    {
        Console.Error.WriteLine("Every formula plan item must include Presentation MathML in `mathMl`; raw text or LaTeX must not be inserted directly into m:oMath.");
        foreach (var item in missingMathMl) Console.Error.WriteLine($"Missing mathMl for: {item}");
        return 4;
    }

    var xslPath = ResolveMml2OmmlXsl(options.TryGetValue("xsl", out var xslPathValue) ? xslPathValue : null);
    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils {commandName} {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    RemoveUpdateFieldsOnOpen(doc);
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

    var formulaResult = ConvertFormulaTextsToOfficeMath(body, formulas, spec => CreateOfficeMathFromMathMl(spec, xslPath));
    applied.Add($"math: replaced {formulaResult.ConvertedCount} formula occurrence(s) with OMML generated from MathML in {formulaResult.ParagraphCount} paragraph(s)");
    applied.Add($"math: used MML2OMML stylesheet `{xslPath}`");
    foreach (var item in formulaResult.Unmatched) skipped.Add($"math pattern not found: {item}");
    doc.MainDocumentPart!.Document.Save();

    WriteAppliedReport(reportPath, "Replace Formulas With MathML OMML Report", docxPath, author, lockPath, applied, skipped);
    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return skipped.Count == 0 ? 0 : 9;
}

static void WriteAppliedReport(string? reportPath, string title, string docxPath, string author, string lockPath, IReadOnlyList<string> applied, IReadOnlyList<string> skipped)
{
    if (reportPath is null)
    {
        return;
    }

    Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
    var builder = new StringBuilder();
    builder.AppendLine($"# {title}");
    builder.AppendLine();
    builder.AppendLine($"- DOCX: `{docxPath}`");
    builder.AppendLine($"- Autor: `{author}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine();
    builder.AppendLine("## Applied");
    foreach (var item in applied) builder.AppendLine($"- {item}");
    builder.AppendLine();
    builder.AppendLine("## Skipped / Review");
    if (skipped.Count == 0) builder.AppendLine("- Nenhum item.");
    foreach (var item in skipped) builder.AppendLine($"- {item}");
    File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
}

static (int ConvertedCount, int ParagraphCount, IReadOnlyList<string> Unmatched) ConvertFormulaTextsToOfficeMath(
    Body body,
    IReadOnlyList<FormulaSpec> formulas,
    Func<FormulaSpec, OpenXmlElement> createMathElement)
{
    var matched = new HashSet<string>(StringComparer.Ordinal);
    var converted = 0;
    var paragraphCount = 0;

    foreach (var paragraph in body.Descendants<Paragraph>().ToList())
    {
        if (HasDrawing(paragraph))
        {
            continue;
        }

        var existingMathReplacements = ReplaceLinearTextOfficeMath(paragraph, formulas, matched, createMathElement);
        if (existingMathReplacements > 0)
        {
            converted += existingMathReplacements;
            paragraphCount++;
        }

        var text = ParagraphText(paragraph);
        if (string.IsNullOrWhiteSpace(text))
        {
            continue;
        }

        var replacements = FindFormulaReplacements(text, formulas);
        if (replacements.Count == 0)
        {
            continue;
        }

        RebuildParagraphWithOfficeMath(paragraph, text, replacements, createMathElement);
        foreach (var replacement in replacements)
        {
            matched.Add(replacement.Spec.RequiredText);
        }
        converted += replacements.Count;
        paragraphCount++;
    }

    var unmatched = formulas.Where(f => !matched.Contains(f.RequiredText)).Select(f => f.RequiredText).ToList();
    return (converted, paragraphCount, unmatched);
}

static List<FormulaReplacement> FindFormulaReplacements(string text, IReadOnlyList<FormulaSpec> formulas)
{
    var replacements = new List<FormulaReplacement>();
    var occupied = new bool[text.Length];
    foreach (var spec in formulas.OrderByDescending(f => f.RequiredText.Length))
    {
        var start = 0;
        while (start < text.Length)
        {
            var index = text.IndexOf(spec.RequiredText, start, StringComparison.Ordinal);
            if (index < 0)
            {
                break;
            }

            var length = spec.RequiredText.Length;
            if (!occupied.Skip(index).Take(length).Any(x => x))
            {
                replacements.Add(new FormulaReplacement(index, length, spec));
                for (var i = index; i < index + length; i++)
                {
                    occupied[i] = true;
                }
            }

            start = index + length;
        }
    }

    return replacements.OrderBy(r => r.Start).ToList();
}

static int ReplaceLinearTextOfficeMath(
    Paragraph paragraph,
    IReadOnlyList<FormulaSpec> formulas,
    ISet<string> matched,
    Func<FormulaSpec, OpenXmlElement> createMathElement)
{
    var converted = 0;
    foreach (var math in paragraph.Descendants<M.OfficeMath>().ToList())
    {
        var mathText = Regex.Replace(math.InnerText, "\\s+", " ").Trim();
        var auditText = SerializeOfficeMathForAudit(math);
        var spec = formulas.FirstOrDefault(f =>
            string.Equals(f.RequiredText, mathText, StringComparison.Ordinal)
            || string.Equals(f.RequiredText, auditText, StringComparison.Ordinal));
        if (spec is null)
        {
            continue;
        }

        math.InsertAfterSelf(createMathElement(spec));
        math.Remove();
        matched.Add(spec.RequiredText);
        converted++;
    }

    return converted;
}

static void RebuildParagraphWithOfficeMath(
    Paragraph paragraph,
    string text,
    IReadOnlyList<FormulaReplacement> replacements,
    Func<FormulaSpec, OpenXmlElement> createMathElement)
{
    var properties = paragraph.ParagraphProperties?.CloneNode(true);
    paragraph.RemoveAllChildren();
    if (properties is not null)
    {
        paragraph.Append(properties);
    }

    var cursor = 0;
    foreach (var replacement in replacements)
    {
        if (replacement.Start > cursor)
        {
            paragraph.Append(CreateTextRun(text[cursor..replacement.Start]));
        }

        paragraph.Append(createMathElement(replacement.Spec));
        cursor = replacement.Start + replacement.Length;
    }

    if (cursor < text.Length)
    {
        paragraph.Append(CreateTextRun(text[cursor..]));
    }
}

static OpenXmlElement CreateLinearOfficeMath(FormulaSpec spec)
{
    return new M.OfficeMath(
        new M.Run(
            new M.Text(spec.Latex) { Space = SpaceProcessingModeValues.Preserve }));
}

static (int PromotedCount, IReadOnlyList<string> Warnings) PromoteLinearOfficeMathToProfessional(Body body, IReadOnlyList<FormulaSpec> formulas)
{
    var promoted = 0;
    var warnings = new List<string>();
    var byLatex = formulas
        .Where(f => !string.IsNullOrWhiteSpace(f.Latex))
        .GroupBy(f => f.Latex, StringComparer.Ordinal)
        .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);

    foreach (var math in body.Descendants<M.OfficeMath>().ToList())
    {
        if (!IsLinearLatexOfficeMath(math, out var latexText))
        {
            continue;
        }

        if (!byLatex.TryGetValue(latexText, out var spec))
        {
            continue;
        }

        try
        {
            var professional = CreateProfessionalOfficeMathFromLatex(spec);
            math.InsertAfterSelf(professional);
            math.Remove();
            promoted++;
        }
        catch (Exception ex)
        {
            warnings.Add($"math promotion fallback kept linear for `{spec.RequiredText}`: {ex.Message}");
        }
    }

    return (promoted, warnings);
}

static bool IsLinearLatexOfficeMath(M.OfficeMath math, out string latexText)
{
    latexText = "";
    if (math.ChildElements.Count != 1 || math.FirstChild is not M.Run run)
    {
        return false;
    }

    var texts = run.Elements<M.Text>().ToList();
    if (texts.Count != 1)
    {
        return false;
    }

    latexText = texts[0].Text ?? "";
    return !string.IsNullOrWhiteSpace(latexText);
}

static OpenXmlElement CreateProfessionalOfficeMathFromLatex(FormulaSpec spec)
{
    var tokens = TokenizeLatex(spec.Latex);
    var parser = new LatexMathParser(tokens);
    var expression = parser.Parse();
    var omml = new M.OfficeMath();
    foreach (var child in expression.ToOpenXml())
    {
        omml.Append(child);
    }

    return omml;
}

static string SerializeOfficeMathForAudit(M.OfficeMath math)
{
    var text = SerializeMathElement(math);
    return Regex.Replace(text, "\\s+", " ").Trim();
}

static string SerializeMathElement(OpenXmlElement element) => element.LocalName switch
{
    "oMath" or "oMathPara" or "e" or "num" or "den" or "sub" or "sup" or "deg" or "fName" =>
        string.Concat(element.ChildElements.Select(SerializeMathElement)),
    "r" => string.Concat(element.ChildElements.Select(SerializeMathElement)),
    "t" => element.InnerText,
    "sSub" => $"{SerializeChildContainer(element, "e")}_{{{SerializeChildContainer(element, "sub")}}}",
    "sSup" => $"{SerializeChildContainer(element, "e")}^{{{SerializeChildContainer(element, "sup")}}}",
    "sSubSup" => $"{SerializeChildContainer(element, "e")}_{{{SerializeChildContainer(element, "sub")}}}^{{{SerializeChildContainer(element, "sup")}}}",
    "f" => $"\\frac{{{SerializeChildContainer(element, "num")}}}{{{SerializeChildContainer(element, "den")}}}",
    "nary" => $"{SerializeNaryOperator(element)}{SerializeNaryLimits(element)} {SerializeChildContainer(element, "e")}".Trim(),
    "d" => $"{SerializeDelimiterChar(element, "begChr", "(")}{SerializeChildContainer(element, "e")}{SerializeDelimiterChar(element, "endChr", ")")}",
    "func" => $"{SerializeChildContainer(element, "fName")}{SerializeChildContainer(element, "e")}",
    "acc" => SerializeChildContainer(element, "e"),
    "rad" => $"\\sqrt{{{SerializeChildContainer(element, "e")}}}",
    _ => element.ChildElements.Count == 0 ? element.InnerText : string.Concat(element.ChildElements.Select(SerializeMathElement))
};

static string SerializeChildContainer(OpenXmlElement element, string localName) =>
    string.Concat(element.ChildElements.Where(c => c.LocalName == localName).Select(SerializeMathElement));

static string SerializeNaryOperator(OpenXmlElement element)
{
    var chr = element.Descendants().FirstOrDefault(e => e.LocalName == "chr")?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value;
    return chr switch
    {
        "âˆ‘" => "\\sum",
        "âˆ" => "\\prod",
        _ => string.IsNullOrWhiteSpace(chr) ? "\\sum" : chr
    };
}

static string SerializeNaryLimits(OpenXmlElement element)
{
    var sub = SerializeChildContainer(element, "sub");
    var sup = SerializeChildContainer(element, "sup");
    var builder = new StringBuilder();
    if (!string.IsNullOrWhiteSpace(sub))
    {
        builder.Append("_{").Append(sub).Append('}');
    }

    if (!string.IsNullOrWhiteSpace(sup))
    {
        builder.Append("^{").Append(sup).Append('}');
    }

    return builder.ToString();
}

static string SerializeDelimiterChar(OpenXmlElement element, string localName, string fallback)
{
    var prop = element.ChildElements.FirstOrDefault(c => c.LocalName == "dPr");
    var value = prop?.ChildElements.FirstOrDefault(c => c.LocalName == localName)?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value;
    return string.IsNullOrWhiteSpace(value) ? fallback : value;
}

static OpenXmlElement CreateOfficeMathFromMathMl(FormulaSpec spec, string mml2OmmlXslPath)
{
    var ommlXml = TransformMathMlToOmml(spec.MathMl, mml2OmmlXslPath);
    var root = XDocument.Parse(ommlXml).Root ?? throw new InvalidOperationException("MML2OMML transform returned empty XML.");
    var rootName = root.Name.LocalName;
    if (rootName == "oMath")
    {
        return new M.OfficeMath(root.ToString(SaveOptions.DisableFormatting));
    }

    var math = root.Descendants().FirstOrDefault(e => e.Name.LocalName == "oMath");
    if (math is not null)
    {
        return new M.OfficeMath(math.ToString(SaveOptions.DisableFormatting));
    }

    throw new InvalidOperationException("MML2OMML transform did not produce m:oMath.");
}

static string TransformMathMlToOmml(string mathMl, string mml2OmmlXslPath)
{
    var transform = new XslCompiledTransform();
    var settings = new XsltSettings(enableDocumentFunction: false, enableScript: false);
    transform.Load(mml2OmmlXslPath, settings, new XmlUrlResolver());

    using var reader = XmlReader.Create(new StringReader(mathMl), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
    using var output = new StringWriter(CultureInfo.InvariantCulture);
    using var writer = XmlWriter.Create(output, transform.OutputSettings);
    transform.Transform(reader, writer);
    return output.ToString();
}

static string ResolveMml2OmmlXsl(string? configuredPath)
{
    if (!string.IsNullOrWhiteSpace(configuredPath))
    {
        var path = Path.GetFullPath(configuredPath);
        if (!File.Exists(path)) throw new FileNotFoundException("MML2OMML.XSL not found.", path);
        return path;
    }

    var candidates = new[]
    {
        @"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"
    };

    var found = candidates.FirstOrDefault(File.Exists);
    if (found is not null) return found;
    throw new FileNotFoundException("MML2OMML.XSL not found. Pass --xsl with the Office MML2OMML.XSL path.");
}

static (bool applied, string message) ReplaceFigureImageAndCaption(WordprocessingDocument doc, Body body, FigureReplacement fix)
{
    if (!File.Exists(fix.ImagePath))
    {
        return (false, $"figure `{fix.DrawingName}`: image file not found `{fix.ImagePath}`");
    }

    var drawing = body.Descendants<Drawing>()
        .FirstOrDefault(d =>
        {
            var inlineProps = d.Inline?.DocProperties;
            var anchorProps = d.Anchor?.GetFirstChild<DocProperties>();
            return string.Equals(inlineProps?.Name?.Value, fix.DrawingName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(anchorProps?.Name?.Value, fix.DrawingName, StringComparison.OrdinalIgnoreCase);
        });
    if (drawing is null)
    {
        return (false, $"figure `{fix.DrawingName}`: drawing not found");
    }

    var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
    var relationshipId = blip?.Embed?.Value;
    if (string.IsNullOrWhiteSpace(relationshipId))
    {
        return (false, $"figure `{fix.DrawingName}`: image relationship not found");
    }

    if (doc.MainDocumentPart!.GetPartById(relationshipId) is not ImagePart imagePart)
    {
        return (false, $"figure `{fix.DrawingName}`: image part not found");
    }

    using (var stream = File.OpenRead(fix.ImagePath))
    {
        imagePart.FeedData(stream);
    }

    var pixelDimensions = TryReadPngDimensions(fix.ImagePath);
    if (pixelDimensions is not null)
    {
        var widthEmu = 14L * 360000L;
        var heightEmu = (long)Math.Round(widthEmu * pixelDimensions.Value.Height / (double)pixelDimensions.Value.Width);
        SetDrawingExtent(drawing, widthEmu, heightEmu);
    }

    if (drawing.Inline?.DocProperties is not null)
    {
        drawing.Inline.DocProperties.Description = fix.CaptionText;
    }
    var anchorProperties = drawing.Anchor?.GetFirstChild<DocProperties>();
    if (anchorProperties is not null)
    {
        anchorProperties.Description = fix.CaptionText;
    }

    var imageParagraph = drawing.Ancestors<Paragraph>().FirstOrDefault();
    var caption = imageParagraph?.ElementsBefore().OfType<Paragraph>().LastOrDefault();
    if (caption is not null)
    {
        SetParagraphStyle(caption, "Figura");
        CenterParagraph(caption);
        ReplaceParagraphPlainText(caption, fix.CaptionText);
    }

    var source = imageParagraph?.ElementsAfter().OfType<Paragraph>().FirstOrDefault();
    if (source is not null && ParagraphText(source).StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase))
    {
        SetParagraphStyle(source, "legenda");
        CenterParagraph(source);
        ReplaceParagraphPlainText(source, fix.SourceText);
    }

    return (true, $"figure `{fix.DrawingName}`: replaced image, caption and source");
}

static void SetDrawingExtent(Drawing drawing, long widthEmu, long heightEmu)
{
    if (drawing.Inline is not null)
    {
        drawing.Inline.Extent ??= new Extent();
        drawing.Inline.Extent.Cx = widthEmu;
        drawing.Inline.Extent.Cy = heightEmu;
    }

    if (drawing.Anchor is not null)
    {
        var extent = drawing.Anchor.GetFirstChild<Extent>();
        if (extent is not null)
        {
            extent.Cx = widthEmu;
            extent.Cy = heightEmu;
        }
    }

    var picExtents = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
    if (picExtents is not null)
    {
        picExtents.Cx = widthEmu;
        picExtents.Cy = heightEmu;
    }
}

static int EnableUpdateFieldsOnOpen(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils enable-update-fields-on-open {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    EnsureUpdateFieldsOnOpen(doc);
    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Enable Update Fields On Open Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine("- Regra: gravar `w:updateFields` em `word/settings.xml` para o Word atualizar campos na abertura.");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    Console.WriteLine("APPLY settings: enabled `w:updateFields` on open");
    return 0;
}

static int DisableUpdateFieldsOnOpen(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils disable-update-fields-on-open {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    RemoveUpdateFieldsOnOpen(doc);
    doc.MainDocumentPart!.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Disable Update Fields On Open Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine("- Regra: remover `w:updateFields` de `word/settings.xml` para evitar aviso do Word sobre campos que podem referenciar outros arquivos.");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    Console.WriteLine("APPLY settings: removed `w:updateFields` on open");
    return 0;
}

static IReadOnlyList<string> NormalizeFigureIndentInDocument(WordprocessingDocument doc, Body body)
{
    var applied = new List<string>();
    var figureStyleChanged = NormalizeStyleParagraphIndent(doc, "Figura");
    if (figureStyleChanged)
    {
        applied.Add("style `Figura`: paragraph indentation set to zero");
    }

    var sourceStyleChanged = NormalizeStyleParagraphIndent(doc, "legenda");
    if (sourceStyleChanged)
    {
        applied.Add("style `legenda`: paragraph indentation set to zero");
    }

    var paragraphs = body.Elements<Paragraph>().ToList();
    var changedParagraphs = 0;
    for (var i = 0; i < paragraphs.Count; i++)
    {
        var paragraph = paragraphs[i];
        var isFigureCaption = string.Equals(ParagraphStyleId(paragraph), "Figura", StringComparison.OrdinalIgnoreCase);
        var isFigureImage = HasDrawing(paragraph)
            && i > 0
            && string.Equals(ParagraphStyleId(paragraphs[i - 1]), "Figura", StringComparison.OrdinalIgnoreCase);
        var isFigureSource = ParagraphText(paragraph).StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase)
            && i > 0
            && (HasDrawing(paragraphs[i - 1])
                || string.Equals(ParagraphStyleId(paragraphs[Math.Max(0, i - 2)]), "Figura", StringComparison.OrdinalIgnoreCase));

        if (!isFigureCaption && !isFigureImage && !isFigureSource)
        {
            continue;
        }

        SetParagraphIndentZero(paragraph);
        changedParagraphs++;
    }

    applied.Add($"paragraphs: indentation set to zero on {changedParagraphs} figure caption/image/source paragraph(s)");
    return applied;
}

static bool NormalizeStyleParagraphIndent(WordprocessingDocument doc, string styleId)
{
    var style = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
        .Elements<Style>()
        .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
    if (style is null)
    {
        return false;
    }

    style.StyleParagraphProperties ??= new StyleParagraphProperties();
    SetIndentZero(style.StyleParagraphProperties);
    doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Save();
    return true;
}

static void SetParagraphIndentZero(Paragraph paragraph)
{
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    SetIndentZero(paragraph.ParagraphProperties);
}

static void SetIndentZero(OpenXmlElement paragraphProperties)
{
    foreach (var old in paragraphProperties.Elements<Indentation>().ToList()) old.Remove();
    var indentation = new Indentation { Left = "0", Right = "0", FirstLine = "0", Hanging = "0" };
    OpenXmlElement? insertBefore = paragraphProperties.Elements<Justification>().FirstOrDefault();
    insertBefore ??= paragraphProperties.Elements<ParagraphMarkRunProperties>().FirstOrDefault();
    if (insertBefore is not null)
    {
        paragraphProperties.InsertBefore(indentation, insertBefore);
    }
    else
    {
        paragraphProperties.Append(indentation);
    }
}

static void SetRunFontsOnly(OpenXmlElement runProperties, string targetFont)
{
    var fonts = runProperties.Elements<RunFonts>().FirstOrDefault();
    if (fonts is not null)
    {
        fonts.Remove();
    }

    fonts = new RunFonts();
    var firstChild = runProperties.ChildElements.FirstOrDefault();
    if (firstChild is not null)
    {
        runProperties.InsertBefore(fonts, firstChild);
    }
    else
    {
        runProperties.Append(fonts);
    }

    fonts.Ascii = targetFont;
    fonts.HighAnsi = targetFont;
    fonts.EastAsia = targetFont;
    fonts.ComplexScript = targetFont;
}

static IReadOnlyList<string> RepairArticleFigures(Body body)
{
    var applied = new List<string>();
    var blocks = body.Elements<OpenXmlElement>().ToList();
    for (var i = 0; i < blocks.Count - 1; i++)
    {
        if (blocks[i] is not Paragraph figureParagraph || !HasDrawing(figureParagraph))
        {
            continue;
        }

        if (blocks[i + 1] is not Paragraph captionParagraph)
        {
            continue;
        }

        var captionText = ParagraphText(captionParagraph);
        if (!Regex.IsMatch(captionText, @"^\s*Figura\s+\d+\s*[-â€“â€”]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            continue;
        }

        var title = Regex.Replace(captionText, @"^\s*Figura\s+\d+\s*[-â€“â€”]\s*", "", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Trim();
        var drawings = figureParagraph.Descendants<Drawing>().Select(d => (Drawing)d.CloneNode(true)).ToList();
        SetParagraphStyle(captionParagraph, "Figura");
        CenterParagraph(captionParagraph);
        ReplaceParagraphWithTitleAndDrawings(captionParagraph, title, drawings);

        figureParagraph.Remove();
        applied.Add(title);
        blocks = body.Elements<OpenXmlElement>().ToList();
        i = Math.Max(-1, i - 1);
    }

    foreach (var paragraph in body.Elements<Paragraph>()
        .Where(p => string.Equals(ParagraphStyleId(p), "Figura", StringComparison.OrdinalIgnoreCase) && HasDrawing(p))
        .ToList())
    {
        var title = Regex.Replace(ParagraphText(paragraph).Trim(), @"^\s*Figura\s+\d+\s*[-â€“â€”]\s*", "", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            continue;
        }

        var drawings = paragraph.Descendants<Drawing>().Select(d => (Drawing)d.CloneNode(true)).ToList();
        CenterParagraph(paragraph);
        ReplaceParagraphWithTitleAndDrawings(paragraph, title, drawings);
    }

    var paragraphs = body.Elements<Paragraph>().ToList();
    for (var i = 0; i < paragraphs.Count - 1; i++)
    {
        var captionParagraph = paragraphs[i];
        if (!string.Equals(ParagraphStyleId(captionParagraph), "Figura", StringComparison.OrdinalIgnoreCase))
        {
            continue;
        }

        var imageParagraph = paragraphs[i + 1];
        if (!HasDrawing(imageParagraph))
        {
            continue;
        }

        var title = ParagraphText(captionParagraph).Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            continue;
        }

        var drawings = imageParagraph.Descendants<Drawing>().Select(d => (Drawing)d.CloneNode(true)).ToList();
        CenterParagraph(captionParagraph);
        ReplaceParagraphWithTitleAndDrawings(captionParagraph, title, drawings);
        imageParagraph.Remove();
    }

    foreach (var paragraph in body.Elements<Paragraph>().Where(p => HasDrawing(p) && !string.Equals(ParagraphStyleId(p), "Figura", StringComparison.OrdinalIgnoreCase)))
    {
        NormalizeImageParagraph(paragraph);
    }

    return applied;
}

static void ReplaceParagraphWithTitleAndDrawings(Paragraph paragraph, string title, IReadOnlyList<Drawing> drawings)
{
    ReplaceParagraphPlainText(paragraph, title);
    if (drawings.Count == 0)
    {
        return;
    }

    paragraph.Append(new Run(new Break()));
    foreach (var drawing in drawings)
    {
        paragraph.Append(new Run(ConvertDrawingToInline(drawing)));
    }
}

static void NormalizeImageParagraph(Paragraph paragraph)
{
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    var props = paragraph.ParagraphProperties;
    foreach (var old in props.Elements<SpacingBetweenLines>().ToList()) old.Remove();
    foreach (var old in props.Elements<Justification>().ToList()) old.Remove();
    props.Append(new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
    props.Append(new Justification { Val = JustificationValues.Center });
}

static (bool applied, string reason) RepairArticleTablePair(Body body, string title)
{
    var blocks = body.Elements<OpenXmlElement>().ToList();
    for (var i = 0; i < blocks.Count - 3; i++)
    {
        if (blocks[i] is not Paragraph titleParagraph)
        {
            continue;
        }

        if (!Normalize(ParagraphText(titleParagraph)).Equals(Normalize(title), StringComparison.Ordinal))
        {
            continue;
        }

        var firstTable = blocks.Skip(i + 1).OfType<Table>().FirstOrDefault();
        if (firstTable is null)
        {
            return (false, "caption found, but no table found after it");
        }

        var firstIndex = blocks.IndexOf(firstTable);
        var secondTable = blocks.Skip(firstIndex + 1).OfType<Table>().FirstOrDefault();
        if (secondTable is null || secondTable.ElementsBefore().OfType<Paragraph>().Any(p => ParagraphText(p).StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase)))
        {
            SetParagraphStyle(titleParagraph, "Tabela");
            CenterParagraph(titleParagraph);
            var rebuilt = RebuildAcademicTable(firstTable);
            firstTable.InsertAfterSelf(rebuilt);
            firstTable.Remove();
            var source = rebuilt.ElementsAfter().OfType<Paragraph>().FirstOrDefault();
            if (source is not null && ParagraphText(source).StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase))
            {
                SetParagraphStyle(source, "legenda");
            }

            return (true, "single table found; applied style `Tabela`, source below and width=100% autofit");
        }

        SetParagraphStyle(titleParagraph, "Tabela");
        CenterParagraph(titleParagraph);
        firstTable.Remove();
        var rebuiltSecond = RebuildAcademicTable(secondTable);
        secondTable.InsertAfterSelf(rebuiltSecond);
        secondTable.Remove();
        var next = rebuiltSecond.ElementsAfter().OfType<Paragraph>().FirstOrDefault();
        if (next is not null && ParagraphText(next).StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase))
        {
            SetParagraphStyle(next, "legenda");
        }

        return (true, "removed first duplicated table, kept formatted table, applied style `Tabela`, source below and width=100% autofit");
    }

    return (false, "caption not found");
}

static void ApplyAcademicTableWindowFit(Table table)
{
    var properties = table.GetFirstChild<TableProperties>();
    if (properties is null)
    {
        properties = new TableProperties();
        table.PrependChild(properties);
    }

    foreach (var old in properties.Elements<TableStyle>().ToList()) old.Remove();
    properties.PrependChild(new TableStyle { Val = "tabelauerj" });

    properties.TableWidth ??= new TableWidth();
    properties.TableWidth.Type = TableWidthUnitValues.Dxa;
    properties.TableWidth.Width = "9000";

    properties.TableJustification ??= new TableJustification();
    properties.TableJustification.Val = TableRowAlignmentValues.Center;

    foreach (var old in properties.Elements<TableLayout>().ToList()) old.Remove();
    foreach (var old in properties.Elements<TableLook>().ToList()) old.Remove();
    foreach (var old in properties.Elements<TableIndentation>().ToList()) old.Remove();
    foreach (var old in properties.Elements<TableCellMarginDefault>().ToList()) old.Remove();
    foreach (var old in properties.Elements<TableBorders>().ToList()) old.Remove();
    EnsureTableBorders(properties);
    properties.Append(new TableLayout { Type = TableLayoutValues.Fixed });
    properties.Append(new TableCellMarginDefault(
        new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
        new TableCellLeftMargin { Width = 80, Type = TableWidthValues.Dxa },
        new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
        new TableCellRightMargin { Width = 80, Type = TableWidthValues.Dxa }));
    foreach (var old in table.Descendants<ConditionalFormatStyle>().ToList()) old.Remove();

    var columnWidths = CalculateContentBasedColumnWidths(table);
    if (columnWidths.Count > 0)
    {
        table.RemoveAllChildren<TableGrid>();
        var grid = new TableGrid();
        foreach (var width in columnWidths)
        {
            grid.Append(new GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) });
        }

        var propertiesIndex = table.Elements().TakeWhile(e => e is TableProperties).Count();
        table.InsertAt(grid, propertiesIndex);
    }

    foreach (var row in table.Elements<TableRow>())
    {
        var columnIndex = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            var width = columnWidths.Count == 0
                ? 4500
                : Enumerable.Range(columnIndex, Math.Min(span, columnWidths.Count - columnIndex))
                    .Select(i => columnWidths[i])
                    .DefaultIfEmpty(columnWidths.Last())
                    .Sum();
            cell.TableCellProperties ??= new TableCellProperties();
            foreach (var old in cell.TableCellProperties.Elements<TableCellFitText>().ToList()) old.Remove();
            foreach (var old in cell.TableCellProperties.Elements<TextDirection>().ToList()) old.Remove();
            foreach (var old in cell.TableCellProperties.Elements<NoWrap>().ToList()) old.Remove();
            cell.TableCellProperties.TableCellWidth ??= new TableCellWidth();
            cell.TableCellProperties.TableCellWidth.Type = TableWidthUnitValues.Dxa;
            cell.TableCellProperties.TableCellWidth.Width = width.ToString(CultureInfo.InvariantCulture);
            columnIndex += span;
        }
    }

    foreach (var paragraph in table.Descendants<Paragraph>())
    {
        NormalizeTableParagraph(paragraph);
    }
}

static void EnsureTableBorders(TableProperties properties)
{
    properties.Append(new TableBorders(
        new TopBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
        new LeftBorder { Val = BorderValues.Nil },
        new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
        new RightBorder { Val = BorderValues.Nil },
        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 2U, Color = "000000" },
        new InsideVerticalBorder { Val = BorderValues.Nil }));
}

static Table RebuildAcademicTable(Table source)
{
    var rows = source.Elements<TableRow>()
        .Select(row => row.Elements<TableCell>()
            .Select(cell => Regex.Replace(string.Join(" ", cell.Descendants<Text>().Select(t => t.Text)), "\\s+", " ").Trim())
            .ToList())
        .Where(row => row.Count > 0)
        .ToList();

    var table = new Table();
    var maxColumns = rows.Select(r => r.Count).DefaultIfEmpty(0).Max();
    var widths = CalculateContentBasedColumnWidthsFromRows(rows);
    table.Append(new TableProperties(
        new TableStyle { Val = "tabelauerj" },
        new TableWidth { Width = "9000", Type = TableWidthUnitValues.Dxa },
        new TableJustification { Val = TableRowAlignmentValues.Center },
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
            new LeftBorder { Val = BorderValues.Nil },
            new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
            new RightBorder { Val = BorderValues.Nil },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 2U, Color = "000000" },
            new InsideVerticalBorder { Val = BorderValues.Nil }),
        new TableLayout { Type = TableLayoutValues.Fixed },
        new TableCellMarginDefault(
            new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellLeftMargin { Width = 80, Type = TableWidthValues.Dxa },
            new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellRightMargin { Width = 80, Type = TableWidthValues.Dxa })));

    var grid = new TableGrid();
    foreach (var width in widths)
    {
        grid.Append(new GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) });
    }
    table.Append(grid);

    foreach (var rowValues in rows)
    {
        var row = new TableRow();
        for (var i = 0; i < maxColumns; i++)
        {
            var width = widths.Count > i ? widths[i] : 4500;
            var cell = new TableCell(new TableCellProperties(new TableCellWidth
            {
                Width = width.ToString(CultureInfo.InvariantCulture),
                Type = TableWidthUnitValues.Dxa
            }));
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "dados" }));
            NormalizeTableParagraph(paragraph);
            paragraph.Append(CreateTextRun(i < rowValues.Count ? rowValues[i] : ""));
            cell.Append(paragraph);
            row.Append(cell);
        }
        table.Append(row);
    }

    return table;
}

static List<int> CalculateContentBasedColumnWidths(Table table)
{
    var rows = table.Elements<TableRow>().ToList();
    var maxColumns = rows
        .Select(row => row.Elements<TableCell>().Sum(cell => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1))
        .DefaultIfEmpty(0)
        .Max();
    if (maxColumns <= 0)
    {
        return [];
    }

    var weights = Enumerable.Repeat(1.0, maxColumns).ToArray();
    foreach (var row in rows)
    {
        var columnIndex = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            var text = Regex.Replace(string.Join(" ", cell.Descendants<Text>().Select(t => t.Text)), "\\s+", " ").Trim();
            var score = Math.Clamp(text.Length, 4, 80);
            if (span <= 1 && columnIndex < weights.Length)
            {
                weights[columnIndex] = Math.Max(weights[columnIndex], score);
            }
            else
            {
                var distributed = score / Math.Max(1, span);
                for (var i = columnIndex; i < Math.Min(weights.Length, columnIndex + span); i++)
                {
                    weights[i] = Math.Max(weights[i], distributed);
                }
            }

            columnIndex += span;
        }
    }

    const int totalWidth = 9000;
    const int minWidth = 1100;
    var widths = weights.Select(weight => (int)Math.Round(totalWidth * weight / weights.Sum())).ToList();
    for (var i = 0; i < widths.Count; i++)
    {
        widths[i] = Math.Max(minWidth, widths[i]);
    }

    while (widths.Sum() > totalWidth && widths.Any(w => w > minWidth))
    {
        var i = widths.IndexOf(widths.Max());
        widths[i]--;
    }

    while (widths.Sum() < totalWidth)
    {
        var i = widths.IndexOf(widths.Max());
        widths[i]++;
    }

    return widths;
}

static List<int> CalculateContentBasedColumnWidthsFromRows(IReadOnlyList<IReadOnlyList<string>> rows)
{
    var maxColumns = rows.Select(row => row.Count).DefaultIfEmpty(0).Max();
    if (maxColumns <= 0)
    {
        return [];
    }

    var weights = Enumerable.Repeat(1.0, maxColumns).ToArray();
    foreach (var row in rows)
    {
        for (var i = 0; i < row.Count; i++)
        {
            weights[i] = Math.Max(weights[i], Math.Clamp(row[i].Length, 4, 80));
        }
    }

    return NormalizeColumnWeights(weights);
}

static List<int> NormalizeColumnWeights(double[] weights)
{
    const int totalWidth = 9000;
    const int minWidth = 1100;
    var widths = weights.Select(weight => (int)Math.Round(totalWidth * weight / weights.Sum())).ToList();
    for (var i = 0; i < widths.Count; i++)
    {
        widths[i] = Math.Max(minWidth, widths[i]);
    }

    while (widths.Sum() > totalWidth && widths.Any(w => w > minWidth))
    {
        var i = widths.IndexOf(widths.Max());
        widths[i]--;
    }

    while (widths.Sum() < totalWidth)
    {
        var i = widths.IndexOf(widths.Max());
        widths[i]++;
    }

    return widths;
}

static void NormalizeTableParagraph(Paragraph paragraph)
{
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    var props = paragraph.ParagraphProperties;

    foreach (var old in props.Elements<SpacingBetweenLines>().ToList()) old.Remove();
    foreach (var old in props.Elements<Indentation>().ToList()) old.Remove();

    var spacing = new SpacingBetweenLines
    {
        Before = "0",
        After = "0",
        Line = "240",
        LineRule = LineSpacingRuleValues.Auto
    };
    var indentation = new Indentation { Left = "0", Right = "0", FirstLine = "0", Hanging = "0" };

    OpenXmlElement? insertBefore = props.Elements<Justification>().FirstOrDefault();
    insertBefore ??= props.Elements<ParagraphMarkRunProperties>().FirstOrDefault();
    if (insertBefore is not null)
    {
        props.InsertBefore(spacing, insertBefore);
        props.InsertBefore(indentation, insertBefore);
    }
    else
    {
        props.Append(spacing);
        props.Append(indentation);
    }
}

static Drawing ConvertDrawingToInline(Drawing drawing)
{
    if (drawing.Inline is not null)
    {
        return new Drawing((Inline)drawing.Inline.CloneNode(true));
    }

    var anchor = drawing.Anchor;
    if (anchor is null)
    {
        return drawing;
    }

    var inline = new Inline(
        (Extent?)anchor.GetFirstChild<Extent>()?.CloneNode(true) ?? new Extent { Cx = 0L, Cy = 0L },
        (EffectExtent?)anchor.GetFirstChild<EffectExtent>()?.CloneNode(true) ?? new EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        (DocProperties?)anchor.GetFirstChild<DocProperties>()?.CloneNode(true) ?? new DocProperties { Id = 1U, Name = "figure" },
        (NonVisualGraphicFrameDrawingProperties?)anchor.GetFirstChild<NonVisualGraphicFrameDrawingProperties>()?.CloneNode(true) ?? new NonVisualGraphicFrameDrawingProperties(),
        (DocumentFormat.OpenXml.Drawing.Graphic?)anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Graphic>()?.CloneNode(true) ?? new DocumentFormat.OpenXml.Drawing.Graphic())
    {
        DistanceFromTop = 0U,
        DistanceFromBottom = 0U,
        DistanceFromLeft = 0U,
        DistanceFromRight = 0U
    };

    return new Drawing(inline);
}

static void ReplaceParagraphPlainText(Paragraph paragraph, string text)
{
    var properties = paragraph.ParagraphProperties?.CloneNode(true);
    paragraph.RemoveAllChildren();
    if (properties is not null)
    {
        paragraph.Append(properties);
    }

    paragraph.Append(CreateTextRun(text));
}

static void SetParagraphStyle(Paragraph paragraph, string styleId)
{
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    if (paragraph.ParagraphProperties.ParagraphStyleId is null)
    {
        paragraph.ParagraphProperties.PrependChild(new ParagraphStyleId { Val = styleId });
    }
    else
    {
        paragraph.ParagraphProperties.ParagraphStyleId.Val = styleId;
    }
}

static void CenterParagraph(Paragraph paragraph)
{
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    var justification = paragraph.ParagraphProperties.GetFirstChild<Justification>();
    if (justification is null)
    {
        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Center });
    }
    else
    {
        justification.Val = JustificationValues.Center;
    }
}

static bool ReorderFontanaBeforeGneiting(Body body)
{
    var paragraphs = body.Elements<Paragraph>().ToList();
    var gneiting = paragraphs.FirstOrDefault(p => ParagraphText(p).StartsWith("GNEITING,", StringComparison.OrdinalIgnoreCase));
    var fontana = paragraphs.FirstOrDefault(p => ParagraphText(p).StartsWith("FONTANA,", StringComparison.OrdinalIgnoreCase));
    if (gneiting is null || fontana is null)
    {
        return false;
    }

    if (paragraphs.IndexOf(fontana) < paragraphs.IndexOf(gneiting))
    {
        return false;
    }

    fontana.Remove();
    gneiting.InsertBeforeSelf(fontana);
    return true;
}

static string BuildArticleAbntLayoutReport(
    string docxPath,
    string lockPath,
    string author,
    IReadOnlyList<string> applied,
    IReadOnlyList<string> skipped)
{
    var builder = new StringBuilder();
    builder.AppendLine("# Reparo ABNT/UERJ do artigo");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{docxPath}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine($"- Autor da etapa: `{author}`");
    builder.AppendLine("- Utilitario/codigo usado: `docx-utils`, comando `repair-article-abnt-layout` (.NET + Open XML)");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine();
    builder.AppendLine("## Aplicado");
    foreach (var item in applied)
    {
        builder.AppendLine($"- {item}");
    }

    builder.AppendLine();
    builder.AppendLine("## Nao aplicado / revisar");
    if (skipped.Count == 0)
    {
        builder.AppendLine("- Nenhum item.");
    }
    else
    {
        foreach (var item in skipped)
        {
            builder.AppendLine($"- {item}");
        }
    }

    builder.AppendLine();
    builder.AppendLine("## Observacoes");
    builder.AppendLine("- As figuras foram movidas para o paragrafo da legenda, com estilo `Figura`, centralizacao e wrap `topBottom`.");
    builder.AppendLine("- As tabelas duplicadas anteriores foram removidas mantendo a segunda tabela de cada par, conforme a revisao visual do artigo.");
    builder.AppendLine("- Tabelas remanescentes usam estilo `tabelauerj`, largura 100% e `autofit`; paragrafos internos usam estilo `dados`.");
    builder.AppendLine("- As fontes permanecem em paragrafos proprios abaixo de figuras/tabelas, com estilo `legenda`.");
    return builder.ToString();
}

static int RepairRefNumberOnly(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("lock", out var lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils repair-ref-number-only {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        RemoveUpdateFieldsOnOpen(doc);
        var refFieldsUpdated = 0;
        var resultRunsUpdated = 0;
        foreach (var beginRun in doc.MainDocumentPart!.Document.Descendants<Run>().Where(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin).ToList())
        {
            var fieldRuns = CollectComplexFieldRuns(beginRun);
            if (fieldRuns.Count == 0)
            {
                continue;
            }

            var instruction = string.Concat(fieldRuns.SelectMany(r => r.Elements<FieldCode>()).Select(c => c.Text));
            var match = Regex.Match(instruction, @"\bREF\s+(xref_(fig|tab)_\d+)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                continue;
            }

            var normalizedInstruction = Regex.Replace(instruction, "\\s+\\\\r\\b", " \\\\n", RegexOptions.CultureInvariant);
            if (!Regex.IsMatch(normalizedInstruction, "(^|\\s)\\\\n(\\s|$)", RegexOptions.CultureInvariant))
            {
                normalizedInstruction = Regex.Replace(normalizedInstruction, "\\s+\\\\h\\b", " \\\\n \\\\h", RegexOptions.CultureInvariant);
            }

            foreach (var code in fieldRuns.SelectMany(r => r.Elements<FieldCode>()))
            {
                code.Text = normalizedInstruction;
                break;
            }

            refFieldsUpdated++;

            var resultText = string.Concat(fieldRuns
                .SkipWhile(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
                .Skip(1)
                .TakeWhile(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.End)
                .SelectMany(r => r.Descendants<Text>())
                .Select(t => t.Text));
            var number = Regex.Match(resultText, @"\d+").Value;
            if (!string.IsNullOrWhiteSpace(number))
            {
                var resultTexts = fieldRuns
                    .SkipWhile(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
                    .Skip(1)
                    .TakeWhile(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.End)
                    .SelectMany(r => r.Descendants<Text>())
                    .ToList();
                for (var i = 0; i < resultTexts.Count; i++)
                {
                    resultTexts[i].Text = i == 0 ? number : "";
                    resultTexts[i].Space = SpaceProcessingModeValues.Preserve;
                }

                resultRunsUpdated++;
            }
        }

        applied.Add($"updated REF fields from label-returning switches to number-only switch: {refFieldsUpdated}");
        applied.Add($"updated cached REF results to numeric text: {resultRunsUpdated}");
        doc.MainDocumentPart.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        File.WriteAllText(reportPath, BuildRefNumberOnlyReport(docxPath, lockPath, author, applied), Encoding.UTF8);
    }

    foreach (var item in applied)
    {
        Console.WriteLine($"APPLY {item}");
    }

    return 0;
}

static List<Paragraph> GetAllParagraphs(WordprocessingDocument doc)
{
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    return body.Elements<Paragraph>().ToList();
}

static Paragraph? FindCaptionParagraph(IReadOnlyList<Paragraph> paragraphs, CrossrefCaptionSpec caption)
{
    var matches = paragraphs
        .Where(p => string.Equals(ParagraphStyleId(p), caption.StyleId, StringComparison.OrdinalIgnoreCase))
        .Where(p => Normalize(ParagraphText(p)).Equals(Normalize(caption.AnchorText), StringComparison.Ordinal))
        .ToList();
    if (caption.Occurrence is int occurrence && occurrence > 0 && matches.Count >= occurrence)
    {
        return matches[occurrence - 1];
    }

    return matches.Count == 1 ? matches[0] : null;
}

static Paragraph? FindCaptionParagraphAllowingManualPrefix(IReadOnlyList<Paragraph> paragraphs, CrossrefCaptionSpec caption)
{
    var matches = paragraphs
        .Where(p => string.Equals(ParagraphStyleId(p), caption.StyleId, StringComparison.OrdinalIgnoreCase))
        .Where(p =>
        {
            var text = StripManualCaptionPrefix(ParagraphText(p), caption);
            return Normalize(text).Equals(Normalize(caption.AnchorText), StringComparison.Ordinal);
        })
        .ToList();
    if (caption.Occurrence is int occurrence && occurrence > 0 && matches.Count >= occurrence)
    {
        return matches[occurrence - 1];
    }

    return matches.Count == 1 ? matches[0] : null;
}

static string StripManualCaptionPrefix(string text, CrossrefCaptionSpec caption) =>
    Regex.Replace(
        text,
        $"^{Regex.Escape(caption.Label)}\\s+\\d+\\s*[-â€“â€”]\\s*",
        "",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

static bool RemoveLeadingManualCaptionPrefix(Paragraph paragraph, CrossrefCaptionSpec caption)
{
    var firstContent = paragraph.Elements<OpenXmlElement>()
        .FirstOrDefault(e => e is not ParagraphProperties and not BookmarkStart and not BookmarkEnd);
    if (firstContent is InsertedRun inserted && Regex.IsMatch(ElementText(inserted), $"^{Regex.Escape(caption.Label)}\\s+\\d+\\s*[-â€“â€”]\\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
    {
        inserted.Remove();
        return true;
    }

    return false;
}

static void RemoveBookmarkByName(WordprocessingDocument doc, string bookmarkName)
{
    var starts = doc.MainDocumentPart?.Document.Descendants<BookmarkStart>()
        .Where(b => string.Equals(b.Name?.Value, bookmarkName, StringComparison.Ordinal))
        .ToList() ?? [];
    foreach (var start in starts)
    {
        var id = start.Id?.Value;
        start.Remove();
        if (id is null)
        {
            continue;
        }

        foreach (var end in doc.MainDocumentPart!.Document.Descendants<BookmarkEnd>().Where(b => b.Id?.Value == id).ToList())
        {
            end.Remove();
        }
    }
}

static void AddParagraphBookmark(Paragraph paragraph, string bookmarkName, ref int bookmarkId)
{
    var id = bookmarkId.ToString(CultureInfo.InvariantCulture);
    bookmarkId++;

    var start = new BookmarkStart { Id = id, Name = bookmarkName };
    var end = new BookmarkEnd { Id = id };
    var first = paragraph.Elements<OpenXmlElement>().FirstOrDefault(e => e is not ParagraphProperties);
    if (first is null)
    {
        paragraph.Append(start);
        paragraph.Append(end);
        return;
    }

    paragraph.InsertBefore(start, first);
    paragraph.Append(end);
}

static string EffectiveNumberingSummary(WordprocessingDocument doc, Paragraph paragraph)
{
    var props = paragraph.ParagraphProperties;
    var numbering = props?.NumberingProperties;
    var styleId = ParagraphStyleId(paragraph);
    var style = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
        .Elements<Style>()
        .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
    var styleNumbering = style?.StyleParagraphProperties?.NumberingProperties;
    var numberingId = numbering?.NumberingId?.Val is null
        ? ""
        : numbering.NumberingId.Val.Value.ToString(CultureInfo.InvariantCulture);
    var numberingLevel = numbering?.NumberingLevelReference?.Val is null
        ? ""
        : numbering.NumberingLevelReference.Val.Value.ToString(CultureInfo.InvariantCulture);
    var styleNumberingId = styleNumbering?.NumberingId?.Val is null
        ? ""
        : styleNumbering.NumberingId.Val.Value.ToString(CultureInfo.InvariantCulture);
    var styleNumberingLevel = styleNumbering?.NumberingLevelReference?.Val is null
        ? ""
        : styleNumbering.NumberingLevelReference.Val.Value.ToString(CultureInfo.InvariantCulture);
    var effectiveNumberingId = !string.IsNullOrWhiteSpace(numberingId) ? numberingId : styleNumberingId;
    var effectiveNumberingLevel = !string.IsNullOrWhiteSpace(numberingLevel) ? numberingLevel : styleNumberingLevel;
    var format = ResolveNumberingFormat(doc, effectiveNumberingId, effectiveNumberingLevel);
    return string.IsNullOrWhiteSpace(format) ? "" : $"effectiveNumberingId={effectiveNumberingId}; {format}";
}

static bool BookmarkExists(WordprocessingDocument doc, string name) =>
    doc.MainDocumentPart?.Document.Descendants<BookmarkStart>()
        .Any(b => string.Equals(b.Name?.Value, name, StringComparison.Ordinal)) == true;

static int NextBookmarkId(WordprocessingDocument doc)
{
    var max = doc.MainDocumentPart?.Document.Descendants<BookmarkStart>()
        .Select(b => int.TryParse(b.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : 0)
        .DefaultIfEmpty(0)
        .Max() ?? 0;
    return max + 1;
}

static void InsertCaptionPrefix(Paragraph paragraph, CrossrefCaptionSpec caption, RevisionMetadata metadata, ref int bookmarkId)
{
    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    insertion.Append(CreateTextRun($"{caption.Label} "));
    insertion.Append(new BookmarkStart { Id = bookmarkId.ToString(CultureInfo.InvariantCulture), Name = caption.Bookmark });
    foreach (var run in CreateComplexFieldRuns($" SEQ {caption.Sequence} \\r {caption.Number} \\* ARABIC ", caption.Number.ToString(CultureInfo.InvariantCulture)))
    {
        insertion.Append(run);
    }

    insertion.Append(new BookmarkEnd { Id = bookmarkId.ToString(CultureInfo.InvariantCulture) });
    bookmarkId++;
    insertion.Append(CreateTextRun(" - "));

    var first = paragraph.Elements<OpenXmlElement>().FirstOrDefault(e => e is not ParagraphProperties);
    if (first is null)
    {
        paragraph.Append(insertion);
    }
    else
    {
        paragraph.InsertBefore(insertion, first);
    }
}

static IReadOnlyList<CrossrefRunPart>? BuildPlannedParagraph(
    string originalText,
    CrossrefParagraphReplacement replacement,
    ISet<string> bookmarkNames,
    List<string> skipped)
{
    var pieces = new List<CrossrefRunPart>();
    var cursor = 0;
    foreach (var item in replacement.Replacements)
    {
        var index = originalText.IndexOf(item.OldText, cursor, StringComparison.Ordinal);
        if (index < 0)
        {
            skipped.Add($"{replacement.Id}: old text not found after cursor ({item.OldText})");
            return null;
        }

        if (index > cursor)
        {
            pieces.Add(CrossrefRunPart.TextPart(originalText[cursor..index]));
        }

        foreach (var part in item.Parts)
        {
            if (!string.IsNullOrEmpty(part.Ref) && !bookmarkNames.Contains(part.Ref))
            {
                skipped.Add($"{replacement.Id}: missing bookmark {part.Ref}");
                return null;
            }

            pieces.Add(part);
        }

        cursor = index + item.OldText.Length;
    }

    if (cursor < originalText.Length)
    {
        pieces.Add(CrossrefRunPart.TextPart(originalText[cursor..]));
    }

    return pieces;
}

static void ReplaceParagraphTracked(
    Paragraph paragraph,
    string originalText,
    IReadOnlyList<CrossrefRunPart> newParts,
    RevisionMetadata metadata)
{
    var properties = paragraph.ParagraphProperties?.CloneNode(true);
    var directChildren = paragraph.ChildElements.ToList();
    var replaceableIndices = directChildren
        .Select((child, index) => new { child, index })
        .Where(x => IsReplaceableParagraphChild(x.child))
        .Select(x => x.index)
        .ToList();

    var leadingPreserved = new List<OpenXmlElement>();
    var trailingPreserved = new List<OpenXmlElement>();

    if (replaceableIndices.Count > 0)
    {
        var firstReplaceable = replaceableIndices.Min();
        var lastReplaceable = replaceableIndices.Max();

        leadingPreserved.AddRange(directChildren
            .Take(firstReplaceable)
            .Where(child => child is not ParagraphProperties)
            .Select(child => child.CloneNode(true)));

        trailingPreserved.AddRange(directChildren
            .Skip(lastReplaceable + 1)
            .Where(child => child is not ParagraphProperties)
            .Select(child => child.CloneNode(true)));
    }

    paragraph.RemoveAllChildren();
    if (properties is not null)
    {
        paragraph.Append(properties);
    }

    foreach (var child in leadingPreserved)
    {
        paragraph.Append(child);
    }

    var deletion = new DeletedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };
    var deletedRun = new Run();
    deletedRun.Append(new DeletedText(originalText) { Space = SpaceProcessingModeValues.Preserve });
    deletion.Append(deletedRun);
    paragraph.Append(deletion);

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    foreach (var part in newParts)
    {
        if (!string.IsNullOrEmpty(part.Text))
        {
            insertion.Append(CreateTextRun(part.Text));
        }
        else if (!string.IsNullOrEmpty(part.Latex))
        {
            insertion.Append(CreateProfessionalOfficeMathFromLatex(
                new FormulaSpec(part.Latex, part.Latex, "", false, null)));
        }
        else if (!string.IsNullOrEmpty(part.Ref))
        {
            var instruction = !string.IsNullOrWhiteSpace(part.FieldInstruction)
                ? part.FieldInstruction!
                : $" REF {part.Ref} \\h ";
            foreach (var run in CreateComplexFieldRuns(instruction, part.Result ?? ""))
            {
                insertion.Append(run);
            }
        }
    }

    paragraph.Append(insertion);

    foreach (var child in trailingPreserved)
    {
        paragraph.Append(child);
    }
}

static bool IsReplaceableParagraphChild(OpenXmlElement child)
{
    if (child is ParagraphProperties
        || child is CommentRangeStart
        || child is CommentRangeEnd
        || child is BookmarkStart
        || child is BookmarkEnd)
    {
        return false;
    }

    if (child is Run run)
    {
        return !IsCommentReferenceOnlyRun(run);
    }

    return child is InsertedRun
        || child is DeletedRun
        || child is Hyperlink
        || child is SimpleField
        || child is SdtRun;
}

static bool IsCommentReferenceOnlyRun(Run run)
{
    if (!run.Descendants<CommentReference>().Any())
    {
        return false;
    }

    var hasMeaningfulText = run.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text));
    var hasFieldCode = run.Descendants<FieldCode>().Any(fc => !string.IsNullOrWhiteSpace(fc.Text));
    var hasDeletedText = run.Descendants<DeletedText>().Any(dt => !string.IsNullOrWhiteSpace(dt.Text));
    return !hasMeaningfulText && !hasFieldCode && !hasDeletedText;
}

static Run CreateTextRun(string text)
{
    var run = new Run();
    run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    return run;
}

static IEnumerable<Run> CreateComplexFieldRuns(string instruction, string result)
{
    var begin = new Run();
    begin.Append(new FieldChar { FieldCharType = FieldCharValues.Begin });
    yield return begin;

    var instr = new Run();
    instr.Append(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve });
    yield return instr;

    var separate = new Run();
    separate.Append(new FieldChar { FieldCharType = FieldCharValues.Separate });
    yield return separate;

    yield return CreateTextRun(result);

    var end = new Run();
    end.Append(new FieldChar { FieldCharType = FieldCharValues.End });
    yield return end;
}

static string BuildCrossrefReport(
    string docxPath,
    CrossrefPlan plan,
    string lockPath,
    string author,
    IReadOnlyList<string> applied,
    IReadOnlyList<string> skipped)
{
    var builder = new StringBuilder();
    builder.AppendLine("# RelatÃ³rio de aplicaÃ§Ã£o de referÃªncias cruzadas");
    builder.AppendLine();
    builder.AppendLine($"- VersÃ£o de entrada: `{plan.InputPath}`");
    builder.AppendLine($"- VersÃ£o de saÃ­da: `{docxPath}`");
    builder.AppendLine("- UtilitÃ¡rio/cÃ³digo usado: `docx-utils`, comando `apply-crossrefs` (.NET + Open XML)");
    builder.AppendLine($"- Plano usado: `{plan.PlanPath}`");
    builder.AppendLine($"- Lock usado: `{lockPath}`");
    builder.AppendLine($"- Autor das revisÃµes: `{author}`");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine();
    builder.AppendLine("## Chamadas convertidas e alvos criados");
    foreach (var item in applied)
    {
        builder.AppendLine($"- {item}");
    }

    builder.AppendLine();
    builder.AppendLine("## Chamadas puladas por ambiguidade ou validaÃ§Ã£o");
    if (skipped.Count == 0)
    {
        builder.AppendLine("- Nenhuma.");
    }
    else
    {
        foreach (var item in skipped)
        {
            builder.AppendLine($"- {item}");
        }
    }

    builder.AppendLine();
    builder.AppendLine("## ValidaÃ§Ã£o executada");
    builder.AppendLine("- Executar apÃ³s a aplicaÃ§Ã£o: `docx-utils validate <saida.docx>`.");
    builder.AppendLine("- Executar apÃ³s a aplicaÃ§Ã£o: `docx-utils structure-audit <saida.docx> --out <json>`.");
    builder.AppendLine();
    builder.AppendLine("## Riscos remanescentes");
    builder.AppendLine("- O arquivo de entrada nÃ£o continha campos `SEQ` em legendas de tabela/figura; o utilitÃ¡rio inseriu campos novos apenas nas legendas planejadas e revalidadas.");
    builder.AppendLine("- Campos `SEQ`/`REF` foram gravados com resultado atual no XML; o Word pode recalcular visualmente os campos ao atualizar o documento.");
    builder.AppendLine("- A Lista de Tabelas e a Lista de IlustraÃ§Ãµes nÃ£o foram atualizadas nesta rodada.");
    return builder.ToString();
}

static string BuildStyleCaptionRepairReport(
    string docxPath,
    CrossrefPlan plan,
    string lockPath,
    string author,
    IReadOnlyList<string> applied,
    IReadOnlyList<string> skipped)
{
    var builder = new StringBuilder();
    builder.AppendLine("# Reparo de legendas numeradas por estilo");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{docxPath}`");
    builder.AppendLine($"- Plano: `{plan.PlanPath}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine($"- Autor das revisoes: `{author}`");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine();
    builder.AppendLine("## Padrao aplicado");
    builder.AppendLine();
    builder.AppendLine("- Legendas de figuras e tabelas permanecem em paragrafos com estilos `Figura` e `Tabela`.");
    builder.AppendLine("- A numeracao visivel vem da numeracao automatica do estilo, nao de prefixo textual manual nem de campo `SEQ` inserido na legenda.");
    builder.AppendLine("- Bookmarks `xref_fig_*` e `xref_tab_*` foram reposicionados no paragrafo de legenda numerado.");
    builder.AppendLine("- Campos `REF` existentes foram ajustados com `\\r` para referenciar o numero do paragrafo numerado.");
    builder.AppendLine("- O documento foi marcado com `w:updateFields`, para recalcular listas automaticas e referencias ao abrir no Word.");
    builder.AppendLine();
    builder.AppendLine("## Aplicado");
    foreach (var item in applied)
    {
        builder.AppendLine($"- {item}");
    }

    if (skipped.Count > 0)
    {
        builder.AppendLine();
        builder.AppendLine("## Nao aplicado / revisar");
        foreach (var item in skipped)
        {
            builder.AppendLine($"- {item}");
        }
    }

    return builder.ToString();
}

static Table? FindTableAfterCaption(IReadOnlyList<OpenXmlElement> blocks, string captionText)
{
    for (var i = 0; i < blocks.Count - 1; i++)
    {
        if (blocks[i] is not Paragraph paragraph)
        {
            continue;
        }

        if (!string.Equals(ParagraphStyleId(paragraph), "Tabela", StringComparison.OrdinalIgnoreCase))
        {
            continue;
        }

        if (!Normalize(ParagraphText(paragraph)).Equals(Normalize(captionText), StringComparison.Ordinal))
        {
            continue;
        }

        return blocks.Skip(i + 1).OfType<Table>().FirstOrDefault();
    }

    return null;
}

static TableReplacementSelectionResult SelectTableForReplacement(Body body, ReplaceTableSpec spec)
{
    var blocks = body.ChildElements
        .Where(e => e is Paragraph or Table)
        .ToList();

    var tables = new List<TableReplacementSelectionResult>();
    var ordinal = 0;
    for (var i = 0; i < blocks.Count; i++)
    {
        if (blocks[i] is not Table table)
        {
            continue;
        }

        ordinal++;
        var previousParagraphText = string.Empty;
        for (var j = i - 1; j >= 0; j--)
        {
            if (blocks[j] is Paragraph paragraph)
            {
                previousParagraphText = ParagraphText(paragraph);
                break;
            }
        }

        var nextParagraphText = string.Empty;
        for (var j = i + 1; j < blocks.Count; j++)
        {
            if (blocks[j] is Paragraph paragraph)
            {
                nextParagraphText = ParagraphText(paragraph);
                break;
            }
        }

        tables.Add(new TableReplacementSelectionResult(
            true,
            "",
            table,
            ordinal,
            i + 1,
            FirstNonEmptyTableText(table),
            previousParagraphText,
            nextParagraphText,
            ResolveReplacementTableColumnWidths(table, GetCurrentTableColumnCount(table)),
            ""));
    }

    IEnumerable<TableReplacementSelectionResult> matches = tables;
    if (spec.Ordinal is int ordinalSelector and > 0)
    {
        matches = matches.Where(t => t.Ordinal == ordinalSelector);
    }

    if (spec.BlockIndex is int blockIndexSelector and > 0)
    {
        matches = matches.Where(t => t.BlockIndex == blockIndexSelector);
    }

    if (spec.Block is int blockSelector and > 0)
    {
        matches = matches.Where(t => t.BlockIndex == blockSelector);
    }

    if (!string.IsNullOrWhiteSpace(spec.FirstCellText))
    {
        matches = matches.Where(t => NormalizedStartsWith(t.FirstCellText, spec.FirstCellText));
    }

    if (!string.IsNullOrWhiteSpace(spec.PreviousParagraphPrefix))
    {
        matches = matches.Where(t => NormalizedStartsWith(t.PreviousParagraphText, spec.PreviousParagraphPrefix));
    }

    if (!string.IsNullOrWhiteSpace(spec.NextParagraphPrefix))
    {
        matches = matches.Where(t => NormalizedStartsWith(t.NextParagraphText, spec.NextParagraphPrefix));
    }

    var filtered = matches.ToList();
    if (filtered.Count == 0)
    {
        return new TableReplacementSelectionResult(false, BuildReplacementTableNotFoundReason(spec, tables), null, 0, 0, "", "", "", [], "");
    }

    if (filtered.Count > 1)
    {
        return new TableReplacementSelectionResult(false, BuildReplacementTableAmbiguousReason(spec, filtered), null, 0, 0, "", "", "", [], "");
    }

    var selected = filtered[0];
    return new TableReplacementSelectionResult(
        true,
        "",
        selected.Table,
        selected.Ordinal,
        selected.BlockIndex,
        selected.FirstCellText,
        selected.PreviousParagraphText,
        selected.NextParagraphText,
        selected.ColumnWidths,
        ResolveReplacementCellStyleId(selected.Table!, spec));
}

static string BuildReplacementTableNotFoundReason(ReplaceTableSpec spec, IReadOnlyList<TableReplacementSelectionResult> tables)
{
    var selector = DescribeTableSelector(spec);
    if (tables.Count == 0)
    {
        return $"{selector}: no tables found";
    }

    return $"{selector}: no matching table";
}

static string BuildReplacementTableAmbiguousReason(ReplaceTableSpec spec, IReadOnlyList<TableReplacementSelectionResult> tables)
{
    var selector = DescribeTableSelector(spec);
    return $"{selector}: matched {tables.Count} tables";
}

static string DescribeTableSelector(ReplaceTableSpec spec)
{
    var parts = new List<string>();
    if (spec.Ordinal is int ordinal and > 0)
    {
        parts.Add($"ordinal={ordinal}");
    }
    if (spec.BlockIndex is int blockIndex and > 0)
    {
        parts.Add($"blockIndex={blockIndex}");
    }
    if (spec.Block is int block and > 0)
    {
        parts.Add($"block={block}");
    }
    if (!string.IsNullOrWhiteSpace(spec.FirstCellText))
    {
        parts.Add($"firstCellText={spec.FirstCellText}");
    }
    if (!string.IsNullOrWhiteSpace(spec.PreviousParagraphPrefix))
    {
        parts.Add($"previousParagraphPrefix={spec.PreviousParagraphPrefix}");
    }
    if (!string.IsNullOrWhiteSpace(spec.NextParagraphPrefix))
    {
        parts.Add($"nextParagraphPrefix={spec.NextParagraphPrefix}");
    }

    return parts.Count == 0 ? "selector" : $"selector({string.Join(", ", parts)})";
}

static IReadOnlyList<int> ResolveReplacementTableColumnWidths(Table table, int columnCount)
{
    if (columnCount <= 0)
    {
        return [];
    }

    var widths = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
        .Select(c => ParseTableWidth(c.Width?.Value))
        .Where(width => width > 0)
        .ToList() ?? [];

    if (widths.Count == 0)
    {
        var firstRow = table.Elements<TableRow>().FirstOrDefault();
        if (firstRow is not null)
        {
            widths = firstRow.Elements<TableCell>()
                .Select(cell => ParseTableWidth(cell.TableCellProperties?.TableCellWidth?.Width?.Value))
                .Where(width => width > 0)
                .ToList();
        }
    }

    if (widths.Count == 0)
    {
        return Enumerable.Repeat(Math.Max(1200, 9000 / Math.Max(1, columnCount)), columnCount).ToList();
    }

    if (widths.Count == columnCount)
    {
        return widths;
    }

    if (widths.Count > columnCount)
    {
        return widths.Take(columnCount).ToList();
    }

    var extended = new List<int>(widths);
    var fillWidth = widths.LastOrDefault(Math.Max(1200, 9000 / Math.Max(1, columnCount)));
    while (extended.Count < columnCount)
    {
        extended.Add(fillWidth);
    }

    return extended;
}

static int ParseTableWidth(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
    {
        return 0;
    }

    return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
        ? parsed
        : 0;
}

static int GetCurrentTableColumnCount(Table table)
{
    var row = table.Elements<TableRow>().FirstOrDefault();
    if (row is null)
    {
        return table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
    }

    var columns = row.Elements<TableCell>()
        .Sum(cell => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
    if (columns > 0)
    {
        return columns;
    }

    return table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
}

static string ResolveReplacementCellStyleId(Table table, ReplaceTableSpec spec)
{
    if (!string.IsNullOrWhiteSpace(spec.CellStyleId))
    {
        return spec.CellStyleId;
    }

    return table.Descendants<Paragraph>()
        .Select(ParagraphStyleId)
        .FirstOrDefault(style => !string.IsNullOrWhiteSpace(style))
        ?? "";
}

static void ReplaceTableRows(Table table, ReplaceTableSpec spec, IReadOnlyList<int> columnWidths, string cellStyleId)
{
    var columnCount = spec.Rows.Select(row => row?.Count ?? 0).DefaultIfEmpty(0).Max();
    var effectiveWidths = columnWidths.Count == columnCount
        ? columnWidths
        : ResolveReplacementTableColumnWidths(table, columnCount);

    foreach (var row in table.Elements<TableRow>().ToList())
    {
        row.Remove();
    }

    foreach (var oldGrid in table.Elements<TableGrid>().ToList())
    {
        oldGrid.Remove();
    }

    var grid = new TableGrid();
    foreach (var width in effectiveWidths)
    {
        grid.Append(new GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) });
    }

    var properties = table.GetFirstChild<TableProperties>();
    if (properties is not null)
    {
        table.InsertAfter(grid, properties);
    }
    else
    {
        table.PrependChild(grid);
    }

    foreach (var rowValues in spec.Rows)
    {
        var row = new TableRow();
        for (var i = 0; i < columnCount; i++)
        {
            var width = effectiveWidths.Count > i ? effectiveWidths[i] : effectiveWidths.LastOrDefault(2400);
            var cell = new TableCell(new TableCellProperties(new TableCellWidth
            {
                Width = width.ToString(CultureInfo.InvariantCulture),
                Type = TableWidthUnitValues.Dxa
            }));
            var paragraph = new Paragraph();
            if (!string.IsNullOrWhiteSpace(cellStyleId))
            {
                paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = cellStyleId }));
            }
            NormalizeTableParagraph(paragraph);
            paragraph.Append(CreateTextRun(i < rowValues.Count ? rowValues[i] ?? "" : ""));
            cell.Append(paragraph);
            row.Append(cell);
        }

        table.Append(row);
    }
}

static int ApplyParagraphStyleInsideTable(Table table, string styleId)
{
    var changed = 0;
    foreach (var paragraph in table.Descendants<Paragraph>())
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        var current = paragraph.ParagraphProperties.ParagraphStyleId;
        if (string.Equals(current?.Val?.Value, styleId, StringComparison.Ordinal))
        {
            NormalizeTableParagraph(paragraph);
            continue;
        }

        if (current is null)
        {
            paragraph.ParagraphProperties.PrependChild(new ParagraphStyleId { Val = styleId });
        }
        else
        {
            current.Val = styleId;
        }

        NormalizeTableParagraph(paragraph);
        changed++;
    }

    return changed;
}

static bool RemoveSourceTextFromCaption(Paragraph paragraph, string sourceText, RevisionMetadata metadata)
{
    foreach (var text in paragraph.Descendants<Text>().ToList())
    {
        var index = text.Text.IndexOf(sourceText, StringComparison.Ordinal);
        if (index < 0)
        {
            continue;
        }

        var before = text.Text[..index].TrimEnd();
        text.Text = before;
        text.Space = SpaceProcessingModeValues.Preserve;

        var run = text.Ancestors<Run>().FirstOrDefault();
        if (run?.Parent is OpenXmlElement parent)
        {
            var deletion = new DeletedRun
            {
                Id = metadata.NextRevisionId(),
                Author = metadata.Author,
                Date = metadata.DateUtc
            };
            var deletedRun = new Run();
            deletedRun.Append(new DeletedText(sourceText) { Space = SpaceProcessingModeValues.Preserve });
            deletion.Append(deletedRun);
            parent.InsertAfter(deletion, run);
        }

        return true;
    }

    var fullText = ParagraphText(paragraph);
    if (!fullText.Contains(sourceText, StringComparison.Ordinal))
    {
        return false;
    }

    var firstText = paragraph.Descendants<Text>().FirstOrDefault();
    if (firstText is null)
    {
        return false;
    }

    firstText.Text = fullText.Replace(sourceText, "", StringComparison.Ordinal).TrimEnd();
    firstText.Space = SpaceProcessingModeValues.Preserve;
    return true;
}

static Paragraph CreateTrackedSourceParagraph(string sourceText, string styleId, RevisionMetadata metadata)
{
    var paragraph = new Paragraph();
    paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
    paragraph.ParagraphId = HexBinaryValue.FromString(RandomHex(8));
    paragraph.TextId = HexBinaryValue.FromString(RandomHex(8));

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };
    insertion.Append(CreateTextRun(sourceText));
    paragraph.Append(insertion);
    return paragraph;
}

static string BuildLayoutRepairReport(
    string docxPath,
    string lockPath,
    string author,
    IReadOnlyList<string> applied,
    IReadOnlyList<string> skipped)
{
    var builder = new StringBuilder();
    builder.AppendLine("# Reparo de pendencias de layout");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{docxPath}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine($"- Autor das revisoes: `{author}`");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine();
    builder.AppendLine("## Aplicado");
    foreach (var item in applied)
    {
        builder.AppendLine($"- {item}");
    }

    if (skipped.Count > 0)
    {
        builder.AppendLine();
        builder.AppendLine("## Nao aplicado / revisar");
        foreach (var item in skipped)
        {
            builder.AppendLine($"- {item}");
        }
    }

    return builder.ToString();
}

static List<Run> CollectComplexFieldRuns(Run beginRun)
{
    var parent = beginRun.Parent;
    if (parent is null)
    {
        return [];
    }

    var runs = parent.Elements<Run>().ToList();
    var start = runs.IndexOf(beginRun);
    if (start < 0)
    {
        return [];
    }

    var collected = new List<Run>();
    for (var i = start; i < runs.Count; i++)
    {
        collected.Add(runs[i]);
        if (runs[i].GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End)
        {
            break;
        }
    }

    return collected;
}

static string BuildRefNumberOnlyReport(
    string docxPath,
    string lockPath,
    string author,
    IReadOnlyList<string> applied)
{
    var builder = new StringBuilder();
    builder.AppendLine("# Reparo de referÃªncias cruzadas numÃ©ricas");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{docxPath}`");
    builder.AppendLine($"- Lock: `{lockPath}`");
    builder.AppendLine($"- Autor das revisoes: `{author}`");
    builder.AppendLine($"- Gerado em UTC: `{DateTime.UtcNow:O}`");
    builder.AppendLine();
    builder.AppendLine("## Aplicado");
    foreach (var item in applied)
    {
        builder.AppendLine($"- {item}");
    }

    builder.AppendLine();
    builder.AppendLine("## ObservaÃ§Ã£o");
    builder.AppendLine("- Campos `REF xref_fig_*` e `REF xref_tab_*` devem retornar apenas o nÃºmero da legenda, porque o texto corrido jÃ¡ contÃ©m `figura`, `figuras`, `tabela` ou `tabelas`.");
    return builder.ToString();
}

static int InsertFigures(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<FigureInsertionPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read figure insertion plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    var notes = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils insert-figures {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        var metadata = new RevisionMetadata(author, DateTime.UtcNow);
        var hasFiguraStyle = HasParagraphStyle(doc, "Figura");
        var existingFiguraCaptions = CountExistingFiguraCaptions(doc);

        foreach (var figure in plan.Figures)
        {
            if (string.IsNullOrWhiteSpace(figure.ImagePath) || !File.Exists(figure.ImagePath))
            {
                skipped.Add($"{figure.Id}: image not found ({figure.ImagePath})");
                continue;
            }

            var paragraphs = GetParagraphs(doc);
            ParagraphEntry after;
            ParagraphEntry before;
            try
            {
                after = FindUniqueParagraph(paragraphs, figure.AfterPrefix, $"{figure.Id}/after");
                before = FindUniqueParagraph(paragraphs, figure.BeforePrefix, $"{figure.Id}/before");
            }
            catch (InvalidOperationException ex)
            {
                skipped.Add($"{figure.Id}: anchor lookup failed ({ex.Message})");
                continue;
            }

            if (after.Index >= before.Index)
            {
                skipped.Add($"{figure.Id}: anchor order mismatch (after={after.Index}, before={before.Index})");
                continue;
            }

            var captionPrefix = BuildPresencePrefix(figure.CaptionText ?? "");
            if (!string.IsNullOrWhiteSpace(captionPrefix))
            {
                var alreadyHasCaption = paragraphs
                    .Where(p => p.Index > after.Index && p.Index < before.Index)
                    .Any(p => NormalizedStartsWith(p.Text, captionPrefix));
                if (alreadyHasCaption)
                {
                    skipped.Add($"{figure.Id}: caption already present between anchors");
                    continue;
                }
            }

            var imagePart = doc.MainDocumentPart!.AddImagePart(ImagePartType.Png);
            using (var imageStream = File.OpenRead(figure.ImagePath))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = doc.MainDocumentPart!.GetIdOfPart(imagePart);

            var pixelDimensions = TryReadPngDimensions(figure.ImagePath);
            var widthCm = figure.WidthCm > 0 ? figure.WidthCm : 14.0;
            var widthEmu = (long)Math.Round(widthCm * 360000.0);
            long heightEmu;
            if (pixelDimensions is { Width: > 0, Height: > 0 } px)
            {
                var aspect = (double)px.Height / px.Width;
                heightEmu = (long)Math.Round(widthEmu * aspect);
            }
            else
            {
                // Fallback: assume 4:3 aspect ratio if PNG header could not be parsed.
                heightEmu = (long)Math.Round(widthEmu * 0.75);
                notes.Add($"{figure.Id}: PNG dimensions unavailable, used 4:3 fallback");
            }

            existingFiguraCaptions++;
            var captionNumber = existingFiguraCaptions;

            var imageParagraph = CreateInsertedFigureParagraph(
                relationshipId,
                figure.Id,
                figure.CaptionText ?? "",
                widthEmu,
                heightEmu,
                hasFiguraStyle,
                metadata);
            var captionParagraph = CreateInsertedFigureCaption(
                figure.CaptionText ?? "",
                captionNumber,
                hasFiguraStyle,
                metadata);
            var sourceParagraph = CreateInsertedFigureSource(
                figure.SourceText ?? "",
                metadata);

            after.Paragraph.InsertAfterSelf(imageParagraph);
            imageParagraph.InsertAfterSelf(captionParagraph);
            captionParagraph.InsertAfterSelf(sourceParagraph);

            applied.Add($"{figure.Id}: inserted after P[{after.Index}] caption #{captionNumber}");
        }

        doc.MainDocumentPart!.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Insert Figures Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        builder.AppendLine();
        builder.AppendLine("## Aplicado");
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        if (skipped.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("## Ignorado / revisar");
            foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        }
        if (notes.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("## ObservaÃ§Ãµes");
            foreach (var item in notes) builder.AppendLine($"- {item}");
            builder.AppendLine("- NumeraÃ§Ã£o de legendas usa contagem literal (parÃ¡grafos com estilo `Figura` jÃ¡ existentes + 1). Caso o documento utilize campos `SEQ` automÃ¡ticos, abrir no Word e atualizar campos para reconciliar numeraÃ§Ã£o.");
        }
        else
        {
            builder.AppendLine();
            builder.AppendLine("## ObservaÃ§Ãµes");
            builder.AppendLine("- NumeraÃ§Ã£o de legendas usa contagem literal (parÃ¡grafos com estilo `Figura` jÃ¡ existentes + 1). Caso o documento utilize campos `SEQ` automÃ¡ticos, abrir no Word e atualizar campos para reconciliar numeraÃ§Ã£o.");
        }
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int InsertComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CommentInsertionPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read comment insertion plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils insert-comments {DateTime.UtcNow:O}\n"));
    lockStream.Flush();

    WordprocessingDocument doc;
    try
    {
        // Reopen the DOCX only after the exclusive lock is held.
        doc = WordprocessingDocument.Open(docxPath, true);
    }
    catch (IOException ex)
    {
        Console.Error.WriteLine($"Unable to open DOCX for writing: {ex.Message}");
        return 8;
    }

    using (doc)
    {
        EnsureTrackRevisions(doc);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new Comments();
        var nextCommentId = NextCommentId(commentsPart.Comments);
        var nowUtc = DateTime.UtcNow;

        foreach (var entry in plan.Comments)
        {
            if (string.IsNullOrWhiteSpace(entry.AnchorPrefix))
            {
                skipped.Add($"{entry.Id}: missing anchorPrefix");
                continue;
            }

            var paragraphs = GetParagraphs(doc);
            ParagraphEntry target;
            try
            {
                target = FindUniqueParagraph(paragraphs, entry.AnchorPrefix, $"{entry.Id}/anchor");
            }
            catch (InvalidOperationException ex)
            {
                skipped.Add($"{entry.Id}: anchor lookup failed ({ex.Message})");
                continue;
            }

            var commentId = nextCommentId++.ToString(CultureInfo.InvariantCulture);

            var commentElement = new Comment
            {
                Id = commentId,
                Author = author,
                Date = nowUtc,
            };
            ApplyCommentInitials(commentElement);
            var commentBody = new Paragraph();
            var commentRun = new Run();
            commentRun.Append(new Text(entry.CommentText ?? "") { Space = SpaceProcessingModeValues.Preserve });
            commentBody.Append(commentRun);
            commentElement.Append(commentBody);
            commentsPart.Comments!.Append(commentElement);

            var rangeStart = new CommentRangeStart { Id = commentId };
            var rangeEnd = new CommentRangeEnd { Id = commentId };
            var referenceRun = new Run();
            referenceRun.Append(new RunProperties(new RunStyle { Val = "CommentReference" }));
            referenceRun.Append(new CommentReference { Id = commentId });

            if (!string.IsNullOrWhiteSpace(entry.AnchorContains)
                && TrySplitRunsForAnchor(target.Paragraph, entry.AnchorContains!, out var splitFirstRun, out var splitLastRun))
            {
                splitFirstRun!.InsertBeforeSelf(rangeStart);
                splitLastRun!.InsertAfterSelf(rangeEnd);
                rangeEnd.InsertAfterSelf(referenceRun);
            }
            else
            {
                var firstChild = target.Paragraph.Elements<OpenXmlElement>().FirstOrDefault(e => e is not ParagraphProperties);
                if (firstChild is null)
                {
                    target.Paragraph.Append(rangeStart);
                    target.Paragraph.Append(rangeEnd);
                    target.Paragraph.Append(referenceRun);
                }
                else
                {
                    target.Paragraph.InsertBefore(rangeStart, firstChild);
                    target.Paragraph.Append(rangeEnd);
                    target.Paragraph.Append(referenceRun);
                }
            }

            applied.Add($"{entry.Id}: comment id={commentId} on P[{target.Index}]");
        }

        commentsPart.Comments!.Save();
        mainPart.Document.Save();
    }

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Insert Comments Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        builder.AppendLine();
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static ParagraphEntry? FindEquationParagraphAfterAnchor(
    IReadOnlyList<ParagraphEntry> bodyParagraphs,
    ParagraphEntry anchor)
{
    var anchorPosition = -1;
    for (var index = 0; index < bodyParagraphs.Count; index++)
    {
        if (ReferenceEquals(bodyParagraphs[index].Paragraph, anchor.Paragraph))
        {
            anchorPosition = index;
            break;
        }
    }
    if (anchorPosition < 0)
    {
        return null;
    }

    for (var index = anchorPosition + 1; index < bodyParagraphs.Count; index++)
    {
        var candidate = bodyParagraphs[index];
        if (candidate.Paragraph.Ancestors<Table>().Any())
        {
            continue;
        }

        if (HasMath(candidate.Paragraph) || string.Equals(ParagraphStyleId(candidate.Paragraph), "equao", StringComparison.OrdinalIgnoreCase))
        {
            return candidate;
        }

        if (!string.IsNullOrWhiteSpace(candidate.Text))
        {
            break;
        }
    }

    return null;
}

static bool IsFollowingEquationAnchorMarker(string? anchorContains) =>
    string.Equals(anchorContains, "__FOLLOWING_EQUATION__", StringComparison.Ordinal);

static ParagraphEntry FindCommentAnchorParagraph(
    WordprocessingDocument doc,
    string anchorPrefix,
    string label,
    string? anchorContains = null,
    int? occurrence = null)
{
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    var bodyParagraphs = body.Elements<Paragraph>()
        .Select((p, index) => new ParagraphEntry(p, index, ParagraphText(p)))
        .ToList();
    var textParagraphs = bodyParagraphs
        .Where(p => !string.IsNullOrWhiteSpace(p.Text))
        .ToList();

    var directMatches = textParagraphs
        .Where(p => NormalizedStartsWith(p.Text, anchorPrefix))
        .ToList();
    if (directMatches.Count > 0)
    {
        var useFollowingEquation = IsFollowingEquationAnchorMarker(anchorContains);
        if (occurrence is int directOccurrence)
        {
            if (directOccurrence < 1 || directOccurrence > directMatches.Count)
            {
                throw new InvalidOperationException($"Anchor occurrence {directOccurrence} out of range for {label} (matches={directMatches.Count}).");
            }

            var selected = directMatches[directOccurrence - 1];
            if (!useFollowingEquation)
            {
                return selected;
            }

            return FindEquationParagraphAfterAnchor(bodyParagraphs, selected)
                ?? throw new InvalidOperationException($"No following equation paragraph found for {label}.");
        }

        if (directMatches.Count != 1)
        {
            throw new InvalidOperationException($"Expected exactly one match for {label}; got {directMatches.Count}.");
        }

        var direct = directMatches[0];
        if (!useFollowingEquation)
        {
            return direct;
        }

        return FindEquationParagraphAfterAnchor(bodyParagraphs, direct)
            ?? throw new InvalidOperationException($"No following equation paragraph found for {label}.");
    }

    var contextMatches = textParagraphs
        .Where(p => NormalizedStartsWith(p.Text, anchorPrefix))
        .Select(context => FindEquationParagraphAfterAnchor(bodyParagraphs, context))
        .Where(candidate => candidate is not null)
        .Select(candidate => candidate!)
        .DistinctBy(candidate => candidate.Index)
        .ToList();

    if (occurrence is int contextOccurrence)
    {
        if (contextOccurrence < 1 || contextOccurrence > contextMatches.Count)
        {
            throw new InvalidOperationException($"Equation-following anchor occurrence {contextOccurrence} out of range for {label} (matches={contextMatches.Count}).");
        }

        return contextMatches[contextOccurrence - 1];
    }

    if (contextMatches.Count != 1)
    {
        throw new InvalidOperationException($"Expected exactly one equation-following match for {label}; got {contextMatches.Count}.");
    }

    return contextMatches[0];
}

static bool ReanchorCommentOnParagraph(MainDocumentPart mainPart, string commentId, ParagraphEntry target, string? anchorContains)
{
    var alreadyAnchored = target.Paragraph.Descendants<CommentRangeStart>()
        .Any(c => c.Id?.Value == commentId);
    if (alreadyAnchored)
    {
        return false;
    }

    RemoveCommentMarkers(mainPart, commentId);

    var rangeStart = new CommentRangeStart { Id = commentId };
    var rangeEnd = new CommentRangeEnd { Id = commentId };
    var referenceRun = new Run();
    referenceRun.Append(new RunProperties(new RunStyle { Val = "CommentReference" }));
    referenceRun.Append(new CommentReference { Id = commentId });

    if (!string.IsNullOrWhiteSpace(anchorContains)
        && TrySplitRunsForAnchor(target.Paragraph, anchorContains, out var splitFirstRun, out var splitLastRun))
    {
        splitFirstRun!.InsertBeforeSelf(rangeStart);
        splitLastRun!.InsertAfterSelf(rangeEnd);
        rangeEnd.InsertAfterSelf(referenceRun);
    }
    else
    {
        var firstChild = target.Paragraph.Elements<OpenXmlElement>().FirstOrDefault(e => e is not ParagraphProperties);
        if (firstChild is null)
        {
            target.Paragraph.Append(rangeStart);
            target.Paragraph.Append(rangeEnd);
            target.Paragraph.Append(referenceRun);
        }
        else
        {
            target.Paragraph.InsertBefore(rangeStart, firstChild);
            target.Paragraph.Append(rangeEnd);
            target.Paragraph.Append(referenceRun);
        }
    }

    return true;
}

static int ReanchorComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CommentReanchorPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read comment reanchor plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils reanchor-comments {DateTime.UtcNow:O} author={author}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
    var comments = mainPart.WordprocessingCommentsPart?.Comments;
    if (comments is null)
    {
        Console.Error.WriteLine("Comments part not found.");
        return 7;
    }

    foreach (var entry in plan.Anchors)
    {
        if (string.IsNullOrWhiteSpace(entry.CommentId))
        {
            skipped.Add($"{entry.Id}: missing commentId");
            continue;
        }

        var comment = comments.Elements<Comment>().FirstOrDefault(c => c.Id?.Value == entry.CommentId);
        if (comment is null)
        {
            skipped.Add($"{entry.Id}: comment id={entry.CommentId} not found");
            continue;
        }

        if (string.IsNullOrWhiteSpace(entry.AnchorPrefix))
        {
            skipped.Add($"{entry.Id}: missing anchorPrefix");
            continue;
        }

        ParagraphEntry target;
        try
        {
            target = FindCommentAnchorParagraph(doc, entry.AnchorPrefix, $"{entry.Id}/anchor", entry.AnchorContains, entry.Occurrence);
        }
        catch (InvalidOperationException ex)
        {
            skipped.Add($"{entry.Id}: anchor lookup failed ({ex.Message})");
            continue;
        }

        var runAnchorContains = IsFollowingEquationAnchorMarker(entry.AnchorContains)
            ? null
            : entry.AnchorContains;
        if (!ReanchorCommentOnParagraph(mainPart, entry.CommentId, target, runAnchorContains))
        {
            skipped.Add($"{entry.Id}: comment id={entry.CommentId} already anchored on P[{target.Index}]");
            continue;
        }

        applied.Add($"{entry.Id}: comment id={entry.CommentId} anchored on P[{target.Index}]");
    }

    comments.Save();
    mainPart.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Reanchor Comments Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        builder.AppendLine();
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int AnswerComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CommentAnswerPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read comment answer plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    var nowUtc = DateTime.UtcNow;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils answer-comments {DateTime.UtcNow:O} author={author}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
    var comments = mainPart.WordprocessingCommentsPart?.Comments;
    if (comments is null)
    {
        Console.Error.WriteLine("Comments part not found.");
        return 7;
    }

    foreach (var entry in plan.Answers)
    {
        if (string.IsNullOrWhiteSpace(entry.CommentId))
        {
            skipped.Add($"{entry.Id}: missing commentId");
            continue;
        }

        var comment = comments.Elements<Comment>().FirstOrDefault(c => c.Id?.Value == entry.CommentId);
        if (comment is null)
        {
            skipped.Add($"{entry.Id}: comment id={entry.CommentId} not found");
            continue;
        }

        var responseBody = string.IsNullOrWhiteSpace(entry.ResponseText)
            ? ""
            : $"Resposta ({author}, {nowUtc:yyyy-MM-dd}): {entry.ResponseText.Trim()}";

        if (string.IsNullOrWhiteSpace(responseBody))
        {
            skipped.Add($"{entry.Id}: empty response");
            continue;
        }

        var alreadyPresent = comment.Elements<Paragraph>()
            .Select(ParagraphText)
            .Any(text => string.Equals(text, responseBody, StringComparison.Ordinal));
        if (alreadyPresent)
        {
            skipped.Add($"{entry.Id}: response already present on comment id={entry.CommentId}");
            continue;
        }

        var paragraph = new Paragraph(
            new Run(
                new Text(responseBody) { Space = SpaceProcessingModeValues.Preserve }));
        comment.Append(paragraph);
        applied.Add($"{entry.Id}: response appended to comment id={entry.CommentId}");
    }

    comments.Save();
    mainPart.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Answer Comments Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        builder.AppendLine();
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static Paragraph CreateCommentBodyParagraph(string text, string paraId)
{
    var paragraph = new Paragraph
    {
        ParagraphId = paraId,
        TextId = "77777777"
    };
    paragraph.Append(new ParagraphProperties(
        new ParagraphStyleId { Val = "Textodecomentrio" },
        new Indentation { FirstLine = "0" },
        new Justification { Val = JustificationValues.Left }));

    var annotationRefRun = new Run(
        new RunProperties(new RunStyle { Val = "Refdecomentrio" }),
        new AnnotationReferenceMark());
    var textRun = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    paragraph.Append(annotationRefRun);
    paragraph.Append(textRun);
    return paragraph;
}

static string NextUniqueHexValue(HashSet<string> existing, int hexLength = 8)
{
    while (true)
    {
        var candidate = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture)
            .Substring(0, hexLength)
            .ToUpperInvariant();
        if (existing.Add(candidate))
        {
            return candidate;
        }
    }
}

static int ReplyComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);

    if (!options.TryGetValue("plan", out var planPathValue))
    {
        Console.Error.WriteLine("Missing required option: --plan");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var planPath = Path.GetFullPath(planPathValue);
    if (!File.Exists(planPath))
    {
        Console.Error.WriteLine($"Plan not found: {planPath}");
        return 5;
    }

    var plan = JsonSerializer.Deserialize<CommentReplyPlan>(File.ReadAllText(planPath), new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan is null)
    {
        Console.Error.WriteLine("Unable to read comment reply plan.");
        return 6;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    var applied = new List<string>();
    var skipped = new List<string>();
    var nowUtc = DateTime.UtcNow;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils reply-comments {DateTime.UtcNow:O} author={author}\n"));
    lockStream.Flush();

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
    var commentsPart = mainPart.WordprocessingCommentsPart;
    if (commentsPart?.Comments is null)
    {
        Console.Error.WriteLine("Comments part not found.");
        return 7;
    }

    var commentsExPart = mainPart.WordprocessingCommentsExPart ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
    commentsExPart.CommentsEx ??= new W15.CommentsEx();
    var commentsIdsPart = mainPart.WordprocessingCommentsIdsPart ?? mainPart.AddNewPart<WordprocessingCommentsIdsPart>();
    commentsIdsPart.CommentsIds ??= new W16Cid.CommentsIds();
    var commentsExtensiblePart = mainPart.WordCommentsExtensiblePart ?? mainPart.AddNewPart<WordCommentsExtensiblePart>();
    commentsExtensiblePart.CommentsExtensible ??= new W16Cex.CommentsExtensible();

    var existingCommentParaIds = commentsPart.Comments.Elements<Comment>()
        .SelectMany(comment => comment.Elements<Paragraph>())
        .Select(paragraph => paragraph.ParagraphId?.Value)
        .Where(value => !string.IsNullOrWhiteSpace(value))
        .Select(value => value!)
        .ToHashSet(StringComparer.Ordinal);
    var existingDurableIds = commentsIdsPart.CommentsIds.Elements<W16Cid.CommentId>()
        .Select(item => item.DurableId?.Value)
        .Where(value => !string.IsNullOrWhiteSpace(value))
        .Select(value => value!)
        .ToHashSet(StringComparer.Ordinal);
    var nextCommentId = NextCommentId(commentsPart.Comments);

    foreach (var entry in plan.Replies)
    {
        if (string.IsNullOrWhiteSpace(entry.ParentCommentId))
        {
            skipped.Add($"{entry.Id}: missing parentCommentId");
            continue;
        }

        var parentComment = commentsPart.Comments.Elements<Comment>()
            .FirstOrDefault(comment => comment.Id?.Value == entry.ParentCommentId);
        if (parentComment is null)
        {
            skipped.Add($"{entry.Id}: parent comment id={entry.ParentCommentId} not found");
            continue;
        }

        if (!string.IsNullOrWhiteSpace(entry.AnchorPrefix))
        {
            ParagraphEntry target;
            try
            {
                target = FindCommentAnchorParagraph(doc, entry.AnchorPrefix, $"{entry.Id}/reanchor", entry.AnchorContains, entry.Occurrence);
            }
            catch (InvalidOperationException ex)
            {
                skipped.Add($"{entry.Id}: reanchor failed ({ex.Message})");
                continue;
            }

            var runAnchorContains = IsFollowingEquationAnchorMarker(entry.AnchorContains)
                ? null
                : entry.AnchorContains;
            if (ReanchorCommentOnParagraph(mainPart, entry.ParentCommentId, target, runAnchorContains))
            {
                applied.Add($"{entry.Id}: reanchored parent comment id={entry.ParentCommentId} on P[{target.Index}]");
            }
        }

        var parentParaId = parentComment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
        if (string.IsNullOrWhiteSpace(parentParaId))
        {
            skipped.Add($"{entry.Id}: parent comment id={entry.ParentCommentId} has no paraId");
            continue;
        }

        var commentExByParaId = commentsExPart.CommentsEx.Elements<W15.CommentEx>()
            .Where(item => item.ParaId?.Value is not null)
            .ToDictionary(item => item.ParaId!.Value!, StringComparer.Ordinal);
        var replyAlreadyPresent = commentsPart.Comments.Elements<Comment>()
            .Any(comment =>
            {
                var replyParaId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
                if (string.IsNullOrWhiteSpace(replyParaId))
                {
                    return false;
                }

                if (!commentExByParaId.TryGetValue(replyParaId, out var commentEx))
                {
                    return false;
                }

                return string.Equals(commentEx.ParaIdParent?.Value, parentParaId, StringComparison.Ordinal)
                    && string.Equals(ElementText(comment).Trim(), (entry.ReplyText ?? "").Trim(), StringComparison.Ordinal);
            });
        if (replyAlreadyPresent)
        {
            skipped.Add($"{entry.Id}: reply already present for parent comment id={entry.ParentCommentId}");
            continue;
        }

        var replyText = (entry.ReplyText ?? "").Trim();
        if (string.IsNullOrWhiteSpace(replyText))
        {
            skipped.Add($"{entry.Id}: empty replyText");
            continue;
        }

        var replyCommentId = nextCommentId++.ToString(CultureInfo.InvariantCulture);
        var replyParaId = NextUniqueHexValue(existingCommentParaIds);
        var replyDurableId = NextUniqueHexValue(existingDurableIds);

        var replyComment = new Comment
        {
            Id = replyCommentId,
            Author = author,
            Date = nowUtc
        };
        ApplyCommentInitials(replyComment);
        replyComment.Append(CreateCommentBodyParagraph(replyText, replyParaId));
        commentsPart.Comments.Append(replyComment);

        commentsExPart.CommentsEx.Append(new W15.CommentEx
        {
            ParaId = replyParaId,
            ParaIdParent = parentParaId,
            Done = false
        });
        commentsIdsPart.CommentsIds.Append(new W16Cid.CommentId
        {
            ParaId = replyParaId,
            DurableId = replyDurableId
        });
        commentsExtensiblePart.CommentsExtensible.Append(new W16Cex.CommentExtensible
        {
            DurableId = replyDurableId,
            DateUtc = nowUtc
        });

        applied.Add($"{entry.Id}: reply comment id={replyCommentId} linked to parent comment id={entry.ParentCommentId}");
    }

    commentsPart.Comments.Save();
    commentsExPart.CommentsEx.Save(commentsExPart);
    commentsIdsPart.CommentsIds.Save(commentsIdsPart);
    commentsExtensiblePart.CommentsExtensible.Save(commentsExtensiblePart);
    mainPart.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Reply Comments Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Plano: `{planPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- Aplicados: {applied.Count}");
        builder.AppendLine($"- Ignorados: {skipped.Count}");
        builder.AppendLine();
        foreach (var item in applied) builder.AppendLine($"- APPLY `{item}`");
        foreach (var item in skipped) builder.AppendLine($"- SKIP `{item}`");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var item in applied) Console.WriteLine($"APPLY {item}");
    foreach (var item in skipped) Console.WriteLine($"SKIP {item}");
    return 0;
}

static int RemoveComments(string docxPath, IReadOnlyDictionary<string, string> options)
{
    var author = MutationAuthorResolver.Resolve(docxPath, options);
    if (!options.TryGetValue("ids", out var idsValue) || string.IsNullOrWhiteSpace(idsValue))
    {
        Console.Error.WriteLine("Missing required option: --ids");
        return 4;
    }


    if (!options.TryGetValue("lock", out var lockPathValue) || string.IsNullOrWhiteSpace(lockPathValue))
    {
        Console.Error.WriteLine("Missing required option: --lock");
        return 4;
    }

    var lockPath = Path.GetFullPath(lockPathValue);
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath) ?? ".");
    var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
        ? Path.GetFullPath(reportPathValue)
        : null;

    using var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    lockStream.SetLength(0);
    lockStream.Write(Encoding.UTF8.GetBytes($"docx-utils remove-comments {DateTime.UtcNow:O} author={author}\n"));
    lockStream.Flush();

    var removed = new List<string>();
    var missing = new List<string>();
    var markerCount = 0;

    using var doc = WordprocessingDocument.Open(docxPath, true);
    EnsureTrackRevisions(doc);
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
    var commentsPart = mainPart.WordprocessingCommentsPart;
    var comments = commentsPart?.Comments;
    if (comments is null)
    {
        Console.WriteLine("NO_COMMENTS_PART");
        return 0;
    }

    var existingIds = comments.Elements<Comment>()
        .Select(c => c.Id?.Value)
        .Where(id => !string.IsNullOrWhiteSpace(id))
        .Select(id => id!)
        .ToHashSet(StringComparer.Ordinal);

    var idsToRemove = string.Equals(idsValue.Trim(), "all", StringComparison.OrdinalIgnoreCase)
        ? existingIds
        : idsValue.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .ToHashSet(StringComparer.Ordinal);

    foreach (var id in idsToRemove)
    {
        var comment = comments.Elements<Comment>().FirstOrDefault(c => c.Id?.Value == id);
        if (comment is null)
        {
            missing.Add(id);
            continue;
        }

        comment.Remove();
        removed.Add(id);
    }

    foreach (var part in MainStoryParts(mainPart))
    {
        if (part.RootElement is null) continue;
        foreach (var start in part.RootElement.Descendants<CommentRangeStart>()
            .Where(c => c.Id?.Value is { } id && idsToRemove.Contains(id))
            .ToList())
        {
            start.Remove();
            markerCount++;
        }

        foreach (var end in part.RootElement.Descendants<CommentRangeEnd>()
            .Where(c => c.Id?.Value is { } id && idsToRemove.Contains(id))
            .ToList())
        {
            end.Remove();
            markerCount++;
        }

        foreach (var reference in part.RootElement.Descendants<CommentReference>()
            .Where(c => c.Id?.Value is { } id && idsToRemove.Contains(id))
            .ToList())
        {
            var run = reference.Ancestors<Run>().FirstOrDefault();
            if (run is not null)
            {
                run.Remove();
            }
            else
            {
                reference.Remove();
            }
            markerCount++;
        }

        part.RootElement.Save();
    }

    comments.Save();
    mainPart.Document.Save();

    if (reportPath is not null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
        var builder = new StringBuilder();
        builder.AppendLine("# Remove Comments Report");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{docxPath}`");
        builder.AppendLine($"- Autor: `{author}`");
        builder.AppendLine($"- Lock: `{lockPath}`");
        builder.AppendLine($"- IDs solicitados: `{idsValue}`");
        builder.AppendLine($"- Comentarios removidos: {removed.Count}");
        builder.AppendLine($"- Marcadores removidos: {markerCount}");
        builder.AppendLine($"- IDs ausentes: {missing.Count}");
        foreach (var id in removed) builder.AppendLine($"- REMOVE comment id={id}");
        foreach (var id in missing) builder.AppendLine($"- MISSING comment id={id}");
        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
    }

    foreach (var id in removed) Console.WriteLine($"REMOVE comment id={id}");
    foreach (var id in missing) Console.WriteLine($"MISSING comment id={id}");
    Console.WriteLine($"MARKERS_REMOVED {markerCount}");
    return 0;
}

static bool HasParagraphStyle(WordprocessingDocument doc, string styleId)
{
    var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
    if (styles is null)
    {
        return false;
    }

    return styles.Elements<Style>().Any(s =>
        string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal)
        && (s.Type?.Value == StyleValues.Paragraph || s.Type is null));
}

static void RemoveCommentMarkers(MainDocumentPart mainPart, string commentId)
{
    foreach (var part in MainStoryParts(mainPart))
    {
        if (part.RootElement is null)
        {
            continue;
        }

        foreach (var start in part.RootElement.Descendants<CommentRangeStart>()
            .Where(c => c.Id?.Value == commentId)
            .ToList())
        {
            start.Remove();
        }

        foreach (var end in part.RootElement.Descendants<CommentRangeEnd>()
            .Where(c => c.Id?.Value == commentId)
            .ToList())
        {
            end.Remove();
        }

        foreach (var reference in part.RootElement.Descendants<CommentReference>()
            .Where(c => c.Id?.Value == commentId)
            .ToList())
        {
            var run = reference.Ancestors<Run>().FirstOrDefault();
            if (run is not null)
            {
                run.Remove();
            }
            else
            {
                reference.Remove();
            }
        }

        part.RootElement.Save();
    }
}

static int CountExistingFiguraCaptions(WordprocessingDocument doc)
{
    var body = doc.MainDocumentPart?.Document.Body;
    if (body is null)
    {
        return 0;
    }

    return body.Descendants<Paragraph>()
        .Count(p => string.Equals(
            p.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
            "Figura",
            StringComparison.Ordinal));
}

static (int Width, int Height)? TryReadPngDimensions(string path)
{
    try
    {
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[24];
        var read = stream.Read(header);
        if (read < 24)
        {
            return null;
        }

        // PNG signature is bytes 0..7; IHDR width/height are big-endian uint32 at offsets 16..20 and 20..24.
        if (header[0] != 0x89 || header[1] != 0x50 || header[2] != 0x4E || header[3] != 0x47)
        {
            return null;
        }

        var width = (header[16] << 24) | (header[17] << 16) | (header[18] << 8) | header[19];
        var height = (header[20] << 24) | (header[21] << 16) | (header[22] << 8) | header[23];
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        return (width, height);
    }
    catch
    {
        return null;
    }
}

static Paragraph CreateInsertedFigureParagraph(
    string relationshipId,
    string figureId,
    string captionText,
    long widthEmu,
    long heightEmu,
    bool hasFiguraStyle,
    RevisionMetadata metadata)
{
    var paragraph = new Paragraph();
    paragraph.ParagraphId = HexBinaryValue.FromString(RandomHex(8));
    paragraph.TextId = HexBinaryValue.FromString(RandomHex(8));

    var paragraphProperties = new ParagraphProperties();
    if (hasFiguraStyle)
    {
        paragraphProperties.Append(new ParagraphStyleId { Val = "Figura" });
    }
    paragraphProperties.Append(new Justification { Val = JustificationValues.Center });
    paragraph.Append(paragraphProperties);

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    var run = new Run();
    var drawing = BuildInlineDrawing(relationshipId, figureId, captionText, widthEmu, heightEmu);
    run.Append(drawing);
    insertion.Append(run);
    paragraph.Append(insertion);
    return paragraph;
}

static Drawing BuildInlineDrawing(string relationshipId, string figureId, string description, long widthEmu, long heightEmu)
{
    var safeDescription = string.IsNullOrWhiteSpace(description) ? figureId : description;
    var docPropsId = (uint)Math.Abs(figureId.GetHashCode() % 1_000_000) + 1u;

    var inline = new Inline(
        new Extent { Cx = widthEmu, Cy = heightEmu },
        new EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        new DocProperties { Id = docPropsId, Name = figureId, Description = safeDescription },
        new NonVisualGraphicFrameDrawingProperties(
            new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
        new DocumentFormat.OpenXml.Drawing.Graphic(
            new DocumentFormat.OpenXml.Drawing.GraphicData(
                BuildPicture(relationshipId, figureId, safeDescription, widthEmu, heightEmu))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            }))
    {
        DistanceFromTop = 0U,
        DistanceFromBottom = 0U,
        DistanceFromLeft = 0U,
        DistanceFromRight = 0U
    };

    return new Drawing(inline);
}

static DocumentFormat.OpenXml.Drawing.Pictures.Picture BuildPicture(
    string relationshipId,
    string figureId,
    string description,
    long widthEmu,
    long heightEmu)
{
    var picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = 0U, Name = figureId, Description = description },
            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.PictureLocks { NoChangeAspect = true, NoChangeArrowheads = true })),
        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
            new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId },
            new DocumentFormat.OpenXml.Drawing.SourceRectangle(),
            new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthEmu, Cy = heightEmu }),
            new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList())
            {
                Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
            }));

    return picture;
}

static Paragraph CreateInsertedFigureCaption(
    string captionText,
    int captionNumber,
    bool hasFiguraStyle,
    RevisionMetadata metadata)
{
    var paragraph = new Paragraph();
    paragraph.ParagraphId = HexBinaryValue.FromString(RandomHex(8));
    paragraph.TextId = HexBinaryValue.FromString(RandomHex(8));

    var paragraphProperties = new ParagraphProperties();
    if (hasFiguraStyle)
    {
        paragraphProperties.Append(new ParagraphStyleId { Val = "Figura" });
    }
    paragraph.Append(paragraphProperties);

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    insertion.Append(CreateTextRun($"Figura {captionNumber.ToString(CultureInfo.InvariantCulture)} - {captionText}"));
    paragraph.Append(insertion);
    return paragraph;
}

static Paragraph CreateInsertedFigureSource(string sourceText, RevisionMetadata metadata)
{
    var paragraph = new Paragraph();
    paragraph.ParagraphId = HexBinaryValue.FromString(RandomHex(8));
    paragraph.TextId = HexBinaryValue.FromString(RandomHex(8));

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    insertion.Append(CreateTextRun(sourceText));
    paragraph.Append(insertion);
    return paragraph;
}

static int NextCommentId(Comments comments)
{
    var max = 0;
    foreach (var comment in comments.Elements<Comment>())
    {
        if (comment.Id?.Value is { } idText && int.TryParse(idText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
        {
            if (id > max)
            {
                max = id;
            }
        }
    }

    return max + 1;
}

static bool TrySplitRunsForAnchor(Paragraph paragraph, string anchorContains, out Run? firstRun, out Run? lastRun)
{
    firstRun = null;
    lastRun = null;
    if (string.IsNullOrEmpty(anchorContains))
    {
        return false;
    }

    var runs = paragraph.Elements<Run>().ToList();
    if (runs.Count == 0)
    {
        return false;
    }

    // Build offset map of each Text element in run order so we can locate the anchor.
    var sb = new StringBuilder();
    var textInfos = new List<(Run Run, Text Text, int Offset)>();
    foreach (var run in runs)
    {
        foreach (var text in run.Elements<Text>())
        {
            textInfos.Add((run, text, sb.Length));
            sb.Append(text.Text ?? "");
        }
    }

    var concatenated = sb.ToString();
    var startIndex = concatenated.IndexOf(anchorContains, StringComparison.Ordinal);
    if (startIndex < 0)
    {
        return false;
    }

    var endIndex = startIndex + anchorContains.Length;

    // Split first text element to align with startIndex.
    var startInfo = textInfos.LastOrDefault(t => t.Offset <= startIndex);
    if (startInfo.Run is null)
    {
        return false;
    }

    var startInsideOffset = startIndex - startInfo.Offset;
    var startTextValue = startInfo.Text.Text ?? "";
    if (startInsideOffset > 0 && startInsideOffset < startTextValue.Length)
    {
        var beforeText = startTextValue[..startInsideOffset];
        var afterText = startTextValue[startInsideOffset..];
        startInfo.Text.Text = beforeText;
        startInfo.Text.Space = SpaceProcessingModeValues.Preserve;

        var newRun = (Run)startInfo.Run.CloneNode(false);
        if (startInfo.Run.RunProperties is not null)
        {
            newRun.Append((RunProperties)startInfo.Run.RunProperties.CloneNode(true));
        }
        newRun.Append(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
        startInfo.Run.InsertAfterSelf(newRun);
        firstRun = newRun;
    }
    else if (startInsideOffset == 0)
    {
        firstRun = startInfo.Run;
    }
    else
    {
        // anchor begins exactly at end of a text element; use the next run.
        firstRun = startInfo.Run.ElementsAfter().OfType<Run>().FirstOrDefault();
        if (firstRun is null)
        {
            return false;
        }
    }

    // Recompute infos because tree changed.
    var refreshedRuns = paragraph.Elements<Run>().ToList();
    sb.Clear();
    var refreshedInfos = new List<(Run Run, Text Text, int Offset)>();
    foreach (var run in refreshedRuns)
    {
        foreach (var text in run.Elements<Text>())
        {
            refreshedInfos.Add((run, text, sb.Length));
            sb.Append(text.Text ?? "");
        }
    }

    var endInfo = refreshedInfos.LastOrDefault(t => t.Offset < endIndex);
    if (endInfo.Run is null)
    {
        return false;
    }

    var endInsideOffset = endIndex - endInfo.Offset;
    var endTextValue = endInfo.Text.Text ?? "";
    if (endInsideOffset > 0 && endInsideOffset < endTextValue.Length)
    {
        var beforeText = endTextValue[..endInsideOffset];
        var afterText = endTextValue[endInsideOffset..];
        endInfo.Text.Text = beforeText;
        endInfo.Text.Space = SpaceProcessingModeValues.Preserve;

        var newRun = (Run)endInfo.Run.CloneNode(false);
        if (endInfo.Run.RunProperties is not null)
        {
            newRun.Append((RunProperties)endInfo.Run.RunProperties.CloneNode(true));
        }
        newRun.Append(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
        endInfo.Run.InsertAfterSelf(newRun);
        lastRun = endInfo.Run;
    }
    else
    {
        lastRun = endInfo.Run;
    }

    return firstRun is not null && lastRun is not null;
}

static FileStream OpenSharedRead(string docxPath) =>
    File.Open(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

static IEnumerable<OpenXmlPart> MainStoryParts(MainDocumentPart mainPart)
{
    yield return mainPart;
    foreach (var part in mainPart.HeaderParts) yield return part;
    foreach (var part in mainPart.FooterParts) yield return part;
    if (mainPart.FootnotesPart is not null) yield return mainPart.FootnotesPart;
    if (mainPart.EndnotesPart is not null) yield return mainPart.EndnotesPart;
}

static List<ParagraphEntry> GetParagraphs(WordprocessingDocument doc)
{
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    return body.Elements<Paragraph>()
        .Select((p, index) => new ParagraphEntry(p, index, ParagraphText(p)))
        .Where(p => !string.IsNullOrWhiteSpace(p.Text))
        .ToList();
}

static List<ParagraphEntry> GetAllParagraphEntries(WordprocessingDocument doc)
{
    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    return body.Descendants<Paragraph>()
        .Select((p, index) => new ParagraphEntry(p, index, ParagraphText(p)))
        .Where(p => !string.IsNullOrWhiteSpace(p.Text))
        .ToList();
}

static int RemoveElementsBetweenAnchors(OpenXmlElement start, OpenXmlElement end)
{
    if (!ReferenceEquals(start.Parent, end.Parent))
    {
        throw new InvalidOperationException("Anchor paragraphs must share the same parent.");
    }

    var removed = 0;
    var current = start.NextSibling();
    while (current is not null && !ReferenceEquals(current, end))
    {
        var next = current.NextSibling();
        current.Remove();
        removed++;
        current = next;
    }

    return removed;
}

static ParagraphEntry FindUniqueParagraph(IEnumerable<ParagraphEntry> paragraphs, string prefix, string label)
{
    var matches = paragraphs
        .Where(p => NormalizedStartsWith(p.Text, prefix))
        .ToList();

    if (matches.Count != 1)
    {
        throw new InvalidOperationException($"Expected exactly one match for {label}; got {matches.Count}.");
    }

    return matches[0];
}

static bool ContentAlreadyPresent(IEnumerable<ParagraphEntry> paragraphs, string content)
{
    var contentPrefix = BuildPresencePrefix(content);
    return paragraphs.Any(p => NormalizedStartsWith(p.Text, contentPrefix));
}

static bool BlockTableAlreadyPresent(WordprocessingDocument doc, BlockSpec spec)
{
    foreach (var item in spec.Items.Where(i => i.Kind.Equals("table", StringComparison.OrdinalIgnoreCase)))
    {
        if (item.Rows is null || item.Rows.Count == 0)
        {
            continue;
        }

        if (TableWithSameRowsExists(doc, item.Rows))
        {
            return true;
        }
    }

    return false;
}

static bool TableWithSameRowsExists(WordprocessingDocument doc, IReadOnlyList<IReadOnlyList<string?>> rows)
{
    var expected = NormalizeTableRows(rows);
    if (expected.Count == 0)
    {
        return false;
    }

    var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");
    foreach (var table in body.Descendants<Table>())
    {
        var actual = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>()
                .Select(cell => Normalize(string.Join(" ", cell.Descendants<Text>().Select(t => t.Text))))
                .ToList())
            .Where(row => row.Any(cell => !string.IsNullOrWhiteSpace(cell)))
            .ToList();

        if (RowsEqual(expected, actual))
        {
            return true;
        }
    }

    return false;
}

static List<List<string>> NormalizeTableRows(IReadOnlyList<IReadOnlyList<string?>> rows) =>
    rows.Select(row => row.Select(cell => Normalize(cell ?? "")).ToList())
        .Where(row => row.Any(cell => !string.IsNullOrWhiteSpace(cell)))
        .ToList();

static bool RowsEqual(IReadOnlyList<IReadOnlyList<string>> left, IReadOnlyList<IReadOnlyList<string>> right)
{
    if (left.Count != right.Count)
    {
        return false;
    }

    for (var i = 0; i < left.Count; i++)
    {
        if (left[i].Count != right[i].Count)
        {
            return false;
        }

        for (var j = 0; j < left[i].Count; j++)
        {
            if (!string.Equals(left[i][j], right[i][j], StringComparison.Ordinal))
            {
                return false;
            }
        }
    }

    return true;
}

static string BuildPresencePrefix(string content)
{
    var trimmed = Regex.Replace(content, "\\s+", " ").Trim();
    var cut = trimmed.IndexOf('.', StringComparison.Ordinal);
    return cut > 30 ? trimmed[..(cut + 1)] : trimmed[..Math.Min(trimmed.Length, 120)];
}

static Paragraph CreateInsertedParagraph(Paragraph template, string text, RevisionMetadata metadata)
{
    var paragraph = new Paragraph();

    if (template.ParagraphProperties is not null)
    {
        paragraph.Append((ParagraphProperties)template.ParagraphProperties.CloneNode(true));
    }

    paragraph.RsidParagraphAddition = template.RsidParagraphAddition;
    paragraph.RsidParagraphProperties = template.RsidParagraphProperties;
    paragraph.RsidRunAdditionDefault = template.RsidRunAdditionDefault;

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };

    var run = new Run();
    var templateRunProps = template.Descendants<RunProperties>().FirstOrDefault();
    if (templateRunProps is not null)
    {
        run.Append((RunProperties)templateRunProps.CloneNode(true));
    }

    run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    insertion.Append(run);
    paragraph.Append(insertion);
    return paragraph;
}

static Paragraph CreateInsertedParagraphWithStyle(Paragraph template, string text, string? styleId, RevisionMetadata metadata)
{
    var paragraph = CreateInsertedParagraph(template, text, metadata);
    if (!string.IsNullOrWhiteSpace(styleId))
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        var current = paragraph.ParagraphProperties.ParagraphStyleId;
        if (current is null)
        {
            paragraph.ParagraphProperties.PrependChild(new ParagraphStyleId { Val = styleId });
        }
        else
        {
            current.Val = styleId;
        }
    }
    return paragraph;
}

static Paragraph CreateInsertedEquationParagraph(Paragraph template, string latex, string styleId, RevisionMetadata metadata)
{
    if (string.IsNullOrWhiteSpace(latex))
    {
        throw new InvalidOperationException("Equation item requires `latex`.");
    }

    var paragraph = new Paragraph();
    if (template.ParagraphProperties is not null)
    {
        paragraph.Append((ParagraphProperties)template.ParagraphProperties.CloneNode(true));
    }

    paragraph.ParagraphProperties ??= new ParagraphProperties();
    var currentStyle = paragraph.ParagraphProperties.ParagraphStyleId;
    if (currentStyle is null)
    {
        paragraph.ParagraphProperties.PrependChild(new ParagraphStyleId { Val = styleId });
    }
    else
    {
        currentStyle.Val = styleId;
    }

    paragraph.RsidParagraphAddition = template.RsidParagraphAddition;
    paragraph.RsidParagraphProperties = template.RsidParagraphProperties;
    paragraph.RsidRunAdditionDefault = template.RsidRunAdditionDefault;

    var insertion = new InsertedRun
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    };
    insertion.Append(CreateProfessionalOfficeMathFromLatex(new FormulaSpec(latex, latex, "", false, null)));
    paragraph.Append(insertion);
    return paragraph;
}

static void EnsureEquationTabs(ParagraphProperties properties, Paragraph paragraph)
{
    var (centerPosition, rightPosition) = GetEquationTabPositions(paragraph);
    var tabs = properties.GetFirstChild<Tabs>();
    if (tabs is not null)
    {
        tabs.Remove();
    }

    tabs = new Tabs(
        new TabStop { Val = TabStopValues.Center, Position = centerPosition },
        new TabStop { Val = TabStopValues.Right, Position = rightPosition });
    properties.Append(tabs);
}

static (int centerPosition, int rightPosition) GetEquationTabPositions(Paragraph paragraph)
{
    const int defaultPageWidth = 11906;
    const int defaultMargin = 1701;

    var section = paragraph.Ancestors<Body>().FirstOrDefault()?
        .Elements<SectionProperties>()
        .LastOrDefault()
        ?? paragraph.Descendants<SectionProperties>().LastOrDefault();
    var pageSize = section?.GetFirstChild<PageSize>();
    var pageMargin = section?.GetFirstChild<PageMargin>();

    var pageWidth = (int)(pageSize?.Width?.Value ?? (uint)defaultPageWidth);
    var leftMargin = (int)(pageMargin?.Left?.Value ?? (uint)defaultMargin);
    var rightMargin = (int)(pageMargin?.Right?.Value ?? (uint)defaultMargin);
    var usableWidth = Math.Max(2000, pageWidth - leftMargin - rightMargin);
    var center = leftMargin + usableWidth / 2;
    var rightPos = leftMargin + usableWidth;
    return (center, rightPos);
}

static Table CreateTrackedTable(BlockItemSpec item, RevisionMetadata metadata)
{
    var table = new Table();
    table.Append(new TableProperties(
        new TableStyle { Val = item.TableStyleId ?? "tabelauerj" },
        new TableWidth { Width = "9000", Type = TableWidthUnitValues.Dxa },
        new TableJustification { Val = TableRowAlignmentValues.Center },
        new TableLayout { Type = TableLayoutValues.Fixed }));

    var columnCount = item.Rows?.Select(r => r.Count).DefaultIfEmpty(0).Max() ?? 0;
    if (columnCount > 0)
    {
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new GridColumn { Width = "2400" });
        }
        table.Append(grid);
    }

    var cellPct = columnCount > 0 ? Math.Max(1, 5000 / columnCount).ToString(CultureInfo.InvariantCulture) : "2500";
    foreach (var rowValues in item.Rows ?? [])
    {
        var row = new TableRow();
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            var value = columnIndex < rowValues.Count ? rowValues[columnIndex] : "";
            var cell = new TableCell();
            cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = cellPct }));
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = item.CellStyleId ?? "dados" }));
            var insertion = new InsertedRun
            {
                Id = metadata.NextRevisionId(),
                Author = metadata.Author,
                Date = metadata.DateUtc
            };
            var run = new Run(new Text(value ?? "") { Space = SpaceProcessingModeValues.Preserve });
            insertion.Append(run);
            paragraph.Append(insertion);
            cell.Append(paragraph);
            row.Append(cell);
        }
        table.Append(row);
    }

    ApplyAcademicTableWindowFit(table);
    return table;
}

static void EnsureTrackRevisions(WordprocessingDocument doc)
{
    EnableTrackRevisionsOnly(doc);
    EnsureCanonicalStyles(doc, GetDefaultCanonicalStylesDirectory());
}

static void EnableTrackRevisionsOnly(WordprocessingDocument doc)
{
    var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart ?? doc.MainDocumentPart?.AddNewPart<DocumentSettingsPart>();
    if (settingsPart is null)
    {
        throw new InvalidOperationException("Unable to access document settings.");
    }

    settingsPart.Settings ??= new Settings();
    foreach (var existing in settingsPart.Settings.Elements<TrackRevisions>().ToList())
    {
        existing.Remove();
    }

    settingsPart.Settings.AddChild(new TrackRevisions(), true);
    settingsPart.Settings.Save();
}

static IReadOnlyList<string> EnsureCanonicalStyles(WordprocessingDocument doc, string sourceDir)
{
    var applied = new List<string>();
    var stylesPath = Path.Combine(sourceDir, "styles.xml");
    if (!File.Exists(stylesPath))
    {
        return applied;
    }

    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart not found.");
    var canonicalStyles = new Styles(File.ReadAllText(stylesPath, Encoding.UTF8));
    SanitizeStrictStyleCompatibility(canonicalStyles);
    var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
    stylesPart.Styles ??= new Styles();
    var targetStyles = stylesPart.Styles;

    targetStyles.DocDefaults?.Remove();
    var docDefaults = canonicalStyles.DocDefaults?.CloneNode(true);
    if (docDefaults is not null)
    {
        targetStyles.PrependChild(docDefaults);
        applied.Add("styles: canonical docDefaults copied");
    }

    targetStyles.LatentStyles?.Remove();
    var latentStyles = canonicalStyles.LatentStyles?.CloneNode(true);
    if (latentStyles is not null)
    {
        var insertBefore = targetStyles.Elements<Style>().FirstOrDefault();
        if (insertBefore is null)
        {
            targetStyles.Append(latentStyles);
        }
        else
        {
            targetStyles.InsertBefore(latentStyles, insertBefore);
        }
        applied.Add("styles: canonical latentStyles copied");
    }

    var styleIds = ReadCanonicalStyleIds(sourceDir);
    if (styleIds.Count == 0)
    {
        styleIds = canonicalStyles.Elements<Style>()
            .Select(s => s.StyleId?.Value ?? "")
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }
    styleIds = ExpandStyleDependencies(canonicalStyles, styleIds);

    var copiedStyles = 0;
    foreach (var styleId in styleIds.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
    {
        var canonical = canonicalStyles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
        if (canonical is null)
        {
            continue;
        }

        foreach (var existing in targetStyles.Elements<Style>()
            .Where(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase))
            .ToList())
        {
            existing.Remove();
        }

        targetStyles.Append(canonical.CloneNode(true));
        copiedStyles++;
    }
    applied.Add($"styles: canonical definitions copied for {copiedStyles} used/dependency style(s)");
    stylesPart.Styles.Save();

    var numberingPath = Path.Combine(sourceDir, "numbering.xml");
    if (File.Exists(numberingPath))
    {
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(File.ReadAllText(numberingPath, Encoding.UTF8));
        numberingPart.Numbering.Save();
        applied.Add("numbering: canonical numbering definitions copied");
    }

    var tableStyleChanges = 0;
    foreach (var table in mainPart.Document.Descendants<Table>())
    {
        if (EnsureTableStyle(table, "tabelauerj"))
        {
            tableStyleChanges++;
        }
    }
    if (tableStyleChanges > 0)
    {
        applied.Add($"tables: applied canonical table style `tabelauerj` to {tableStyleChanges} table(s)");
    }
    EnsureTableStyleDisplayName(doc, "tabelauerj", "tabela_uerj");
    applied.Add("styles: canonical table style `tabelauerj` display name set to `tabela_uerj`");

    return applied;
}

static IReadOnlyList<string> SyncStylesFromSourceDocx(WordprocessingDocument targetDoc, WordprocessingDocument sourceDoc)
{
    var applied = new List<string>();

    var targetMainPart = targetDoc.MainDocumentPart ?? throw new InvalidOperationException("Target MainDocumentPart not found.");
    var sourceMainPart = sourceDoc.MainDocumentPart ?? throw new InvalidOperationException("Source MainDocumentPart not found.");

    if (sourceMainPart.StyleDefinitionsPart?.Styles is not null)
    {
        var targetStylesPart = targetMainPart.StyleDefinitionsPart ?? targetMainPart.AddNewPart<StyleDefinitionsPart>();
        using var sourceStylesStream = sourceMainPart.StyleDefinitionsPart.GetStream(FileMode.Open, FileAccess.Read);
        targetStylesPart.FeedData(sourceStylesStream);
        targetStylesPart.Styles?.Save();
        applied.Add($"styles: copied full StyleDefinitionsPart ({sourceMainPart.StyleDefinitionsPart.Styles.Elements<Style>().Count()} style(s))");
    }

    if (sourceMainPart.StylesWithEffectsPart is not null)
    {
        var targetEffectsPart = targetMainPart.StylesWithEffectsPart ?? targetMainPart.AddNewPart<StylesWithEffectsPart>();
        using var sourceEffectsStream = sourceMainPart.StylesWithEffectsPart.GetStream(FileMode.Open, FileAccess.Read);
        targetEffectsPart.FeedData(sourceEffectsStream);
        applied.Add("styles: copied StylesWithEffectsPart");
    }

    if (sourceMainPart.NumberingDefinitionsPart?.Numbering is not null)
    {
        var targetNumberingPart = targetMainPart.NumberingDefinitionsPart ?? targetMainPart.AddNewPart<NumberingDefinitionsPart>();
        using var sourceNumberingStream = sourceMainPart.NumberingDefinitionsPart.GetStream(FileMode.Open, FileAccess.Read);
        targetNumberingPart.FeedData(sourceNumberingStream);
        targetNumberingPart.Numbering?.Save();
        applied.Add($"numbering: copied full NumberingDefinitionsPart ({sourceMainPart.NumberingDefinitionsPart.Numbering.Elements<NumberingInstance>().Count()} numbering instance(s))");
    }

    EnsureTableStyleDisplayName(targetDoc, "tabelauerj", "tabela_uerj");
    applied.Add("styles: table style `tabelauerj` display name normalized to `tabela_uerj`");
    return applied;
}

static bool EnsureTableStyle(Table table, string styleId)
{
    var properties = table.GetFirstChild<TableProperties>();
    if (properties is null)
    {
        properties = new TableProperties();
        table.PrependChild(properties);
    }

    var current = properties.Elements<TableStyle>().FirstOrDefault();
    var hasSingleCorrectStyle = current is not null
        && string.Equals(current.Val?.Value, styleId, StringComparison.OrdinalIgnoreCase)
        && properties.Elements<TableStyle>().Count() == 1
        && properties.FirstChild == current;
    if (hasSingleCorrectStyle)
    {
        return false;
    }

    foreach (var old in properties.Elements<TableStyle>().ToList())
    {
        old.Remove();
    }
    properties.PrependChild(new TableStyle { Val = styleId });
    return true;
}

static void EnsureTableStyleDisplayName(WordprocessingDocument doc, string styleId, string styleName)
{
    var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
    if (styles is null)
    {
        return;
    }

    var style = styles.Elements<Style>()
        .FirstOrDefault(s => s.Type?.Value == StyleValues.Table
            && string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
    if (style is null)
    {
        return;
    }

    style.StyleName ??= new StyleName();
    style.StyleName.Val = styleName;
    styles.Save();
}

static void SanitizeStrictStyleCompatibility(Styles styles)
{
    const string wordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    foreach (var indentation in styles.Descendants<Indentation>())
    {
        indentation.RemoveAttribute("start", wordNs);
        indentation.RemoveAttribute("end", wordNs);
    }

    RemoveWord2010Ligatures(styles);
}

static int RemoveWord2010Ligatures(OpenXmlElement root)
{
    const string word2010Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    var ligatures = root
        .Descendants<OpenXmlUnknownElement>()
        .Where(element => element.LocalName == "ligatures" && string.Equals(element.NamespaceUri, word2010Ns, StringComparison.Ordinal))
        .ToList();

    foreach (var ligature in ligatures)
    {
        ligature.Remove();
    }

    return ligatures.Count;
}

static int RemoveWord2010LigaturesFromStylesPart(StyleDefinitionsPart stylesPart)
{
    if (stylesPart.Styles is { } styles)
    {
        var removed = RemoveWord2010Ligatures(styles);
        if (removed > 0)
        {
            styles.Save();
        }

        return removed;
    }

    const string word2010Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    using var readStream = stylesPart.GetStream(FileMode.Open, FileAccess.Read);
    using var reader = new StreamReader(readStream, Encoding.UTF8, true, leaveOpen: false);
    var rawXml = reader.ReadToEnd();
    if (string.IsNullOrWhiteSpace(rawXml))
    {
        return 0;
    }

    var xml = XDocument.Parse(rawXml, LoadOptions.PreserveWhitespace);
    var ligatures = xml.Descendants(XName.Get("ligatures", word2010Ns)).ToList();
    foreach (var ligature in ligatures)
    {
        ligature.Remove();
    }

    if (ligatures.Count == 0)
    {
        return 0;
    }

    var updatedXml = xml.Declaration is null
        ? xml.ToString(SaveOptions.DisableFormatting)
        : xml.Declaration + xml.ToString(SaveOptions.DisableFormatting);
    using var writeStream = new MemoryStream(Encoding.UTF8.GetBytes(updatedXml));
    stylesPart.FeedData(writeStream);
    return ligatures.Count;
}

static HashSet<string> ReadCanonicalStyleIds(string sourceDir)
{
    var manifestPath = Path.Combine(sourceDir, "used-styles-manifest.json");
    var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    if (!File.Exists(manifestPath))
    {
        return result;
    }

    using var doc = JsonDocument.Parse(File.ReadAllText(manifestPath, Encoding.UTF8));
    foreach (var propertyName in new[] { "paragraphStyles", "characterStyles", "tableStyles" })
    {
        if (!doc.RootElement.TryGetProperty(propertyName, out var array) || array.ValueKind != JsonValueKind.Array)
        {
            continue;
        }

        foreach (var item in array.EnumerateArray())
        {
            var value = item.GetString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                result.Add(value);
            }
        }
    }

    return result;
}

static HashSet<string> ExpandStyleDependencies(Styles styles, HashSet<string> initialStyleIds)
{
    var result = new HashSet<string>(initialStyleIds, StringComparer.OrdinalIgnoreCase);
    var changed = true;
    while (changed)
    {
        changed = false;
        foreach (var styleId in result.ToList())
        {
            var style = styles.Elements<Style>()
                .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
            if (style is null)
            {
                continue;
            }

            foreach (var dependency in new[]
            {
                style.BasedOn?.Val?.Value,
                style.NextParagraphStyle?.Val?.Value,
                style.LinkedStyle?.Val?.Value
            })
            {
                if (!string.IsNullOrWhiteSpace(dependency) && result.Add(dependency))
                {
                    changed = true;
                }
            }
        }
    }

    return result;
}

static string GetDefaultCanonicalStylesDirectory()
{
    var current = new DirectoryInfo(AppContext.BaseDirectory);
    while (current is not null)
    {
        if (File.Exists(Path.Combine(current.FullName, "DocxOpenXmlTools.csproj")))
        {
            return Path.Combine(current.FullName, "CanonicalStyles");
        }

        current = current.Parent;
    }

    return @"D:\docx_utils\DocxOpenXmlTools\CanonicalStyles";
}

static void RemoveUpdateFieldsOnOpen(WordprocessingDocument doc)
{
    var settings = doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
    if (settings is null)
    {
        return;
    }

    foreach (var existing in settings.Elements<UpdateFieldsOnOpen>().ToList())
    {
        existing.Remove();
    }

    settings.Save();
}

static void EnsureUpdateFieldsOnOpen(WordprocessingDocument doc)
{
    var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart ?? doc.MainDocumentPart?.AddNewPart<DocumentSettingsPart>();
    if (settingsPart is null)
    {
        throw new InvalidOperationException("Unable to access document settings.");
    }

    settingsPart.Settings ??= new Settings();
    foreach (var existing in settingsPart.Settings.Elements<UpdateFieldsOnOpen>().ToList())
    {
        existing.Remove();
    }

    settingsPart.Settings.AddChild(new UpdateFieldsOnOpen { Val = true }, true);
    settingsPart.Settings.Save();
}

static string ParagraphText(Paragraph paragraph)
{
    var pieces = paragraph
        .Descendants<Text>()
        .Select(t => t.Text)
        .Where(t => !string.IsNullOrEmpty(t));
    return Regex.Replace(string.Concat(pieces), "\\s+", " ").Trim();
}

static string ParagraphAuditDisplayText(Paragraph paragraph)
{
    var text = ParagraphText(paragraph);
    if (!string.IsNullOrWhiteSpace(text))
    {
        return text;
    }

    var mathText = MathOnlyText(paragraph);
    return string.IsNullOrWhiteSpace(mathText) ? text : $"[MATH] {mathText}";
}

static string ParagraphStyleId(Paragraph paragraph) =>
    paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";

static bool HasDrawing(Paragraph paragraph) => paragraph.Descendants<Drawing>().Any();

static bool HasMath(Paragraph paragraph) => ContainsMath(paragraph);

static RunningTextEligibilityResult RunningTextEligibility(Paragraph paragraph, string text, bool inReferences, bool inDeclarations)
{
    if (inReferences)
    {
        return new(false, "references section");
    }

    if (string.IsNullOrWhiteSpace(text))
    {
        return new(false, "empty");
    }

    if (paragraph.Ancestors<Table>().Any())
    {
        return new(false, "table paragraph");
    }

    if (HasDrawing(paragraph) || ContainsVisualObject(paragraph))
    {
        return new(false, "drawing/object");
    }

    if (HasMath(paragraph))
    {
        return new(false, "math object");
    }

    var styleId = ParagraphStyleId(paragraph);
    if (styleId.Equals("Figura", StringComparison.OrdinalIgnoreCase)
        || styleId.Equals("Tabela", StringComparison.OrdinalIgnoreCase)
        || styleId.Equals("legenda", StringComparison.OrdinalIgnoreCase)
        || styleId.Equals("dados", StringComparison.OrdinalIgnoreCase))
    {
        return new(false, $"protected style {styleId}");
    }

    if (IsReferencesHeading(text) || IsDeclarationsHeading(text))
    {
        return new(false, "section heading");
    }

    if (Regex.IsMatch(text, @"^\d+(\.\d+)*\s+\p{Lu}", RegexOptions.CultureInvariant) && IsMostlyUpper(text))
    {
        return new(false, "numbered section heading");
    }

    if (Regex.IsMatch(text, @"^(Figura|Tabela)\s+\d+\s+-", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
    {
        return new(false, "caption");
    }

    if (text.StartsWith("Fonte:", StringComparison.OrdinalIgnoreCase))
    {
        return new(false, "source paragraph");
    }

    var isKeywordParagraph = text.StartsWith("Palavras-chave:", StringComparison.OrdinalIgnoreCase)
        || text.StartsWith("Keywords:", StringComparison.OrdinalIgnoreCase);
    if (!isKeywordParagraph && text.Length < 80)
    {
        return new(false, inDeclarations ? "short declaration" : "short non-running paragraph");
    }

    return new(true, "running text");
}

static bool ApplyNormalTimesStyle(Paragraph paragraph, string targetStyleId, string targetFont, string targetHalfPoints, RevisionMetadata metadata)
{
    var changed = false;
    paragraph.ParagraphProperties ??= new ParagraphProperties();
    var paragraphProperties = paragraph.ParagraphProperties;
    var currentStyle = paragraphProperties.ParagraphStyleId?.Val?.Value ?? "";
    if (!string.Equals(currentStyle, targetStyleId, StringComparison.Ordinal))
    {
        AddParagraphPropertiesChange(paragraphProperties, metadata);
        if (paragraphProperties.ParagraphStyleId is null)
        {
            paragraphProperties.PrependChild(new ParagraphStyleId { Val = targetStyleId });
        }
        else
        {
            paragraphProperties.ParagraphStyleId.Val = targetStyleId;
        }
        changed = true;
    }

    foreach (var run in paragraph.Descendants<Run>().Where(r => r.Descendants<Text>().Any()).ToList())
    {
        run.RunProperties ??= new RunProperties();
        var runProperties = run.RunProperties;
        if (!RunMatchesFont(runProperties, targetFont, targetHalfPoints))
        {
            AddRunPropertiesChange(runProperties, metadata);
            SetRunFont(runProperties, targetFont, targetHalfPoints);
            changed = true;
        }
    }

    return changed;
}

static void AddParagraphPropertiesChange(ParagraphProperties paragraphProperties, RevisionMetadata metadata)
{
    if (paragraphProperties.Elements<ParagraphPropertiesChange>().Any())
    {
        return;
    }

    var previous = new PreviousParagraphProperties();
    foreach (var child in paragraphProperties.ChildElements.Where(c => c is not ParagraphPropertiesChange and not ParagraphMarkRunProperties).ToList())
    {
        previous.Append(child.CloneNode(true));
    }

    paragraphProperties.Append(new ParagraphPropertiesChange(previous)
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    });
}

static void AddRunPropertiesChange(RunProperties runProperties, RevisionMetadata metadata)
{
    if (runProperties.Elements<RunPropertiesChange>().Any())
    {
        return;
    }

    var previous = new PreviousRunProperties();
    foreach (var child in runProperties.ChildElements.Where(c => c is not RunPropertiesChange).ToList())
    {
        previous.Append(child.CloneNode(true));
    }

    runProperties.Append(new RunPropertiesChange(previous)
    {
        Id = metadata.NextRevisionId(),
        Author = metadata.Author,
        Date = metadata.DateUtc
    });
}

static bool RunMatchesFont(RunProperties runProperties, string targetFont, string targetHalfPoints)
{
    var fonts = runProperties.RunFonts;
    var size = runProperties.FontSize?.Val?.Value ?? "";
    var csSize = runProperties.FontSizeComplexScript?.Val?.Value ?? "";
    return fonts is not null
        && string.Equals(fonts.Ascii?.Value ?? "", targetFont, StringComparison.Ordinal)
        && string.Equals(fonts.HighAnsi?.Value ?? "", targetFont, StringComparison.Ordinal)
        && string.Equals(fonts.EastAsia?.Value ?? "", targetFont, StringComparison.Ordinal)
        && string.Equals(fonts.ComplexScript?.Value ?? "", targetFont, StringComparison.Ordinal)
        && string.Equals(size, targetHalfPoints, StringComparison.Ordinal)
        && string.Equals(csSize, targetHalfPoints, StringComparison.Ordinal);
}

static void SetRunFont(RunProperties runProperties, string targetFont, string targetHalfPoints)
{
    runProperties.RunFonts ??= new RunFonts();
    runProperties.RunFonts.Ascii = targetFont;
    runProperties.RunFonts.HighAnsi = targetFont;
    runProperties.RunFonts.EastAsia = targetFont;
    runProperties.RunFonts.ComplexScript = targetFont;
    runProperties.FontSize ??= new FontSize();
    runProperties.FontSize.Val = targetHalfPoints;
    runProperties.FontSizeComplexScript ??= new FontSizeComplexScript();
    runProperties.FontSizeComplexScript.Val = targetHalfPoints;
    MoveRunPropertiesChangeToEnd(runProperties);
}

static bool RevertOwnFormattingRevisions(Paragraph paragraph, string author)
{
    var reverted = false;
    foreach (var change in paragraph.ParagraphProperties?.Elements<ParagraphPropertiesChange>().ToList() ?? [])
    {
        if (!string.Equals(change.Author?.Value ?? "", author, StringComparison.Ordinal))
        {
            continue;
        }

        var pPr = paragraph.ParagraphProperties!;
        var previous = change.GetFirstChild<PreviousParagraphProperties>();
        pPr.RemoveAllChildren();
        if (previous is not null)
        {
            foreach (var child in previous.ChildElements)
            {
                pPr.Append(child.CloneNode(true));
            }
        }
        reverted = true;
    }

    foreach (var runProperties in paragraph.Descendants<RunProperties>().ToList())
    {
        foreach (var change in runProperties.Elements<RunPropertiesChange>().ToList())
        {
            if (!string.Equals(change.Author?.Value ?? "", author, StringComparison.Ordinal))
            {
                continue;
            }

            var previous = change.GetFirstChild<PreviousRunProperties>();
            runProperties.RemoveAllChildren();
            if (previous is not null)
            {
                foreach (var child in previous.ChildElements)
                {
                    runProperties.Append(child.CloneNode(true));
                }
            }
            reverted = true;
        }
    }

    return reverted;
}

static void NormalizeFormattingRevisionMarkup(WordprocessingDocument doc)
{
    foreach (var pPrChange in doc.MainDocumentPart!.Document.Descendants<ParagraphPropertiesChange>())
    {
        foreach (var previous in pPrChange.Elements<PreviousParagraphProperties>())
        {
            foreach (var invalidParagraphMarkRunProperties in previous.Elements<ParagraphMarkRunProperties>().ToList())
            {
                invalidParagraphMarkRunProperties.Remove();
            }
        }
    }

    foreach (var runProperties in doc.MainDocumentPart.Document.Descendants<RunProperties>())
    {
        MoveRunPropertiesChangeToEnd(runProperties);
    }
}

static void MoveRunPropertiesChangeToEnd(RunProperties runProperties)
{
    var changes = runProperties.Elements<RunPropertiesChange>().ToList();
    foreach (var change in changes)
    {
        change.Remove();
    }

    foreach (var change in changes)
    {
        runProperties.Append(change);
    }
}

static bool ContainsVisualObject(OpenXmlElement element) =>
    element.Descendants().Any(e => e.LocalName is "pict" or "object" or "shape" or "AlternateContent");

static bool IsReferencesHeading(string text) =>
    Normalize(text).Equals(Normalize("REFERÃŠNCIAS"), StringComparison.Ordinal);

static bool IsDeclarationsHeading(string text) =>
    Normalize(text).Equals(Normalize("DECLARAÃ‡Ã•ES"), StringComparison.Ordinal);

static bool IsMostlyUpper(string text)
{
    var letters = text.Where(char.IsLetter).ToList();
    if (letters.Count == 0)
    {
        return false;
    }

    var upper = letters.Count(char.IsUpper);
    return upper / (double)letters.Count >= 0.8;
}

static string TruncateForReport(string text)
{
    var normalized = Regex.Replace(text, "\\s+", " ").Trim();
    return normalized.Length <= 180 ? normalized : normalized[..180] + "...";
}

static string ResolveNumberingFormat(WordprocessingDocument doc, string numberingId, string numberingLevel)
{
    if (string.IsNullOrWhiteSpace(numberingId))
    {
        return "";
    }

    var numbering = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
    if (numbering is null)
    {
        return "";
    }

    var num = numbering.Elements<NumberingInstance>()
        .FirstOrDefault(n => string.Equals(n.NumberID?.Value.ToString(CultureInfo.InvariantCulture), numberingId, StringComparison.Ordinal));
    var abstractNumId = num?.AbstractNumId?.Val?.Value;
    if (abstractNumId is null)
    {
        return "";
    }

    var level = string.IsNullOrWhiteSpace(numberingLevel) ? 0 : int.Parse(numberingLevel, CultureInfo.InvariantCulture);
    var abstractNum = numbering.Elements<AbstractNum>()
        .FirstOrDefault(n => n.AbstractNumberId?.Value == abstractNumId.Value);
    var lvl = abstractNum?.Elements<Level>()
        .FirstOrDefault(l => l.LevelIndex?.Value == level)
        ?? abstractNum?.Elements<Level>().FirstOrDefault();

    if (lvl is null)
    {
        return $"abstractNumId={abstractNumId.Value}";
    }

    var format = lvl.NumberingFormat?.Val?.Value.ToString() ?? "";
    var text = lvl.LevelText?.Val?.Value ?? "";
    var start = lvl.StartNumberingValue?.Val?.Value.ToString(CultureInfo.InvariantCulture) ?? "";
    return $"abstractNumId={abstractNumId.Value}; level={lvl.LevelIndex?.Value}; start={start}; format={format}; text={text}";
}

static bool ContainsMath(OpenXmlElement element) =>
    element.Descendants().Any(e => e.LocalName is "oMath" or "oMathPara");

static int LastMathOrder(Paragraph paragraph)
{
    var last = -1;
    var order = 0;
    foreach (var element in paragraph.Descendants())
    {
        if (element.LocalName is "oMath" or "oMathPara")
        {
            last = order;
        }

        order++;
    }

    return last;
}

static IEnumerable<ExtractedSeqField> ExtractSeqFields(Paragraph paragraph)
{
    var order = 0;
    var inField = false;
    var inResult = false;
    var beginOrder = -1;
    var instruction = new StringBuilder();
    var result = new StringBuilder();

    foreach (var element in paragraph.Descendants())
    {
        if (element is SimpleField simpleField)
        {
            var simpleInstruction = simpleField.Instruction?.Value ?? "";
            if (IsSeqInstruction(simpleInstruction))
            {
                var resultText = ElementText(simpleField);
                yield return BuildExtractedSeqField(order, simpleInstruction, resultText);
            }
        }

        if (element is FieldChar fieldChar)
        {
            var fieldCharType = fieldChar.FieldCharType?.Value;
            if (fieldCharType == FieldCharValues.Begin)
            {
                inField = true;
                inResult = false;
                beginOrder = order;
                instruction.Clear();
                result.Clear();
            }
            else if (fieldCharType == FieldCharValues.Separate && inField)
            {
                inResult = true;
            }
            else if (fieldCharType == FieldCharValues.End && inField)
            {
                var instructionText = instruction.ToString();
                if (IsSeqInstruction(instructionText))
                {
                    yield return BuildExtractedSeqField(beginOrder, instructionText, result.ToString());
                }

                inField = false;
                inResult = false;
                beginOrder = -1;
                instruction.Clear();
                result.Clear();
            }
        }
        else if (inField && !inResult && element is FieldCode fieldCode)
        {
            instruction.Append(fieldCode.Text);
        }
        else if (inField && inResult && element is Text text)
        {
            result.Append(text.Text);
        }

        order++;
    }
}

static ExtractedSeqField BuildExtractedSeqField(int beginOrder, string instruction, string resultText)
{
    var normalizedInstruction = Regex.Replace(instruction, "\\s+", " ").Trim();
    var sequenceName = ParseSeqName(normalizedInstruction);
    var normalizedResult = Regex.Replace(resultText, "\\s+", " ").Trim();
    var number = int.TryParse(Regex.Match(normalizedResult, @"\d+").Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
        ? parsed
        : (int?)null;

    return new ExtractedSeqField(beginOrder, normalizedInstruction, sequenceName, normalizedResult, number);
}

static bool IsSeqInstruction(string instruction) =>
    Regex.IsMatch(instruction, @"(^|\s)SEQ\s+", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

static string ParseSeqName(string instruction)
{
    var match = Regex.Match(instruction, @"(?:^|\s)SEQ\s+([^\\\s]+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    return match.Success ? match.Groups[1].Value.Trim() : "";
}

static bool IsCaptionSeq(string sequenceName, string paragraphText, string styleId)
{
    var normalizedName = Normalize(sequenceName);
    var normalizedText = Normalize(paragraphText);
    var normalizedStyle = Normalize(styleId);
    return normalizedName is "figura" or "tabela" or "quadro"
        || normalizedStyle.Contains("figura", StringComparison.Ordinal)
        || normalizedStyle.Contains("tabela", StringComparison.Ordinal)
        || Regex.IsMatch(normalizedText, @"^(figura|tabela|quadro)\s+\d+\b", RegexOptions.CultureInvariant);
}

static bool HasParenthesizedSeqResult(string paragraphText, int number) =>
    Regex.IsMatch(paragraphText, @$"\(\s*{number.ToString(CultureInfo.InvariantCulture)}\s*\)\s*$", RegexOptions.CultureInvariant);

static string MathOnlyText(Paragraph paragraph)
{
    var pieces = paragraph
        .Descendants()
        .Where(e => e.LocalName == "oMath")
        .Select(e => e.InnerText)
        .Where(t => !string.IsNullOrEmpty(t));
    return Regex.Replace(string.Concat(pieces), "\\s+", " ").Trim();
}

static ParagraphAuditEntry? PreviousParagraph(IReadOnlyList<ParagraphAuditEntry> paragraphs, int blockIndex, int offset = 1) =>
    paragraphs.Where(p => p.BlockIndex < blockIndex).Reverse().Skip(offset - 1).FirstOrDefault();

static ParagraphAuditEntry? NextParagraph(IReadOnlyList<ParagraphAuditEntry> paragraphs, int blockIndex, int offset = 1) =>
    paragraphs.Where(p => p.BlockIndex > blockIndex).Skip(offset - 1).FirstOrDefault();

static bool LooksLikeTableTitle(string text) =>
    Regex.IsMatch(text, @"^\s*Tabela\s+\d+\s*[-â€“â€”]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

static bool LooksLikeFigureTitle(string text) =>
    Regex.IsMatch(text, @"^\s*(Figura|GrÃ¡fico|Grafico|Quadro|Mapa|Fluxograma|Imagem)\s+\d+\s*[-â€“â€”]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

static bool IsSourceLine(string text) =>
    Regex.IsMatch(text, @"^\s*Fonte\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

static bool HasVerticalBorders(TableBorders? borders) =>
    borders?.LeftBorder is not null
    || borders?.RightBorder is not null
    || borders?.InsideVerticalBorder is not null;

static IReadOnlyDictionary<string, string> BorderSummary(TableBorders? borders)
{
    if (borders is null)
    {
        return new Dictionary<string, string>(StringComparer.Ordinal);
    }

    return new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["top"] = BorderValue(borders.TopBorder),
        ["bottom"] = BorderValue(borders.BottomBorder),
        ["left"] = BorderValue(borders.LeftBorder),
        ["right"] = BorderValue(borders.RightBorder),
        ["insideH"] = BorderValue(borders.InsideHorizontalBorder),
        ["insideV"] = BorderValue(borders.InsideVerticalBorder)
    };
}

static string BorderValue(BorderType? border)
{
    if (border is null)
    {
        return "";
    }

    var value = border.Val?.Value.ToString() ?? "";
    var size = border.Size?.Value.ToString(CultureInfo.InvariantCulture) ?? "";
    var color = border.Color?.Value ?? "";
    return $"{value}; sz={size}; color={color}";
}

static double EmuToCm(long emu) => Math.Round(emu / 360000.0, 2);

static string BuildLayoutAuditMarkdown(LayoutAuditReport report)
{
    var academicTables = report.Tables.Where(t => t.IsAcademicTable).ToList();
    var academicFigures = report.Figures.Where(f => f.IsAcademicFigure).ToList();
    var tableIssues = academicTables.SelectMany(t => t.Issues.Select(i => (t, i))).ToList();
    var figureIssues = academicFigures.SelectMany(f => f.Issues.Select(i => (f, i))).ToList();
    var builder = new StringBuilder();
    builder.AppendLine("# Auditoria de layout de tabelas e figuras");
    builder.AppendLine();
    builder.AppendLine($"- Documento: `{report.DocxPath}`");
    builder.AppendLine($"- Gerado em UTC: `{report.GeneratedAtUtc:O}`");
    builder.AppendLine("- Base normativa usada: Roteiro UERJ/ABNT NBR 14724 para ilustraÃ§Ãµes e tabelas; IBGE para apresentaÃ§Ã£o tabular estatÃ­stica; regra local `.ai/padrao_referencias_abnt_uerj.md`.");
    builder.AppendLine();
    builder.AppendLine("## CritÃ©rios aplicados");
    builder.AppendLine();
    builder.AppendLine("- Tabelas e figuras devem ser mencionadas no texto, numeradas sequencialmente, inseridas no corpo, com tÃ­tulo/legenda acima e fonte abaixo.");
    builder.AppendLine("- Tabelas estatÃ­sticas devem ter tÃ­tulo, corpo, cabeÃ§alho/coluna indicadora quando aplicÃ¡vel, fonte/notas, e desenho tabular sÃ³brio; bordas verticais/laterais devem ser evitadas salvo necessidade semÃ¢ntica.");
    builder.AppendLine("- Figuras nÃ£o precisam de moldura/linha externa por norma; contorno sÃ³ Ã© adequado quando integra o prÃ³prio grÃ¡fico/quadro ou melhora legibilidade sem poluir o layout.");
    builder.AppendLine("- Neste DOCX, legendas sÃ£o parÃ¡grafos com estilos `Figura` e `Tabela`; a numeraÃ§Ã£o vem do estilo, nÃ£o de texto manual.");
    builder.AppendLine();
    builder.AppendLine("## Resumo");
    builder.AppendLine();
    builder.AppendLine($"- Tabelas Word detectadas: {report.TableCount}; tabelas acadÃªmicas: {report.AcademicTableCount}.");
    builder.AppendLine($"- Figuras/desenhos detectados: {report.FigureCount}; figuras acadÃªmicas: {report.AcademicFigureCount}.");
    builder.AppendLine($"- OcorrÃªncias em tabelas a revisar: {tableIssues.Count}.");
    builder.AppendLine($"- OcorrÃªncias em figuras a revisar: {figureIssues.Count}.");
    builder.AppendLine();
    builder.AppendLine("## Tabelas acadÃªmicas");
    builder.AppendLine();
    foreach (var table in academicTables)
    {
        builder.AppendLine($"### Tabela audit ordinal {table.Ordinal}, block {table.BlockIndex}");
        builder.AppendLine();
        builder.AppendLine($"- TÃ­tulo anterior: `{table.PreviousParagraph?.Text ?? ""}`; estilo: `{table.PreviousParagraph?.StyleId ?? ""}`.");
        builder.AppendLine($"- DimensÃ£o: {table.RowCount} linhas x {table.ColumnCount} colunas; cÃ©lulas: {table.CellCount}; largura: `{table.WidthType}:{table.Width}`; alinhamento: `{table.Justification}`; layout: `{table.LayoutType}`.");
        builder.AppendLine($"- Estilo de tabela: `{table.TableStyleId}`; bordas: {string.Join("; ", table.Borders.Select(kv => $"{kv.Key}={kv.Value}"))}; celulas mescladas: {table.HasMergedCells}.");
        builder.AppendLine($"- Estilos de parÃ¡grafo nas cÃ©lulas: {string.Join(", ", table.CellParagraphStyles.Select(kv => $"{kv.Key}={kv.Value}"))}.");
        builder.AppendLine($"- Tamanhos de fonte diretos nas cÃ©lulas: {(table.DirectFontSizes.Count == 0 ? "(sem tamanho direto; herdado de estilos)" : string.Join(", ", table.DirectFontSizes.Select(kv => $"{kv.Key}={kv.Value}")))}.");
        builder.AppendLine($"- Fonte detectada: `{table.SourceCandidate}`.");
        builder.AppendLine(table.Issues.Count == 0
            ? "- DiagnÃ³stico: sem inconsistÃªncia estrutural detectada por Open XML."
            : $"- Revisar: {string.Join("; ", table.Issues)}.");
        builder.AppendLine();
    }

    builder.AppendLine("## Figuras acadÃªmicas");
    builder.AppendLine();
    foreach (var figure in academicFigures)
    {
        builder.AppendLine($"### Figura audit ordinal {figure.Ordinal}, P[{figure.ParagraphIndex}]");
        builder.AppendLine();
        builder.AppendLine($"- Legenda: `{figure.ParagraphText}`; estilo: `{figure.ParagraphStyleId}`.");
        builder.AppendLine($"- Tipo: `{figure.DrawingKind}`; wrap: `{figure.Wrap}`; tamanho: {figure.WidthCm:0.##} cm x {figure.HeightCm:0.##} cm; contorno detectado: {figure.HasOutline}.");
        builder.AppendLine($"- Fonte detectada: `{figure.SourceCandidate}`.");
        builder.AppendLine(figure.Issues.Count == 0
            ? "- DiagnÃ³stico: sem inconsistÃªncia estrutural detectada por Open XML."
            : $"- Revisar: {string.Join("; ", figure.Issues)}.");
        builder.AppendLine();
    }

    return builder.ToString();
}

static string FirstNonEmptyTableText(Table table)
{
    foreach (var cell in table.Descendants<TableCell>())
    {
        var text = Regex.Replace(string.Concat(cell.Descendants<Text>().Select(t => t.Text)), "\\s+", " ").Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text.Length <= 180 ? text : text[..180];
        }
    }

    return "";
}

static List<FieldInfoEntry> ExtractFields(Paragraph paragraph)
{
    var fields = new List<FieldInfoEntry>();
    var runs = paragraph.Elements<Run>().ToList();
    for (var i = 0; i < runs.Count; i++)
    {
        var fldChar = runs[i].GetFirstChild<FieldChar>();
        if (fldChar?.FieldCharType?.Value != FieldCharValues.Begin)
        {
            continue;
        }

        var instruction = new StringBuilder();
        var result = new StringBuilder();
        var inResult = false;
        for (var j = i + 1; j < runs.Count; j++)
        {
            var current = runs[j];
            var currentFldChar = current.GetFirstChild<FieldChar>();
            if (currentFldChar?.FieldCharType?.Value == FieldCharValues.Separate)
            {
                inResult = true;
                continue;
            }

            if (currentFldChar?.FieldCharType?.Value == FieldCharValues.End)
            {
                fields.Add(new FieldInfoEntry(
                    Regex.Replace(instruction.ToString(), "\\s+", " ").Trim(),
                    Regex.Replace(result.ToString(), "\\s+", " ").Trim()));
                i = j;
                break;
            }

            var instrText = current.GetFirstChild<FieldCode>()?.Text;
            if (!string.IsNullOrEmpty(instrText))
            {
                instruction.Append(instrText);
            }

            if (inResult)
            {
                foreach (var text in current.Descendants<Text>())
                {
                    result.Append(text.Text);
                }
            }
        }
    }

    return fields;
}

static string ElementText(OpenXmlElement element)
{
    var pieces = element
        .Descendants<Text>()
        .Select(t => t.Text)
        .Concat(element.Descendants<DeletedText>().Select(t => t.Text))
        .Where(t => !string.IsNullOrEmpty(t));
    return Regex.Replace(string.Concat(pieces), "\\s+", " ").Trim();
}

static void WriteCommentsJson(IReadOnlyList<CommentListEntry> entries)
{
    var payload = new
    {
        comments = entries.Select(entry => new
        {
            id = entry.Id,
            autor = entry.Author,
            conteudo = entry.Text,
            orientacao = "",
            data = entry.Date,
            ancora = string.IsNullOrWhiteSpace(entry.AnchorText) ? null : entry.AnchorText,
            parentCommentId = entry.ParentCommentId
        })
    };
    Console.WriteLine(JsonSerializer.Serialize(payload, CliOptions.JsonOptionsIndented()));
}

static TabularOutput BuildCommentsOutputTable(IReadOnlyList<CommentListEntry> entries, bool includeIndex) =>
    includeIndex
        ? new(
            ["(index)", "id", "autor", "conteudo", "orientacao"],
            entries.Select((entry, index) => new[]
            {
                index.ToString(CultureInfo.InvariantCulture),
                entry.Id,
                entry.Author,
                entry.Text,
                ""
            }).ToList())
        : new(
            ["id", "autor", "conteudo", "orientacao"],
            entries.Select(entry => new[] { entry.Id, entry.Author, entry.Text, "" }).ToList());

static void WriteTabularOutput(TabularOutput table, string format)
{
    if (format is "table" or "console" or "console-table")
    {
        WriteSpectreConsoleTable(table);
        return;
    }

    if (format is "md" or "markdown")
    {
        Console.Write(BuildMarkdownTable(table));
        return;
    }

    throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported tabular output format.");
}

static string BuildMarkdownTable(TabularOutput table)
{
    var builder = new StringBuilder();
    AppendMarkdownRow(builder, table.Headers);
    AppendMarkdownRow(builder, table.Headers.Select(_ => "---").ToArray());
    foreach (var row in table.Rows)
    {
        AppendMarkdownRow(builder, row);
    }

    var markdown = builder.ToString();
    _ = Markdown.Parse(markdown);
    return markdown;
}

static void AppendMarkdownRow(StringBuilder builder, IReadOnlyList<string> cells)
{
    var sanitized = cells
        .Select(cell => EscapeMarkdownTableCell(NormalizeTableCell(cell)));
    builder.Append("| ");
    builder.Append(string.Join(" | ", sanitized));
    builder.AppendLine(" |");
}

static void WriteSpectreConsoleTable(TabularOutput output)
{
    var table = new Spectre.Console.Table
    {
        Border = Spectre.Console.TableBorder.Square,
        ShowRowSeparators = true
    };

    foreach (var header in output.Headers)
    {
        table.AddColumn(new Spectre.Console.TableColumn(EscapeSpectreMarkup(NormalizeTableCell(header))));
    }

    foreach (var row in output.Rows)
    {
        Spectre.Console.TableExtensions.AddRow(table, row.Select(cell => EscapeSpectreMarkup(NormalizeTableCell(cell))).ToArray());
    }

    var console = Spectre.Console.AnsiConsole.Create(new Spectre.Console.AnsiConsoleSettings
    {
        Ansi = Spectre.Console.AnsiSupport.No,
        ColorSystem = Spectre.Console.ColorSystemSupport.NoColors,
        Out = new Spectre.Console.AnsiConsoleOutput(Console.Out)
    });
    console.Write(table);
}

static string NormalizeTableCell(string? value) =>
    Regex.Replace(value ?? "", "\\s+", " ").Trim();

static string EscapeMarkdownTableCell(string value) =>
    value.Replace("\\", "\\\\", StringComparison.Ordinal)
        .Replace("|", "\\|", StringComparison.Ordinal);

static string EscapeSpectreMarkup(string value) =>
    Spectre.Console.Markup.Escape(value);

static string GetRevisionAuthor(OpenXmlElement element) => element switch
{
    InsertedRun inserted => inserted.Author?.Value ?? "",
    DeletedRun deleted => deleted.Author?.Value ?? "",
    _ => ""
};

static string GetRevisionDate(OpenXmlElement element) => element switch
{
    InsertedRun inserted => inserted.Date is null ? "" : inserted.Date.Value.ToString("O", CultureInfo.InvariantCulture),
    DeletedRun deleted => deleted.Date is null ? "" : deleted.Date.Value.ToString("O", CultureInfo.InvariantCulture),
    _ => ""
};

static bool NormalizedStartsWith(string text, string prefix)
{
    var normalizedText = Normalize(text);
    var normalizedPrefix = Normalize(prefix);
    return normalizedText.StartsWith(normalizedPrefix, StringComparison.Ordinal);
}

static string Normalize(string value)
{
    var formD = value.Normalize(NormalizationForm.FormD);
    var builder = new StringBuilder(formD.Length);
    foreach (var c in formD)
    {
        if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
        {
            builder.Append(char.ToLowerInvariant(c));
        }
    }

    return Regex.Replace(builder.ToString(), "\\s+", " ").Trim();
}

static string RandomHex(int digits)
{
    Span<byte> bytes = stackalloc byte[(digits + 1) / 2];
    Random.Shared.NextBytes(bytes);
    return Convert.ToHexString(bytes)[..digits];
}

static IReadOnlyList<LatexToken> TokenizeLatex(string latex)
{
    var tokens = new List<LatexToken>();
    for (var i = 0; i < latex.Length;)
    {
        var c = latex[i];
        if (char.IsWhiteSpace(c))
        {
            i++;
            continue;
        }

        if (c == '\\')
        {
            var j = i + 1;
            while (j < latex.Length && char.IsLetter(latex[j]))
            {
                j++;
            }

            if (j == i + 1 && j < latex.Length)
            {
                j++;
            }

            tokens.Add(new LatexToken(LatexTokenKind.Command, latex[i..j]));
            i = j;
            continue;
        }

        if (char.IsLetterOrDigit(c))
        {
            var j = i + 1;
            while (j < latex.Length && char.IsLetterOrDigit(latex[j]))
            {
                j++;
            }

            tokens.Add(new LatexToken(char.IsDigit(c) ? LatexTokenKind.Number : LatexTokenKind.Identifier, latex[i..j]));
            i = j;
            continue;
        }

        tokens.Add(new LatexToken(c switch
        {
            '{' => LatexTokenKind.LBrace,
            '}' => LatexTokenKind.RBrace,
            '(' => LatexTokenKind.LParen,
            ')' => LatexTokenKind.RParen,
            '[' => LatexTokenKind.LBracket,
            ']' => LatexTokenKind.RBracket,
            '_' => LatexTokenKind.Subscript,
            '^' => LatexTokenKind.Superscript,
            ',' => LatexTokenKind.Comma,
            _ => LatexTokenKind.Symbol
        }, c.ToString()));
        i++;
    }

    return tokens;
}

sealed class LatexMathParser
{
    private readonly IReadOnlyList<LatexToken> _tokens;
    private int _index;

    public LatexMathParser(IReadOnlyList<LatexToken> tokens) => _tokens = tokens;

    public LatexNode Parse()
    {
        var expr = ParseSequence();
        if (_index != _tokens.Count)
        {
            throw new InvalidOperationException($"Unexpected token `{_tokens[_index].Value}` at position {_index}.");
        }

        return expr;
    }

    private LatexNode ParseSequence(params LatexTokenKind[] terminators)
    {
        var items = new List<LatexNode>();
        while (_index < _tokens.Count && !terminators.Contains(Current.Kind))
        {
            items.Add(ParseAtomWithScripts());
        }

        return new SequenceNode(items);
    }

    private LatexNode ParseAtomWithScripts()
    {
        LatexNode node = ParseAtom();
        LatexNode? sub = null;
        LatexNode? sup = null;

        while (_index < _tokens.Count && (Current.Kind == LatexTokenKind.Subscript || Current.Kind == LatexTokenKind.Superscript))
        {
            var kind = Current.Kind;
            _index++;
            var arg = ParseScriptArgument();
            if (kind == LatexTokenKind.Subscript)
            {
                sub = arg;
            }
            else
            {
                sup = arg;
            }
        }

        if (sub is null && sup is null)
        {
            return node;
        }

        return new ScriptNode(node, sub, sup);
    }

    private LatexNode ParseScriptArgument()
    {
        if (_index >= _tokens.Count)
        {
            throw new InvalidOperationException("Unexpected end of LaTeX while reading script argument.");
        }

        if (Current.Kind == LatexTokenKind.LBrace)
        {
            _index++;
            var group = ParseSequence(LatexTokenKind.RBrace);
            Expect(LatexTokenKind.RBrace);
            return group;
        }

        return ParseAtomWithScripts();
    }

    private LatexNode ParseAtom()
    {
        if (_index >= _tokens.Count)
        {
            throw new InvalidOperationException("Unexpected end of LaTeX.");
        }

        var token = Current;
        _index++;
        return token.Kind switch
        {
            LatexTokenKind.LBrace => ParseGrouped(LatexTokenKind.RBrace),
            LatexTokenKind.LParen => new DelimitedNode("(", ")", ParseDelimitedContent(LatexTokenKind.RParen)),
            LatexTokenKind.LBracket => new DelimitedNode("[", "]", ParseDelimitedContent(LatexTokenKind.RBracket)),
            LatexTokenKind.Command => ParseCommand(token.Value),
            LatexTokenKind.Identifier or LatexTokenKind.Number or LatexTokenKind.Symbol or LatexTokenKind.Comma => new SymbolNode(MapLatexSymbol(token.Value)),
            _ => throw new InvalidOperationException($"Unexpected token `{token.Value}`.")
        };
    }

    private LatexNode ParseGrouped(LatexTokenKind terminator)
    {
        var group = ParseSequence(terminator);
        Expect(terminator);
        return group;
    }

    private LatexNode ParseDelimitedContent(LatexTokenKind terminator)
    {
        var content = ParseSequence(terminator);
        Expect(terminator);
        return content;
    }

    private LatexNode ParseCommand(string command) => command switch
    {
        "\\frac" => new FractionNode(ParseRequiredGroup(), ParseRequiredGroup()),
        "\\sum" => new SymbolNode("âˆ‘"),
        "\\cdot" => new SymbolNode("â‹…"),
        "\\times" => new SymbolNode("Ã—"),
        "\\in" => new SymbolNode("âˆˆ"),
        "\\to" => new SymbolNode("â†’"),
        "\\ge" or "\\geq" => new SymbolNode("â‰¥"),
        "\\le" or "\\leq" => new SymbolNode("â‰¤"),
        "\\lambda" => new SymbolNode("Î»"),
        "\\alpha" => new SymbolNode("Î±"),
        "\\beta" => new SymbolNode("Î²"),
        "\\gamma" => new SymbolNode("Î³"),
        "\\delta" => new SymbolNode("Î´"),
        "\\mu" => new SymbolNode("Î¼"),
        "\\sigma" => new SymbolNode("Ïƒ"),
        "\\pi" => new SymbolNode("Ï€"),
        "\\infty" => new SymbolNode("âˆž"),
        "\\{" => new SymbolNode("{"),
        "\\}" => new SymbolNode("}"),
        "\\[" => new SymbolNode("["),
        "\\]" => new SymbolNode("]"),
        "\\lbrace" => new SymbolNode("{"),
        "\\rbrace" => new SymbolNode("}"),
        "\\left" => ParseLeftRightDelimited(),
        "\\right" => new SymbolNode(""),
        _ => new SymbolNode(MapLatexSymbol(command))
    };

    private LatexNode ParseLeftRightDelimited()
    {
        if (_index >= _tokens.Count)
        {
            throw new InvalidOperationException("Missing delimiter after \\left.");
        }

        var open = NormalizeDelimiterToken(_tokens[_index++].Value);
        var content = ParseSequenceUntilRight();
        var close = ")";
        if (_index < _tokens.Count && _tokens[_index].Kind == LatexTokenKind.Command && _tokens[_index].Value == "\\right")
        {
            _index++;
            if (_index < _tokens.Count)
            {
                close = NormalizeDelimiterToken(_tokens[_index++].Value);
            }
        }

        return new DelimitedNode(open, close, content);
    }

    private LatexNode ParseSequenceUntilRight()
    {
        var items = new List<LatexNode>();
        while (_index < _tokens.Count && !(Current.Kind == LatexTokenKind.Command && Current.Value == "\\right"))
        {
            items.Add(ParseAtomWithScripts());
        }

        return new SequenceNode(items);
    }

    private LatexNode ParseRequiredGroup()
    {
        if (_index >= _tokens.Count || Current.Kind != LatexTokenKind.LBrace)
        {
            throw new InvalidOperationException("Expected `{` after command.");
        }

        _index++;
        var group = ParseSequence(LatexTokenKind.RBrace);
        Expect(LatexTokenKind.RBrace);
        return group;
    }

    private void Expect(LatexTokenKind kind)
    {
        if (_index >= _tokens.Count || _tokens[_index].Kind != kind)
        {
            throw new InvalidOperationException($"Expected token `{kind}`.");
        }

        _index++;
    }

    private LatexToken Current => _tokens[_index];

    private static string NormalizeDelimiterToken(string value) => value switch
    {
        "." => "",
        "\\{" or "\\lbrace" => "{",
        "\\}" or "\\rbrace" => "}",
        "\\[" => "[",
        "\\]" => "]",
        _ => value
    };

    private static string MapLatexSymbol(string value) => value switch
    {
        "\\prime" => "'",
        "\\mid" => "|",
        "\\{" => "{",
        "\\}" => "}",
        _ => value
    };
}

abstract record LatexNode
{
    public abstract IEnumerable<OpenXmlElement> ToOpenXml();
}

sealed record SequenceNode(IReadOnlyList<LatexNode> Items) : LatexNode
{
    public override IEnumerable<OpenXmlElement> ToOpenXml() => Items.SelectMany(item => item.ToOpenXml());
}

sealed record SymbolNode(string Text) : LatexNode
{
    public override IEnumerable<OpenXmlElement> ToOpenXml()
    {
        if (string.IsNullOrEmpty(Text))
        {
            yield break;
        }

        yield return new M.Run(new M.Text(Text) { Space = SpaceProcessingModeValues.Preserve });
    }
}

sealed record FractionNode(LatexNode Numerator, LatexNode Denominator) : LatexNode
{
    public override IEnumerable<OpenXmlElement> ToOpenXml()
    {
        yield return new M.Fraction(
            new M.Numerator(Numerator.ToOpenXml()),
            new M.Denominator(Denominator.ToOpenXml()));
    }
}

sealed record ScriptNode(LatexNode Base, LatexNode? Sub, LatexNode? Sup) : LatexNode
{
    public override IEnumerable<OpenXmlElement> ToOpenXml()
    {
        var baseArg = new M.Base(Base.ToOpenXml());
        if (Sub is not null && Sup is not null)
        {
            yield return new M.SubSuperscript(
                baseArg,
                new M.SubArgument(Sub.ToOpenXml()),
                new M.SuperArgument(Sup.ToOpenXml()));
            yield break;
        }

        if (Sub is not null)
        {
            yield return new M.Subscript(
                baseArg,
                new M.SubArgument(Sub.ToOpenXml()));
            yield break;
        }

        if (Sup is not null)
        {
            yield return new M.Superscript(
                baseArg,
                new M.SuperArgument(Sup.ToOpenXml()));
        }
    }
}

sealed record DelimitedNode(string Open, string Close, LatexNode Content) : LatexNode
{
    public override IEnumerable<OpenXmlElement> ToOpenXml()
    {
        var delimiter = new M.Delimiter(
            new M.DelimiterProperties(
                new M.BeginChar { Val = Open },
                new M.EndChar { Val = Close }),
            new M.Base(Content.ToOpenXml()));
        yield return delimiter;
    }
}

sealed record InsertionSpec
{
    public string Id { get; init; } = "";
    public string AfterPrefix { get; init; } = "";
    public string BeforePrefix { get; init; } = "";
    public string Content { get; init; } = "";
    public string StyleSource { get; init; } = "after";
}

sealed record BlockInsertionPlan
{
    public IReadOnlyList<BlockSpec> Blocks { get; init; } = [];
}

sealed record BlockSpec
{
    public string Id { get; init; } = "";
    public string AfterPrefix { get; init; } = "";
    public string BeforePrefix { get; init; } = "";
    public string UniqueText { get; init; } = "";
    public string StyleSource { get; init; } = "after";
    public IReadOnlyList<BlockItemSpec> Items { get; init; } = [];
}

sealed record BlockItemSpec
{
    public string Kind { get; init; } = "paragraph";
    public string? Text { get; init; }
    public string? Latex { get; init; }
    public string? StyleId { get; init; }
    public string? TableStyleId { get; init; }
    public string? CellStyleId { get; init; }
    public IReadOnlyList<IReadOnlyList<string?>>? Rows { get; init; }
}

sealed record ParagraphEditPlan
{
    public IReadOnlyList<ParagraphEditSpec> Edits { get; init; } = [];
}

sealed record ReplaceTablePlan
{
    public IReadOnlyList<ReplaceTableSpec> Tables { get; init; } = [];
}

sealed record ReplaceTableSpec
{
    public string Id { get; init; } = "";
    public int? Ordinal { get; init; }
    public int? Block { get; init; }
    public int? BlockIndex { get; init; }
    public string FirstCellText { get; init; } = "";
    public string PreviousParagraphPrefix { get; init; } = "";
    public string NextParagraphPrefix { get; init; } = "";
    public string? TableStyleId { get; init; }
    public string? CellStyleId { get; init; }
    public IReadOnlyList<IReadOnlyList<string?>> Rows { get; init; } = [];
}

sealed record EquationRewritePlan
{
    public IReadOnlyList<EquationRewriteSpec> Edits { get; init; } = [];
}

sealed record EquationRewriteSpec
{
    public string Id { get; init; } = "";
    public string ParagraphPrefix { get; init; } = "";
    public string? ReplacementText { get; init; }
    public IReadOnlyList<BlockItemSpec> Items { get; init; } = [];
}

sealed record ParagraphEditSpec
{
    public string Id { get; init; } = "";
    public string ParagraphPrefix { get; init; } = "";
    public string? ReplacementText { get; init; }
    public string? StyleId { get; init; }
    public int? Occurrence { get; init; }
}

sealed record ParagraphEntry(Paragraph Paragraph, int Index, string Text);

sealed record TableReplacementSelectionResult(
    bool IsSuccess,
    string Message,
    Table? Table,
    int Ordinal,
    int BlockIndex,
    string FirstCellText,
    string PreviousParagraphText,
    string NextParagraphText,
    IReadOnlyList<int> ColumnWidths,
    string EffectiveCellStyleId);

sealed record RunningTextEligibilityResult(bool IsEligible, string Reason);

sealed record RunningTextStyleReportEntry(int Index, string Status, string Reason, string Text);

sealed record StructureAuditReport(
    string DocxPath,
    DateTime GeneratedAtUtc,
    int ParagraphCount,
    IReadOnlyList<TableAuditEntry> Tables,
    IReadOnlyList<FigureAuditEntry> Figures,
    IReadOnlyList<EquationAuditEntry> Equations);

sealed record LayoutAuditReport(
    string DocxPath,
    DateTime GeneratedAtUtc,
    int TableCount,
    int AcademicTableCount,
    int FigureCount,
    int AcademicFigureCount,
    IReadOnlyList<TableLayoutEntry> Tables,
    IReadOnlyList<FigureLayoutEntry> Figures);

sealed record TableLayoutEntry(
    int Ordinal,
    int BlockIndex,
    bool IsAcademicTable,
    int RowCount,
    int ColumnCount,
    int CellCount,
    string TableStyleId,
    string WidthType,
    string Width,
    string Justification,
    string LayoutType,
    IReadOnlyDictionary<string, string> Borders,
    bool HasVerticalBorders,
    IReadOnlyList<string> GridColumnWidths,
    bool HasMergedCells,
    IReadOnlyDictionary<string, int> CellParagraphStyles,
    IReadOnlyDictionary<string, int> DirectFontSizes,
    IReadOnlyDictionary<string, int> JustificationCounts,
    IReadOnlyDictionary<string, int> CellVerticalAlignments,
    ParagraphAuditEntry? PreviousParagraph,
    ParagraphAuditEntry? NextParagraph,
    ParagraphAuditEntry? Next2Paragraph,
    string SourceCandidate,
    string FirstCellText,
    IReadOnlyList<string> Issues);

sealed record FigureLayoutEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    bool IsAcademicFigure,
    string ParagraphText,
    string ParagraphStyleId,
    string DrawingName,
    string DrawingDescription,
    string DrawingKind,
    string Wrap,
    double WidthCm,
    double HeightCm,
    bool HasOutline,
    bool SimplePosition,
    uint RelativeHeight,
    string SourceCandidate,
    ParagraphAuditEntry? NextParagraph,
    ParagraphAuditEntry? Next2Paragraph,
    IReadOnlyList<string> Issues);

sealed record ParagraphAuditEntry(
    int Index,
    int BlockIndex,
    string Text,
    string StyleId,
    bool HasDrawing,
    bool HasMath);

sealed record MathAuditEntry(
    int ParagraphIndex,
    int MathIndex,
    string ParagraphText,
    string MathText,
    bool LooksLikeLinearLatex);

sealed record MathTextAuditResult(
    string DocxPath,
    DateTime GeneratedAtUtc,
    int ParagraphCount,
    int FlaggedParagraphCount,
    int ParagraphsWithOfficeMath,
    int ParagraphsWithTextMathCandidates,
    int OfficeMathCount,
    IReadOnlyList<MathTextAuditEntry> Paragraphs);

sealed record MathTextAuditEntry(
    int ParagraphIndex,
    string ParagraphText,
    string Container,
    int OfficeMathCount,
    IReadOnlyList<MathTextCandidate> TextMathCandidates);

sealed record MathTextCandidate(
    string Reason,
    string Value);

enum LatexTokenKind
{
    Command,
    Identifier,
    Number,
    Symbol,
    LBrace,
    RBrace,
    LParen,
    RParen,
    LBracket,
    RBracket,
    Subscript,
    Superscript,
    Comma
}

sealed record LatexToken(LatexTokenKind Kind, string Value);

sealed record TableAuditEntry(
    int Ordinal,
    int BlockIndex,
    int RowCount,
    int ColumnCount,
    string StyleId,
    string Width,
    bool HasBorders,
    ParagraphAuditEntry? PreviousParagraph,
    ParagraphAuditEntry? NextParagraph,
    ParagraphAuditEntry? Next2Paragraph,
    bool HasTitleAbove,
    bool HasSourceBelow,
    string SourceCandidate,
    string FirstCellText);

sealed record FigureAuditEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    string ParagraphText,
    string DrawingName,
    string DrawingDescription,
    ParagraphAuditEntry? PreviousParagraph,
    ParagraphAuditEntry? NextParagraph,
    ParagraphAuditEntry? Next2Paragraph,
    bool HasTitleAbove,
    bool HasSourceBelow,
    string SourceCandidate);

sealed record EquationSeqAuditReport(
    string DocxPath,
    DateTime GeneratedAtUtc,
    int MathParagraphCount,
    int SeqFieldCount,
    int CaptionSeqFieldCount,
    int DissertationEquationCount,
    int DistinctDissertationEquationNumberCount,
    IReadOnlyList<MathParagraphAuditEntry> MathParagraphs,
    IReadOnlyList<SeqFieldAuditEntry> SeqFields,
    IReadOnlyList<NumberedDissertationEquationEntry> DissertationEquations);

sealed record MathParagraphAuditEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    string StyleId,
    string ParagraphText,
    IReadOnlyList<SeqFieldSummary> SeqFields,
    bool IsDissertationEquation,
    int? DissertationEquationNumber);

sealed record SeqFieldSummary(
    string Instruction,
    string SequenceName,
    string ResultText,
    int? ResultNumber);

sealed record SeqFieldAuditEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    string StyleId,
    string ParagraphText,
    string Instruction,
    string SequenceName,
    string ResultText,
    int? ResultNumber,
    bool IsCaptionSeq);

sealed record NumberedDissertationEquationEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    string StyleId,
    int Number,
    string MathText,
    string ParagraphText,
    string SeqInstruction,
    ParagraphAuditEntry? PreviousParagraph,
    ParagraphAuditEntry? NextParagraph);

sealed record ExtractedSeqField(
    int BeginOrder,
    string Instruction,
    string SequenceName,
    string ResultText,
    int? ResultNumber);

sealed record CrossrefPlan(
    string InputPath,
    string PlanPath,
    IReadOnlyList<CrossrefCaptionSpec> Captions,
    IReadOnlyList<CrossrefParagraphReplacement> Replacements);

sealed record CrossrefCaptionSpec(
    string Id,
    string StyleId,
    string Label,
    string Sequence,
    int Number,
    string Bookmark,
    string AnchorText,
    int? Occurrence);

sealed record CrossrefParagraphReplacement(
    string Id,
    string ParagraphPrefix,
    IReadOnlyList<CrossrefTextReplacement> Replacements);

sealed record CrossrefTextReplacement(
    string OldText,
    IReadOnlyList<CrossrefRunPart> Parts);

sealed record CrossrefRunPart(string? Text, string? Ref, string? Result, string? FieldInstruction = null, string? Latex = null)
{
    public static CrossrefRunPart TextPart(string text) => new(text, null, null, null, null);
}

sealed record BookmarkPlan(
    string InputPath,
    string PlanPath,
    IReadOnlyList<BookmarkPlanEntry> Bookmarks);

sealed record BookmarkPlanEntry(
    string Id,
    string ParagraphPrefix,
    string Bookmark,
    int? Occurrence);

sealed record EquationAuditEntry(
    int Ordinal,
    int BlockIndex,
    int ParagraphIndex,
    string Text,
    string MathText,
    string StyleId,
    bool HasTrailingNumber,
    ParagraphAuditEntry? PreviousParagraph,
    ParagraphAuditEntry? NextParagraph);

sealed record FieldInfoEntry(string Instruction, string ResultText);

sealed record FigureInsertionPlan
{
    public IReadOnlyList<FigureSpec> Figures { get; init; } = [];
}

sealed record FigureSpec
{
    public string Id { get; init; } = "";
    public string AfterPrefix { get; init; } = "";
    public string BeforePrefix { get; init; } = "";
    public string ImagePath { get; init; } = "";
    public string CaptionText { get; init; } = "";
    public string SourceText { get; init; } = "";
    public double WidthCm { get; init; }
}

sealed record FigureReplacementPlan
{
    public IReadOnlyList<FigureReplacement> Figures { get; init; } = [];
}

sealed record FigureReplacement(
    string DrawingName,
    string ImagePath,
    string CaptionText,
    string SourceText);

sealed record FormulaConversionPlan
{
    public IReadOnlyList<FormulaPlanItem> Formulas { get; init; } = [];
}

sealed record FormulaPlanItem
{
    public string Text { get; init; } = "";
    public string Latex { get; init; } = "";
    public string MathMl { get; init; } = "";
    public bool Display { get; init; }
    public int? Occurrence { get; init; }
}

sealed record FormulaSpec(string RequiredText, string Latex, string MathMl, bool Display, int? Occurrence);

sealed record FormulaReplacement(int Start, int Length, FormulaSpec Spec);

sealed record ReferenceTitleSpec(string Id, string ReferencePrefix, string TitleToEmphasize);

sealed record TabularOutput(IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows);

sealed record CommentListEntry(string Id, string Author, string Date, string? ParentCommentId, string Text, string AnchorText);

sealed record CommentInsertionPlan
{
    public IReadOnlyList<CommentSpec> Comments { get; init; } = [];
}

sealed record CommentReanchorPlan
{
    public IReadOnlyList<CommentAnchorSpec> Anchors { get; init; } = [];
}

sealed record CommentAnswerPlan
{
    public IReadOnlyList<CommentAnswerSpec> Answers { get; init; } = [];
}

sealed record CommentReplyPlan
{
    public IReadOnlyList<CommentReplySpec> Replies { get; init; } = [];
}

sealed record CommentSpec
{
    public string Id { get; init; } = "";
    public string AnchorPrefix { get; init; } = "";
    public string? AnchorContains { get; init; }
    public string CommentText { get; init; } = "";
}

sealed record CommentAnchorSpec
{
    public string Id { get; init; } = "";
    public string CommentId { get; init; } = "";
    public string AnchorPrefix { get; init; } = "";
    public string? AnchorContains { get; init; }
    public int? Occurrence { get; init; }
}

sealed record CommentAnswerSpec
{
    public string Id { get; init; } = "";
    public string CommentId { get; init; } = "";
    public string ResponseText { get; init; } = "";
}

sealed record CommentReplySpec
{
    public string Id { get; init; } = "";
    public string ParentCommentId { get; init; } = "";
    public string ReplyText { get; init; } = "";
    public string? AnchorPrefix { get; init; }
    public string? AnchorContains { get; init; }
    public int? Occurrence { get; init; }
}

sealed class RevisionMetadata(string author, DateTime dateUtc)
{
    private int _counter = 1;

    public string Author { get; } = author;
    public DateTime DateUtc { get; } = dateUtc;

    public StringValue NextRevisionId() => (_counter++).ToString(CultureInfo.InvariantCulture);
}


