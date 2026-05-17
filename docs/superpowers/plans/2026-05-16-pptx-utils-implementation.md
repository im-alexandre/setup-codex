# pptx-utils Implementation Plan

> **For agentic workers:** Execute this plan with `$implement-tdd docs/superpowers/plans/2026-05-16-pptx-utils-implementation.md` or run `$implement-tdd` from the project root to auto-pick the newest plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single `$pptx-utils` skill that makes reliable PPTX template editing the default path through a published .NET/Open XML CLI while preserving useful QA and rendering assets from the current `pptx` skill.

**Architecture:** Create `skills/pptx-utils` as the only user-facing PPTX skill. The default execution path is `bin/pptx-utils/pptx-utils(.exe)`, backed by `src/PptxOpenXmlTools` and tested by `src/PptxOpenXmlTools.Tests`; existing Python scripts from the installed `pptx` skill are copied in as fallback and QA helpers. Mutations are expressed as JSON plans rather than generated JavaScript or manual slide XML edits.

**Tech Stack:** .NET 8, `DocumentFormat.OpenXml`, `ShapeCrawler`, `System.Text.Json`, xUnit, PowerShell/Bash installers, existing Python QA scripts, LibreOffice/Poppler for visual preview.

---

## File Structure

- Create: `skills/pptx-utils/SKILL.md`
  - Responsibility: short operational skill instructions; command-first, published-binary-first.
- Create: `skills/pptx-utils/BACKLOG.md`
  - Responsibility: missing PPTX capabilities discovered during use.
- Create: `skills/pptx-utils/references/plan-contracts.md`
  - Responsibility: human-readable JSON contracts.
- Create: `skills/pptx-utils/references/plan-contracts.json`
  - Responsibility: machine-readable command contracts.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/PptxOpenXmlTools.csproj`
  - Responsibility: CLI project.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`
  - Responsibility: thin command router.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Commands/*.cs`
  - Responsibility: one focused command per file.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Models/*.cs`
  - Responsibility: DTOs for JSON output and plan input.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/PptxOpenXmlTools.Tests.csproj`
  - Responsibility: xUnit tests.
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/*Tests.cs`
  - Responsibility: command-level red/green tests.
- Create: `skills/pptx-utils/scripts/install-pptx-utils.ps1`
  - Responsibility: restore, test, publish and optional skill validation on Windows.
- Create: `skills/pptx-utils/scripts/install-pptx-utils.sh`
  - Responsibility: restore, test, publish and optional skill validation on Linux/WSL.
- Copy: `C:\Users\imale\.agents\skills\pptx\scripts\thumbnail.py` to `skills/pptx-utils/scripts/thumbnail.py`
  - Responsibility: template thumbnail overview.
- Copy: `C:\Users\imale\.agents\skills\pptx\scripts\office\*.py` to `skills/pptx-utils/scripts/office/`
  - Responsibility: fallback unpack/pack/validate and render helper.

## Task 1: Scaffold Skill Directory And Metadata

**Files:**
- Create: `skills/pptx-utils/SKILL.md`
- Create: `skills/pptx-utils/BACKLOG.md`
- Create: `skills/pptx-utils/references/plan-contracts.md`
- Create: `skills/pptx-utils/references/plan-contracts.json`

- [ ] **Step 1: Write the initial skill documentation**

Create `skills/pptx-utils/SKILL.md` with this content:

```markdown
---
name: pptx-utils
description: Editar, inspecionar, validar e renderizar apresentacoes PPTX com foco em templates existentes, usando o binario .NET/Open XML publicado como caminho padrao.
---

# PPTX Utils

Use esta skill para qualquer tarefa envolvendo `.pptx`.

## Regra De Execucao

1. Para edicao de apresentacoes existentes ou templates, use o binario publicado:

   `C:\Users\imale\.codex\skills\pptx-utils\bin\pptx-utils\pptx-utils.exe <comando> <pptx> [opcoes]`

1. Quando o shim estiver no `PATH`, use:

   `pptx-utils <comando> <pptx> [opcoes]`

1. Nao edite XML de slides manualmente salvo fallback, depuracao ou lacuna registrada em `BACKLOG.md`.
1. Nao gere JavaScript/PptxGenJS para editar templates existentes.
1. Use planos JSON para mutacoes.
1. Depois de mutar um PPTX, rode `validate`; para entrega visual, rode tambem `render-preview`.

## Comandos MVP

- `inspect`: lista estrutura, slides, shapes e textos.
- `text-map`: gera mapa de textos editaveis.
- `replace-text`: substitui textos por plano JSON.
- `duplicate-slide`: duplica slide existente.
- `delete-slide`: remove slides.
- `reorder-slides`: reordena slides.
- `replace-image`: substitui imagens por plano JSON.
- `validate`: valida integridade Open XML.
- `render-preview`: renderiza previews usando LibreOffice/Poppler.

## Recursos Herdados

- `scripts/thumbnail.py`: analise visual rapida de templates.
- `scripts/office/soffice.py`: conversao headless para PDF.
- `scripts/office/unpack.py`, `pack.py`, `validate.py`: fallback e depuracao.
- `references/plan-contracts.md`: contratos de planos.

## Fallback Para Criacao Do Zero

PptxGenJS pode ser usado apenas quando nao houver template e a tarefa for criar um deck novo do zero. Esse caminho e secundario e nao deve substituir o fluxo de edicao confiavel.
```

- [ ] **Step 2: Write the backlog seed**

Create `skills/pptx-utils/BACKLOG.md` with:

```markdown
# Backlog pptx-utils

- Criar comando `create-from-template` depois que `inspect`, `text-map` e `replace-text` estiverem estaveis.
- Avaliar suporte a charts editaveis via ShapeCrawler ou Open XML direto.
- Avaliar substituicao futura de `thumbnail.py` por comando integrado quando houver beneficio real.
- Documentar fallback PptxGenJS sem carregar tutorial JS por padrao.
```

- [ ] **Step 3: Write plan contracts**

Create `skills/pptx-utils/references/plan-contracts.md` with:

```markdown
# Contratos De Planos pptx-utils

## replace-text

```json
{
  "replacements": [
    {
      "slideIndex": 1,
      "shapeIndex": 2,
      "find": "Texto antigo",
      "replace": "Texto novo",
      "mode": "exact"
    }
  ]
}
```

`slideIndex` e 1-based. `shapeIndex` e 1-based dentro do slide. `mode` aceita `exact`.

## replace-image

```json
{
  "replacements": [
    {
      "slideIndex": 1,
      "shapeIndex": 4,
      "imagePath": "imagens/nova.png"
    }
  ]
}
```
```

Create `skills/pptx-utils/references/plan-contracts.json` with:

```json
{
  "replace-text": {
    "required": ["replacements"],
    "replacementRequired": ["slideIndex", "shapeIndex", "find", "replace"],
    "modes": ["exact"]
  },
  "replace-image": {
    "required": ["replacements"],
    "replacementRequired": ["slideIndex", "shapeIndex", "imagePath"]
  }
}
```

- [ ] **Step 4: Validate metadata**

Run:

```powershell
python C:\Users\imale\.codex\skills\.system\skill-creator\scripts\quick_validate.py C:\Users\imale\.codex\skills\pptx-utils
```

Expected: validation fails only if the skill validator requires source files not yet present; if it reports YAML/frontmatter errors, fix `SKILL.md` before continuing.

- [ ] **Step 5: Commit**

```powershell
git add skills/pptx-utils/SKILL.md skills/pptx-utils/BACKLOG.md skills/pptx-utils/references/plan-contracts.md skills/pptx-utils/references/plan-contracts.json
git commit -m "feat: scaffold pptx-utils skill"
```

## Task 2: Scaffold .NET Projects With A Failing Help Test

**Files:**
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/PptxOpenXmlTools.csproj`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/PptxOpenXmlTools.Tests.csproj`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/HelpTests.cs`

- [ ] **Step 1: Create failing help test**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/HelpTests.cs`:

```csharp
using System.Diagnostics;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class HelpTests
{
    [Fact]
    public async Task Help_PrintsAvailableCommands()
    {
        var result = await Cli.RunAsync("--help");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("pptx-utils", result.Stdout);
        Assert.Contains("inspect", result.Stdout);
        Assert.Contains("text-map", result.Stdout);
        Assert.Contains("replace-text", result.Stdout);
        Assert.Contains("validate", result.Stdout);
    }
}

internal static class Cli
{
    public static async Task<CliResult> RunAsync(params string[] args)
    {
        var project = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "PptxOpenXmlTools", "PptxOpenXmlTools.csproj"));
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        psi.ArgumentList.Add("run");
        psi.ArgumentList.Add("--project");
        psi.ArgumentList.Add(project);
        psi.ArgumentList.Add("--");
        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        using var process = Process.Start(psi)!;
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return new CliResult(process.ExitCode, stdout, stderr);
    }
}

internal sealed record CliResult(int ExitCode, string Stdout, string Stderr);
```

- [ ] **Step 2: Create test project**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/PptxOpenXmlTools.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>
</Project>
```

- [ ] **Step 3: Run test to verify it fails**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

Expected: FAIL because `PptxOpenXmlTools.csproj` does not exist.

- [ ] **Step 4: Create CLI project**

Create `skills/pptx-utils/src/PptxOpenXmlTools/PptxOpenXmlTools.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
    <PackageReference Include="ShapeCrawler" Version="0.60.0" />
  </ItemGroup>
</Project>
```

Create `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`:

```csharp
using System.Text;

Console.OutputEncoding = Encoding.UTF8;

if (args.Length == 0 || args[0] is "--help" or "-h" or "help")
{
    Console.WriteLine("""
    pptx-utils

    Commands:
      inspect
      text-map
      replace-text
      duplicate-slide
      delete-slide
      reorder-slides
      replace-image
      validate
      render-preview
    """);
    return 0;
}

Console.Error.WriteLine($"Unknown command: {args[0]}");
return 2;
```

- [ ] **Step 5: Run test to verify it passes**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add skills/pptx-utils/src
git commit -m "test: add pptx-utils cli help baseline"
```

## Task 3: Add Test Fixture Builder And Validate Command

**Files:**
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Commands/ValidateCommand.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/PptxFixture.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/ValidateTests.cs`
- Modify: `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`

- [ ] **Step 1: Write fixture builder**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/PptxFixture.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Tests;

internal static class PptxFixture
{
    public static string CreateSingleSlideDeck(string text)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-utils-{Guid.NewGuid():N}.pptx");

        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()),
            CreateTextShape(2U, "TextBox 1", text))));

        var relId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.Append(new SlideIdList(new SlideId { Id = 256U, RelationshipId = relId }));
        presentationPart.Presentation.Save();
        return path;
    }

    private static Shape CreateTextShape(uint id, string name, string text)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text)))));
    }
}
```

- [ ] **Step 2: Write failing validate test**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/ValidateTests.cs`:

```csharp
using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ValidateTests
{
    [Fact]
    public async Task Validate_ReturnsJsonWithZeroOpenXmlErrorsForFixture()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("validate", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.True(doc.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal(0, doc.RootElement.GetProperty("openXmlValidationErrors").GetInt32());
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj --filter Validate
```

Expected: FAIL with unknown command `validate`.

- [ ] **Step 4: Implement validate command**

Create `skills/pptx-utils/src/PptxOpenXmlTools/Commands/ValidateCommand.cs`:

```csharp
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace PptxOpenXmlTools.Commands;

internal static class ValidateCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils validate <pptx> [--format json]");
            return 2;
        }

        var pptx = args[1];
        if (!File.Exists(pptx))
        {
            Console.Error.WriteLine($"File not found: {pptx}");
            return 2;
        }

        using var document = PresentationDocument.Open(pptx, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document).ToList();
        var hasPresentation = document.PresentationPart?.Presentation?.SlideIdList is not null;

        var result = new
        {
            isValid = errors.Count == 0 && hasPresentation,
            openXmlValidationErrors = errors.Count,
            hasPresentation
        };

        Console.WriteLine(JsonSerializer.Serialize(result));
        return result.isValid ? 0 : 1;
    }
}
```

Modify `Program.cs` command dispatch:

```csharp
using System.Text;
using PptxOpenXmlTools.Commands;

Console.OutputEncoding = Encoding.UTF8;

if (args.Length == 0 || args[0] is "--help" or "-h" or "help")
{
    Console.WriteLine("""
    pptx-utils

    Commands:
      inspect
      text-map
      replace-text
      duplicate-slide
      delete-slide
      reorder-slides
      replace-image
      validate
      render-preview
    """);
    return 0;
}

return args[0] switch
{
    "validate" => ValidateCommand.Run(args),
    _ => Unknown(args[0])
};

static int Unknown(string command)
{
    Console.Error.WriteLine($"Unknown command: {command}");
    return 2;
}
```

- [ ] **Step 5: Run tests**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add skills/pptx-utils/src
git commit -m "feat: add pptx validate command"
```

## Task 4: Add Inspect And Text Map

**Files:**
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Commands/InspectCommand.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Commands/TextMapCommand.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/InspectTests.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/TextMapTests.cs`
- Modify: `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`

- [ ] **Step 1: Write failing inspect test**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/InspectTests.cs`:

```csharp
using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class InspectTests
{
    [Fact]
    public async Task Inspect_ReturnsSlideAndShapeCounts()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("inspect", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, doc.RootElement.GetProperty("slideCount").GetInt32());
        Assert.Equal(1, doc.RootElement.GetProperty("slides")[0].GetProperty("textShapeCount").GetInt32());
    }
}
```

- [ ] **Step 2: Write failing text-map test**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/TextMapTests.cs`:

```csharp
using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class TextMapTests
{
    [Fact]
    public async Task TextMap_ReturnsEditableTextWithSlideAndShapeIndexes()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("text-map", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        var item = doc.RootElement.GetProperty("items")[0];
        Assert.Equal(1, item.GetProperty("slideIndex").GetInt32());
        Assert.Equal(1, item.GetProperty("shapeIndex").GetInt32());
        Assert.Equal("Hello", item.GetProperty("text").GetString());
    }
}
```

- [ ] **Step 3: Run tests to verify failure**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj --filter "Inspect|TextMap"
```

Expected: FAIL with unknown commands `inspect` and `text-map`.

- [ ] **Step 4: Implement inspect and text-map**

Create `skills/pptx-utils/src/PptxOpenXmlTools/Commands/InspectCommand.cs`:

```csharp
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Commands;

internal static class InspectCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils inspect <pptx> [--format json]");
            return 2;
        }

        using var document = PresentationDocument.Open(args[1], false);
        var slides = document.PresentationPart!.SlideParts.ToList();
        var result = new
        {
            slideCount = slides.Count,
            slides = slides.Select((slide, index) => new
            {
                slideIndex = index + 1,
                shapeCount = slide.Slide.Descendants<P.Shape>().Count(),
                textShapeCount = slide.Slide.Descendants<P.Shape>().Count(s => s.Descendants<A.Text>().Any())
            })
        };
        Console.WriteLine(JsonSerializer.Serialize(result));
        return 0;
    }
}
```

Create `skills/pptx-utils/src/PptxOpenXmlTools/Commands/TextMapCommand.cs`:

```csharp
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Commands;

internal static class TextMapCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils text-map <pptx> [--format json]");
            return 2;
        }

        using var document = PresentationDocument.Open(args[1], false);
        var items = new List<object>();
        var slideIndex = 0;

        foreach (var slide in document.PresentationPart!.SlideParts)
        {
            slideIndex++;
            var shapeIndex = 0;
            foreach (var shape in slide.Slide.Descendants<P.Shape>())
            {
                var text = string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                shapeIndex++;
                var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";
                items.Add(new
                {
                    slideIndex,
                    shapeIndex,
                    shapeName = name,
                    text,
                    paragraphCount = shape.Descendants<A.Paragraph>().Count(),
                    runCount = shape.Descendants<A.Run>().Count()
                });
            }
        }

        Console.WriteLine(JsonSerializer.Serialize(new { items }));
        return 0;
    }
}
```

Modify dispatch in `Program.cs`:

```csharp
return args[0] switch
{
    "inspect" => InspectCommand.Run(args),
    "text-map" => TextMapCommand.Run(args),
    "validate" => ValidateCommand.Run(args),
    _ => Unknown(args[0])
};
```

- [ ] **Step 5: Run tests**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add skills/pptx-utils/src
git commit -m "feat: add pptx inspect and text-map"
```

## Task 5: Add Replace Text By JSON Plan

**Files:**
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Commands/ReplaceTextCommand.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools/Models/ReplaceTextPlan.cs`
- Create: `skills/pptx-utils/src/PptxOpenXmlTools.Tests/ReplaceTextTests.cs`
- Modify: `skills/pptx-utils/src/PptxOpenXmlTools/Program.cs`

- [ ] **Step 1: Write failing replace-text test**

Create `skills/pptx-utils/src/PptxOpenXmlTools.Tests/ReplaceTextTests.cs`:

```csharp
using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ReplaceTextTests
{
    [Fact]
    public async Task ReplaceText_ReplacesExactTextAndOutputValidates()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Old text");
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(plan, """
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Old text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));

        var textMap = await Cli.RunAsync("text-map", output, "--format", "json");
        using var doc = JsonDocument.Parse(textMap.Stdout);
        Assert.Equal("New text", doc.RootElement.GetProperty("items")[0].GetProperty("text").GetString());

        var validate = await Cli.RunAsync("validate", output, "--format", "json");
        Assert.Equal(0, validate.ExitCode);
    }
}
```

- [ ] **Step 2: Run test to verify failure**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj --filter ReplaceText
```

Expected: FAIL with unknown command `replace-text`.

- [ ] **Step 3: Implement plan model**

Create `skills/pptx-utils/src/PptxOpenXmlTools/Models/ReplaceTextPlan.cs`:

```csharp
namespace PptxOpenXmlTools.Models;

internal sealed class ReplaceTextPlan
{
    public List<ReplaceTextItem> Replacements { get; set; } = [];
}

internal sealed class ReplaceTextItem
{
    public int SlideIndex { get; set; }
    public int ShapeIndex { get; set; }
    public string Find { get; set; } = "";
    public string Replace { get; set; } = "";
    public string Mode { get; set; } = "exact";
}
```

- [ ] **Step 4: Implement replace-text command**

Create `skills/pptx-utils/src/PptxOpenXmlTools/Commands/ReplaceTextCommand.cs`:

```csharp
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using PptxOpenXmlTools.Models;

namespace PptxOpenXmlTools.Commands;

internal static class ReplaceTextCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 6)
        {
            Console.Error.WriteLine("Usage: pptx-utils replace-text <pptx> --plan <json> --output <pptx>");
            return 2;
        }

        var input = args[1];
        var planPath = ValueAfter(args, "--plan");
        var output = ValueAfter(args, "--output");
        if (planPath is null || output is null)
        {
            Console.Error.WriteLine("Missing --plan or --output.");
            return 2;
        }

        File.Copy(input, output, overwrite: true);
        var plan = JsonSerializer.Deserialize<ReplaceTextPlan>(File.ReadAllText(planPath), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;

        using var document = PresentationDocument.Open(output, true);
        var slides = document.PresentationPart!.SlideParts.ToList();

        foreach (var replacement in plan.Replacements)
        {
            if (replacement.Mode != "exact")
            {
                Console.Error.WriteLine($"Unsupported mode: {replacement.Mode}");
                return 2;
            }

            var slide = slides[replacement.SlideIndex - 1];
            var shapes = slide.Slide.Descendants<P.Shape>().Where(s => s.Descendants<A.Text>().Any()).ToList();
            var shape = shapes[replacement.ShapeIndex - 1];
            var textNodes = shape.Descendants<A.Text>().ToList();
            var current = string.Concat(textNodes.Select(t => t.Text));
            if (current != replacement.Find)
            {
                Console.Error.WriteLine($"Text mismatch on slide {replacement.SlideIndex}, shape {replacement.ShapeIndex}.");
                return 1;
            }

            textNodes[0].Text = replacement.Replace;
            foreach (var extra in textNodes.Skip(1))
            {
                extra.Text = "";
            }

            slide.Slide.Save();
        }

        return 0;
    }

    private static string? ValueAfter(string[] args, string option)
    {
        var index = Array.IndexOf(args, option);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }
}
```

Modify dispatch:

```csharp
"replace-text" => ReplaceTextCommand.Run(args),
```

- [ ] **Step 5: Run tests**

Run:

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add skills/pptx-utils/src
git commit -m "feat: add pptx replace-text plan command"
```

## Task 6: Incorporate Existing PPTX Python Helpers

**Files:**
- Copy: `C:\Users\imale\.agents\skills\pptx\scripts\thumbnail.py` to `skills/pptx-utils/scripts/thumbnail.py`
- Copy: `C:\Users\imale\.agents\skills\pptx\scripts\office/` to `skills/pptx-utils/scripts/office/`
- Create: `skills/pptx-utils/scripts/README.md`

- [ ] **Step 1: Copy helper scripts**

Run:

```powershell
New-Item -ItemType Directory -Force -Path skills\pptx-utils\scripts\office | Out-Null
Copy-Item C:\Users\imale\.agents\skills\pptx\scripts\thumbnail.py skills\pptx-utils\scripts\thumbnail.py -Force
Copy-Item C:\Users\imale\.agents\skills\pptx\scripts\office\*.py skills\pptx-utils\scripts\office\ -Force
```

Expected: files copied.

- [ ] **Step 2: Document helper ownership**

Create `skills/pptx-utils/scripts/README.md`:

```markdown
# Scripts Herdados

Estes scripts foram incorporados da skill `pptx` antiga para preservar recursos uteis enquanto o nucleo mutador passa para .NET.

- `thumbnail.py`: visao geral visual de templates.
- `office/soffice.py`: conversao headless para PDF.
- `office/unpack.py`, `pack.py`, `validate.py`: fallback e depuracao.

O fluxo padrao de edicao deve continuar usando `pptx-utils`. Scripts de unpack/pack nao devem ser usados como caminho normal para mutacoes quando houver comando .NET equivalente.
```

- [ ] **Step 3: Smoke test thumbnail script import**

Run:

```powershell
python skills\pptx-utils\scripts\thumbnail.py --help
```

Expected: help text or usage text exits without Python syntax error.

- [ ] **Step 4: Commit**

```powershell
git add skills/pptx-utils/scripts
git commit -m "feat: incorporate pptx helper scripts"
```

## Task 7: Add Installers And Publish Binary

**Files:**
- Create: `skills/pptx-utils/scripts/install-pptx-utils.ps1`
- Create: `skills/pptx-utils/scripts/install-pptx-utils.sh`
- Modify: `skills/pptx-utils/SKILL.md`

- [ ] **Step 1: Create Windows installer**

Create `skills/pptx-utils/scripts/install-pptx-utils.ps1`:

```powershell
$ErrorActionPreference = "Stop"

$SkillRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$Project = Join-Path $SkillRoot "src\PptxOpenXmlTools\PptxOpenXmlTools.csproj"
$Tests = Join-Path $SkillRoot "src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj"
$PublishDir = Join-Path $SkillRoot "bin\pptx-utils"

dotnet restore $Tests
dotnet test $Tests --configuration Release
dotnet publish $Project --configuration Release --output $PublishDir

Write-Host "pptx-utils publicado em $PublishDir"
```

- [ ] **Step 2: Create Bash installer**

Create `skills/pptx-utils/scripts/install-pptx-utils.sh`:

```bash
#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
skill_root="$(cd -- "${script_dir}/.." && pwd)"
project="${skill_root}/src/PptxOpenXmlTools/PptxOpenXmlTools.csproj"
tests="${skill_root}/src/PptxOpenXmlTools.Tests/PptxOpenXmlTools.Tests.csproj"
publish_dir="${skill_root}/bin/pptx-utils"

dotnet restore "$tests"
dotnet test "$tests" --configuration Release
dotnet publish "$project" --configuration Release --output "$publish_dir"

printf 'pptx-utils publicado em %s\n' "$publish_dir"
```

- [ ] **Step 3: Run installer**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File skills\pptx-utils\scripts\install-pptx-utils.ps1
```

Expected: tests pass and `skills\pptx-utils\bin\pptx-utils\pptx-utils.exe` exists.

- [ ] **Step 4: Run published binary smoke test**

Run:

```powershell
skills\pptx-utils\bin\pptx-utils\pptx-utils.exe --help
```

Expected: output contains `pptx-utils`, `inspect`, `replace-text`, `validate`.

- [ ] **Step 5: Commit**

```powershell
git add skills/pptx-utils/scripts skills/pptx-utils/bin skills/pptx-utils/SKILL.md
git commit -m "build: publish pptx-utils binary"
```

## Task 8: Decommission Old PPTX Skill Entry Point

**Files:**
- Inspect: `C:\Users\imale\.agents\skills\pptx`
- Modify or remove according to repository policy: old `pptx` skill entry point.

- [ ] **Step 1: Confirm old skill location and ownership**

Run:

```powershell
Get-Item C:\Users\imale\.agents\skills\pptx
git -C C:\Users\imale\.codex status --short
```

Expected: old skill is outside or inside tracked scope is clear before deleting or replacing.

- [ ] **Step 2: Replace old entry point with migration notice if deletion is unsafe**

If the old skill is outside this repo or deletion would affect unrelated tooling, replace only its `SKILL.md` with a short redirect:

```markdown
---
name: pptx
description: Deprecated local entry point. Use the `pptx-utils` skill for PPTX inspection, editing, validation and rendering.
---

# Deprecated

Use `$pptx-utils` for PPTX work. This entry point exists only to prevent silent use of the old JS/XML-heavy workflow.
```

If the old skill is tracked in this repo and safe to remove, delete it after confirming `pptx-utils` validates.

- [ ] **Step 3: Validate only one operational entry point remains**

Run:

```powershell
rg -n "name: pptx$|name: pptx-utils|PptxGenJS|pptxgenjs" C:\Users\imale\.codex\skills C:\Users\imale\.agents\skills
```

Expected: `pptx-utils` is the operational skill; `pptx` is absent or explicitly deprecated.

- [ ] **Step 4: Commit**

If `C:\Users\imale\.agents\skills\pptx` is outside the current Git repository, do not try to stage it from this repo. Commit only the `pptx-utils` changes here and record the external deprecation action in the final execution report.

```powershell
git add skills/pptx-utils
git commit -m "chore: route pptx work through pptx-utils"
```

## Task 9: Final Validation

**Files:**
- Read: all files under `skills/pptx-utils`

- [ ] **Step 1: Run .NET tests**

```powershell
dotnet test skills\pptx-utils\src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj --configuration Release
```

Expected: PASS.

- [ ] **Step 2: Run installer**

```powershell
powershell -ExecutionPolicy Bypass -File skills\pptx-utils\scripts\install-pptx-utils.ps1
```

Expected: PASS and binary published.

- [ ] **Step 3: Run skill validation**

```powershell
python C:\Users\imale\.codex\skills\.system\skill-creator\scripts\quick_validate.py C:\Users\imale\.codex\skills\pptx-utils
```

Expected: PASS.

- [ ] **Step 4: Run an end-to-end smoke manually**

Create a fixture through tests or use a known local sample, then run:

```powershell
skills\pptx-utils\bin\pptx-utils\pptx-utils.exe inspect sample.pptx --format json
skills\pptx-utils\bin\pptx-utils\pptx-utils.exe text-map sample.pptx --format json
skills\pptx-utils\bin\pptx-utils\pptx-utils.exe validate sample.pptx --format json
```

Expected: all commands exit `0` and print valid JSON.

- [ ] **Step 5: Final commit**

```powershell
git status --short
git add skills/pptx-utils docs/superpowers/specs/2026-05-16-pptx-utils-design.md docs/superpowers/plans/2026-05-16-pptx-utils-implementation.md
git commit -m "docs: plan pptx-utils migration"
```

## Parallelization Notes

- Tasks 1 and 2 are sequential.
- After Task 3, Task 4 can run before Task 5, but Task 5 depends on `text-map` behavior.
- Task 6 can run in parallel with Tasks 3-5 because it only copies helper scripts.
- Task 7 depends on tests and project structure from Tasks 2-5.
- Task 8 must run after `pptx-utils` validates.
- Task 9 is final integration only.

## Self Review

- Spec coverage: skill naming, single entry point, .NET/OpenXML default, helper reuse, JSON plans, validation and QA are covered.
- Placeholder scan: no incomplete implementation markers or open-ended placeholders remain.
- Type consistency: command names match the spec and `SKILL.md`; plan property names match tests and model classes.
