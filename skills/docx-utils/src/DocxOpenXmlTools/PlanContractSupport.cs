using System.Text;
using System.Text.Json;

internal static class PlanContractSupport
{
    private const string PlanContractsFileName = "plan-contracts.json";

    public static int PrintPlanContracts(string? targetCommand, string format)
    {
        if (!IsSupportedFormat(format))
        {
            Console.Error.WriteLine("Unsupported format. Use `markdown` or `json`.");
            return 4;
        }

        var contractsPath = FindPlanContractsPath();
        if (contractsPath is null)
        {
            Console.Error.WriteLine("Unable to locate references/plan-contracts.json.");
            return 5;
        }

        var rawContracts = File.ReadAllText(contractsPath, Encoding.UTF8);
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(rawContracts);
        }
        catch (JsonException ex)
        {
            Console.Error.WriteLine($"Invalid plan contract artifact: {ex.Message}");
            return 6;
        }

        using (document)
        {
            var commands = GetCommands(document.RootElement);

            if (!string.IsNullOrWhiteSpace(targetCommand))
            {
                var command = FindCommand(commands, targetCommand);
                if (command.ValueKind == JsonValueKind.Undefined)
                {
                    Console.Error.WriteLine($"Unknown plan contract: {targetCommand}");
                    return 3;
                }

            return PrintSingleContract(command, format);
        }

            if (string.Equals(format, "json", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine(rawContracts);
                return 0;
            }

            return PrintAllContracts(commands);
        }
    }

    public static PlanValidationResult ValidateInsertBlocksPlan(string planJson)
    {
        try
        {
            using var document = JsonDocument.Parse(planJson);
            var errors = new List<string>();
            ValidateInsertBlocksPlan(document.RootElement, errors);
            return errors.Count == 0
                ? new PlanValidationResult(true, [])
                : new PlanValidationResult(false, errors);
        }
        catch (JsonException ex)
        {
            return new PlanValidationResult(false, [$"Invalid JSON in insert-blocks plan: {ex.Message}"]);
        }
    }

    public static PlanValidationResult ValidateReplaceTablePlan(string planJson)
    {
        try
        {
            using var document = JsonDocument.Parse(planJson);
            var errors = new List<string>();
            ValidateReplaceTablePlan(document.RootElement, errors);
            return errors.Count == 0
                ? new PlanValidationResult(true, [])
                : new PlanValidationResult(false, errors);
        }
        catch (JsonException ex)
        {
            return new PlanValidationResult(false, [$"Invalid JSON in replace-table plan: {ex.Message}"]);
        }
    }

    public static PlanValidationResult ValidateCreateDocxPlan(string planJson)
    {
        try
        {
            using var document = JsonDocument.Parse(planJson);
            var errors = new List<string>();
            ValidateCreateDocxPlan(document.RootElement, errors);
            return errors.Count == 0
                ? new PlanValidationResult(true, [])
                : new PlanValidationResult(false, errors);
        }
        catch (JsonException ex)
        {
            return new PlanValidationResult(false, [$"Invalid JSON in create-docx plan: {ex.Message}"]);
        }
    }

    private static int PrintAllContracts(JsonElement commands)
    {
        var commandList = commands.ValueKind == JsonValueKind.Array
            ? commands.EnumerateArray().ToList()
            : new List<JsonElement>();

        var builder = new StringBuilder();
        builder.AppendLine("# Plan Contracts");
        foreach (var command in commandList)
        {
            AppendMarkdownContract(builder, command);
        }

        Console.Write(builder.ToString());
        return 0;
    }

    private static int PrintSingleContract(JsonElement command, string format)
    {
        if (!IsSupportedFormat(format))
        {
            Console.Error.WriteLine("Unsupported format. Use `markdown` or `json`.");
            return 4;
        }

        if (string.Equals(format, "json", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine(JsonSerializer.Serialize(command, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }

        var builder = new StringBuilder();
        builder.AppendLine("# Plan Contracts");
        AppendMarkdownContract(builder, command);
        Console.Write(builder.ToString());
        return 0;
    }

    private static void AppendMarkdownContract(StringBuilder builder, JsonElement command)
    {
        var name = command.TryGetProperty("name", out var nameElement) && nameElement.ValueKind == JsonValueKind.String
            ? nameElement.GetString() ?? ""
            : "";
        var summary = command.TryGetProperty("summary", out var summaryElement) && summaryElement.ValueKind == JsonValueKind.String
            ? summaryElement.GetString() ?? ""
            : "";

        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        builder.AppendLine();
        builder.AppendLine($"## `{name}`");
        if (!string.IsNullOrWhiteSpace(summary))
        {
            builder.AppendLine(summary);
        }

        if (command.TryGetProperty("planShape", out var planShape) && planShape.ValueKind == JsonValueKind.String)
        {
            builder.AppendLine();
            builder.AppendLine($"Contrato: `{planShape.GetString()}`");
        }

        if (command.TryGetProperty("requiredOptions", out var requiredOptions) && requiredOptions.ValueKind == JsonValueKind.Array)
        {
            builder.AppendLine();
            builder.AppendLine("Opcoes obrigatorias:");
            foreach (var option in requiredOptions.EnumerateArray())
            {
                if (option.ValueKind == JsonValueKind.String)
                {
                    builder.AppendLine($"- `{option.GetString()}`");
                }
            }
        }

        if (command.TryGetProperty("rules", out var rules) && rules.ValueKind == JsonValueKind.Array)
        {
            builder.AppendLine();
            builder.AppendLine("Regras:");
            foreach (var rule in rules.EnumerateArray())
            {
                if (rule.ValueKind == JsonValueKind.String)
                {
                    builder.AppendLine($"- {rule.GetString()}");
                }
            }
        }

        if (command.TryGetProperty("examples", out var examples) && examples.ValueKind == JsonValueKind.Array)
        {
            builder.AppendLine();
            builder.AppendLine("Exemplos:");
            foreach (var example in examples.EnumerateArray())
            {
                if (example.ValueKind == JsonValueKind.String)
                {
                    builder.AppendLine($"- `{example.GetString()}`");
                }
            }
        }
    }

    private static bool IsSupportedFormat(string format) =>
        string.Equals(format, "markdown", StringComparison.OrdinalIgnoreCase)
        || string.Equals(format, "json", StringComparison.OrdinalIgnoreCase);

    private static JsonElement GetCommands(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("commands", out var commands))
        {
            return commands;
        }

        return default;
    }

    private static JsonElement FindCommand(JsonElement commands, string targetCommand)
    {
        if (commands.ValueKind != JsonValueKind.Array)
        {
            return default;
        }

        foreach (var command in commands.EnumerateArray())
        {
            if (command.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            if (command.TryGetProperty("name", out var nameElement)
                && nameElement.ValueKind == JsonValueKind.String
                && string.Equals(nameElement.GetString(), targetCommand, StringComparison.OrdinalIgnoreCase))
            {
                return command;
            }
        }

        return default;
    }

    private static string? FindPlanContractsPath()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "references", PlanContractsFileName);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        return null;
    }

    private static void ValidateInsertBlocksPlan(JsonElement root, List<string> errors)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            errors.Add("insert-blocks plan must be a JSON object with a `blocks` array.");
            return;
        }

        if (!TryGetProperty(root, "blocks", out var blocksElement))
        {
            errors.Add("Missing required field: `blocks`.");
            return;
        }

        if (blocksElement.ValueKind != JsonValueKind.Array)
        {
            errors.Add("Field `blocks` must be an array.");
            return;
        }

        if (blocksElement.GetArrayLength() == 0)
        {
            errors.Add("Field `blocks` must contain at least one block.");
        }

        var blockIndex = 0;
        foreach (var blockElement in blocksElement.EnumerateArray())
        {
            ValidateInsertBlock(blockElement, $"blocks[{blockIndex}]", errors);
            blockIndex++;
        }
    }

    private static void ValidateInsertBlock(JsonElement blockElement, string path, List<string> errors)
    {
        if (blockElement.ValueKind != JsonValueKind.Object)
        {
            errors.Add($"{path} must be a JSON object.");
            return;
        }

        ValidateRequiredString(blockElement, "id", path, errors);
        ValidateRequiredString(blockElement, "afterPrefix", path, errors);
        ValidateRequiredString(blockElement, "beforePrefix", path, errors);

        if (TryGetProperty(blockElement, "styleSource", out var styleSourceElement)
            && styleSourceElement.ValueKind == JsonValueKind.String)
        {
            var styleSource = styleSourceElement.GetString();
            if (!string.IsNullOrWhiteSpace(styleSource)
                && !styleSource.Equals("after", StringComparison.OrdinalIgnoreCase)
                && !styleSource.Equals("before", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"{path}.styleSource must be `after` or `before`.");
            }
        }

        if (!TryGetProperty(blockElement, "items", out var itemsElement))
        {
            errors.Add($"Missing required field: `{path}.items`.");
            return;
        }

        if (itemsElement.ValueKind != JsonValueKind.Array)
        {
            errors.Add($"{path}.items must be an array.");
            return;
        }

        if (itemsElement.GetArrayLength() == 0)
        {
            errors.Add($"{path}.items must contain at least one item.");
        }

        var itemIndex = 0;
        foreach (var itemElement in itemsElement.EnumerateArray())
        {
            ValidateInsertBlockItem(itemElement, $"{path}.items[{itemIndex}]", errors);
            itemIndex++;
        }
    }

    private static void ValidateInsertBlockItem(JsonElement itemElement, string path, List<string> errors)
    {
        if (itemElement.ValueKind != JsonValueKind.Object)
        {
            errors.Add($"{path} must be a JSON object.");
            return;
        }

        var kind = ValidateRequiredString(itemElement, "kind", path, errors);
        if (!string.IsNullOrWhiteSpace(kind)
            && !kind.Equals("paragraph", StringComparison.OrdinalIgnoreCase)
            && !kind.Equals("table", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"{path}.kind must be `paragraph` or `table`.");
            return;
        }

        if (kind.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
        {
            var hasText = TryGetStringValue(itemElement, "text", out var textValue) && !string.IsNullOrWhiteSpace(textValue);
            var hasLatex = TryGetStringValue(itemElement, "latex", out var latexValue) && !string.IsNullOrWhiteSpace(latexValue);
            if (!hasText && !hasLatex)
            {
                errors.Add($"{path} with kind `paragraph` must include `text` or `latex`.");
            }
        }
        else if (kind.Equals("table", StringComparison.OrdinalIgnoreCase))
        {
            ValidateRows(itemElement, path, errors);
        }
    }

    private static void ValidateReplaceTablePlan(JsonElement root, List<string> errors)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            errors.Add("replace-table plan must be a JSON object with a `tables` array.");
            return;
        }

        if (!TryGetProperty(root, "tables", out var tablesElement))
        {
            errors.Add("Missing required field: `tables`.");
            return;
        }

        if (tablesElement.ValueKind != JsonValueKind.Array)
        {
            errors.Add("Field `tables` must be an array.");
            return;
        }

        if (tablesElement.GetArrayLength() == 0)
        {
            errors.Add("Field `tables` must contain at least one table.");
        }

        var tableIndex = 0;
        foreach (var tableElement in tablesElement.EnumerateArray())
        {
            ValidateReplaceTableItem(tableElement, $"tables[{tableIndex}]", errors);
            tableIndex++;
        }
    }

    private static void ValidateCreateDocxPlan(JsonElement root, List<string> errors)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            errors.Add("create-docx plan must be a JSON object.");
            return;
        }

        ValidateRequiredString(root, "title", "plan", errors);

        if (!TryGetProperty(root, "paragraphs", out var paragraphsElement))
        {
            errors.Add("Missing required field: `paragraphs`.");
        }
        else if (paragraphsElement.ValueKind != JsonValueKind.Array)
        {
            errors.Add("Field `paragraphs` must be an array.");
        }
        else
        {
            if (paragraphsElement.GetArrayLength() == 0)
            {
                errors.Add("Field `paragraphs` must contain at least one paragraph.");
            }

            var paragraphIndex = 0;
            foreach (var paragraphElement in paragraphsElement.EnumerateArray())
            {
                if (paragraphElement.ValueKind != JsonValueKind.String)
                {
                    errors.Add($"paragraphs[{paragraphIndex}] must be a string.");
                }
                else if (string.IsNullOrWhiteSpace(paragraphElement.GetString()))
                {
                    errors.Add($"paragraphs[{paragraphIndex}] must not be blank.");
                }
                paragraphIndex++;
            }
        }

        if (TryGetProperty(root, "subtitles", out var subtitlesElement))
        {
            ValidateStringArray(subtitlesElement, "subtitles", errors, allowEmpty: true);
        }

        if (TryGetProperty(root, "sections", out var sectionsElement))
        {
            if (sectionsElement.ValueKind != JsonValueKind.Array)
            {
                errors.Add("Field `sections` must be an array.");
            }
            else
            {
                var sectionIndex = 0;
                foreach (var sectionElement in sectionsElement.EnumerateArray())
                {
                    if (sectionElement.ValueKind != JsonValueKind.Object)
                    {
                        errors.Add($"sections[{sectionIndex}] must be a JSON object.");
                        sectionIndex++;
                        continue;
                    }

                    ValidateRequiredString(sectionElement, "heading", $"sections[{sectionIndex}]", errors);
                    if (TryGetProperty(sectionElement, "level", out var levelElement))
                    {
                        if (levelElement.ValueKind != JsonValueKind.Number || !levelElement.TryGetInt32(out var level) || level <= 0)
                        {
                            errors.Add($"sections[{sectionIndex}].level must be a positive integer.");
                        }
                    }

                    if (TryGetProperty(sectionElement, "paragraphs", out var sectionParagraphs))
                    {
                        ValidateStringArray(sectionParagraphs, $"sections[{sectionIndex}].paragraphs", errors, allowEmpty: true);
                    }

                    sectionIndex++;
                }
            }
        }

        if (TryGetProperty(root, "references", out var referencesElement))
        {
            ValidateStringArray(referencesElement, "references", errors, allowEmpty: true);
        }
    }

    private static void ValidateStringArray(JsonElement element, string path, List<string> errors, bool allowEmpty)
    {
        if (element.ValueKind != JsonValueKind.Array)
        {
            errors.Add($"Field `{path}` must be an array.");
            return;
        }

        if (!allowEmpty && element.GetArrayLength() == 0)
        {
            errors.Add($"Field `{path}` must not be empty.");
        }

        var index = 0;
        foreach (var item in element.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
            {
                errors.Add($"{path}[{index}] must be a string.");
            }
            else if (string.IsNullOrWhiteSpace(item.GetString()))
            {
                errors.Add($"{path}[{index}] must not be blank.");
            }
            index++;
        }
    }

    private static void ValidateReplaceTableItem(JsonElement tableElement, string path, List<string> errors)
    {
        if (tableElement.ValueKind != JsonValueKind.Object)
        {
            errors.Add($"{path} must be a JSON object.");
            return;
        }

        ValidateRequiredString(tableElement, "id", path, errors);

        var hasSelector = false;
        if (TryGetPositiveInteger(tableElement, "ordinal", path, errors, out _))
        {
            hasSelector = true;
        }
        if (TryGetPositiveInteger(tableElement, "block", path, errors, out _))
        {
            hasSelector = true;
        }
        if (TryGetPositiveInteger(tableElement, "blockIndex", path, errors, out _))
        {
            hasSelector = true;
        }
        if (TryGetStringValue(tableElement, "firstCellText", out var firstCellText) && !string.IsNullOrWhiteSpace(firstCellText))
        {
            hasSelector = true;
        }
        if (TryGetStringValue(tableElement, "previousParagraphPrefix", out var previousParagraphPrefix) && !string.IsNullOrWhiteSpace(previousParagraphPrefix))
        {
            hasSelector = true;
        }
        if (TryGetStringValue(tableElement, "nextParagraphPrefix", out var nextParagraphPrefix) && !string.IsNullOrWhiteSpace(nextParagraphPrefix))
        {
            hasSelector = true;
        }

        if (!hasSelector)
        {
            errors.Add($"{path} must define at least one selector using `ordinal`, `block`, `blockIndex`, `firstCellText`, `previousParagraphPrefix`, or `nextParagraphPrefix`.");
        }

        ValidateRows(tableElement, path, errors);
    }

    private static void ValidateRows(JsonElement itemElement, string path, List<string> errors)
    {
        if (!TryGetProperty(itemElement, "rows", out var rowsElement))
        {
            errors.Add($"Missing required field: `{path}.rows`.");
            return;
        }

        if (rowsElement.ValueKind != JsonValueKind.Array)
        {
            errors.Add($"{path}.rows must be an array.");
            return;
        }

        if (rowsElement.GetArrayLength() == 0)
        {
            errors.Add($"{path}.rows must contain at least one row.");
        }

        var rowIndex = 0;
        foreach (var rowElement in rowsElement.EnumerateArray())
        {
            if (rowElement.ValueKind != JsonValueKind.Array)
            {
                errors.Add($"{path}.rows[{rowIndex}] must be an array of cells.");
                rowIndex++;
                continue;
            }

            if (rowElement.GetArrayLength() == 0)
            {
                errors.Add($"{path}.rows[{rowIndex}] must contain at least one cell.");
                rowIndex++;
                continue;
            }

            var cellIndex = 0;
            foreach (var cellElement in rowElement.EnumerateArray())
            {
                if (cellElement.ValueKind != JsonValueKind.String)
                {
                    errors.Add($"{path}.rows[{rowIndex}][{cellIndex}] must be a string.");
                }
                cellIndex++;
            }

            rowIndex++;
        }
    }

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement property)
    {
        if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out property))
        {
            return true;
        }

        property = default;
        return false;
    }

    private static string ValidateRequiredString(JsonElement element, string propertyName, string path, List<string> errors)
    {
        if (!TryGetProperty(element, propertyName, out var property))
        {
            errors.Add($"Missing required field: `{path}.{propertyName}`.");
            return "";
        }

        if (property.ValueKind != JsonValueKind.String)
        {
            errors.Add($"{path}.{propertyName} must be a string.");
            return "";
        }

        var value = property.GetString() ?? "";
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"{path}.{propertyName} must not be blank.");
        }

        return value;
    }

    private static bool TryGetStringValue(JsonElement element, string propertyName, out string value)
    {
        value = "";
        if (!TryGetProperty(element, propertyName, out var property) || property.ValueKind != JsonValueKind.String)
        {
            return false;
        }

        value = property.GetString() ?? "";
        return true;
    }

    private static bool TryGetPositiveInteger(JsonElement element, string propertyName, string path, List<string> errors, out int value)
    {
        value = 0;
        if (!TryGetProperty(element, propertyName, out var property))
        {
            return false;
        }

        if (property.ValueKind != JsonValueKind.Number || !property.TryGetInt32(out value) || value <= 0)
        {
            errors.Add($"{path}.{propertyName} must be a positive integer.");
            value = 0;
            return false;
        }

        return true;
    }
}

internal sealed record PlanValidationResult(bool IsValid, IReadOnlyList<string> Errors);
