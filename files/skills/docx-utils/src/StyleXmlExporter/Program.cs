using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: StyleXmlExporter <input.docx> <output-directory>");
    return 1;
}

var docxPath = Path.GetFullPath(args[0]);
var outDir = Path.GetFullPath(args[1]);

if (!File.Exists(docxPath))
{
    Console.Error.WriteLine($"DOCX not found: {docxPath}");
    return 2;
}

Directory.CreateDirectory(outDir);
var styleDir = Path.Combine(outDir, "styles");
var effectsDir = Path.Combine(outDir, "stylesWithEffects");
Directory.CreateDirectory(styleDir);
Directory.CreateDirectory(effectsDir);

using var doc = WordprocessingDocument.Open(docxPath, false);
var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("No main document part.");
var stylesPart = mainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("No style definitions part.");

var stylesXml = ReadPartXml(stylesPart);
File.WriteAllText(Path.Combine(outDir, "styles.xml"), stylesXml);

var stylesDoc = XDocument.Parse(stylesXml, LoadOptions.PreserveWhitespace);
var w = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

var manifest = new List<object>();
var styleIndex = 0;

foreach (var style in stylesDoc.Root?.Elements(w + "style") ?? Enumerable.Empty<XElement>())
{
    styleIndex++;
    var type = style.Attribute(w + "type")?.Value ?? "unknown";
    var styleId = style.Attribute(w + "styleId")?.Value ?? $"style_{styleIndex:0000}";
    var name = style.Element(w + "name")?.Attribute(w + "val")?.Value ?? styleId;
    var fileName = $"{styleIndex:0000}_{Sanitize(type)}_{Sanitize(styleId)}_{Sanitize(name)}.xml";
    var relativePath = Path.Combine("styles", fileName).Replace('\\', '/');

    WriteElement(Path.Combine(outDir, relativePath), style);
    manifest.Add(new
    {
        source = "styles.xml",
        index = styleIndex,
        type,
        styleId,
        name,
        file = relativePath
    });
}

var docDefaults = stylesDoc.Root?.Element(w + "docDefaults");
if (docDefaults is not null)
{
    WriteElement(Path.Combine(outDir, "docDefaults.xml"), docDefaults);
}

var latentStyles = stylesDoc.Root?.Element(w + "latentStyles");
if (latentStyles is not null)
{
    WriteElement(Path.Combine(outDir, "latentStyles.xml"), latentStyles);
}

var effectsPart = mainPart.StylesWithEffectsPart;
var effectsCount = 0;
if (effectsPart is not null)
{
    var effectsXml = ReadPartXml(effectsPart);
    File.WriteAllText(Path.Combine(outDir, "stylesWithEffects.xml"), effectsXml);
    var effectsDoc = XDocument.Parse(effectsXml, LoadOptions.PreserveWhitespace);

    foreach (var style in effectsDoc.Root?.Elements(w + "style") ?? Enumerable.Empty<XElement>())
    {
        effectsCount++;
        var type = style.Attribute(w + "type")?.Value ?? "unknown";
        var styleId = style.Attribute(w + "styleId")?.Value ?? $"style_{effectsCount:0000}";
        var name = style.Element(w + "name")?.Attribute(w + "val")?.Value ?? styleId;
        var fileName = $"{effectsCount:0000}_{Sanitize(type)}_{Sanitize(styleId)}_{Sanitize(name)}.xml";
        var relativePath = Path.Combine("stylesWithEffects", fileName).Replace('\\', '/');

        WriteElement(Path.Combine(outDir, relativePath), style);
    }
}

var manifestObject = new
{
    generatedAt = DateTimeOffset.Now.ToString("O"),
    sourceDocx = docxPath,
    outputDirectory = outDir,
    styleDefinitionsPart = new
    {
        fullXml = "styles.xml",
        docDefaults = docDefaults is null ? null : "docDefaults.xml",
        latentStyles = latentStyles is null ? null : "latentStyles.xml",
        styleCount = styleIndex,
        styles = manifest
    },
    stylesWithEffectsPart = effectsPart is null
        ? null
        : new
        {
            fullXml = "stylesWithEffects.xml",
            styleCount = effectsCount,
            directory = "stylesWithEffects"
        }
};

File.WriteAllText(
    Path.Combine(outDir, "manifest.json"),
    JsonSerializer.Serialize(manifestObject, new JsonSerializerOptions { WriteIndented = true })
);

Console.WriteLine($"Exported {styleIndex} styles from styles.xml to {styleDir}");
if (effectsPart is not null)
{
    Console.WriteLine($"Exported {effectsCount} styles from stylesWithEffects.xml to {effectsDir}");
}
Console.WriteLine($"Manifest: {Path.Combine(outDir, "manifest.json")}");
return 0;

static string ReadPartXml(OpenXmlPart part)
{
    using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
    using var reader = new StreamReader(stream);
    return reader.ReadToEnd();
}

static void WriteElement(string path, XElement element)
{
    Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
    var doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(element));
    File.WriteAllText(path, doc.ToString(SaveOptions.DisableFormatting));
}

static string Sanitize(string value)
{
    var sanitized = Regex.Replace(value.Trim(), @"[^\p{L}\p{Nd}._-]+", "_");
    sanitized = sanitized.Trim('_');
    if (string.IsNullOrWhiteSpace(sanitized))
    {
        return "unnamed";
    }

    return sanitized.Length <= 80 ? sanitized : sanitized[..80];
}
