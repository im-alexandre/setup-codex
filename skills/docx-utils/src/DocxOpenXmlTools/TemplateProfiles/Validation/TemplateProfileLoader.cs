using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace DocxOpenXmlTools.TemplateProfiles;

internal static class TemplateProfileLoader
{
    public static TemplateProfileLoadResult Load(string profilePath)
    {
        if (string.IsNullOrWhiteSpace(profilePath))
        {
            return new TemplateProfileLoadResult
            {
                Errors = ["Profile path is required."]
            };
        }

        var normalizedProfilePath = Path.GetFullPath(profilePath);
        var errors = new List<string>();

        if (!File.Exists(normalizedProfilePath))
        {
            errors.Add($"Profile not found: {normalizedProfilePath}");
            return new TemplateProfileLoadResult
            {
                ProfilePath = normalizedProfilePath,
                Errors = errors
            };
        }

        string json;
        try
        {
            json = File.ReadAllText(normalizedProfilePath, Encoding.UTF8);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            errors.Add($"Unable to read profile: {normalizedProfilePath}. {ex.Message}");
            return new TemplateProfileLoadResult
            {
                ProfilePath = normalizedProfilePath,
                Errors = errors
            };
        }

        TemplateProfileDocument? profile;
        try
        {
            profile = JsonSerializer.Deserialize<TemplateProfileDocument>(json, ReadJsonOptions());
        }
        catch (JsonException ex)
        {
            errors.Add($"Profile invalid: unable to parse JSON in {normalizedProfilePath}. {ex.Message}");
            return new TemplateProfileLoadResult
            {
                ProfilePath = normalizedProfilePath,
                Errors = errors
            };
        }

        if (profile is null)
        {
            errors.Add($"Profile invalid: JSON root is empty in {normalizedProfilePath}.");
            return new TemplateProfileLoadResult
            {
                ProfilePath = normalizedProfilePath,
                Errors = errors
            };
        }

        profile = profile with
        {
            Template = profile.Template ?? new TemplateProfileTemplate(),
            Regions = profile.Regions ?? []
        };

        var resolvedTemplatePath = ResolveTemplatePath(normalizedProfilePath, profile.Template.SourceFile);
        return new TemplateProfileLoadResult
        {
            ProfilePath = normalizedProfilePath,
            Profile = profile,
            ResolvedTemplatePath = resolvedTemplatePath,
            Errors = errors
        };
    }

    private static string ResolveTemplatePath(string profilePath, string? sourceFile)
    {
        if (string.IsNullOrWhiteSpace(sourceFile))
        {
            return string.Empty;
        }

        if (Path.IsPathRooted(sourceFile))
        {
            return Path.GetFullPath(sourceFile);
        }

        var profileDirectory = Path.GetDirectoryName(Path.GetFullPath(profilePath)) ?? ".";
        return Path.GetFullPath(Path.Combine(profileDirectory, sourceFile));
    }

    private static JsonSerializerOptions ReadJsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };
}

internal sealed record TemplateProfileLoadResult
{
    public string ProfilePath { get; init; } = string.Empty;

    public TemplateProfileDocument? Profile { get; init; }

    public string ResolvedTemplatePath { get; init; } = string.Empty;

    public IReadOnlyList<string> Errors { get; init; } = [];

    public bool IsLoaded => Profile is not null && Errors.Count == 0;
}
