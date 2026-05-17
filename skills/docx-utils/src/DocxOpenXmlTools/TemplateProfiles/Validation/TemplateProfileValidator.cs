using System.Security.Cryptography;

namespace DocxOpenXmlTools.TemplateProfiles;

internal static class TemplateProfileValidator
{
    private static readonly string[] RequiredRegions = ["title", "abstract", "references"];

    public static TemplateProfileValidationResult Validate(string profilePath)
    {
        var loadResult = TemplateProfileLoader.Load(profilePath);
        var errors = new List<string>(loadResult.Errors);

        if (loadResult.Profile is null)
        {
            return new TemplateProfileValidationResult
            {
                IsValid = false,
                Profile = null,
                ResolvedTemplatePath = loadResult.ResolvedTemplatePath,
                Errors = errors
            };
        }

        var profile = loadResult.Profile;
        ValidateTemplateMetadata(loadResult.ProfilePath, profile, loadResult.ResolvedTemplatePath, errors);
        ValidateRequiredRegions(profile, errors);

        if (!string.IsNullOrWhiteSpace(loadResult.ResolvedTemplatePath) && File.Exists(loadResult.ResolvedTemplatePath))
        {
            ValidateTemplateHash(profile, loadResult.ResolvedTemplatePath, errors);
        }
        else if (!string.IsNullOrWhiteSpace(loadResult.ResolvedTemplatePath))
        {
            errors.Add($"Template not found: {loadResult.ResolvedTemplatePath}. The profile sourceFile must point to a DOCX file relative to the profile.");
        }

        return new TemplateProfileValidationResult
        {
            IsValid = errors.Count == 0,
            Profile = profile,
            ResolvedTemplatePath = loadResult.ResolvedTemplatePath,
            Errors = errors
        };
    }

    private static void ValidateTemplateMetadata(string profilePath, TemplateProfileDocument profile, string resolvedTemplatePath, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(profile.Template.Name))
        {
            errors.Add($"Perfil invalido em {profilePath}: template.name e obrigatorio.");
        }

        if (string.IsNullOrWhiteSpace(profile.Template.SourceFile))
        {
            errors.Add($"Perfil invalido em {profilePath}: template.sourceFile e obrigatorio e deve ser relativo ao arquivo de perfil.");
        }

        if (string.IsNullOrWhiteSpace(profile.Template.Sha256))
        {
            errors.Add($"Perfil invalido em {profilePath}: template.sha256 e obrigatorio.");
        }

        if (profile.Template.ProfileVersion <= 0)
        {
            errors.Add($"Perfil invalido em {profilePath}: template.profileVersion deve ser maior que zero.");
        }

        if (!string.IsNullOrWhiteSpace(profile.Template.SourceFile) && string.IsNullOrWhiteSpace(resolvedTemplatePath))
        {
            errors.Add($"Perfil invalido em {profilePath}: template.sourceFile nao pode ser resolvido.");
        }
    }

    private static void ValidateRequiredRegions(TemplateProfileDocument profile, List<string> errors)
    {
        foreach (var requiredRegion in RequiredRegions)
        {
            if (!profile.Regions.Any(region => string.Equals(region.Role, requiredRegion, StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add($"Perfil invalido: regiao obrigatoria '{requiredRegion}' ausente (required region region). Adicione uma regiao com role '{requiredRegion}' ao profile.canonical.json.");
            }
        }
    }

    private static void ValidateTemplateHash(TemplateProfileDocument profile, string resolvedTemplatePath, List<string> errors)
    {
        var actualHash = ComputeSha256(resolvedTemplatePath);
        if (!string.Equals(actualHash, profile.Template.Sha256, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"Perfil invalido: sha256/hash divergente para o template '{resolvedTemplatePath}'. Esperado {profile.Template.Sha256}, encontrado {actualHash}.");
        }
    }

    private static string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}

internal sealed record TemplateProfileValidationResult
{
    public bool IsValid { get; init; }

    public TemplateProfileDocument? Profile { get; init; }

    public string ResolvedTemplatePath { get; init; } = string.Empty;

    public IReadOnlyList<string> Errors { get; init; } = [];
}
