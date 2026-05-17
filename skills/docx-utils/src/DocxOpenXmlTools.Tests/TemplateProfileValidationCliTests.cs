using System.Text.Json;
using System.Text;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateProfileValidationCliTests
{
    [Fact]
    public async Task Validate_template_profile_accepts_canonical_profile_with_matching_hash()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);

        var result = await TemplateCliTestSupport.RunToolsAsync($"validate-template-profile \"{profilePath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Profile valido", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Validate_template_profile_rejects_invalid_hash_and_reports_the_mismatch()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteInvalidProfile(tempDir.Path, templatePath);

        var result = await TemplateCliTestSupport.RunToolsAsync($"validate-template-profile \"{profilePath}\"");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("sha256", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("invalido", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Validate_template_profile_detects_schema_problem_when_required_region_is_missing()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var profilePath = Path.Combine(tempDir.Path, "profile.missing-region.json");

        var profile = new
        {
            template = new
            {
                name = "template-profile-sintetico",
                sourceFile = Path.GetFileName(templatePath),
                sha256 = TemplateCliTestSupport.ComputeSha256(templatePath),
                profileVersion = 1
            },
            regions = new object[]
            {
                new
                {
                    role = "title",
                    templateBlockId = "p-0002",
                    replaceWith = "source.title",
                    preserveFormatting = true
                }
            }
        };

        await File.WriteAllTextAsync(profilePath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);

        var result = await TemplateCliTestSupport.RunToolsAsync($"validate-template-profile \"{profilePath}\"");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("references", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("region", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
    }
}
