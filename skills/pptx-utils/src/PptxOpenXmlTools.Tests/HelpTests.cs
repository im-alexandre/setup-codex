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
    private static readonly Lazy<Task> BuildAppOnce = new(BuildAppAsync);

    public static async Task<CliResult> RunAsync(params string[] args)
    {
        var project = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "PptxOpenXmlTools", "PptxOpenXmlTools.csproj"));
        await BuildAppOnce.Value;

        var dll = Path.Combine(Path.GetDirectoryName(project)!, "bin", "Release", "net8.0", "pptx-utils.dll");
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        psi.ArgumentList.Add(dll);
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

    private static async Task BuildAppAsync()
    {
        var project = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "PptxOpenXmlTools", "PptxOpenXmlTools.csproj"));
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        psi.ArgumentList.Add("build");
        psi.ArgumentList.Add(project);
        psi.ArgumentList.Add("--configuration");
        psi.ArgumentList.Add("Release");
        psi.ArgumentList.Add("--no-restore");

        using var process = Process.Start(psi)!;
        await process.StandardOutput.ReadToEndAsync();
        await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Failed to build app project: {project}");
        }
    }
}

internal sealed record CliResult(int ExitCode, string Stdout, string Stderr);
