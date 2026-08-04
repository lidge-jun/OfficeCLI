using System.Text.Json.Nodes;
using OfficeCli;

namespace OfficeCli.Tests;

public sealed class ConvertCommandTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(
        Path.GetTempPath(), $"officecli-convert-tests-{Guid.NewGuid():N}");

    public ConvertCommandTests() => Directory.CreateDirectory(_tempDir);

    [Fact]
    public void Convert_RejectsInputOutputIdentityBeforeLaunchingLibreOffice()
    {
        var source = Path.Combine(_tempDir, "same.docx");
        File.WriteAllText(source, "source-sentinel");

        var (exitCode, stdout) = Invoke(
            ["convert", source, "--to", "docx", "--output", source, "--force", "--json"]);

        Assert.Equal(1, exitCode);
        Assert.Equal("source-sentinel", File.ReadAllText(source));
        var root = JsonNode.Parse(stdout)!;
        Assert.Contains("must be different", root["message"]!.GetValue<string>());
    }

    [Fact]
    public void Convert_RejectsExistingOutputWithoutForce()
    {
        var source = Path.Combine(_tempDir, "source.docx");
        var output = Path.Combine(_tempDir, "existing.pdf");
        File.WriteAllText(source, "source-sentinel");
        File.WriteAllText(output, "output-sentinel");

        var (exitCode, stdout) = Invoke(
            ["convert", source, "--to", "pdf", "--output", output, "--json"]);

        Assert.Equal(1, exitCode);
        Assert.Equal("output-sentinel", File.ReadAllText(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.Contains("already exists", root["message"]!.GetValue<string>());
        Assert.Contains("--force", root["message"]!.GetValue<string>());
    }

    [Fact]
    public void ConvertHelp_ExposesExplicitForceOption()
    {
        var (exitCode, stdout) = Invoke(["convert", "--help"]);

        Assert.Equal(0, exitCode);
        Assert.Contains("--force", stdout);
        Assert.Contains("Replace an existing output file", stdout);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    private static (int ExitCode, string Stdout) Invoke(string[] args)
    {
        var originalOut = Console.Out;
        var originalError = Console.Error;
        using var output = new StringWriter();
        using var error = new StringWriter();
        Console.SetOut(output);
        Console.SetError(error);
        try
        {
            var exitCode = CommandBuilder.BuildRootCommand().Parse(args).Invoke();
            return (exitCode, output.ToString() + error.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
    }
}
