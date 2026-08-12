// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Tests;

public class BatchPowerShellDiagnosticsTests
{
    [Fact]
    public void InlineJsonWithStrippedQuotesNamesPowerShellAndSafeInputPath()
    {
        var result = Invoke([
            "batch", "unused.xlsx", "--commands",
            "[{command:set,path:/Sheet1/A1,props:{value:hi}}]"
        ]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("appears to have lost its quotes", result.Stderr);
        Assert.Contains("common in PowerShell", result.Stderr);
        Assert.Contains("--input <file>", result.Stderr);
        Assert.DoesNotContain("invalid start of a property name", result.Stderr);
    }

    [Fact]
    public void BatchHelpLeadsWithPortableInputAndWarnsAboutInlineQuoting()
    {
        var result = Invoke(["batch", "--help"]);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Prefer --input <file> (or stdin) for portable scripts", result.Stdout);
        Assert.Contains("in PowerShell, use --input to avoid lost quotes", result.Stdout);
        Assert.True(
            result.Stdout.IndexOf("--input <file>", StringComparison.Ordinal)
            < result.Stdout.IndexOf("--commands", StringComparison.Ordinal),
            "portable --input guidance should appear before --commands");
    }

    private static (int ExitCode, string Stdout, string Stderr) Invoke(string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        var originalOut = Console.Out;
        var originalError = Console.Error;
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        Console.SetOut(stdout);
        Console.SetError(stderr);
        try
        {
            var exitCode = root.Parse(args).Invoke();
            return (exitCode, stdout.ToString(), stderr.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
    }
}
