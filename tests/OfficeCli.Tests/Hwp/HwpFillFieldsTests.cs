using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

/// <summary>
/// Coverage for the batch <c>fill-fields</c> bridge command.
/// </summary>
/// <remarks>
/// These drive the REAL Rust sidecar rather than a shell fake: the whole point
/// of the command is that one process call performs N mutations with a single
/// parse/serialize cycle, and a fake that echoes JSON would prove none of that.
/// They also run out-of-process, so they avoid the Console.Out redirect race
/// documented in 007_sidecar_flakiness.md. They skip cleanly when the debug
/// sidecar has not been built.
/// </remarks>
public class HwpFillFieldsTests
{
    private static string? LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        return null;
    }

    private static string? LocateApiBridge()
        => LocateRepoFile("src/rhwp-field-bridge/target/debug/rhwp-field-bridge");

    private static string? LocateFixture()
        => LocateRepoFile("tests/fixtures/hwp/rhwp-fields/field-01.hwp");

    private static (int ExitCode, string Stdout, string Stderr) RunBridge(string bridge, string[] args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = bridge,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var a in args) psi.ArgumentList.Add(a);

        using var proc = Process.Start(psi)!;
        var stdout = proc.StandardOutput.ReadToEnd();
        var stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit();
        return (proc.ExitCode, stdout, stderr);
    }

    private static string TempOut() =>
        Path.Combine(Path.GetTempPath(), $"fill-fields-{Guid.NewGuid():N}.hwp");

    [Fact]
    public void FillFields_SetsEveryNamedFieldInOnePass()
    {
        var bridge = LocateApiBridge();
        var fixture = LocateFixture();
        if (bridge is null || fixture is null) return;

        var output = TempOut();
        try
        {
            var (exit, stdout, _) = RunBridge(bridge,
            [
                "fill-fields", "--format", "hwp", "--input", fixture, "--output", output,
                "--set", "회사명=리지에이아이",
                "--set", "작성자=전준",
                "--set", "부서명=개발팀",
                "--json"
            ]);

            Assert.Equal(0, exit);
            var root = JsonNode.Parse(stdout)!;
            Assert.Equal(3, root["filled"]!.GetValue<int>());
            Assert.Equal(3, root["requested"]!.GetValue<int>());
            Assert.True(File.Exists(output));

            // Read back from the WRITTEN file: the mutation must survive
            // serialization, not merely be reported as applied.
            var (readExit, readOut, _) = RunBridge(bridge,
                ["list-fields", "--format", "hwp", "--input", output, "--json"]);
            Assert.Equal(0, readExit);
            var fields = JsonNode.Parse(readOut)!["fields"]!.AsArray();

            string? ValueOf(string name) => fields
                .FirstOrDefault(f => f!["name"]!.GetValue<string>() == name
                                     && !string.IsNullOrEmpty(f["value"]!.GetValue<string>()))
                ?["value"]!.GetValue<string>();

            Assert.Equal("리지에이아이", ValueOf("회사명"));
            Assert.Equal("전준", ValueOf("작성자"));
            Assert.Equal("개발팀", ValueOf("부서명"));
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void FillFields_StrictModeWritesNothingWhenAnyFieldFails()
    {
        var bridge = LocateApiBridge();
        var fixture = LocateFixture();
        if (bridge is null || fixture is null) return;

        var output = TempOut();
        try
        {
            var (exit, _, _) = RunBridge(bridge,
            [
                "fill-fields", "--format", "hwp", "--input", fixture, "--output", output,
                "--set", "회사명=good",
                "--set", "존재하지않는필드=bad",
                "--json"
            ]);

            Assert.NotEqual(0, exit);
            // The point of atomicity: a form must not be left half-filled, which
            // would look complete to a reader while silently missing values.
            Assert.False(File.Exists(output));
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void FillFields_NonStrictSkipsUnknownAndReportsIt()
    {
        var bridge = LocateApiBridge();
        var fixture = LocateFixture();
        if (bridge is null || fixture is null) return;

        var output = TempOut();
        try
        {
            var (exit, stdout, _) = RunBridge(bridge,
            [
                "fill-fields", "--format", "hwp", "--input", fixture, "--output", output,
                "--strict", "false",
                "--set", "회사명=good",
                "--set", "없는필드=skipped",
                "--json"
            ]);

            Assert.Equal(0, exit);
            var root = JsonNode.Parse(stdout)!;
            Assert.Equal(1, root["filled"]!.GetValue<int>());
            Assert.Equal(2, root["requested"]!.GetValue<int>());
            Assert.Contains(root["warnings"]!.AsArray(),
                w => w!.GetValue<string>().Contains("없는필드"));
            Assert.True(File.Exists(output));
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void FillFields_PreservesEqualsSignInsideValues()
    {
        var bridge = LocateApiBridge();
        var fixture = LocateFixture();
        if (bridge is null || fixture is null) return;

        var output = TempOut();
        try
        {
            // Only the FIRST '=' separates name from value.
            var (exit, _, _) = RunBridge(bridge,
            [
                "fill-fields", "--format", "hwp", "--input", fixture, "--output", output,
                "--set", "이메일=a=b@c.com", "--json"
            ]);
            Assert.Equal(0, exit);

            var (_, readOut, _) = RunBridge(bridge,
                ["get-field", "--format", "hwp", "--input", output, "--name", "이메일", "--json"]);
            Assert.Equal("a=b@c.com",
                JsonNode.Parse(readOut)!["field"]!["value"]!.GetValue<string>());
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void FillFields_RejectsMalformedAssignment()
    {
        var bridge = LocateApiBridge();
        var fixture = LocateFixture();
        if (bridge is null || fixture is null) return;

        var output = TempOut();
        try
        {
            var (exit, _, stderr) = RunBridge(bridge,
            [
                "fill-fields", "--format", "hwp", "--input", fixture, "--output", output,
                "--set", "no-equals-sign", "--json"
            ]);

            Assert.NotEqual(0, exit);
            Assert.Contains("name=value", stderr);
            Assert.False(File.Exists(output));
        }
        finally { File.Delete(output); }
    }
}
