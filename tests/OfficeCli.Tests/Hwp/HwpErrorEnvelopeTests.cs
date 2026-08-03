using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

/// <summary>
/// Covers the HWP structured error envelope, including the case an audit
/// flagged as untested: failures raised BEFORE SafeSaveRunner returns carry no
/// transaction, so the envelope must degrade cleanly rather than assume one.
/// </summary>
public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void HwpFailureBeforeSafeSave_OmitsTransactionButKeepsStructuredError()
    {
        if (OperatingSystem.IsWindows()) return;

        // No sidecar configured: the failure happens during runtime discovery,
        // long before any SafeSave transaction exists.
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", "/nonexistent/bridge");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", "/nonexistent/api");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", "/nonexistent/rhwp");
        try
        {
            var input = CreateInput(".hwp");
            var (exitCode, stdout) = InvokeOfficeCli(["view", input, "text", "--json"]);

            Assert.NotEqual(0, exitCode);
            var root = JsonNode.Parse(stdout)!;
            Assert.False(root["success"]!.GetValue<bool>());

            // Structured error survives even with no transaction attached.
            var error = root["error"]!;
            Assert.False(string.IsNullOrWhiteSpace(error["code"]!.GetValue<string>()));

            // The envelope must NOT invent a transaction node here. Attaching one
            // would tell the caller a write was attempted when none was.
            Assert.Null(root["data"]?["transaction"]);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", null);
            Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", null);
            Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", null);
        }
    }

    [Fact]
    public void HwpSafeSaveFailure_AttachesTransactionWithFailedChecks()
    {
        if (OperatingSystem.IsWindows()) return;

        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        try
        {
            var input = CreateInput(".hwp");
            var output = CreateOutput(".hwp");

            // Replace a string the document does not contain: semantic-delta
            // must fail and SafeSave must refuse to publish.
            var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=absent-needle-xyz",
                "--prop", "value=replacement",
                "--prop", $"output={output}",
                "--json"
            ]);

            Assert.NotEqual(0, exitCode);
            var root = JsonNode.Parse(stdout)!;
            Assert.False(root["success"]!.GetValue<bool>());

            var transaction = root["data"]?["transaction"];
            Assert.NotNull(transaction);
            Assert.False(transaction!["ok"]!.GetValue<bool>());

            var checks = transaction["checks"]!.AsArray();
            Assert.Contains(checks, c => c!["name"]!.GetValue<string>() == "semantic-delta"
                                         && !c["ok"]!.GetValue<bool>());
            // The source must survive a refused write.
            Assert.Contains(checks, c => c!["name"]!.GetValue<string>() == "source-preserved"
                                         && c["ok"]!.GetValue<bool>());
            Assert.False(File.Exists(output));
        }
        finally
        {
            Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", null);
            Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", null);
        }
    }
}
