using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void OfficeCliGetProviderRhwp_ReturnsBinaryHwpProviderMetadata()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "get", input, "/provider/rhwp",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.DoesNotContain("hwp_generic_handler_unsupported", stdout);
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal("hwp", root["data"]!["format"]!.GetValue<string>());
        Assert.Equal("/provider/rhwp", root["data"]!["path"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.Equal("experimental", root["data"]!["status"]!.GetValue<string>());
        Assert.True(root["data"]!["bridgeAvailable"]!.GetValue<bool>());
        Assert.True(root["data"]!["apiSidecarAvailable"]!.GetValue<bool>());
        Assert.True(root["data"]!["rhwpAvailable"]!.GetValue<bool>());
        Assert.True(root["data"]!["readRenderAvailable"]!.GetValue<bool>());
        Assert.True(root["data"]!["mutationAvailable"]!.GetValue<bool>());
    }
}
