using FluentAssertions;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class ResidentExitCodeTests : IDisposable
{
    private readonly string _pptxPath;

    public ResidentExitCodeTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
    }

    public void Dispose()
    {
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    [Fact]
    public async Task TryResident_FailedCommand_ShouldReturnNonZeroExitCode()
    {
        // Arrange: create a blank pptx and start a resident server
        BlankDocCreator.Create(_pptxPath);
        using var server = new ResidentServer(_pptxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());

        // Give the server a moment to start listening
        await Task.Delay(200);

        // Act: send a command that will fail (get on a non-existent path)
        var request = new ResidentRequest
        {
            Command = "get",
            Args = { ["path"] = "/slide[999]", ["depth"] = "1" }
        };
        var response = ResidentClient.TrySend(_pptxPath, request);

        // Assert: the response should exist and have a non-zero exit code
        response.Should().NotBeNull("resident server should respond");
        response!.ExitCode.Should().NotBe(0,
            "a failed command should return non-zero exit code, but currently the server always returns 0");
    }

    [Fact]
    public async Task TryResident_SuccessfulCommand_ShouldReturnZeroExitCode()
    {
        // Arrange
        BlankDocCreator.Create(_pptxPath);
        using var server = new ResidentServer(_pptxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());
        await Task.Delay(200);

        // Act: send a valid command (view the file)
        var request = new ResidentRequest
        {
            Command = "view",
            Args = { ["mode"] = "text" }
        };
        var response = ResidentClient.TrySend(_pptxPath, request);

        // Assert: successful command should return exit code 0
        response.Should().NotBeNull();
        response!.ExitCode.Should().Be(0);
    }

    [Fact]
    public async Task TryResident_ExitCodeNotPropagated_BugDemo()
    {
        // This test demonstrates the bug: CommandBuilder.TryResident returns bool,
        // so the caller always does "return;" (exit code 0) even when
        // the resident command failed with ExitCode=1.

        // Arrange
        BlankDocCreator.Create(_pptxPath);
        using var server = new ResidentServer(_pptxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());
        await Task.Delay(200);

        // Act: send a command that will fail (invalid path triggers ArgumentException)
        var request = new ResidentRequest
        {
            Command = "get",
            Args = { ["path"] = "/slide[999]", ["depth"] = "1" }
        };
        var response = ResidentClient.TrySend(_pptxPath, request);

        // The server correctly sets ExitCode=1 for failed commands
        response.Should().NotBeNull();
        response!.ExitCode.Should().Be(1, "server returns ExitCode=1 on failure");

        // BUG: CommandBuilder.TryResident only returns bool (was resident running?).
        // It discards response.ExitCode. The caller then does "return;" = exit code 0.
        //
        // Current: bool TryResident(...) → if (TryResident(...)) return;  // always exit 0
        // Should:  int? TryResident(...) → var rc = TryResident(...); if (rc.HasValue) return rc.Value;
        var tryResidentReturnValue = CommandBuilder.TryResident(_pptxPath, req =>
        {
            req.Command = "get";
            req.Args["path"] = "/slide[999]";
            req.Args["depth"] = "1";
        });

        // FIXED: TryResident now returns int? — non-null means "handled", value is exit code.
        tryResidentReturnValue.Should().NotBeNull("resident handled the request");
        tryResidentReturnValue!.Value.Should().Be(1, "failed command should propagate exit code 1");
    }
}
