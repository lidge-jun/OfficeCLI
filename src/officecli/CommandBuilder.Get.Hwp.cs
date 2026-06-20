// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static bool TryHandleHwpProviderRhwpGet(
        string filePath,
        string path,
        bool json,
        out int exitCode)
    {
        exitCode = 0;
        if (!string.Equals(Path.GetExtension(filePath), ".hwp", StringComparison.OrdinalIgnoreCase)
            || !string.Equals(path, "/provider/rhwp", StringComparison.OrdinalIgnoreCase))
            return false;

        if (!File.Exists(filePath))
            throw new CliException($"File not found: {filePath}")
            {
                Code = "file_not_found",
                Suggestion = "Check the file path. Use an absolute path or a path relative to the current directory.",
                Help = "officecli hwp doctor --json"
            };

        var runtime = HwpRuntimeProbe.Probe();
        var data = new JsonObject
        {
            ["format"] = "hwp",
            ["path"] = "/provider/rhwp",
            ["engine"] = HwpCapabilityConstants.EngineRhwpBridge,
            ["status"] = "experimental",
            ["bridgeAvailable"] = runtime.BridgeAvailable,
            ["apiSidecarAvailable"] = runtime.ApiAvailable,
            ["rhwpAvailable"] = runtime.RhwpAvailable,
            ["readRenderAvailable"] = runtime.ReadRenderAvailable,
            ["mutationAvailable"] = runtime.MutationAvailable,
            ["nativeOpAvailable"] = runtime.NativeOpAvailable,
            ["bridgePath"] = runtime.BridgePath,
            ["apiSidecarPath"] = runtime.ApiPath,
            ["rhwpPath"] = runtime.RhwpPath,
            ["supportedOps"] = new JsonArray(runtime.ApiCommands.Select(command => JsonValue.Create(command)).ToArray()),
            ["doctorCommand"] = "officecli hwp doctor --json",
            ["capabilityCommand"] = "officecli capabilities --json",
            ["note"] = "Binary .hwp provider metadata only; generic HWP DOM get remains unsupported."
        };

        if (json)
        {
            Console.WriteLine(OutputFormatter.WrapEnvelope(data.ToJsonString(OutputFormatter.PublicJsonOptions)));
        }
        else
        {
            Console.WriteLine("provider: rhwp");
            Console.WriteLine("format: hwp");
            Console.WriteLine($"engine: {HwpCapabilityConstants.EngineRhwpBridge}");
            Console.WriteLine("status: experimental");
            Console.WriteLine($"bridgeAvailable: {runtime.BridgeAvailable}");
            Console.WriteLine($"apiSidecarAvailable: {runtime.ApiAvailable}");
            Console.WriteLine($"rhwpAvailable: {runtime.RhwpAvailable}");
            Console.WriteLine($"readRenderAvailable: {runtime.ReadRenderAvailable}");
            Console.WriteLine($"mutationAvailable: {runtime.MutationAvailable}");
            Console.WriteLine($"nativeOpAvailable: {runtime.NativeOpAvailable}");
        }

        exitCode = 0;
        return true;
    }
}
