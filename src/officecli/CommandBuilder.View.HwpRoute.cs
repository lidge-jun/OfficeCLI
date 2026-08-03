// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// HWP/HWPX view routing. Extracted from the fork's CommandBuilder.View.cs so
// that upstream's View.cs only gains a dispatch call, not 300 lines of body.

using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static void AddHwpViewOption(Dictionary<string, string> args, string key, int? value)
    {
        if (value.HasValue)
            args[key] = value.Value.ToString();
    }

    private static int HandleHwpView(
        string filePath,
        HwpFormat format,
        string mode,
        string? pageFilter,
        bool json,
        string? fieldName = null,
        int? fieldId = null,
        string? outArg = null,
        IReadOnlyDictionary<string, string>? viewArgs = null,
        string? nativeOp = null,
        string[]? nativeArgs = null)
    {
        var modeKey = mode.Trim().ToLowerInvariant();
        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        var operation = HwpViewOperationForMode(modeKey);

        if (!HwpEngineSelector.IsExperimentalBridgeEnabled()
            && !HwpEngineSelector.CanUseInstalledRuntime(formatKey, operation))
        {
            var label = format == HwpFormat.Hwp ? "Binary .hwp" : "HWPX";
            throw new HwpEngineException(
                $"{label} bridge view requires packaged rhwp sidecars or OFFICECLI_HWP_ENGINE=rhwp-experimental.",
                HwpCapabilityConstants.ReasonBridgeNotEnabled,
                "Run ./dev-install.sh, or set OFFICECLI_HWP_ENGINE=rhwp-experimental and install rhwp-officecli-bridge.",
                [
                    HwpCapabilityConstants.OperationReadText,
                    HwpCapabilityConstants.OperationRenderSvg,
                    HwpCapabilityConstants.OperationRenderPng,
                    HwpCapabilityConstants.OperationExportPdf,
                    HwpCapabilityConstants.OperationExportMarkdown,
                    HwpCapabilityConstants.OperationThumbnail,
                    HwpCapabilityConstants.OperationDocumentInfo,
                    HwpCapabilityConstants.OperationDiagnostics,
                    HwpCapabilityConstants.OperationDumpControls,
                    HwpCapabilityConstants.OperationDumpPages,
                    HwpCapabilityConstants.OperationListFields,
                    HwpCapabilityConstants.OperationReadField,
                    HwpCapabilityConstants.OperationReadTableCell,
                    HwpCapabilityConstants.OperationScanCells,
                    HwpCapabilityConstants.OperationNativeRead
                ],
                formatKey,
                operation,
                HwpCapabilityConstants.EngineNone,
                HwpCapabilityConstants.ModeNone);
        }

        var engine = HwpEngineSelector.GetEngine(formatKey, operation);
        var fileInfo = new FileInfo(filePath);
        var ct = CancellationToken.None;

        if (modeKey is "text" or "t")
        {
            var request = new HwpReadRequest(format, filePath, fileInfo.Length, json);
            var result = engine.ReadTextAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = new System.Text.Json.Nodes.JsonObject
                    {
                        ["text"] = result.Text,
                        ["engine"] = result.Engine,
                        ["engineVersion"] = result.EngineVersion
                    },
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Text);
            }
            return 0;
        }

        if (modeKey is "svg" or "g")
        {
            var outDir = Path.Combine(Path.GetTempPath(), $"officecli_hwp_svg_{Guid.NewGuid():N}");
            Directory.CreateDirectory(outDir);
            var request = new HwpRenderRequest(
                format, filePath, outDir,
                pageFilter ?? "all", fileInfo.Length, json);
            var result = engine.RenderSvgAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var pagesArr = new System.Text.Json.Nodes.JsonArray();
                foreach (var p in result.Pages)
                    pagesArr.Add((System.Text.Json.Nodes.JsonNode?)new System.Text.Json.Nodes.JsonObject
                    {
                        ["page"] = p.Page, ["path"] = p.SvgPath, ["sha256"] = p.Sha256
                    });
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = new System.Text.Json.Nodes.JsonObject
                    {
                        ["pages"] = pagesArr,
                        ["manifest"] = result.ManifestPath,
                        ["engine"] = result.Engine,
                        ["engineVersion"] = result.EngineVersion
                    },
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                foreach (var p in result.Pages)
                    Console.WriteLine($"Page {p.Page}: {p.SvgPath}");
            }
            return 0;
        }

        if (modeKey is "png" or "pdf" or "markdown" or "md" or "thumbnail" or "info" or "diagnostics" or "diag" or "dump" or "controls" or "pages" or "dump-pages" or "table-cell" or "cell" or "tables" or "cells" or "native" or "native-op")
        {
            var args = new Dictionary<string, string>(StringComparer.Ordinal);
            if (viewArgs != null)
                foreach (var entry in viewArgs)
                    args[entry.Key] = entry.Value;
            string bridgeCommand;
            var effectiveOperation = operation ?? HwpCapabilityConstants.OperationReadText;
            if (modeKey is "png")
            {
                bridgeCommand = "render-png";
                args["--out-dir"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.Combine(Path.GetTempPath(), $"officecli_hwp_png_{Guid.NewGuid():N}");
                args["--page"] = pageFilter ?? "all";
                Directory.CreateDirectory(args["--out-dir"]);
            }
            else if (modeKey is "pdf")
            {
                bridgeCommand = "export-pdf";
                args["--output"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.GetFullPath(Path.ChangeExtension(filePath, ".pdf"));
                args["--page"] = pageFilter ?? "all";
            }
            else if (modeKey is "markdown" or "md")
            {
                bridgeCommand = "export-markdown";
                args["--page"] = pageFilter ?? "all";
            }
            else if (modeKey is "thumbnail")
            {
                bridgeCommand = "thumbnail";
                args["--output"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.Combine(Path.GetTempPath(), $"officecli_hwp_thumbnail_{Guid.NewGuid():N}.png");
            }
            else if (modeKey is "info")
            {
                bridgeCommand = "document-info";
            }
            else if (modeKey is "diagnostics" or "diag")
            {
                bridgeCommand = "diagnostics";
            }
            else if (modeKey is "dump" or "controls")
            {
                bridgeCommand = "dump-controls";
            }
            else if (modeKey is "pages" or "dump-pages")
            {
                bridgeCommand = "dump-pages";
                if (!string.IsNullOrWhiteSpace(pageFilter))
                    args["--page"] = pageFilter;
            }
            else if (modeKey is "table-cell" or "cell")
            {
                bridgeCommand = "get-cell-text";
            }
            else if (modeKey is "native" or "native-op")
            {
                if (string.IsNullOrWhiteSpace(nativeOp))
                    throw new HwpEngineException(
                        "HWP native view requires --op <rhwp-native-op>.",
                        HwpCapabilityConstants.ReasonUnsupportedOperation,
                        "Example: officecli view input.hwp native --op get-style-list --json",
                        [HwpCapabilityConstants.OperationNativeRead],
                        formatKey,
                        HwpCapabilityConstants.OperationNativeRead,
                        HwpCapabilityConstants.EngineRhwpBridge,
                        HwpCapabilityConstants.ModeExperimental);
                ValidateHwpNativeViewRequest(formatKey, nativeOp, nativeArgs ?? Array.Empty<string>());
                bridgeCommand = "native-op";
                args["--op"] = nativeOp;
                foreach (var (key, value) in ParsePropsArray(nativeArgs ?? Array.Empty<string>()))
                {
                    var normalized = key.StartsWith("--", StringComparison.Ordinal) ? key : $"--{key}";
                    args[normalized] = value;
                }
            }
            else
            {
                bridgeCommand = "scan-cells";
            }

            var request = new HwpJsonViewRequest(format, filePath, fileInfo.Length, effectiveOperation, bridgeCommand, args, json);
            var result = engine.ViewJsonAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var data = (System.Text.Json.Nodes.JsonObject)result.Data.DeepClone();
                data["engine"] = result.Engine;
                data["engineVersion"] = result.EngineVersion;
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = data,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else if (result.Data["markdown"]?.GetValue<string>() is { } markdown)
            {
                Console.WriteLine(markdown);
            }
            else if (result.Data["dump"]?.GetValue<string>() is { } dump)
            {
                Console.WriteLine(dump);
            }
            else if (result.Data["pdf"]?["path"]?.GetValue<string>() is { } pdfPath)
            {
                Console.WriteLine(pdfPath);
            }
            else
            {
                Console.WriteLine(result.Data.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        if (modeKey is "fields")
        {
            var request = new HwpFieldListRequest(format, filePath, fileInfo.Length, json);
            var result = engine.ListFieldsAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = result.Fields.DeepClone(),
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Fields.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        if (modeKey is "field")
        {
            var request = new HwpFieldReadRequest(format, filePath, fieldName, fieldId, fileInfo.Length, json);
            var result = engine.ReadFieldAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = result.Field.DeepClone(),
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Field.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        throw new HwpEngineException(
            $"{formatKey} bridge view mode '{mode}' is not supported. Use text, svg, png, pdf, markdown, thumbnail, info, diagnostics, dump, pages, fields, field, table-cell, tables, or native.",
            HwpCapabilityConstants.ReasonUnsupportedOperation,
            null,
            [
                HwpCapabilityConstants.OperationReadText,
                HwpCapabilityConstants.OperationRenderSvg,
                HwpCapabilityConstants.OperationRenderPng,
                HwpCapabilityConstants.OperationExportPdf,
                HwpCapabilityConstants.OperationExportMarkdown,
                HwpCapabilityConstants.OperationThumbnail,
                HwpCapabilityConstants.OperationDocumentInfo,
                HwpCapabilityConstants.OperationDiagnostics,
                HwpCapabilityConstants.OperationDumpControls,
                HwpCapabilityConstants.OperationDumpPages,
                HwpCapabilityConstants.OperationListFields,
                HwpCapabilityConstants.OperationReadField,
                HwpCapabilityConstants.OperationReadTableCell,
                HwpCapabilityConstants.OperationScanCells,
                HwpCapabilityConstants.OperationNativeRead
            ],
            formatKey,
            null,
            HwpCapabilityConstants.EngineRhwpBridge,
            HwpCapabilityConstants.ModeExperimental);
    }
}
