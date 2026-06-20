// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildNativeOpsCommand(Option<bool> jsonOption)
    {
        var cmd = new Command("native-ops", "List all available rhwp native operations for HWP/HWPX documents");
        cmd.Add(jsonOption);

        cmd.SetAction(result => SafeRun(() =>
        {
            var json = result.GetValue(jsonOption);
            var ops = GetNativeOpsCatalog();

            if (json)
            {
                Console.WriteLine(JsonSerializer.Serialize(ops, new JsonSerializerOptions { WriteIndented = true }));
            }
            else
            {
                foreach (var cat in ops)
                {
                    Console.WriteLine($"\n  {cat.Category} ({cat.Operations.Length} ops)");
                    Console.WriteLine($"  {"─".PadRight(50, '─')}");
                    foreach (var op in cat.Operations)
                    {
                        var tag = op.FirstClass ? " ★" : "";
                        Console.WriteLine($"    {op.Name,-40} {op.Kind}{tag}");
                    }
                }
                Console.WriteLine($"\n  ★ = first-class command available");
                Console.WriteLine($"  Others: officecli set file.hwp /native-op --prop op=<name> --prop output=<path>");
            }
            return 0;
        }));

        return cmd;
    }

    private static NativeOpsCategory[] GetNativeOpsCatalog()
    {
        var firstClass = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "read-text", "render-svg", "render-png", "export-pdf", "export-markdown",
            "thumbnail", "document-info", "diagnostics", "dump-controls", "dump-pages",
            "list-fields", "read-field", "read-table-cell", "scan-cells",
            "fill-field", "replace-text", "insert-text", "set-table-cell",
            "create-blank", "convert-to-editable", "save-as-hwp",
        };

        return new[]
        {
            new NativeOpsCategory("Text Operations", new NativeOp[]
            {
                new("delete-text", "mutation"), new("split-paragraph", "mutation"),
                new("merge-paragraph", "mutation"), new("insert-paragraph", "mutation"),
                new("delete-paragraph", "mutation"), new("insert-page-break", "mutation"),
                new("insert-column-break", "mutation"), new("set-column-def", "mutation"),
                new("insert-text", "mutation", true), new("replace-text", "mutation", true),
                new("read-text", "read", true), new("search-all-text", "read"),
            }),
            new NativeOpsCategory("Table Operations", new NativeOp[]
            {
                new("create-table", "mutation"), new("create-table-ex", "mutation"),
                new("insert-table-row", "mutation"), new("insert-table-column", "mutation"),
                new("delete-table-row", "mutation"), new("delete-table-column", "mutation"),
                new("merge-table-cells", "mutation"), new("split-table-cell", "mutation"),
                new("split-table-cell-into", "mutation"), new("split-table-cells-in-range", "mutation"),
                new("delete-table-control", "mutation"),
                new("set-table-cell", "mutation", true), new("read-table-cell", "read", true),
                new("scan-cells", "read", true),
            }),
            new NativeOpsCategory("Style & Format Operations", new NativeOp[]
            {
                new("get-char-properties-at", "read"), new("get-para-properties-at", "read"),
                new("get-style-list", "read"), new("get-style-detail", "read"),
                new("update-style", "mutation"), new("update-style-shapes", "mutation"),
                new("create-style", "mutation"), new("delete-style", "mutation"),
                new("apply-char-format", "mutation"), new("apply-char-format-in-cell", "mutation"),
                new("apply-para-format", "mutation"), new("apply-para-format-in-cell", "mutation"),
                new("apply-style", "mutation"), new("apply-cell-style", "mutation"),
                new("get-numbering-list", "read"), new("get-bullet-list", "read"),
                new("ensure-default-numbering", "mutation"),
                new("find-or-create-font-id", "mutation"),
                new("set-numbering-restart", "mutation"),
                new("set-page-hide", "mutation"), new("get-page-hide", "read"),
            }),
            new NativeOpsCategory("Header & Footer Operations", new NativeOp[]
            {
                new("get-header-footer", "read"), new("create-header-footer", "mutation"),
                new("delete-header-footer", "mutation"), new("get-header-footer-list", "read"),
                new("get-header-footer-para-info", "read"),
                new("navigate-header-footer-by-page", "read"),
                new("toggle-hide-header-footer", "mutation"),
                new("get-para-properties-in-hf", "read"), new("apply-para-format-in-hf", "mutation"),
                new("insert-field-in-hf", "mutation"), new("apply-hf-template", "mutation"),
                new("insert-text-in-header-footer", "mutation"),
                new("delete-text-in-header-footer", "mutation"),
                new("split-paragraph-in-header-footer", "mutation"),
                new("merge-paragraph-in-header-footer", "mutation"),
            }),
            new NativeOpsCategory("Shape & Object Operations", new NativeOp[]
            {
                new("insert-picture", "mutation"), new("get-picture-properties", "read"),
                new("set-picture-properties", "mutation"), new("delete-picture-control", "mutation"),
                new("create-shape-control", "mutation"), new("get-shape-properties", "read"),
                new("set-shape-properties", "mutation"), new("delete-shape-control", "mutation"),
                new("change-shape-z-order", "mutation"), new("move-line-endpoint", "mutation"),
                new("group-shapes", "mutation"), new("ungroup-shape", "mutation"),
            }),
            new NativeOpsCategory("Footnote & Equation Operations", new NativeOp[]
            {
                new("get-equation-properties", "read"), new("set-equation-properties", "mutation"),
                new("render-equation-preview", "read"),
                new("insert-footnote", "mutation"), new("get-footnote-info", "read"),
                new("insert-text-in-footnote", "mutation"), new("delete-text-in-footnote", "mutation"),
                new("split-paragraph-in-footnote", "mutation"), new("merge-paragraph-in-footnote", "mutation"),
            }),
        };
    }

    private record NativeOpsCategory(string Category, NativeOp[] Operations);
    private record NativeOp(string Name, string Kind, bool FirstClass = false);
}
