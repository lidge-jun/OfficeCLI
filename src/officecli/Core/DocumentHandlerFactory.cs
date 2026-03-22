// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers;

namespace OfficeCli.Core;

public static class DocumentHandlerFactory
{
    public static IDocumentHandler Open(string filePath, bool editable = false)
    {
        if (!File.Exists(filePath))
            throw new CliException($"File not found: {filePath}")
            {
                Code = "file_not_found",
                Suggestion = "Check the file path. Use an absolute path or a path relative to the current directory.",
                Help = "officecli create <path> --type docx|xlsx|pptx"
            };

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        try
        {
            return ext switch
            {
                ".docx" => new WordHandler(filePath, editable),
                ".xlsx" => new ExcelHandler(filePath, editable),
                ".pptx" => new PowerPointHandler(filePath, editable),
                _ => throw new CliException($"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx")
                {
                    Code = "unsupported_type",
                    ValidValues = [".docx", ".xlsx", ".pptx"]
                }
            };
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException($"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }
}
