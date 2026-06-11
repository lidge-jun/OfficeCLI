// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using System.Text.Json;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildConvertCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Source document (.docx, .xlsx, .pptx, .hwpx)" };
        var toOpt = new Option<string>("--to") { Description = "Target format (pdf, docx, xlsx, pptx, html, txt)", Required = true };
        var outputOpt = new Option<FileInfo?>("--output") { Description = "Output file path (default: same name with new extension)" };

        var cmd = new Command("convert", "Convert document to another format using LibreOffice");
        cmd.Add(fileArg);
        cmd.Add(toOpt);
        cmd.Add(outputOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var to = result.GetValue(toOpt)!.ToLowerInvariant().TrimStart('.');
            var output = result.GetValue(outputOpt);
            var json = result.GetValue(jsonOption);

            if (!file.Exists)
            {
                if (json) Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = $"File not found: {file.FullName}" }));
                else Console.Error.WriteLine($"Error: file not found: {file.FullName}");
                return 1;
            }

            var soffice = FindSoffice();
            if (soffice == null)
            {
                var msg = "LibreOffice not found. Install LibreOffice and ensure 'soffice' is in PATH.";
                if (json) Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = msg }));
                else Console.Error.WriteLine($"Error: {msg}");
                return 1;
            }

            var outputDir = output?.Directory?.FullName ?? file.Directory!.FullName;
            var outputName = output?.Name ?? Path.ChangeExtension(file.Name, $".{to}");
            var outputPath = Path.Combine(outputDir, outputName);

            // LibreOffice writes to --outdir with auto-generated name, then we rename.
            var tempDir = Path.Combine(Path.GetTempPath(), $"officecli-convert-{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = soffice,
                    ArgumentList = { "--headless", "--convert-to", to, "--outdir", tempDir, file.FullName },
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using var proc = Process.Start(psi)!;
                proc.WaitForExit(60_000);

                if (proc.ExitCode != 0)
                {
                    var stderr = proc.StandardError.ReadToEnd().Trim();
                    if (json) Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = $"LibreOffice exited with code {proc.ExitCode}", detail = stderr }));
                    else Console.Error.WriteLine($"Error: LibreOffice exited with code {proc.ExitCode}\n{stderr}");
                    return 1;
                }

                // Find the generated file in tempDir
                var converted = Directory.GetFiles(tempDir).FirstOrDefault();
                if (converted == null)
                {
                    if (json) Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "LibreOffice produced no output" }));
                    else Console.Error.WriteLine("Error: LibreOffice produced no output");
                    return 1;
                }

                // Move to final destination
                if (File.Exists(outputPath)) File.Delete(outputPath);
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
                File.Move(converted, outputPath);

                if (json)
                    Console.WriteLine(JsonSerializer.Serialize(new { success = true, source = file.FullName, output = outputPath, format = to }));
                else
                    Console.WriteLine($"Converted: {file.Name} → {outputName}");

                return 0;
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }));

        return cmd;
    }

    private static string? FindSoffice()
    {
        var candidates = new[]
        {
            "soffice",
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice",
            "/usr/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            @"C:\Program Files\LibreOffice\program\soffice.exe",
        };

        foreach (var c in candidates)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = c, Arguments = "--version", RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                using var p = Process.Start(psi);
                if (p != null) { p.WaitForExit(5000); if (p.ExitCode == 0) return c; }
            }
            catch { }
        }
        return null;
    }
}
