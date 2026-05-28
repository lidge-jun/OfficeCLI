// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildWatchCommand(Option<bool> jsonOption)
    {
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx, .hwpx)" };
        var watchPortOpt = new Option<int>("--port") { Description = "HTTP port for preview server" };
        watchPortOpt.DefaultValueFactory = _ => 26315;
        var externalOpt = new Option<bool>("--external") { Description = "Also detect external file edits via filesystem watcher" };

        var watchCommand = new Command("watch", "Start a live preview server that refreshes on document changes. Use --external to also detect edits from other tools.");
        watchCommand.Add(watchFileArg);
        watchCommand.Add(watchPortOpt);
        watchCommand.Add(externalOpt);

        // Subcommands — operate against the running watch process via named-pipe IPC.
        // These were previously top-level (`mark`, `unmark`, `get-marks`, `goto`);
        // grouped under `watch` to reflect that they only function while a watch
        // session is alive. The top-level forms remain registered as hidden BC
        // aliases (see CommandBuilder.cs).
        watchCommand.Add(BuildMarkCommand(jsonOption, "mark"));
        watchCommand.Add(BuildUnmarkMarkCommand(jsonOption, "unmark"));
        watchCommand.Add(BuildGetMarksCommand(jsonOption, "marks"));
        watchCommand.Add(BuildGotoCommand(jsonOption, "goto"));

        watchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(watchFileArg)!;
            var port = result.GetValue(watchPortOpt);
            var external = result.GetValue(externalOpt);

            // Render initial HTML: ask the resident process if one is running,
            // otherwise open the file directly as a fallback.
            string? initialHtml = null;
            if (file.Exists)
            {
                // Try resident first — avoids file lock conflict.
                // Json=true makes resident return raw HTML via Console.Write;
                // the resident then wraps it in a JSON envelope { "success":true, "message":"<html>..." }.
                var resp = ResidentClient.TrySend(file.FullName,
                    new ResidentRequest { Command = "view", Args = new() { ["mode"] = "html" }, Json = true },
                    connectTimeoutMs: 2000);
                if (resp is { ExitCode: 0 } && !string.IsNullOrEmpty(resp.Stdout))
                {
                    try
                    {
                        using var doc = System.Text.Json.JsonDocument.Parse(resp.Stdout);
                        if (doc.RootElement.TryGetProperty("message", out var msg))
                            initialHtml = msg.GetString();
                    }
                    catch { /* parse failed — fall through to direct open */ }
                }
                else
                {
                    // No resident — open directly
                    try
                    {
                        using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                        if (handler is OfficeCli.Handlers.PowerPointHandler ppt)
                            initialHtml = ppt.ViewAsHtml();
                        else if (handler is OfficeCli.Handlers.ExcelHandler excel)
                            initialHtml = excel.ViewAsHtml();
                        else if (handler is OfficeCli.Handlers.WordHandler word)
                            initialHtml = word.ViewAsHtml();
                        else if (handler is OfficeCli.Handlers.HwpxHandler hwpx)
                            initialHtml = hwpx.ViewAsHtml();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Warning: initial render failed — preview will show 'Waiting for first update' until the next document change.");
                        Console.Error.WriteLine($"  {ex.GetType().Name}: {ex.Message}");
                        if (Environment.GetEnvironmentVariable("OFFICECLI_DEBUG") == "1" && ex.StackTrace != null)
                            Console.Error.WriteLine(ex.StackTrace);
                    }
                }
            }

            using var cts = new CancellationTokenSource();

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);

            FileSystemWatcher? fsWatcher = null;
            if (external && file.Exists)
            {
                var debounceTimer = new System.Threading.Timer(_ =>
                {
                    try
                    {
                        string? html = null;
                        var resp = ResidentClient.TrySend(file.FullName,
                            new ResidentRequest { Command = "view", Args = new() { ["mode"] = "html" }, Json = true },
                            connectTimeoutMs: 2000);
                        if (resp is { ExitCode: 0 } && !string.IsNullOrEmpty(resp.Stdout))
                        {
                            try
                            {
                                using var doc = System.Text.Json.JsonDocument.Parse(resp.Stdout);
                                if (doc.RootElement.TryGetProperty("message", out var msg))
                                    html = msg.GetString();
                            }
                            catch { }
                        }
                        if (html == null)
                        {
                            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                            if (handler is OfficeCli.Handlers.PowerPointHandler ppt) html = ppt.ViewAsHtml();
                            else if (handler is OfficeCli.Handlers.ExcelHandler excel) html = excel.ViewAsHtml();
                            else if (handler is OfficeCli.Handlers.WordHandler word) html = word.ViewAsHtml();
                            else if (handler is OfficeCli.Handlers.HwpxHandler hwpx) html = hwpx.ViewAsHtml();
                        }
                        if (html != null)
                        {
                            WatchNotifier.NotifyIfWatching(file.FullName, new WatchMessage
                            {
                                Action = "full",
                                FullHtml = html,
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"[external-watch] re-render failed: {ex.Message}");
                    }
                }, null, Timeout.Infinite, Timeout.Infinite);

                fsWatcher = new FileSystemWatcher(file.Directory!.FullName, file.Name)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size,
                    EnableRaisingEvents = true,
                };
                fsWatcher.Changed += (_, _) => debounceTimer.Change(500, Timeout.Infinite);
                Console.Error.WriteLine($"[watch] external file monitoring enabled for {file.Name}");
            }

            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
            fsWatcher?.Dispose();
            return 0;
        }));

        return watchCommand;
    }

    private static Command BuildUnwatchCommand()
    {
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx, .hwpx)" };
        var unwatchCommand = new Command("unwatch", "Stop the watch preview server for the document");
        unwatchCommand.Add(unwatchFileArg);

        unwatchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(unwatchFileArg)!;
            if (WatchNotifier.SendClose(file.FullName))
                Console.WriteLine($"Watch stopped for {file.Name}");
            else
                Console.Error.WriteLine($"No watch running for {file.Name}");
            return 0;
        }));

        return unwatchCommand;
    }
}
