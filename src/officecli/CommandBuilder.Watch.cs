// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildWatchCommand()
    {
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
        var watchPortOpt = new Option<int>("--port") { Description = "HTTP port for preview server" };
        watchPortOpt.DefaultValueFactory = _ => 18080;

        var watchCommand = new Command("watch", "Start a live preview server that auto-refreshes when the document changes");
        watchCommand.Add(watchFileArg);
        watchCommand.Add(watchPortOpt);

        watchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(watchFileArg)!;
            var port = result.GetValue(watchPortOpt);

            // Render initial HTML from existing file content
            string? initialHtml = null;
            if (file.Exists)
            {
                try
                {
                    using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                    if (handler is OfficeCli.Handlers.PowerPointHandler ppt)
                        initialHtml = ppt.ViewAsHtml();
                    else if (handler is OfficeCli.Handlers.ExcelHandler excel)
                        initialHtml = excel.ViewAsHtml();
                    else if (handler is OfficeCli.Handlers.WordHandler word)
                        initialHtml = word.ViewAsHtml();
                }
                catch { /* ignore — will show waiting page */ }
            }

            using var cts = new CancellationTokenSource();

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);
            // Signal handling (SIGTERM / SIGINT / SIGHUP / SIGQUIT) is
            // now registered inside WatchServer.RunAsync via
            // PosixSignalRegistration, which runs BEFORE the .NET runtime
            // begins its shutdown sequence (on a healthy ThreadPool).
            // That path runs StopAsync to completion — including
            // TcpListener.Stop() (the only reliable way to unstick
            // AcceptTcpClientAsync on macOS) and the CoreFxPipe_ socket
            // cleanup (BUG-BT-003) — before calling Environment.Exit.
            //
            // The older Console.CancelKeyPress + ProcessExit combo was
            // unreliable: SIGINT would cancel _cts but the TCP accept
            // loop did not honour cancellation on macOS, hanging the
            // process for 15+ seconds; ProcessExit ran during runtime
            // teardown when ThreadPool was already unwinding, so the
            // socket cleanup silently skipped.
            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
            return 0;
        }));

        return watchCommand;
    }

    private static Command BuildUnwatchCommand()
    {
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
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
