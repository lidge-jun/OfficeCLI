// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

// Internal commands (spawned as separate processes, not user-facing)
if (args.Length == 1 && args[0] == "__update-check__")
{
    OfficeCli.Core.UpdateChecker.RunRefresh();
    return 0;
}

// Config command: officecli config <key> [value]
if (args.Length >= 2 && args[0] == "config")
{
    OfficeCli.Core.CliLogger.LogCommand(args);
    OfficeCli.Core.UpdateChecker.HandleConfigCommand(args.Skip(1).ToArray());
    return 0;
}

// Log command
OfficeCli.Core.CliLogger.LogCommand(args);

// Non-blocking update check: spawns background upgrade if stale
if (Environment.GetEnvironmentVariable("OFFICECLI_SKIP_UPDATE") != "1")
    OfficeCli.Core.UpdateChecker.CheckInBackground();

var rootCommand = OfficeCli.CommandBuilder.BuildRootCommand();

if (args.Length == 0)
{
    rootCommand.Parse("--help").Invoke();
    return 0;
}

// Handle help commands (docx/xlsx/pptx) before System.CommandLine parsing
// so that --help also shows our custom output instead of the default help
if (OfficeCli.HelpCommands.TryHandle(args))
    return 0;

var parseResult = rootCommand.Parse(args);
return parseResult.Invoke();
