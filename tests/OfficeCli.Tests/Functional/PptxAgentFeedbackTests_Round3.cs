// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for PPT bugs found in Agent A Round 3.
/// Bug 1: autoFit=auto/shrink silently ignored
/// Bug 2: Connector lineDash ordering in ln element
/// </summary>
public class PptxAgentFeedbackTests_Round3 : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxAgentFeedbackTests_Round3()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Bug 1: autoFit=auto/shrink ====================

    [Fact]
    public void SetAutoFitAuto_ShouldApplyNormalAutoFit()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });

        _handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "auto" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit");
        node.Format["autoFit"].ToString().Should().Be("normal",
            "autoFit=auto should map to NormalAutoFit which reads back as 'normal'");
    }

    [Fact]
    public void SetAutoFitShrink_ShouldApplyNormalAutoFit()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });

        _handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "shrink" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit");
        node.Format["autoFit"].ToString().Should().Be("normal");
    }

    [Fact]
    public void SetAutoFitAuto_PersistsAfterReopen()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "auto" });

        Reopen();

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["autoFit"].ToString().Should().Be("normal");
    }
}
