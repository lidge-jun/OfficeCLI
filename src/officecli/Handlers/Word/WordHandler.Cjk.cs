// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// DORMANT (wp5 decision): these WordHandler CJK helpers have no callers in
// this port. Their call sites live in fork-side Word/PPT host modifications
// that were NOT ported -- upstream has since rewritten those paths, and
// forcing the old call sites back in would regress upstream behavior for a
// benefit no HWP route needs (HWPX Korean handling lives in
// HwpxHandler.Korean.cs and Core/CjkHelper.cs, both of which ARE wired).
//
// Kept rather than deleted because they are the reference implementation for
// CJK run segmentation if OOXML Korean support is picked up later. Classified
// explicitly so this is a recorded decision, not unexplained dead code.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private static List<Run> BuildSegmentedRuns(string text, RunProperties? template = null)
    {
        var segments = CjkHelper.SegmentText(text);
        if (segments.Count == 0)
            segments = new List<(string text, CjkScript script)> { ("", CjkScript.None) };

        var runs = new List<Run>();
        foreach (var (segmentText, script) in segments)
        {
            var rPr = template?.CloneNode(true) as RunProperties ?? new RunProperties();
            if (script != CjkScript.None)
                CjkHelper.ApplyToWordRun(rPr, script);
            else
                CjkHelper.ClearWordRunCjk(rPr);

            // Route segment text through AppendTextWithBreaks so that '\n' (w:br)
            // and '\t' (w:tab) inside a CJK segment round-trip through Word
            // instead of collapsing to a space — matches upstream plain-run path.
            var run = new Run(rPr);
            AppendTextWithBreaks(run, segmentText);
            runs.Add(run);
        }

        return runs;
    }
}
