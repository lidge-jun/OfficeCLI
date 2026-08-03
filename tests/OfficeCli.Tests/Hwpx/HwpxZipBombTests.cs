using System.IO.Compression;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Hwpx;

/// <summary>
/// Decompression-bomb coverage for HWPX.
/// </summary>
/// <remarks>
/// Two distinct boundaries, and only the first was ever guarded upstream:
///
///   intact central directory -> DocumentHandlerFactory.GuardDecompressionBomb
///   damaged central directory -> HwpxHandler.TryRecoverBrokenZip
///
/// The salvage path is the dangerous one: it reads the entire file and inflates
/// every entry it finds, so a crafted broken archive bypassed the guard that
/// only ever ran on parseable directories.
/// </remarks>
public class HwpxZipBombTests
{
    private static string TempPath(string ext)
        => Path.Combine(Path.GetTempPath(), $"ocx-bomb-{Guid.NewGuid():N}{ext}");

    /// <summary>A valid zip whose entries inflate far beyond their stored size.</summary>
    private static string CreateValidBomb(int entries, int megabytesEach)
    {
        var path = TempPath(".hwpx");
        var zeros = new byte[megabytesEach * 1024 * 1024];
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            for (var i = 0; i < entries; i++)
            {
                var entry = zip.CreateEntry($"Contents/section{i}.xml", CompressionLevel.SmallestSize);
                using var s = entry.Open();
                s.Write(zeros, 0, zeros.Length);
            }
        }
        return path;
    }

    /// <summary>Truncate the tail so the central directory is unreadable.</summary>
    private static string BreakCentralDirectory(string zipPath)
    {
        var bytes = File.ReadAllBytes(zipPath);
        var broken = TempPath(".hwpx");
        // Drop the trailing directory; local file headers survive, which is
        // exactly what pushes the open into the recovery path.
        File.WriteAllBytes(broken, bytes.AsSpan(0, (int)(bytes.Length * 0.80)).ToArray());
        return broken;
    }

    [Fact]
    public void IntactBomb_IsRejectedByTheDecompressionGuard()
    {
        // ~1.2 GB inflated from a few KB stored.
        var bomb = CreateValidBomb(entries: 120, megabytesEach: 10);
        try
        {
            var ex = Assert.ThrowsAny<Exception>(() => DocumentHandlerFactory.Open(bomb));
            Assert.Contains("bomb", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally { File.Delete(bomb); }
    }

    [Fact]
    public void DamagedBomb_IsRejectedByTheRecoveryPath()
    {
        var bomb = CreateValidBomb(entries: 120, megabytesEach: 10);
        var broken = BreakCentralDirectory(bomb);
        try
        {
            // Assert the SPECIFIC guard, not merely "something threw". A generic
            // ThrowsAny still passed with the inflate cap removed, which means it
            // was proving nothing about the recovery bounds.
            var ex = Assert.ThrowsAny<Exception>(() => DocumentHandlerFactory.Open(broken));
            var code = (ex as CliException)?.Code;
            Assert.True(
                code == "zip_bomb" || ex.Message.Contains("bomb", StringComparison.OrdinalIgnoreCase),
                $"expected a decompression-bomb rejection, got {ex.GetType().Name}: {ex.Message}");
        }
        finally { File.Delete(bomb); File.Delete(broken); }
    }

    [Fact]
    public void DamagedBomb_WithHugeDeclaredEntrySize_HitsThePerEntryGuard()
    {
        // Forge a local file header declaring an absurd uncompressed size. This
        // is the shape the per-entry and ratio checks exist for: the declared
        // size is attacker-controlled and is read before anything is inflated.
        var path = TempPath(".hwpx");
        using (var fs = File.Create(path))
        using (var w = new BinaryWriter(fs))
        {
            var name = "Contents/section0.xml"u8.ToArray();
            var payload = new byte[64];
            w.Write(0x04034b50u);              // local file header signature
            w.Write((ushort)20);               // version needed
            w.Write((ushort)0);                // flags
            w.Write((ushort)8);                // method: deflate
            w.Write((ushort)0); w.Write((ushort)0); // time/date
            w.Write(0u);                       // crc32
            w.Write((uint)payload.Length);     // compressed size
            w.Write(uint.MaxValue - 1);        // declared uncompressed size: absurd
            w.Write((ushort)name.Length);
            w.Write((ushort)0);                // extra length
            w.Write(name);
            w.Write(payload);
        }

        try
        {
            var ex = Assert.ThrowsAny<Exception>(() => DocumentHandlerFactory.Open(path));
            var code = (ex as CliException)?.Code;
            Assert.True(
                code == "zip_bomb" || ex.Message.Contains("bomb", StringComparison.OrdinalIgnoreCase),
                $"expected the per-entry guard to fire, got {ex.GetType().Name}: {ex.Message}");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void OversizedInput_IsRejectedBeforeAllocation()
    {
        // Drive the actual rejection path. The previous version only compared
        // two constants to each other, which is true by construction and proves
        // nothing about the guard -- false-confidence coverage.
        //
        // Build a file that exceeds MaxRecoveryInputBytes and is NOT a readable
        // zip, so open falls into recovery and must refuse before allocating it.
        var path = TempPath(".hwpx");
        var oversized = DocumentLimits.MaxRecoveryInputBytes + (1024 * 1024);
        using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            // Sparse where the filesystem allows it: set the length rather than
            // writing 257 MiB of zeros.
            fs.SetLength(oversized);
            fs.Position = 0;
            fs.Write("PK\u0003\u0004"u8);   // local header signature, truncated archive
        }

        try
        {
            var ex = Assert.ThrowsAny<Exception>(() => DocumentHandlerFactory.Open(path));
            var code = (ex as CliException)?.Code;
            Assert.True(
                code == "zip_bomb" || ex.Message.Contains("bomb", StringComparison.OrdinalIgnoreCase)
                    || ex.Message.Contains("recovery limit", StringComparison.OrdinalIgnoreCase),
                $"expected the pre-allocation input cap to fire, got {ex.GetType().Name}: {ex.Message}");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CumulativeExpansion_IsBoundedAcrossEntries()
    {
        // Many individually-acceptable entries whose SUM exceeds the aggregate
        // cap. Each declares a modest size so the per-entry guard stays quiet
        // and only the cumulative bound can stop it.
        var bomb = CreateValidBomb(entries: 200, megabytesEach: 8);
        var broken = BreakCentralDirectory(bomb);
        try
        {
            var ex = Assert.ThrowsAny<Exception>(() => DocumentHandlerFactory.Open(broken));
            var code = (ex as CliException)?.Code;
            Assert.True(
                code == "zip_bomb" || ex.Message.Contains("bomb", StringComparison.OrdinalIgnoreCase),
                $"expected a cumulative-expansion rejection, got {ex.GetType().Name}: {ex.Message}");
        }
        finally { File.Delete(bomb); File.Delete(broken); }
    }
}
