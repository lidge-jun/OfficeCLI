// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml.Linq;
using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler : IDocumentHandler
{
    private readonly HwpxDocument _doc;
    private double? _baseFontSizePt; // Plan 99.9.I3: cached base font size for heading ratio
    private readonly string _filePath;
    private readonly bool _editable;
    private readonly Stream _stream;
    private bool _dirty;
    private readonly HashSet<string> _deletedBinData = new();

    public HwpxHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _editable = editable;
        Stream? stream = null;
        ZipArchive? archive = null;
        try
        {
            stream = new FileStream(filePath, FileMode.Open,
                editable ? FileAccess.ReadWrite : FileAccess.Read,
                FileShare.ReadWrite);
            archive = new ZipArchive(stream,
                editable ? ZipArchiveMode.Update : ZipArchiveMode.Read);
            _doc = LoadDocument(archive);
            _stream = stream;
        }
        catch (InvalidDataException)
        {
            archive?.Dispose();
            stream?.Dispose();

            // Plan 99.9.I2: Broken ZIP recovery — scan for Local File Headers
            stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            try
            {
                _doc = TryRecoverBrokenZip(stream);
                _stream = stream;
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }
        catch
        {
            archive?.Dispose();
            stream?.Dispose();
            throw;
        }
    }

    private static HwpxDocument LoadDocument(ZipArchive archive)
    {
        // Plan 99.9.E1: Path traversal defense
        foreach (var entry in archive.Entries)
        {
            var name = entry.FullName;
            if (string.IsNullOrEmpty(name)) continue;
            if (name.Contains('\0') ||
                name.StartsWith('/') || name.StartsWith('\\') ||
                (name.Length >= 2 && name[1] == ':') ||
                name.Split('/', '\\').Any(seg => seg == ".."))
            {
                throw new InvalidDataException(
                    $"Suspicious ZIP entry path detected: '{name}'. " +
                    "Path traversal or absolute path entries are not allowed.");
            }
            if ((entry.ExternalAttributes & 0xF0000000) == 0xA0000000)
            {
                throw new InvalidDataException(
                    $"Symlink ZIP entry detected: '{name}'. Symlinks are not allowed.");
            }
        }

        // Plan 99.9.E2: ZIP bomb precheck
        const int MaxEntries = 1000;
        const long MaxUncompressedBytes = 200L * 1024 * 1024; // 200MB
        const double MaxCompressionRatio = 100.0;

        if (archive.Entries.Count > MaxEntries)
            throw new InvalidDataException(
                $"ZIP entry count ({archive.Entries.Count}) exceeds safety limit ({MaxEntries}).");

        long totalUncompressed = 0;
        foreach (var entry in archive.Entries)
        {
            if (entry.Length < 0 || totalUncompressed > MaxUncompressedBytes - entry.Length)
                throw new InvalidDataException(
                    $"Total uncompressed size exceeds safety limit ({MaxUncompressedBytes / (1024*1024)}MB).");
            totalUncompressed += entry.Length;
            if (entry.CompressedLength > 0)
            {
                double ratio = (double)entry.Length / entry.CompressedLength;
                if (ratio > MaxCompressionRatio)
                    throw new InvalidDataException(
                        $"ZIP entry '{entry.FullName}' has suspicious compression ratio ({ratio:F1}:1).");
            }
            else if (entry.Length > 0)
            {
                throw new InvalidDataException(
                    $"ZIP entry '{entry.FullName}' has zero compressed size but non-zero length — suspicious.");
            }
        }
        if (totalUncompressed > MaxUncompressedBytes)
            throw new InvalidDataException(
                $"Total uncompressed size ({totalUncompressed / (1024*1024)}MB) exceeds safety limit ({MaxUncompressedBytes / (1024*1024)}MB).");

        var doc = new HwpxDocument { Archive = archive };

        // Plan 80: Rootfile-aware loading via HwpxManifest
        // Tries: container.xml → rootfile → OPF manifest → conventional fallback
        var manifest = HwpxManifest.Parse(archive);
        doc.RootfilePath = manifest.RootfilePath;

        // Load manifest doc (for SaveManifest and validation)
        var manifestPath = manifest.RootfilePath ?? "Contents/content.hpf";
        var hpfEntry = archive.GetEntry(manifestPath);
        if (hpfEntry != null)
        {
            using var hpfStream = hpfEntry.Open();
            doc.ManifestDoc = LoadAndNormalize(hpfStream);
            doc.ManifestEntryPath = hpfEntry.FullName;
        }

        // Load header
        if (!string.IsNullOrEmpty(manifest.HeaderPath))
        {
            var headerEntry = archive.GetEntry(manifest.HeaderPath);
            if (headerEntry != null)
            {
                doc.HeaderEntryPath = headerEntry.FullName;
                using var stream = headerEntry.Open();
                doc.Header = LoadAndNormalize(stream);
            }
        }

        // Fallback: conventional header path
        if (doc.Header == null)
        {
            var headerEntry = archive.GetEntry("Contents/header.xml");
            if (headerEntry != null)
            {
                doc.HeaderEntryPath = headerEntry.FullName;
                using var stream = headerEntry.Open();
                doc.Header = LoadAndNormalize(stream);
            }
        }

        // Load sections from manifest-discovered paths
        int idx = 0;
        foreach (var sectionPath in manifest.SectionPaths)
        {
            var entry = archive.GetEntry(sectionPath);
            if (entry == null) continue;
            using var s = entry.Open();
            doc.Sections.Add(new HwpxSection
            {
                Index = idx++,
                EntryPath = entry.FullName,
                Document = LoadAndNormalize(s)
            });
        }

        // Fallback: try section0.xml, section1.xml, ...
        if (doc.Sections.Count == 0)
        {
            for (int i = 0; i < 100; i++)
            {
                var entry = archive.GetEntry($"Contents/section{i}.xml");
                if (entry == null) break;
                using var s = entry.Open();
                doc.Sections.Add(new HwpxSection
                {
                    Index = i,
                    EntryPath = entry.FullName,
                    Document = LoadAndNormalize(s)
                });
            }
        }

        if (doc.Sections.Count == 0)
            throw new InvalidOperationException("No sections found in HWPX document");

        return doc;
    }

    // --- Helper: read ZIP entry, normalize HWPML 2016→2011 namespaces, then parse ---
    private static XDocument LoadAndNormalize(Stream stream)
    {
        using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
        var raw = reader.ReadToEnd();
        foreach (var (old, canonical) in HwpxNs.LegacyToCanonical)
            raw = raw.Replace(old, canonical, StringComparison.Ordinal);

        // Plan 99.9.E5: XXE defense via secure parser settings
        var settings = new System.Xml.XmlReaderSettings
        {
            DtdProcessing = System.Xml.DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersFromEntities = 0
        };
        using var stringReader = new StringReader(raw);
        using var xmlReader = System.Xml.XmlReader.Create(stringReader, settings);
        return XDocument.Load(xmlReader);
    }

    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        // HWPX binary extraction not yet implemented
        return false;
    }

    // Plan 99.9.I2: Broken ZIP recovery — scan Local File Headers
    private static HwpxDocument TryRecoverBrokenZip(Stream stream)
    {
        // Copy at most <paramref name="limit"/> bytes, then fail. DeflateStream
        // will happily produce gigabytes from a few KB otherwise.
        static void CopyBounded(Stream source, Stream destination, long limit)
        {
            var buffer = new byte[81920];
            long written = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                written += read;
                if (written > limit)
                    throw new CliException(
                        "Damaged HWPX entry inflates beyond the per-entry recovery limit; "
                        + "rejected as a potential decompression bomb.")
                    { Code = "zip_bomb" };
                destination.Write(buffer, 0, read);
            }
        }

        stream.Position = 0;
        // CONSISTENCY(dos-hardening): the central-directory bomb guard cannot
        // help here — a corrupt archive is exactly what lands in this salvage
        // path, and salvage reads the whole file before it can inspect anything.
        // Bound the input BEFORE allocating it.
        if (stream.Length > DocumentLimits.MaxRecoveryInputBytes)
            throw new CliException(
                $"Cannot recover {stream.Length / (1024 * 1024)} MiB damaged HWPX: exceeds the "
                + $"{DocumentLimits.MaxRecoveryInputBytes / (1024 * 1024)} MiB recovery limit.")
            {
                Code = "zip_bomb",
                Suggestion = "Repair the file with Hancom Office, or extract the parts you need separately."
            };

        var data = new byte[stream.Length];
        stream.ReadExactly(data);

        const uint LocalFileHeader = 0x04034b50;
        var recovered = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        long totalUncompressed = 0;

        int pos = 0;
        while (pos + 30 < data.Length)
        {
            uint sig = BitConverter.ToUInt32(data, pos);
            if (sig != LocalFileHeader) { pos++; continue; }

            ushort compMethod = BitConverter.ToUInt16(data, pos + 8);
            uint compSize = BitConverter.ToUInt32(data, pos + 18);
            uint uncompSize = BitConverter.ToUInt32(data, pos + 22);
            ushort nameLen = BitConverter.ToUInt16(data, pos + 26);
            ushort extraLen = BitConverter.ToUInt16(data, pos + 28);

            // Entry count: a salvage scan can otherwise walk millions of headers.
            if (recovered.Count >= DocumentLimits.MaxZipEntries)
                throw new CliException(
                    $"Damaged HWPX declares more than {DocumentLimits.MaxZipEntries} entries; "
                    + "rejected as a potential decompression bomb.")
                { Code = "zip_bomb" };

            // The declared sizes are attacker-controlled, so check them before
            // they are used for arithmetic or allocation. compSize is later cast
            // to int for the span, so it must fit as well.
            if (uncompSize > DocumentLimits.MaxPerEntryUncompressedBytes
                || compSize > int.MaxValue
                || (compSize > 0 && uncompSize / Math.Max(1u, compSize) > DocumentLimits.MaxCompressionRatio))
                throw new CliException(
                    "Damaged HWPX contains an entry whose declared size or compression ratio "
                    + "exceeds safe limits; rejected as a potential decompression bomb.")
                { Code = "zip_bomb" };

            int headerEnd = pos + 30 + nameLen + extraLen;
            if (headerEnd + compSize > data.Length) break;

            var entryName = System.Text.Encoding.UTF8.GetString(data, pos + 30, nameLen);
            var compData = data.AsSpan(headerEnd, (int)compSize);

            try
            {
                byte[] entryData;
                if (compMethod == 0) // STORED
                {
                    entryData = compData.ToArray();
                }
                else if (compMethod == 8) // DEFLATE
                {
                    using var compStream = new System.IO.Compression.DeflateStream(
                        new MemoryStream(compData.ToArray()),
                        System.IO.Compression.CompressionMode.Decompress);
                    using var outStream = new MemoryStream();
                    // Bounded copy: the declared uncompressed size is a hint, not
                    // a promise, so cap what we will actually inflate.
                    CopyBounded(compStream, outStream, DocumentLimits.MaxPerEntryUncompressedBytes);
                    entryData = outStream.ToArray();
                }
                else
                {
                    pos = headerEnd + (int)compSize;
                    continue;
                }

                // Many individually-acceptable entries can still exhaust memory
                // in aggregate.
                totalUncompressed += entryData.LongLength;
                if (totalUncompressed > DocumentLimits.MaxRecoveryTotalUncompressedBytes)
                    throw new CliException(
                        "Damaged HWPX expands beyond the total recovery limit; rejected as a "
                        + "potential decompression bomb.")
                    { Code = "zip_bomb" };

                if (!recovered.ContainsKey(entryName))
                    recovered[entryName] = entryData;
            }
            catch { /* skip unreadable entry */ }

            pos = headerEnd + (int)compSize;
        }

        if (!recovered.Keys.Any(k => k.Contains("section", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException(
                "Broken ZIP recovery failed: no section XML found in recovered entries.");

        // Rebuild as in-memory ZIP for the standard loader
        var memStream = new MemoryStream();
        using (var newZip = new ZipArchive(memStream, ZipArchiveMode.Create, true))
        {
            foreach (var (name, bytes) in recovered)
            {
                var entry = newZip.CreateEntry(name);
                using var s = entry.Open();
                s.Write(bytes);
            }
        }

        memStream.Position = 0;
        var archive = new ZipArchive(memStream, ZipArchiveMode.Read);
        return LoadDocument(archive);
    }

    public void Save()
    {
        if (!_dirty || !_editable) return;
        _doc.Archive.Dispose();
        _stream.Flush();
        _dirty = false;
    }

    public void Dispose()
    {
        _doc.Archive.Dispose();
        _stream.Dispose();
    }
}
