// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Standardized validation error/warning codes (aligned with kordoc v2.2.6).
/// </summary>
/// <remarks>
/// Kept in its own file rather than inside IDocumentHandler.cs so the HWPX
/// layer adds a file instead of growing an upstream one — the additive-only
/// rule that keeps this fork mergeable.
///
/// STATUS (settled in wp6): these constants are NOT the codes HWPX emits, and
/// they should not become them. HwpxHandler.Validate emits lowercase snake_case
/// ("bindata_missing"), and tests assert those exact strings — see
/// HwpxValidationTests, which checks ErrorType == "package_version_missing" and
/// "bindata_orphan". Substituting the uppercase constants would change the
/// observable error_type on the CLI/JSON surface and break those assertions.
///
/// So the wp3 note ("wire these in wp6 behind golden fixtures") resolved the
/// other way once the fixtures existed to check against: the emitted contract
/// wins, and these remain a kordoc-alignment reference for cross-tool mapping.
/// Deleting them would lose that mapping; renaming the emitted codes would
/// break callers. Documented rather than silently kept.
///
/// Severity likewise stays default-Error at every construction site.
/// HwpxPackageValidator.IsPackageBlocking classifies by error type instead,
/// special-casing package_version_missing and non-strict bindata_orphan. The
/// severity property is honored where read; it is not doing classification work.
/// </remarks>
public static class ValidationCodes
{
    // Errors (critical — document may not open correctly)
    public const string Encrypted = "ENCRYPTED";
    public const string DrmProtected = "DRM_PROTECTED";
    public const string ZipBomb = "ZIP_BOMB";
    public const string Corrupted = "CORRUPTED";
    public const string NoSections = "NO_SECTIONS";
    public const string ZipEmpty = "ZIP_EMPTY";
    public const string ZipCorrupt = "ZIP_CORRUPT";
    public const string OpfMissing = "OPF_MISSING";
    public const string XmlMalformed = "XML_MALFORMED";
    public const string IdRefOrphan = "IDREF_ORPHAN";
    public const string TableStructure = "TABLE_STRUCTURE";
    public const string BinDataMissing = "BINDATA_MISSING";
    public const string BinDataOrphan = "BINDATA_ORPHAN";
    public const string FieldPairMismatch = "FIELD_PAIR_MISMATCH";
    public const string SectionMismatch = "SECTION_MISMATCH";

    // Warnings (non-critical — document opens but may have issues)
    public const string TruncatedTable = "TRUNCATED_TABLE";
    public const string MalformedXml = "MALFORMED_XML_MINOR";
    public const string PartialParse = "PARTIAL_PARSE";
    public const string NamespaceMissing = "NAMESPACE_MISSING";
    public const string NamespaceMismatch = "NAMESPACE_MISMATCH";
    public const string StaleIdRef = "STALE_IDREF";
    public const string EmptySection = "EMPTY_SECTION";
    public const string LargeFile = "LARGE_FILE";
    public const string DeprecatedElement = "DEPRECATED_ELEMENT";
    public const string MergedCellOverlap = "MERGED_CELL_OVERLAP";
}
