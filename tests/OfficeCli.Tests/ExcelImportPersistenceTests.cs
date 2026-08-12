using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli;

namespace OfficeCli.Tests;

public sealed class ExcelImportPersistenceTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(
        Path.GetTempPath(), $"officecli-import-tests-{Guid.NewGuid():N}");

    public ExcelImportPersistenceTests() => Directory.CreateDirectory(_tempDir);

    [Fact]
    public void Import_PersistsParsedCellsAfterTheCliHandlerIsDisposed()
    {
        var workbook = Path.Combine(_tempDir, "persisted.xlsx");
        var source = Path.Combine(_tempDir, "source.csv");
        BlankDocCreator.Create(workbook, locale: "en-US");
        File.WriteAllText(source, "a,b\nx,y\n", new UTF8Encoding(false));

        var (exitCode, stdout) = Invoke(
            ["import", workbook, "/Sheet1", source, "--header", "--json"]);

        Assert.Equal(0, exitCode);
        var response = JsonNode.Parse(stdout)!;
        Assert.True(response["success"]!.GetValue<bool>());
        Assert.Contains("4 cells affected", response["message"]!.GetValue<string>());

        using var document = SpreadsheetDocument.Open(workbook, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet
            ?? throw new InvalidDataException("Worksheet root is missing");
        var cells = (worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidDataException("SheetData is missing"))
            .Descendants<Cell>()
            .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellValue?.Text);

        Assert.Equal(4, cells.Count);
        Assert.Equal("a", cells["A1"]);
        Assert.Equal("b", cells["B1"]);
        Assert.Equal("x", cells["A2"]);
        Assert.Equal("y", cells["B2"]);
    }

    [Fact]
    public void Import_DoesNotReportSuccessWhenNoCellsWereAffected()
    {
        var workbook = Path.Combine(_tempDir, "empty.xlsx");
        var source = Path.Combine(_tempDir, "empty.csv");
        BlankDocCreator.Create(workbook, locale: "en-US");
        File.WriteAllText(source, ",\n,\n", new UTF8Encoding(false));

        var (exitCode, stdout) = Invoke(
            ["import", workbook, "/Sheet1", source, "--json"]);

        Assert.Equal(1, exitCode);
        var response = JsonNode.Parse(stdout)!;
        Assert.False(response["success"]!.GetValue<bool>());
        Assert.Contains("produced no cells", response["error"]!["error"]!.GetValue<string>());

        using var document = SpreadsheetDocument.Open(workbook, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet
            ?? throw new InvalidDataException("Worksheet root is missing");
        var cells = (worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidDataException("SheetData is missing"))
            .Descendants<Cell>();
        Assert.Empty(cells);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    private static (int ExitCode, string Stdout) Invoke(string[] args)
    {
        var originalOut = Console.Out;
        var originalError = Console.Error;
        using var output = new StringWriter();
        using var error = new StringWriter();
        Console.SetOut(output);
        Console.SetError(error);
        try
        {
            var exitCode = CommandBuilder.BuildRootCommand().Parse(args).Invoke();
            return (exitCode, output.ToString() + error.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
    }
}
