using System.Text.RegularExpressions;
using CustomDocWriter.DocumentWriter;
using Microsoft.AspNetCore.Components;
using OfficeOpenXml;

namespace CustomDocWriter.Components.ExcelDocumentFields;

public abstract partial class ExcelDocumentFieldBase : ComponentBase, IExcelDocumentDataWriter
{
    [Parameter]
    public required string FieldId { get; set; }

    [Parameter]
    public required string FieldLabel { get; set; }

    [Parameter]
    public required string FieldName { get; set; }

    [Parameter]
    public required bool FieldDisabled { get; set; }

    [Parameter]
    public required ExcelDocumentWriter Writer { get; set; }

    [Parameter]
    public required int SheetIndex { get; set; }

    protected override void OnAfterRender(bool firstRender)
    {
        base.OnAfterRender(firstRender);

        if (firstRender)
        {
            Writer.RegisterDocumentDataWriter(this);
        }
    }

    public async Task WriteDocumentDataAsync(ExcelPackage document)
    {
        var sheetIndex = document.Compatibility.IsWorksheets1Based ? SheetIndex + 1 : SheetIndex;
        var sheet = document.Workbook.Worksheets[sheetIndex];

        await WriteAsync(sheet);
    }

    protected abstract Task WriteAsync(ExcelWorksheet sheet);

    protected static int GetExcelColumnIndexFromLabel(string columnLabel)
    {
        int index = 0;
        foreach (char c in columnLabel)
        {
            index = index * 26 + char.ToUpper(c) - 'A' + 1;
        }
        return index;
    }

    protected static string GetExcelColumnLabel(int columnIndex)
    {
        string result = string.Empty;

        while (columnIndex > 0)
        {
            columnIndex--; // Adjust because the index is 1-based.
            result = (char)('A' + (columnIndex % 26)) + result;
            columnIndex /= 26;
        }

        return result;
    }

    protected static (int, int) ParseCell(string cell)
    {
        var matches = CellRegex().Match(cell);
        if (!matches.Success)
            throw new InvalidOperationException("Invalid cell parameter.");

        var x = GetExcelColumnIndexFromLabel(matches.Groups[1].Value);
        var y = int.Parse(matches.Groups[2].Value);
        return (x, y);
    }

    [GeneratedRegex("([a-zA-Z]+)([0-9]+)")]
    private static partial Regex CellRegex();
}