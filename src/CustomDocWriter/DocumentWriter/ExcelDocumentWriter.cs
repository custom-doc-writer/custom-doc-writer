using OfficeOpenXml;

namespace CustomDocWriter.DocumentWriter;

public class ExcelDocumentWriter : DocumentWriterBase<ExcelPackage>
{
    protected override async Task<ExcelPackage> CreateDocumentAsync(Stream source)
    {
        var document = new ExcelPackage();
        await document.LoadAsync(source);

        return document;
    }

    protected override Task<MemoryStream> SaveDocumentAsync(ExcelPackage document)
    {
        var memoryStream = new MemoryStream();
        document.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);

        return Task.FromResult(memoryStream);
    }
}