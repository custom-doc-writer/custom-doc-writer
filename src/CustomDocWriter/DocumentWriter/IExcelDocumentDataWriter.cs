using OfficeOpenXml;

namespace CustomDocWriter.DocumentWriter;

public interface IExcelDocumentDataWriter : IDocumentDataWriter<ExcelPackage>
{
}