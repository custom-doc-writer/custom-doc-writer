namespace CustomDocWriter.DocumentWriter;

public interface IDocumentDataWriter<TDocument>
{
    Task WriteDocumentDataAsync(TDocument document);
}