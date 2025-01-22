namespace CustomDocWriter.DocumentWriter;

public interface IDocumentWriter<TDocument>
{
    void RegisterDocumentDataWriter(IDocumentDataWriter<TDocument> writer);
    Task<MemoryStream> WriteDocumentAsync(Stream source);
}