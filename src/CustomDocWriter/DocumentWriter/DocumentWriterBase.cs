namespace CustomDocWriter.DocumentWriter;

public abstract class DocumentWriterBase<TDocument> : IDocumentWriter<TDocument>
{
    private readonly List<IDocumentDataWriter<TDocument>> _writers = [];

    public void RegisterDocumentDataWriter(IDocumentDataWriter<TDocument> writer)
    {
        _writers.Add(writer);
    }

    public async Task<MemoryStream> WriteDocumentAsync(Stream source)
    {
        var document = await CreateDocumentAsync(source);
        foreach (var writer in _writers)
        {
            await writer.WriteDocumentDataAsync(document);
        }

        return await SaveDocumentAsync(document);
    }

    protected abstract Task<TDocument> CreateDocumentAsync(Stream source);
    protected abstract Task<MemoryStream> SaveDocumentAsync(TDocument document);
}