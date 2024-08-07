namespace OfficeForge;

public interface IDocumentReader<TModel>
{
    TModel Read(Stream stream);
    TModel Read(string path);
}

public interface IDocumentWriter<TModel>
{
    void Write(TModel model, Stream stream);
    void Write(TModel model, string path);
}
