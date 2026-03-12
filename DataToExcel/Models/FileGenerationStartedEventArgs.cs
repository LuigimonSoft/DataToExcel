namespace DataToExcel.Models;

public sealed class FileGenerationStartedEventArgs : EventArgs
{
    public FileGenerationStartedEventArgs(string blobName, int fileIndex)
    {
        BlobName = blobName;
        FileIndex = fileIndex;
    }

    public string BlobName { get; }
    public int FileIndex { get; }
}
