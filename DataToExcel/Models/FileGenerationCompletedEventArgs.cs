namespace DataToExcel.Models;

public sealed class FileGenerationCompletedEventArgs : EventArgs
{
    public FileGenerationCompletedEventArgs(string blobName, int fileIndex, BlobUploadResult uploadResult)
    {
        BlobName = blobName;
        FileIndex = fileIndex;
        UploadResult = uploadResult;
    }

    public string BlobName { get; }
    public int FileIndex { get; }
    public BlobUploadResult UploadResult { get; }
}
