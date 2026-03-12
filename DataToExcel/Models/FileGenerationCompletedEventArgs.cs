namespace DataToExcel.Models;

public sealed class FileGenerationCompletedEventArgs : EventArgs
{
    public FileGenerationCompletedEventArgs(string fileName, int fileIndex, BlobUploadResult uploadResult)
    {
        FileName = fileName;
        FileIndex = fileIndex;
        UploadResult = uploadResult;
    }

    public string FileName { get; }
    public int FileIndex { get; }
    public BlobUploadResult UploadResult { get; }
}
