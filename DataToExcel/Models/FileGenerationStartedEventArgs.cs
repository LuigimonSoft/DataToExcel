namespace DataToExcel.Models;

public sealed class FileGenerationStartedEventArgs : EventArgs
{
    public FileGenerationStartedEventArgs(string fileName, int fileIndex)
    {
        FileName = fileName;
        FileIndex = fileIndex;
    }

    public string FileName { get; }
    public int FileIndex { get; }
}
