namespace DataToExcel.Models;

public class RepositoryResponse<T> where T : class
{
    public T? Data { get; set; }
    public string? ErrorMessage { get; set; }
    public bool IsSuccess { get; set; }

    public RepositoryResponse() { }
    public RepositoryResponse(T data)
    {
        Data = data;
        IsSuccess = true;
    }
}
