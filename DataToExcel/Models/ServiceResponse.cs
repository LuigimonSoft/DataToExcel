namespace DataToExcel.Models;

public class ServiceResponse<T> where T : class
{
    public T? Data { get; set; }
    public string? ErrorMessage { get; set; }
    public bool IsSuccess { get; set; }

    public ServiceResponse() { }
    public ServiceResponse(T data)
    {
        Data = data;
        IsSuccess = true;
    }
}
