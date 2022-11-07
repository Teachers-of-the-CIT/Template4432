namespace Template4432.Interfaces
{
    public interface IJsonDataService
    {
        (bool success, int count) ImportJsonData(string json);
    }
}