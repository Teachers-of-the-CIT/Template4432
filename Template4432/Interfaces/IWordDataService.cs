using Microsoft.Office.Interop.Word;

namespace Template4432.Interfaces
{
    public interface IWordDataService
    {
        Document ExportToWord();
    }
}