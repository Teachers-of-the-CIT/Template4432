using Microsoft.Office.Interop.Excel;
using Template4432.Models.Base;

namespace Template4432.Interfaces
{
    public interface IExcelDataService
    {
        void LoadWorkbook(string fileName);
        (bool success, int count) ImportEntitiesFromWorkbook(string fileName);
        Workbook ExportEntities();
    }
}