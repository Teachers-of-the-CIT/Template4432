using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Template4432.Application.Base;
using Template4432.Contexts;
using Template4432.Enums;
using Template4432.Interfaces;
using Template4432.Models;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Template4432.Application
{
    public class SkiServiceService : EntityService<SkiService>, IExcelDataService, IJsonDataService, IWordDataService
    {
        private readonly ExcelApplication _excel;

        private readonly Dictionary<string, int> _columnsImport = new Dictionary<string, int>()
        {
            {"ID", 0},
            {"Наименование услуги", 1},
            {"Вид услуги", 2},
            {"Код услуги", 3},
            {"Стоимость, руб.  за час", 4}
        };

        private readonly Expression<Func<SkiService, SkiServiceType>> _skiServiceTypeSelector = service => service.ServiceType;

        public SkiServiceService(ApplicationContext context, ExcelApplication excel) : base(context)
        {
            _excel = excel;
        }

        public void LoadWorkbook(string fileName)
        {
            _excel.Workbooks.Open(fileName);
        }

        public (bool, int) ImportEntitiesFromWorkbook(string fileName)
        {
            LoadWorkbook(fileName);

            Worksheet worksheet = _excel.Worksheets[1];

            List<SkiService> skiServices = new List<SkiService>();

            Range lastCell = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

            int columnsCount = lastCell.Column;
            int rowCount = lastCell.Row;

            string[,] rawCells = new string[rowCount, columnsCount];

            for (var j = 0; j < columnsCount; j++)
            {
                for (var i = 0; i < rowCount; i++)
                {
                    rawCells[i, j] = worksheet.Cells[i + 1, j + 1].Text();
                }
            }

            _excel.Workbooks[1].Close(false, Type.Missing, Type.Missing);
            _excel.Quit();

            for (int row = 1; row < rowCount; row++)
            {
                try
                {
                    int id = int.Parse(rawCells[row, _columnsImport["ID"]]);
                    string serviceName = rawCells[row, _columnsImport["Наименование услуги"]];
                    string serviceType = rawCells[row, _columnsImport["Вид услуги"]];
                    string serviceCode = rawCells[row, _columnsImport["Код услуги"]];
                    decimal price = decimal.Parse(rawCells[row, _columnsImport["Стоимость, руб.  за час"]]);

                    SkiService skiService = new SkiService(id, serviceName, serviceCode, serviceType, price);

                    skiServices.Add(skiService);
                }
                catch
                {
                    return (false, 0);
                }
            }

            return AddToDatabase(skiServices);
        }

        public Workbook ExportEntities()
        {
            List<IGrouping<SkiServiceType, SkiService>> skiServices = ReadAsQueryable()
                .OrderBy(service => service.PriceForHour)
                .GroupBy(service => service.ServiceType)
                .ToList();

            Workbook workbook = _excel.Workbooks.Add();
            
            foreach (IGrouping<SkiServiceType,SkiService> servicesByType in skiServices)
            {
                Worksheet worksheet = (Worksheet) workbook.Worksheets.Add();
                worksheet.Name = servicesByType.Key.ToString();
                
                worksheet.Cells[1, 1] = "Id";
                worksheet.Cells[1, 2] = "Название услуги";
                worksheet.Cells[1, 3] = "Стоимость";
                
                worksheet.Cells[1, 1].Font.Bold = true;
                worksheet.Cells[1, 2].Font.Bold = true;
                worksheet.Cells[1, 3].Font.Bold = true;
                
                int index = 2;
                foreach (SkiService skiService in servicesByType)
                {
                    worksheet.Cells[index, 1] = skiService.Id.ToString();
                    worksheet.Cells[index, 2] = skiService.ServiceName;
                    worksheet.Cells[index, 3] = skiService.PriceForHour;
                    
                    index++;
                }

                worksheet.Columns.AutoFit();
            }

            _excel.Visible = true;

            return workbook;
        }

        private (bool, int) AddToDatabase(List<SkiService> services)
        {
            try
            {
                _dbSet.AddRange(services);

                _context.SaveChanges();

                return (true, services.Count);
            }
            catch
            {
                return (false, 0);
            }
        }
        
        public (bool, int) ImportJsonData(string json)
        {
            SkiService[] skiServices = JsonConvert.DeserializeObject<SkiService[]>(json);

            return AddToDatabase(skiServices?.ToList());
        }

        public Document ExportToWord()
        {
            List<IGrouping<SkiServiceType, SkiService>> skiServicesGrouped = ReadAsQueryable()
                .OrderBy(service => service.PriceForHour)
                .GroupBy(service => service.ServiceType)
                .ToList();

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document document = word.Documents.Add();
            
            foreach (IGrouping<SkiServiceType, SkiService> skiServiceByType in skiServicesGrouped)
            {
                List<SkiService> skiServices = skiServiceByType.ToList();
                
                Paragraph paragraph = document.Paragraphs.Add();
                paragraph.set_Style("Заголовок 1");

                var range = paragraph.Range;
                range.Text = skiServiceByType.Key.ToString();

                range.InsertParagraphAfter();

                #region Таблица
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;

                var ordersTable = document.Tables.Add(tableRange, skiServices.Count(), 3);
                ordersTable.Borders.InsideLineStyle =WdLineStyle.wdLineStyleSingle;
                ordersTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

                ordersTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                #region Заголовок таблицы
                ordersTable.Rows[1].Range.Bold = 1;
                ordersTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                ordersTable.Cell(1, 1).Range.Text = "Id";
                ordersTable.Cell(1, 2).Range.Text = "Наименование услуги";
                ordersTable.Cell(1, 3).Range.Text = "Стоимость";
                #endregion

                #region Заполнение таблицы
                var row = 2;
                foreach (var order in skiServices)
                {
                    ordersTable.Cell(row, 1).Range.Text = order.Id.ToString();
                    ordersTable.Cell(row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 2).Range.Text = order.ServiceName.ToString();
                    ordersTable.Cell(row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 3).Range.Text = order.PriceForHour.ToString();
                    ordersTable.Cell(row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    row++;
                }
                #endregion
                #endregion

                document.Words.Last.InsertBreak(WdBreakType.wdSectionBreakNextPage);
            }

            word.Visible = true;

            return document;
        }
    }
}