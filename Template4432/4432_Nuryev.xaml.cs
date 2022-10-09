using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Nuryev.xaml
    /// </summary>
    public partial class _4432_Nuryev : System.Windows.Window
    {
        public _4432_Nuryev()
        {
            InitializeComponent();
        }
        private const int _sheetsCount = 6;
        private void ExcelImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++) 
            { 
                for (int i = 0; i < _rows; i++) 
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (LR2ISRPOEntities lr2isrpoEntities = new LR2ISRPOEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    int nullColumn = 0;
                    for (int j = 0; j < _columns; j++)
                    {
                        if (String.IsNullOrEmpty(list[i, j]))
                            nullColumn++;
                    }
                    if (nullColumn == _columns)
                    { continue; }
                    lr2isrpoEntities.Table.Add(new Table()
                    {
                        OrderCode = list[i, 1],
                        CreatingDate = list[i, 2],
                        Time = list[i, 3],
                        ClientCode = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        DateOfClosing = list[i, 7],
                        RentTime = list[i, 8]
                    });
                }
                lr2isrpoEntities.SaveChanges();
            }
        }

        private void ExcelExport_Click(object sender, RoutedEventArgs e)
        {
            List<Table> Tables;
            using (LR2ISRPOEntities lr2isrpoEntities = new LR2ISRPOEntities())
            {
                Tables = lr2isrpoEntities.Table.ToList()
                        .OrderBy(s => s.Time)
                        .ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            var timedevision = Tables
                        .OrderBy(o => o.RentTime)
                        .GroupBy(s => s.RentTime)
                        .ToDictionary(g => g.Key, g => g.Select(s => new { s.id, s.OrderCode, s.CreatingDate, s.ClientCode, s.Services, s.RentTime })
                        .ToArray());
            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                var worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Врем.пр. {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;

                var data = i == 0 ? timedevision.Where(w => w.Key.Equals("120 минут") || w.Key.Equals("2 часа"))
                : i == 1 ? timedevision.Where(w => w.Key.Equals("240 минут") || w.Key.Equals("4 часа")) : i == 2 ? timedevision.Where(w => w.Key.Equals("360 минут") || w.Key.Equals("6 часов"))
                : i == 3 ? timedevision.Where(w => w.Key.Equals("480 минут") || w.Key.Equals("8 часов")) : i == 4 ? timedevision.Where(w => w.Key.Equals("600 минут") || w.Key.Equals("10 часов"))
                : i == 5 ? timedevision.Where(w => w.Key.Equals("720 минут") || w.Key.Equals("12 часов")) : timedevision;

                foreach (var Times in data)
                {
                    foreach (var Devision in Times.Value)
                    {
                        if (Devision.RentTime == Times.Key)
                        {
                            worksheet.Cells[1][startRowIndex] = Devision.id;
                            worksheet.Cells[2][startRowIndex] = Devision.OrderCode;
                            worksheet.Cells[3][startRowIndex] = Devision.CreatingDate;
                            worksheet.Cells[4][startRowIndex] = Devision.ClientCode;
                            worksheet.Cells[5][startRowIndex] = Devision.Services;
                            startRowIndex++;
                        }
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
