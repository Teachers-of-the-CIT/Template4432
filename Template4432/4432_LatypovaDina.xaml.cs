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
using System.Windows.Navigation;
using System.Windows.Shapes;

using Excel = Microsoft.Office.Interop.Excel;
namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_LatypovaDina.xaml
    /// </summary>
    public partial class _4432_LatypovaDina : System.Windows.Window
    {
        public _4432_LatypovaDina()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (2.xlsx)|*.xlsx",
                Title = "Выберите файл БД"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application objWorkExcel = new Excel.Application();
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
            var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                }

            }
            objWorkBook.Close(false, Type.Missing, Type.Missing);
            objWorkExcel.Quit();
            GC.Collect();
            using (ISRPOEntities2 iSRPOEntities = new ISRPOEntities2())
            {
                for (int i = 0; i < _rows; i++)
                {
                    int nullColumn = 0;
                    for (int j = 0; j < _columns; j++)
                    {
                        if (String.IsNullOrEmpty(list[i, j]))
                        {
                            nullColumn++;
                        }
                    }
                    if (nullColumn == _columns)
                    {
                        continue;
                    }
                    iSRPOEntities.Zakaz.Add(new Zakaz()
                    {
                        Kod_zakaza = list[i, 1],
                        Data_zakaza = list[i, 2],
                        Vremya = list[i, 3],
                        Kod_klienta = list[i, 4],
                        Uslugi = list[i, 5],
                        Status = list[i, 6],
                        Data_zakriritiya = list[i, 7],
                        Vremya_prokata = list[i, 8]
                    });
                }
                iSRPOEntities.SaveChanges();
            }
            
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Zakaz> zakaz;
            using (ISRPOEntities2 iSRPOEntities2 = new ISRPOEntities2())
            {
                zakaz = iSRPOEntities2.Zakaz.ToList().OrderBy(s => s.Data_zakaza).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = zakaz.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            var vivod = zakaz
                        .OrderBy(o => o.Data_zakaza)
                        .GroupBy(s => s.Data_zakaza)
                        .ToDictionary(g => g.Key, g => g.Select(s => new { s.ID, s.Kod_zakaza, s.Kod_klienta, s.Uslugi, s.Data_zakaza })
                        .ToArray());
            for (int i = 0; i < zakaz.Count; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.Item[i + 1];
                worksheet.Name =
                $"Статус {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Код клиента";
                worksheet.Cells[4][startRowIndex] = "Услуги";
                startRowIndex++;
                var data = i == 0 ? vivod.Where(w => w.Key.Equals("12.03.2022"))
                : i == 1 ? vivod.Where(w => w.Key.Equals("21.03.2022")) : i == 2 ? vivod.Where(w => w.Key.Equals("09.04.2022")) : vivod;
                foreach (var Data_zakaza in data)
                {
                    foreach (var DZ in Data_zakaza.Value)
                    {
                        if (DZ.Data_zakaza == Data_zakaza.Key)
                        {
                            worksheet.Cells[1][startRowIndex] = DZ.ID;
                            worksheet.Cells[2][startRowIndex] = DZ.Kod_zakaza;
                            worksheet.Cells[3][startRowIndex] = DZ.Kod_klienta;
                            worksheet.Cells[4][startRowIndex] = DZ.Uslugi;
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
