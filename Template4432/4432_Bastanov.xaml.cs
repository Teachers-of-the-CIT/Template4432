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
    /// Логика взаимодействия для _4432_Bastanov.xaml
    /// </summary>  
    public partial class _4432_Bastanov : Window
    {
        public _4432_Bastanov()
        {
            InitializeComponent();
        }
        private const int _sheetsCount = 3;

        private void Imp_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (2.xlsx)|*.xlsx",
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
            using (MacroSocietyEntities macroSocietyEntities = new MacroSocietyEntities())
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
                    { 
                        continue; 
                    }
                    macroSocietyEntities.Prokat_Bastanov.Add(new Prokat_Bastanov()
                    {
                        Code_order = list[i, 1],
                        Data_order = list[i, 2],
                        Time_oder = list[i, 3],
                        Code_client = list[i, 4],
                        Service = list[i, 5],
                        Status = list[i, 6],
                        Data_close = list[i, 7],
                        Time_prokat = list[i, 8]
                    });
                }
                macroSocietyEntities.SaveChanges();
            }
        }

        private void Exp_Click(object sender, RoutedEventArgs e)
        {           
            List<Prokat_Bastanov> allProkat;
            using (MacroSocietyEntities usersEntities = new MacroSocietyEntities())
            {
                allProkat =
                usersEntities.Prokat_Bastanov.ToList().OrderBy(s =>
                s.Status).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            var statusvision = allProkat
                        .OrderBy(o => o.Status)
                        .GroupBy(s => s.Status)
                        .ToDictionary(g => g.Key, g => g.Select(s => new { s.Id, s.Code_order, s.Data_order, s.Code_client, s.Service, s.Status })
                        .ToArray());
            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.Item[i + 1];
                worksheet.Name =
                $"Статус {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;
                var data = i == 0 ? statusvision.Where(w => w.Key.Equals("Новая"))
                : i == 1 ? statusvision.Where(w => w.Key.Equals("В прокате")) : i == 2 ? statusvision.Where(w => w.Key.Equals("Закрыта")) : statusvision;

                foreach (var Status in data)
                {
                    foreach (var St in Status.Value)
                    {
                        if (St.Status == Status.Key)
                        {
                            worksheet.Cells[1][startRowIndex] = St.Id;
                            worksheet.Cells[2][startRowIndex] = St.Code_order;
                            worksheet.Cells[3][startRowIndex] = St.Data_order;
                            worksheet.Cells[4][startRowIndex] = St.Code_client;
                            worksheet.Cells[5][startRowIndex] = St.Service;
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


