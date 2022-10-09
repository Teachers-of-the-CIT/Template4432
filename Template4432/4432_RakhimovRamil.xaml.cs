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
    /// Логика взаимодействия для _4432_RakhimovRamil.xaml
    /// </summary>
    public partial class _4432_RakhimovRamil : Window
    {
        public _4432_RakhimovRamil()
        {
            InitializeComponent();
        }

        private void Button_Import_Click(object sender, RoutedEventArgs e)
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using(var db = new ISRPODBEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    int nullCollumns = 0;
                    for (int j = 0; j < _columns; j++)
                    {
                        if (String.IsNullOrEmpty(list[i, j]))
                            nullCollumns++;
                    }
                    if (nullCollumns == _columns)
                        continue;
                    db.Client.Add(new Client()
                    {
                        FIO = list[i, 0],
                        UserId = Convert.ToInt32(list[i, 1]),
                        BirthDate = list[i, 2],
                        Index = Convert.ToInt32(list[i, 3]),
                        City = list[i, 4],
                        Street = list[i, 5],           
                        House = Convert.ToInt32(list[i, 6]),
                        Apartment = Convert.ToInt32(list[i, 7]),
                        Email = list[i, 8]
                    }); 
                }
                db.SaveChanges();
            }
        }

        private void Button_Export_Click(object sender, RoutedEventArgs e)
        {
            List<Client> allClients;
            List<string> allStreets;
            using (var isrpoEntities = new ISRPODBEntities())
            {
                allClients = (from c in isrpoEntities.Client
                              orderby c.FIO
                              select c).ToList();
                allStreets = (from s in allClients
                              group s by s.Street into g
                              select g.Key).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStreets.Count;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allStreets.Count; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allStreets[i];
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "E-mail";
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;
                foreach(Client client in allClients)
                {
                    if (client.Street != allStreets[i])
                        continue;
                    worksheet.Cells[1][startRowIndex] = client.UserId;
                    worksheet.Cells[2][startRowIndex] = client.FIO;
                    worksheet.Cells[3][startRowIndex] = client.Email;
                    startRowIndex++;
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];             
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
