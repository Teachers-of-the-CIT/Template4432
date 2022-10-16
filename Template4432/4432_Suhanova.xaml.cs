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
    /// Логика взаимодействия для _4432_Suhanova.xaml
    /// </summary>
    public partial class _4432_Suhanova : Window
    {
        public _4432_Suhanova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true)) return;

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

            using (isrpo_lr2Entities db = new isrpo_lr2Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.data.Add(new data()
                    {
                        id = int.Parse(list[i, 0]),
                        name_service = list[i, 1],
                        kind_service = list[i, 2],
                        id_service = list[i, 3],
                        cost = int.Parse(list[i, 4])
                    });
                }
                db.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            var listData = new List<data>();

            using (isrpo_lr2Entities db = new isrpo_lr2Entities())
            {
                listData = db.data.ToList().OrderBy(x => x.kind_service).ToList();
            }
            var allKindServise = listData.GroupBy(x => x.kind_service).ToList();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allKindServise.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allKindServise.Count(); i++) 
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allKindServise[i].Key);
                worksheet.Cells[1][startRowIndex] = "id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                startRowIndex++;
                foreach (var item in listData)
                {
                    if (item.kind_service == allKindServise[i].Key)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1],
                        worksheet.Cells[2][1]];
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        worksheet.Cells[1][startRowIndex] = item.id;
                        worksheet.Cells[2][startRowIndex] = item.name_service;
                        worksheet.Cells[3][startRowIndex] = item.cost;
                        startRowIndex++;
                    }
                    else
                    {
                        continue;
                    }
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
