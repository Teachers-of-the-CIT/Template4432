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
    /// Interaction logic for _4432_Latypov.xaml
    /// </summary>
    public partial class _4432_Latypov : System.Windows.Window
    {
        private const int _sheetsCount = 3;
        public _4432_Latypov()
        {
            InitializeComponent();
        }

        //private void Import_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog ofd = new OpenFileDialog()
        //    {
        //        DefaultExt = "*.xls;*.xlsx",
        //        Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
        //        Title = "Выберите файл базы данных"
        //    };
        //    if (!(ofd.ShowDialog() == true))
        //        return;
        //    lock (ofd) // lock the thread against "crirtical section"
        //    {
        //        string[,] list;

        //        var ObjWorkExcel = new Excel.Application();

        //        var ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);

        //        var ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

        //        var lastCell = ObjWorkSheet
        //            .Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

        //        var _columns = lastCell.Column;

        //        var _rows = lastCell.Row;

        //        list = new string[_rows, _columns];

        //        for (int j = 0; j < _columns; j++)
        //            for (int i = 0; i < _rows; i++)
        //                list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

        //        ObjWorkBook.Close(false, Type.Missing, Type.Missing);

        //        ObjWorkExcel.Quit();

        //        GC.Collect();
        //        using (SecondISRPOLabaEntities entities = new SecondISRPOLabaEntities())
        //        {
        //            for (int i = 1; i < _rows; i++)
        //            {
        //                entities.Services.Add(new Services()
        //                {
        //                    id = int.Parse(list[i, 0]),
        //                    name = list[i, 1],
        //                    type = list[i, 2],
        //                    code = list[i, 3],
        //                    cost = decimal.Parse(list[i, 4])
        //                });
        //            }
        //            entities.SaveChanges();
        //        }
        //    }
        //}

        //private void Export_Click(object sender, RoutedEventArgs e)
        //{
        //    Task.Run(() => //async runtume - anti screen lock
        //    {

        //        List<Services> allServices;
        //        using (SecondISRPOLabaEntities usersEntities = new SecondISRPOLabaEntities())
        //        {
        //            allServices = usersEntities.Services.ToList()
        //                .OrderBy(s => s.name)
        //                .ToList();
        //        }
        //        var app = new Excel.Application();
        //        app.SheetsInNewWorkbook = _sheetsCount;
        //        Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
        //        #region convenient grouping structure
        //        var studentsCategories = allServices
        //                .OrderBy(o => o.cost)
        //                .GroupBy(s => s.cost)
        //                .ToDictionary(g => g.Key, g => g.Select(s => new { s.id, s.name, s.type, s.cost })
        //                .ToArray());
        //        #endregion
        //        for (int i = 0; i < _sheetsCount; i++)
        //        {
        //            int startRowIndex = 1;
        //            var worksheet = app.Worksheets.Item[i + 1];
        //            #region headers
        //            worksheet.Name = $"Категория {i + 1}";
        //            worksheet.Cells[1][startRowIndex] = "Id";
        //            worksheet.Cells[2][startRowIndex] = "Название услуги";
        //            worksheet.Cells[3][startRowIndex] = "Вид услуги";
        //            worksheet.Cells[4][startRowIndex] = "Стоимость";
        //            startRowIndex++;
        //            #endregion

        //            #region spliting by condition
        //            var data = i == 0 ? studentsCategories.Where(w => w.Key.HasValue && w.Key.Value >= 0 && w.Key.Value <= 350)
        //                : i == 1 ? studentsCategories.Where(w => w.Key.HasValue && w.Key.Value >= 250 && w.Key.Value <= 800)
        //                : i == 2 ? studentsCategories.Where(w => w.Key.HasValue && w.Key.Value >= 800) : studentsCategories;
        //            #endregion

        //            foreach (var students in data)
        //            {
        //                foreach (var student in students.Value)
        //                {
        //                    if (student.cost == students.Key)
        //                    {
        //                        #region fill the fields
        //                        worksheet.Cells[1][startRowIndex] = student.id;
        //                        worksheet.Cells[2][startRowIndex] = student.name;
        //                        worksheet.Cells[3][startRowIndex] = student.type;
        //                        worksheet.Cells[4][startRowIndex] = student.cost;
        //                        startRowIndex++;
        //                        #endregion
        //                    }
        //                }
        //            }
        //            worksheet.Columns.AutoFit();
        //        }
        //        app.Visible = true;
        //    });

        //}
    }
}
