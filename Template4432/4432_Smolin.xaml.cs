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
using Microsoft.Win32;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template4432
{
    
    /// <summary>
    /// Логика взаимодействия для _4432_Smolin.xaml
    /// </summary>
    public partial class _4432_Smolin : Window
    {
        LR2IsrpoEntities db = new LR2IsrpoEntities();
        public _4432_Smolin()
        {
            InitializeComponent();

        }

        private void SmolinImportBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                    Title = "Выберите файл"
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
                for (int i = 1; i < _rows; i++)
                {
                    db.Worker.Add(new Worker()
                    {
                        WorkerID = int.Parse(list[i, 0]),
                        Post = list[i, 1],
                        Full_name = list[i, 2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        Last_authorization = list[i, 5],
                        Type_authorization = list[i, 6],
                    });
                }
                db.SaveChanges();
                MessageBox.Show("Данные успешно ипортированы.");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");

            }
        }

        private void SmolinExportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Worker> allWorker;
            
                allWorker = db.Worker.ToList().OrderBy(work => work.Type_authorization).ToList();
            
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 2;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var EmployeeEntryTypes = allWorker.OrderBy(o => o.Type_authorization).GroupBy(g => g.Type_authorization).ToDictionary(d => d.Key, d => d.Select(g => new { g.WorkerID, g.Post, g.Login }).ToArray());

            for (int i = 0; i < 2; i++)
            {
                int startRowIndex = 1;
                var worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Тип входа {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Код сотрудника";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;

                
                    List<Worker> workeer;
                    workeer = (from em in db.Worker select em).ToList<Worker>();
                    foreach (var work in workeer)
                    {
                        if (i == 0)
                        {
                            if (work.Type_authorization == "Успешно")
                            {
                                worksheet.Cells[1][startRowIndex] = work.WorkerID;
                                worksheet.Cells[2][startRowIndex] = work.Post;
                                worksheet.Cells[3][startRowIndex] = work.Login;
                                startRowIndex++;
                            }
                        }
                        else if (i == 1)
                        {
                            if (work.Type_authorization == "Неуспешно")
                            {
                                worksheet.Cells[1][startRowIndex] = work.WorkerID;
                                worksheet.Cells[2][startRowIndex] = work.Post;
                                worksheet.Cells[3][startRowIndex] = work.Login;
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
