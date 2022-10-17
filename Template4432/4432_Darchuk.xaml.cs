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
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Darchuk.xaml
    /// </summary>
    public partial class _4432_Darchuk : Window
    {
        public _4432_Darchuk()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
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

            using (ISRPOLab2ExcelEntities1 db = new ISRPOLab2ExcelEntities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.Employee.Add(new Employee()
                    {
                        EmployeeID = list[i, 0],
                        EmployeePosition = list[i, 1],
                        EmployeeFIO = list[i, 2],
                        EmployeeLogin = list[i, 3],
                        EmployeePassword = list[i, 4],
                        EmployeeLastEntry = list[i, 5],
                        EmployeeTypeEntry = list[i, 6],
                    });
                }
                db.SaveChanges();
            }
            MessageBox.Show("Данные успешно ипортированы.");
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Employee> allEmployee;
            using (ISRPOLab2ExcelEntities1 db = new ISRPOLab2ExcelEntities1())
            {
                allEmployee = db.Employee.ToList().OrderBy(emp => emp.EmployeeTypeEntry).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 2;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var EmployeeEntryTypes = allEmployee.OrderBy(o => o.EmployeeTypeEntry).GroupBy(g => g.EmployeeTypeEntry).ToDictionary(d => d.Key, d => d.Select(g => new { g.EmployeeID, g.EmployeePosition, g.EmployeeLogin }).ToArray());

            for (int i = 0; i < 2; i++)
            {
                int startRowIndex = 1;
                var worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Тип входа {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;
                
                using (ISRPOLab2ExcelEntities1 db = new ISRPOLab2ExcelEntities1())
                {
                    List<Employee> employee;
                    employee = (from em in db.Employee select em).ToList<Employee>();
                    foreach (var emp in employee)
                    {
                        if (i == 0)
                        {
                            if (emp.EmployeeTypeEntry == "Успешно")
                            {
                                worksheet.Cells[1][startRowIndex] = emp.EmployeeID;
                                worksheet.Cells[2][startRowIndex] = emp.EmployeePosition;
                                worksheet.Cells[3][startRowIndex] = emp.EmployeeLogin;
                                startRowIndex++;
                            }
                        }
                        else if (i == 1)
                        {
                            if (emp.EmployeeTypeEntry == "Неуспешно")
                            {
                                worksheet.Cells[1][startRowIndex] = emp.EmployeeID;
                                worksheet.Cells[2][startRowIndex] = emp.EmployeePosition;
                                worksheet.Cells[3][startRowIndex] = emp.EmployeeLogin;
                                startRowIndex++;
                            }
                        }
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void BtnClearDatabase_Click(object sender, RoutedEventArgs e)
        {
            using (ISRPOLab2ExcelEntities1 db = new ISRPOLab2ExcelEntities1())
            {
                List<Employee> employee;
                employee = (from em in db.Employee select em).ToList<Employee>();
                foreach (var emp in employee)
                {
                    db.Employee.Remove(emp);
                }
                db.SaveChanges();
            }
            MessageBox.Show("База данных успешно очищена.");
        }

        private async void BtnImportJson_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "Файл Json (*.json)|*.json|Text files (*.txt)|*.txt",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            FileStream inStream = File.OpenRead(ofd.FileName);

            /*Employee? employeeJson = JsonSerializer.Deserialize<Employee>(inStream);
            MessageBox.Show(employeeJson.EmployeeID);*/
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
