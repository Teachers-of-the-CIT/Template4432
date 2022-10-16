using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
    /// Логика взаимодействия для _4432_Zaripov.xaml
    /// </summary>
    public partial class _4432_Zaripov : Window
    {

        public _4432_Zaripov()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ButtonImp_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel(Spisok.xlsx) | *.xlsx",
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
            for(int j = 0; j < _columns; j++)
            {
                for(int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;                  
                }               
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                var countItems = usersEntities.tableIsrpo2.ToList().Count();
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.tableIsrpo2.Add(new tableIsrpo2()
                    {
                        id = ++countItems,
                        role_person = list[i,0],
                        fullName = list[i,1],
                        loginPerson = list[i,2],
                        passwordPerson = list[i,3]                     
                    });
                }
                usersEntities.SaveChanges();
            }
            MessageBox.Show("Успешно!");
        }

        public string HashPassword(string password)
        {
            byte[] salt;
            byte[] buffer2;
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }
            using (Rfc2898DeriveBytes bytes = new Rfc2898DeriveBytes(password, 0x10, 0x3e8))
            {
                salt = bytes.Salt;
                buffer2 = bytes.GetBytes(0x20);
            }
            byte[] dst = new byte[0x31];
            Buffer.BlockCopy(salt, 0, dst, 1, 0x10);
            Buffer.BlockCopy(buffer2, 0, dst, 0x11, 0x20);
            return Convert.ToBase64String(dst);
        }

        private void ButtonIks_Click(object sender, RoutedEventArgs e)
        {
            List<tableIsrpo2> allItems;
            List<string> allItemsGroupRole;

            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                allItems = usersEntities.tableIsrpo2.ToList().OrderBy(f => f.role_person).ToList();
                allItemsGroupRole = allItems.Select(f=>f.role_person).Distinct().ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = allItemsGroupRole.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                for(int i = 0; i < allItemsGroupRole.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(allItemsGroupRole[i]);
                    worksheet.Cells[1][startRowIndex] = "Login";
                    worksheet.Cells[2][startRowIndex] = "Password";
                    worksheet.Cells[3][startRowIndex] = "Hash password";
                    startRowIndex++;
                    foreach(var us in allItems)
                    {
                        if (us.role_person == allItemsGroupRole[i])
                        {
                            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            headerRange.Font.Italic = true;
                            worksheet.Cells[1][startRowIndex] = us.loginPerson;
                            worksheet.Cells[2][startRowIndex] = us.passwordPerson;
                            worksheet.Cells[3][startRowIndex] = HashPassword(us.passwordPerson);
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
                                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                                            Excel.XlLineStyle.xlContinuous;
                    worksheet.Columns.AutoFit();                                
                }
                app.Visible = true;            
            }
        }
    }
}
