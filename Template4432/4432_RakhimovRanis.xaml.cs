using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
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
    /// Логика взаимодействия для _4432_RakhimovRanis.xaml
    /// </summary>
    public partial class _4432_RakhimovRanis : System.Windows.Window
    {
        public _4432_RakhimovRanis()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (3.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            var _columns = lastCell.Column;
            var _rows = lastCell.Row;
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

            using (LR2ISRPOEntities entities = new LR2ISRPOEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (String.IsNullOrEmpty(list[i, 1]))
                    {
                        break;
                    }
                    var currentdate = DateTime.Now;
                    var bd = DateTime.ParseExact(list[i, 2], "dd.MM.yyyy", null);
                    double differentYear = currentdate.Subtract(bd).TotalDays;
                    int days = Convert.ToInt32(differentYear);
                    int age = days / 365;

                    entities.ClientData.Add(new ClientData()
                    {
                        FIO = list[i, 0],
                        ClientCode = list[i, 1],
                        Birthdate = list[i, 2],
                        Index = list[i, 3],
                        City = list[i, 4],
                        Street = list[i, 5],
                        House = list[i, 6],
                        Flat = list[i, 7],
                        Email = list[i, 8]
                    });
                    entities.Clients.Add(new Clients()
                    {
                        FIO = list[i, 0],
                        ClientCod = list[i, 1],
                        Email = list[i, 8],
                        Age = age

                    });
                }
                entities.SaveChanges();
                MessageBox.Show("Success");
            }
        }

        private void Ecsport_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;
            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                allClients = usersEntities.Clients.ToList();

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = 3;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                int CategoryNow = 1;
                for (int i = 0; i < 3; i++)
                {
                    List<Clients> needtoremove = new List<Clients>();
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = "Категория" + (i + 1).ToString();
                    worksheet.Cells[1][startRowIndex] = "Код клиента";
                    worksheet.Cells[2][startRowIndex] = "ФИО";
                    worksheet.Cells[3][startRowIndex] = "Email";
                    startRowIndex++;
                    if (CategoryNow == 1)
                    {
                        foreach (Clients client in allClients)
                        {
                            if (client.Age >= 20 && client.Age <= 29)
                            {
                                worksheet.Cells[1][startRowIndex] = client.ClientCod;
                                worksheet.Cells[2][startRowIndex] = client.FIO;
                                worksheet.Cells[3][startRowIndex] = client.Email;
                                startRowIndex++;
                                needtoremove.Add(client);
                            }
                        }
                    }
                    

                    if (CategoryNow == 2)
                    {
                        foreach (Clients client in allClients)
                        {
                            if (client.Age >= 30 && client.Age <= 39)
                            {
                                worksheet.Cells[1][startRowIndex] = client.ClientCod;
                                worksheet.Cells[2][startRowIndex] = client.FIO;
                                worksheet.Cells[3][startRowIndex] = client.Email;
                                startRowIndex++;
                                needtoremove.Add(client);
                            }
                        }
                    }
                    if (CategoryNow == 3)
                    {
                        foreach (Clients client in allClients)
                        {
                            if (client.Age >= 40)
                            {
                                worksheet.Cells[1][startRowIndex] = client.ClientCod;
                                worksheet.Cells[2][startRowIndex] = client.FIO;
                                worksheet.Cells[3][startRowIndex] = client.Email;
                                startRowIndex++;
                            }
                        }
                    }
                    CategoryNow++;
                    for (int j = 0; j < needtoremove.Count(); j++)
                    {
                        allClients.Remove(needtoremove[j]);
                    }
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }
            
        }
    }
}
