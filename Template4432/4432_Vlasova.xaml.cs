using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlTypes;
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
using System.Globalization;
using Microsoft.Win32;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Vlasova.xaml
    /// </summary>
    public partial class _4432_Vlasova : Window
    {
        public _4432_Vlasova()
        {
            InitializeComponent();
        }

        private void importButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls; *.xlsx",
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
            for (int j = 0; j< _columns; j++)
            {
                for (int i = 0; i< _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                for (int i = 1; i<_rows; i++)
                {
                    if(String.IsNullOrEmpty(list[i,1]))
                    {
                        break;
                    }
                    var currentDate = DateTime.Now;
                    var birthdaydate = DateTime.ParseExact(list[i, 2], "MM.dd.yyyy", CultureInfo.CurrentCulture);
                    double differencebetweendates = currentDate.Subtract(birthdaydate).TotalDays;
                    int exactage = Convert.ToInt32(differencebetweendates);
                    int age =exactage/365;
                    usersEntities.C3xlsx.Add(new C3xlsx()
                    {
                        NSP = list[i, 0],
                        Client_id = list[i, 1],
                        Birthdate = list[i,2],
                        Index = list[i, 3],
                        City = list[i, 4],
                        Street = list[i, 5],
                        House = list[i, 6],
                        Flat = list[i, 7],
                        Email = list[i, 8]
                    });
                    usersEntities.Clients.Add(new Clients()
                    {
                        NSP = list[i, 0],
                        Client_Id = list[i, 1],
                        Email = list[i, 8],
                        Age = age
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("Данные успешно были записаны");
            }
        }

        private void exportButton_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;
            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                allClients = usersEntities.Clients.ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = 3;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                int CategoryNow = 1;
                for (int i = 0; i<3; i++)
                {
                    List<Clients> needtoremove = new List<Clients>();
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = "Категория" + (i + 1).ToString();
                    worksheet.Cells[1][startRowIndex] = "Код клиента";
                    worksheet.Cells[2][startRowIndex] = "ФИО";
                    worksheet.Cells[3][startRowIndex] = "Email";
                    startRowIndex++;
                    if(CategoryNow == 1)
                    {
                        foreach(Clients client in allClients)
                        {
                            if(client.Age >=20 && client.Age <= 29)
                            {
                                worksheet.Cells[1][startRowIndex] = client.Client_Id;
                                worksheet.Cells[2][startRowIndex] = client.NSP;
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
                                worksheet.Cells[1][startRowIndex] = client.Client_Id;
                                worksheet.Cells[2][startRowIndex] = client.NSP;
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
                                worksheet.Cells[1][startRowIndex] = client.Client_Id;
                                worksheet.Cells[2][startRowIndex] = client.NSP;
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
