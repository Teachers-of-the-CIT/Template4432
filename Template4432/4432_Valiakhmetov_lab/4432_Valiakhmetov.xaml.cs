using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Template4432._4432_Valiakhmetov_lab;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Valiakhmetov.xaml
    /// </summary>
    public partial class _4432_Valiakhmetov : Window
    {
        public _4432_Valiakhmetov()
        {
            InitializeComponent();
        }

        private void btn_import_from_excel_Click(object sender, RoutedEventArgs e)
        {
            // открываем диалог для выбора файла
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            // массив для хранения информации из книги
            string[,] list;

            // открываем выбранную книгу Excel
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorksheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];

            // заполнение массива информацией
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 1; i < _rows; i++)
                {
                    list[i, j] = ObjWorksheet.Cells[i + 1, j + 1].Text;                                   
                }
            }

            // закрываем книгу
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            // добавляем информацию в базу данных
            using (var db = new ClientsExcelDB())
            {
                for (var i = 1; i < _rows; i++)
                {
                    var client = new Clients()
                    {
                        fio = list[i, 0],
                        client_code = list[i, 1],
                        birthday = list[i, 2],
                        index = list[i, 3],
                        city = list[i, 4],
                        street = list[i, 5],
                        house_num = list[i, 6],
                        flat_num = list[i, 7],
                        mail = list[i, 8],
                    };
                    db.Clients.Add(client);
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Успешный импорт");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла ошибка: " + ex.Message);
                }
                
            }
        }

        private void btn_export_to_excel_Click(object sender, RoutedEventArgs e)
        {
            // категории: по улице проживания. дополнительно отсортировать ФИО в алфавитном порядке. формат: код клиента | фио | email
            // список для хранения отсортированных по ФИО данных
            var sortedByFIO = new List<Clients>();
            // список названий всех улиц (будущие категории)
            var allStreetNames = new List<string>();

            // заполнение списка информацией из бд
            using (ClientsExcelDB db = new ClientsExcelDB())
            {
                sortedByFIO = db.Clients.ToList().OrderBy(f => f.fio).ToList();                
            }

            // убираем все пустые строчки
            sortedByFIO.RemoveAll(s => string.IsNullOrWhiteSpace(s.fio));

            // заполняем список названий всех улиц
            foreach (var item in sortedByFIO)
            {
                if (item.street != null && item.street != "")
                {
                    allStreetNames.Add(item.street);
                }                
            }

            // формируем список уникальных значений
            var distinctStreets = allStreetNames.Distinct().ToList();

            // создаем Excel книгу с количеством страниц, равной длине списка улиц
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = distinctStreets.Count();
            var book = app.Workbooks.Add(Type.Missing);
            // для навигации по строчкам в таблице
            var startRowIndex = 1;

            // добавляем на каждый лист имена столбцов
            for (int i = 0; i < distinctStreets.Count(); i++)
            {
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = distinctStreets[i].Replace(" ", "");
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "E-mail";                
            }

            // теперь мы добавляем информацию от 2 клетки, т.к. на первой расположены заголовки
            startRowIndex++;

            // заполняем страницы книги информацией
            foreach (var item in sortedByFIO)
            {
                for (int i = 0; i < distinctStreets.Count(); i++)
                {
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    if (worksheet.Name == item.street.Replace(" ", ""))
                    {
                        while (worksheet.Cells[1][startRowIndex].Text != "")
                        {
                            startRowIndex++;
                        }
                        worksheet.Cells[1][startRowIndex] = item.client_code;
                        worksheet.Cells[2][startRowIndex] = item.fio;
                        worksheet.Cells[3][startRowIndex] = item.mail;
                    }
                }

                startRowIndex = 2;
            }

            // показываем готовую книгу Excel
            app.Visible = true;
        }
    }
}
