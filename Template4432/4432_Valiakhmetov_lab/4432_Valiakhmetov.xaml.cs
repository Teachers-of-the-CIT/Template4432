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
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.IO;
using System.Runtime.InteropServices;

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
                for (int i = 0; i < _rows; i++)
                {
                    if (string.IsNullOrWhiteSpace(list[i,j]) && list[i,j] != "")
                        list[i, j] = ObjWorksheet.Cells[i + 2, j + 1].Text;                                   
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
                    if (list[i, 0] == "")
                        continue;
                    var client = new Clients()
                    {
                        fio = list[i, 0],
                        client_code = list[i, 1],
                        birthday = list[i, 2],
                        index = list[i, 3],
                        city = list[i, 4],
                        street = list[i, 5],
                        house_num = Convert.ToInt32(list[i, 6]),
                        flat_num = Convert.ToInt32(list[i, 7]),
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

        private void btn_export_to_word_Click(object sender, RoutedEventArgs e)
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

            // создание документа Word
            var app = new Word.Application();
            var document = app.Documents.Add();
            var index = 0;

            // создание параграфов
            for (int i = 0; i < distinctStreets.Count; i++)
            {
                int rowsCounter = 0;        

                foreach (var client in sortedByFIO)
                {
                    if (client.street.Replace(" ", "") == distinctStreets[index].Replace(" ",""))
                        rowsCounter++;
                }

                var paragraph = document.Paragraphs.Add();
                var range = paragraph.Range;
                range.Text = distinctStreets[index];
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var streetCategories = document.Tables.Add(tableRange, rowsCounter + 1, 3);

                streetCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                streetCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                streetCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                Word.Range cellRange;
                cellRange = streetCategories.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = streetCategories.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = streetCategories.Cell(1, 3).Range;
                cellRange.Text = "E-mail";

                streetCategories.Rows[1].Range.Bold = 1;
                streetCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                var count = 1;
                
                foreach (var item in sortedByFIO)
                {
                    if (item.street == distinctStreets[index])
                    {
                        cellRange = streetCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.client_code;
                        //cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = streetCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.fio;
                        cellRange = streetCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.mail;

                        count++;
                    }
                }

                index++;
            }                        

            // показываем готовую книгу Excel
            app.Visible = true;
        }

        private async void btn_import_from_json_Click(object sender, RoutedEventArgs e)
        {
            string path = "C:\\Users\\valia\\Source\\Repos\\LR1_Valiakhmetov_Task2_4432\\Template4432\\4432_Valiakhmetov_lab\\data.json";

            using (var fileStream = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (var db = new ClientsExcelDB())
                {
                    var clients = await JsonSerializer.DeserializeAsync<List<Clients>>(fileStream);

                    foreach (var item in clients)
                    {
                        var client = new Clients();

                        client.client_id = item.client_id;
                        client.fio = item.fio;
                        client.client_code = item.client_code;
                        client.birthday = item.birthday;
                        client.index = item.index;
                        client.city = item.city;
                        client.street = item.street;
                        client.house_num = item.house_num;
                        client.flat_num = item.flat_num;
                        client.mail = item.mail;

                        db.Clients.Add(client);
                    }
                    try
                    {
                        db.SaveChanges();
                        MessageBox.Show("Данные импортированы успешно!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
    }
}
