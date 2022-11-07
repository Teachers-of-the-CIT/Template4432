using System;
using System.IO;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Template4432.Application;
using Template4432.Contexts;
using Window = System.Windows.Window;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace Template4432.Forms
{
    public partial class _4432_Abramov : Window
    {
        private readonly SkiServiceService _skiServiceService;

        public _4432_Abramov(ApplicationContext context)
        {
            ExcelApplication excel = new ExcelApplication();
            
            _skiServiceService = new SkiServiceService(context, excel);
            
            InitializeComponent();
        }

        private void ImportButton_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл для импорта"
            };

            bool? showDialogResult = openFileDialog.ShowDialog();
            
            if (!showDialogResult.HasValue)
                return;
            
            if (!showDialogResult.Value)
                return;

            string fileName = openFileDialog.FileName;

            (bool importResult, int count) = _skiServiceService.ImportEntitiesFromWorkbook(fileName);

            if (!importResult)
            {
                MessageBox.Show("Неудача при импорте. Смотрите правильность", "Внимание", MessageBoxButton.OK, MessageBoxImage.Error);
                
                return;
            }

            MessageBox.Show($"Импорт успешен, загружено {count} сущностей", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportForExcelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Workbook workbook = _skiServiceService.ExportEntities();

            string fileName = Directory.GetCurrentDirectory() + $"{Guid.NewGuid()}.xls";

            try
            {
                workbook.SaveAs(fileName, ".xls");
            }
            catch { }
        }
    }
}