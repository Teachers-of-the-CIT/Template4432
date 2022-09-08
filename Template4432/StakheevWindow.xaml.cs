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

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для StakheevWindow.xaml
    /// </summary>
    public partial class StakheevWindow : Window
    {
        public StakheevWindow()
        {
            InitializeComponent();
        }

        private void StakheevWindowClose(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
