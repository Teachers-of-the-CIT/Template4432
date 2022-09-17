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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BnFirst_Click(object sender, RoutedEventArgs e)
        {
            MingalievWindow mingalievWindow = new MingalievWindow();
            mingalievWindow.ShowDialog();
        }

        private void BnNineteenStakheevVadim4432_Click(object sender, RoutedEventArgs e)
        {
            StakheevWindow stakheevWindow = new StakheevWindow();
            stakheevWindow.ShowDialog();
        }

        private void BnThirteen_Click(object sender, RoutedEventArgs e)
        {
            _4432_Naumkina naumkinaWindow = new _4432_Naumkina();
            naumkinaWindow.ShowDialog();
        }
        private void BnTen_Click(object sender, RoutedEventArgs e)
        {
            _4432_Zaripov zaripovWindow = new _4432_Zaripov();
            zaripovWindow.ShowDialog();
        }


        private void BnSixteen_Click(object sender, RoutedEventArgs e)
        {
            _4432_RakhimovRanis rakhimovWindow = new _4432_RakhimovRanis();
            rakhimovWindow.ShowDialog();
        }
        private void BnTwentyfour_Click(object sender, RoutedEventArgs e)
        {
            _4432_Fedorova fedorovaWindow = new _4432_Fedorova();
            fedorovaWindow.ShowDialog();
        }
        private void BnTwentyFive_Click(object sender, RoutedEventArgs e)
        {
            new _4432_Sharipov().ShowDialog();
        }
        private void BnSeventh_Click(object sender, RoutedEventArgs e)
        {
            _4432_Vlasova vlasovaWindow = new _4432_Vlasova();
            vlasovaWindow.ShowDialog();
        }
        private void BnFourteenth_Click(object sender, RoutedEventArgs e)
        {
            _4432_Nuryev nuryevwindow = new _4432_Nuryev();
            nuryevwindow.ShowDialog();
        }

        private void Bn12_Click(object sender, RoutedEventArgs e)
        {
            _4432_LatypovaDina _4432_LatypovaDina = new _4432_LatypovaDina();
            _4432_LatypovaDina.ShowDialog();
        }

        private void BnEleven_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
