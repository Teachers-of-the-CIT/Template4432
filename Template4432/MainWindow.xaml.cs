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
using Template4432.Forms;

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

        private void BnThird_Click(object sender, RoutedEventArgs e)
        {
            _4432_Valiakhmetov valiakhmetov = new _4432_Valiakhmetov();
            valiakhmetov.ShowDialog();
        }
        private void BnEleven_Click(object sender, RoutedEventArgs e)
        {
            new _4432_Latypov().Show();
        }

        private void BnSecond_Click(object sender, RoutedEventArgs e)
        {
            new _4432_Abramov().Show();
        }
        
        private void BnNine_Click(object sender, RoutedEventArgs e)
        {
            _4432_Darchuk Darchuk = new _4432_Darchuk();
            this.Close();
            Darchuk.ShowDialog();
        }

        private void BnTwenty_Click(object sender, RoutedEventArgs e)
        {
            new _4432_Suhanova().Show();
        }
        private void BnFour_Click(object sender, RoutedEventArgs e)
        {
            _4432_Bastanov bastanov = new _4432_Bastanov();
            bastanov.ShowDialog();            
        }
        private void Btn_Fakhrutdinov_form_click(object sender, RoutedEventArgs e)
        {
            _4432_Fakhrutdinov_Marat Fakhrutdinov_Marat_window = new _4432_Fakhrutdinov_Marat();
            Fakhrutdinov_Marat_window.ShowDialog();
        }

        private void BnFifteen_Click(object sender, RoutedEventArgs e)
        {
            new _4432_RakhimovRamil().Show();
        }

        private void BnUniversal(object sender, RoutedEventArgs e)
        {
            Button b = (Button)sender;
            Window w = null;

            if (b.Content.ToString() == "Ашрафзянов Марат")
                w = new _4432_Ashrafzianov();
            if (w != null)
            {
                w.Show();
                this.Visibility = Visibility.Hidden;
                w.Closed += (_s, _e) => {
                    this.Visibility = Visibility.Visible;
                };
            }
        }
    }
}
