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
using System.Configuration;
using System.IO;

namespace ComandoCentralLauncher
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public string AppendMsg
        {
            set { textBox.AppendText(value); }
        }

        public MainWindow()
        {
            //InitializeComponent();
            StartComandoCentral();
        }

        public async void StartComandoCentral()
        {
            InitializeComponent();
            //MessageBox.Show("Click para continuar!");
            InstallUpdateCC installupdatecc = new InstallUpdateCC(this);

            if (await installupdatecc.isMSAccessInstalled())
            {
                if (!await installupdatecc.isCCBinInstalled())
                {
                    WindowsTokenEvaluation.CheckElevation();
                    await installupdatecc.installCCBin();
                }
                
                await installupdatecc.UpdateCC();
                
                LaunchAccess.Launch();
                Application.Current.Shutdown();
            }
        }
    }
}
