using System;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace Diser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static ASTRALib.IRastr Rastr;
        //private static Rastr.Load(ASTRALib.RG_KOD.RG_REPL, path, "");
        private static ASTRALib.ITable Node;
        private static ASTRALib.ITable Vetv;
        private static ASTRALib.ICol numberBus; //Номер Узла
        private static ASTRALib.ICol powerActiveLoad; //активная мощность нагрузки.
        private static ASTRALib.ICol powerRectiveLoad; //реактивная мощность нагрузки.
        private static ASTRALib.ICol powerActiveGeneration; //активная мощность генерации.
        private static ASTRALib.ICol powerRectiveGeneration; //реактивная мощность генерации.
        private static ASTRALib.ICol voltageCalc; //Расчётное напряжение.
        private static ASTRALib.ICol voltageAngle; //Расчётный угол.

        private static string path;
        private static int trackBar1_delta = 50;
        public MainWindow()
        {
            InitializeComponent();
            slider1.Minimum = 100;
            slider1.Maximum = 1000;
            slider1.TickFrequency = trackBar1_delta;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл режима rg2 (.rg2) | *.rg2|All files| *.*";
            if (openDialog.ShowDialog() == true)
            {
                path = openDialog.FileName;
            }
            if (path != null)
            {
                Rastr = new ASTRALib.Rastr();
                Rastr.Load(ASTRALib.RG_KOD.RG_REPL, path, "");
                Node = Rastr.Tables.Item("node");
                Vetv = Rastr.Tables.Item("vetv");
                numberBus = Node.Cols.Item("ny"); //Номер Узла
                powerActiveLoad = Node.Cols.Item("pn"); //активная мощность нагрузки.
                powerRectiveLoad = Node.Cols.Item("qn"); //реактивная мощность нагрузки.
                powerActiveGeneration = Node.Cols.Item("pg"); //активная мощность нагрузки.
                powerRectiveGeneration = Node.Cols.Item("qg"); //реактивная мощность нагрузки.
                voltageCalc = Node.Cols.Item("vras"); //Расчётное напряжение.
                voltageAngle = Node.Cols.Item("delta"); //Расчётный угол.

                openModel.Text = "Модель загружена!";
                openModel.Foreground = new SolidColorBrush(Colors.Green);
            }
        }

        private void slider1_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            test.Text = slider1.Value.ToString();
            double trackBar1_value = slider1.Value;

            if (path != null)
            {

                //Отправка исходных данных
                powerActiveGeneration.set_ZN(1, slider1.Value);

                //Расчет режима
                calcRegim(Rastr);




                //label10.Text = $"U = {Math.Round(voltageCalc.get_ZN(0), 3)}∠{Math.Round(voltageAngle.get_ZN(0), 3)}";
                //label11.Text = $"U = {Math.Round(voltageCalc.get_ZN(1), 3)}∠{Math.Round(voltageAngle.get_ZN(1), 3)}";
                //label12.Text = $"U = {Math.Round(voltageCalc.get_ZN(2), 3)}∠{Math.Round(voltageAngle.get_ZN(2), 3)}";

                s1.Text = $"S = {Math.Round(powerActiveGeneration.get_ZN(0), 3)}+j·{Math.Round(powerRectiveGeneration.get_ZN(0), 3)}";
                s2.Text = $"S = {Math.Round(powerActiveLoad.get_ZN(1), 3)}+j·{Math.Round(powerRectiveLoad.get_ZN(1), 3)}";
                s3.Text = $"S = {Math.Round(powerActiveGeneration.get_ZN(2), 3)}+j·{Math.Round(powerRectiveGeneration.get_ZN(2), 3)}";


                //Thread.Sleep(50);


            }
            //else
                //MessageBox.Show("Загрузите файл модели!");
        }
        bool calcRegim(ASTRALib.IRastr inRastr)
        {
            ASTRALib.ITable ParamRgm = inRastr.Tables.Item("com_regim");
            ASTRALib.ICol statusRgm = ParamRgm.Cols.Item("status");
            inRastr.rgm(""); //Расчет режима 
            int a = statusRgm.get_ZN(0);
            if (a == 0)
                return true;
            else
                return false;
        }
    }
}
