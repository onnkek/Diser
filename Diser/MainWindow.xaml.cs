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
        private static string pathCenter = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Center.rg2";
        private static string pathRight = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Right.rg2";
        private static string pathLeft = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Left.rg2";
        private static ASTRALib.IRastr Rastr;
        private static ASTRALib.ITable Node;
        private static ASTRALib.ITable Vetv;
        private static ASTRALib.ITable Islands;
        private static ASTRALib.ICol numberBus;              //Номер Узла
        private static ASTRALib.ICol powerActiveLoad;        //активная мощность нагрузки.
        private static ASTRALib.ICol powerRectiveLoad;       //реактивная мощность нагрузки.
        private static ASTRALib.ICol powerActiveGeneration;  //активная мощность генерации.
        private static ASTRALib.ICol powerRectiveGeneration; //реактивная мощность генерации.
        private static ASTRALib.ICol voltageCalc;            //Расчётное напряжение.
        private static ASTRALib.ICol voltageAngle;           //Расчётный угол.
        private static ASTRALib.ICol freq;                   //Частота.

        private static int trackBar1_delta = 50;
        public MainWindow()
        {
            InitializeComponent();
            slider1.Minimum = 100;
            slider1.Maximum = 1000;
            slider1.TickFrequency = trackBar1_delta;
            centerF.Text = $"";
            sg6.Text = $"";
            rightF.Text = $"";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LoadModel(pathCenter);
            // Управление ВПТ
            powerActiveLoad.set_ZN(5, 43.9); // 6 узел
            powerActiveLoad.set_ZN(6, 28.1); // 7 узел
            powerActiveLoad.set_ZN(7, 16.2); // 9 узел
            powerActiveLoad.set_ZN(12, 43.9); // 26 узел
            powerActiveLoad.set_ZN(13, 28.1); // 27 узел
            powerActiveLoad.set_ZN(14, 16.2); // 29 узел

            // Управление генерацией


            calcRegim(Rastr);
            centerF.Text = $"Fcenter = {Math.Round(freq.get_ZN(0), 3)}";

            LoadModel(pathRight);
            // Управление ВПТ
            powerActiveGeneration.set_ZN(9, 43.9);  // 1 узел
            powerActiveGeneration.set_ZN(10, 28.1); // 2 узел
            powerActiveGeneration.set_ZN(11, 16.2); // 3 узел

            // Управление генерацией
            //powerActiveGeneration.set_ZN(2, 10);  // 8 узел

            calcRegim(Rastr);
            rightF.Text = $"Fright = {Math.Round(freq.get_ZN(0), 3)}";


            LoadModel(pathLeft);
            // Управление ВПТ
            powerActiveGeneration.set_ZN(9, 43.9);  // 1 узел
            powerActiveGeneration.set_ZN(10, 28.1); // 2 узел
            powerActiveGeneration.set_ZN(11, 16.2); // 3 узел
            
            // Управление генерацией


            calcRegim(Rastr);
            leftF.Text = $"Fleft = {Math.Round(freq.get_ZN(0), 3)}";


            openModel.Text = "Модель загружена!";
            openModel.Foreground = new SolidColorBrush(Colors.Green);
        }


        public void LoadModel(string path)
        {
            Rastr = new ASTRALib.Rastr();
            Rastr.Load(ASTRALib.RG_KOD.RG_REPL, path, "");

            Node = Rastr.Tables.Item("node");
            Vetv = Rastr.Tables.Item("vetv");
            Islands = Rastr.Tables.Item("islands");
            freq = Islands.Cols.Item("f");
            numberBus = Node.Cols.Item("ny"); //Номер Узла
            powerActiveLoad = Node.Cols.Item("pn"); //активная мощность нагрузки.
            powerRectiveLoad = Node.Cols.Item("qn"); //реактивная мощность нагрузки.
            powerActiveGeneration = Node.Cols.Item("pg"); //активная мощность нагрузки.
            powerRectiveGeneration = Node.Cols.Item("qg"); //реактивная мощность нагрузки.
            voltageCalc = Node.Cols.Item("vras"); //Расчётное напряжение.
            voltageAngle = Node.Cols.Item("delta"); //Расчётный угол.
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
