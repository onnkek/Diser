using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

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

        private static int _tolerancePower = 1;
        private static double _fCenter;
        private static double _fRight;
        private static double _fLeft;
        private static List<DCLink> _links = new List<DCLink>();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LoadModel(pathCenter);
            _links.Clear();
            // Create DCLink with default parameters
            _links.Add(new DCLink("DC1", 43.9, 21.8));
            _links.Add(new DCLink("DC2", 28.1, 1.6));
            _links.Add(new DCLink("DC3", 16.2, 3.6));

            // Generation management
            powerActiveLoad.set_ZN(GetIndex(6), _links[0].P); // 6 узел
            powerActiveLoad.set_ZN(GetIndex(7), _links[1].P); // 7 узел
            powerActiveLoad.set_ZN(GetIndex(9), _links[2].P); // 9 узел
            powerActiveLoad.set_ZN(GetIndex(26), 43.9); // 26 узел
            powerActiveLoad.set_ZN(GetIndex(27), 28.1); // 27 узел
            powerActiveLoad.set_ZN(GetIndex(29), 16.2); // 29 узел

            dc56.Text = GetPower(_links[0]);
            dc47.Text = GetPower(_links[1]);
            dc49.Text = GetPower(_links[2]);

            // Change generation in nods 2 and 3 to the normal distribution
            Random rnd = new Random();
            double power = rnd.NextGaussian(20, 20);
            powerActiveGeneration.set_ZN(GetIndex(2), power);
            powerRectiveGeneration.set_ZN(GetIndex(2), power / 2);
            power = rnd.NextGaussian(20, 20);
            powerActiveGeneration.set_ZN(GetIndex(3), power);
            powerRectiveGeneration.set_ZN(GetIndex(3), power / 2);



            calcRegim(Rastr);
            _fCenter = freq.get_ZN(0);
            fCenter.Text = Math.Round(freq.get_ZN(0), 3).ToString();
            load2.Text = GetPower(2, true);
            load3.Text = GetPower(3, true);
            load4.Text = GetPower(4, true);
            load5.Text = GetPower(5, true);
            gen2.Text = GetPower(2, false);
            gen3.Text = GetPower(3, false);


            LoadModel(pathRight);
            
            // Generation management
            powerActiveGeneration.set_ZN(GetIndex(1), _links[0].P);  // 1 узел
            powerActiveGeneration.set_ZN(GetIndex(2), _links[1].P); // 2 узел
            powerActiveGeneration.set_ZN(GetIndex(3), _links[2].P); // 3 узел

            calcRegim(Rastr);
            _fRight = freq.get_ZN(0);
            fRight.Text = Math.Round(freq.get_ZN(0), 3).ToString();

            load6.Text = GetPower(6, true);
            load9.Text = GetPower(9, true);
            load10.Text = GetPower(10, true);
            load11.Text = GetPower(11, true);
            load12.Text = GetPower(12, true);
            load13.Text = GetPower(13, true);
            load14.Text = GetPower(14, true);
            gen6.Text = GetPower(6, false);
            gen8.Text = GetPower(8, false);
            gen9.Text = GetPower(9, false);

            LoadModel(pathLeft);

            // Generation management
            powerActiveGeneration.set_ZN(GetIndex(1), 43.9);  // 1 узел
            powerActiveGeneration.set_ZN(GetIndex(2), 28.1); // 2 узел
            powerActiveGeneration.set_ZN(GetIndex(3), 16.2); // 3 узел
            
            calcRegim(Rastr);
            _fLeft = freq.get_ZN(0);
            fLeft.Text = Math.Round(freq.get_ZN(0), 3).ToString();

            openModel.Text = "Режим рассчитан!";
            openModel.Foreground = new SolidColorBrush(Colors.Green);
        }
        
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            for(int i = 0; i < 20 & Math.Abs(50 - _fCenter) > 0.001; i++)
            {
                double dF = 50 - _fCenter;
                LoadModel(pathCenter);

                // Generation management to maintain frequency
                foreach (var link in _links)
                {
                    link.P = link.P - link.P * dF / 50 / 3;
                    link.Q = link.Q - link.Q * dF / 50 / 3;
                }

                powerActiveLoad.set_ZN(GetIndex(6), _links[0].P); // 6 узел
                powerActiveLoad.set_ZN(GetIndex(7), _links[1].P); // 7 узел
                powerActiveLoad.set_ZN(GetIndex(9), _links[2].P); // 9 узел
                powerActiveLoad.set_ZN(GetIndex(26), 43.9); // 26 узел
                powerActiveLoad.set_ZN(GetIndex(27), 28.1); // 27 узел
                powerActiveLoad.set_ZN(GetIndex(29), 16.2); // 29 узел

                dc56.Text = GetPower(_links[0]);
                dc47.Text = GetPower(_links[1]);
                dc49.Text = GetPower(_links[2]);


                calcRegim(Rastr);
                _fCenter = freq.get_ZN(0);
                fCenter.Text = i + " : " + Math.Round(freq.get_ZN(0), 3).ToString();
                load2.Text = GetPower(2, true);
                load3.Text = GetPower(3, true);
                load4.Text = GetPower(4, true);
                load5.Text = GetPower(5, true);
                gen2.Text = GetPower(2, false);
                gen3.Text = GetPower(3, false);


                //LoadModel(pathRight);
                //// Управление ВПТ

                //powerActiveGeneration.set_ZN(GetIndex(1), P56);  // 1 узел
                //powerActiveGeneration.set_ZN(GetIndex(2), P47); // 2 узел
                //powerActiveGeneration.set_ZN(GetIndex(3), P49); // 3 узел


                //// Управление генерацией
                ////powerActiveGeneration.set_ZN(2, 10);  // 8 узел

                //calcRegim(Rastr);
                //fRight.Text = Math.Round(freq.get_ZN(0), 3).ToString();

                //load6.Text = GetPower(6, true);
                //load9.Text = GetPower(9, true);
                //load10.Text = GetPower(10, true);
                //load11.Text = GetPower(11, true);
                //load12.Text = GetPower(12, true);
                //load13.Text = GetPower(13, true);
                //load14.Text = GetPower(14, true);
                //gen6.Text = GetPower(6, false);
                //gen8.Text = GetPower(8, false);
                //gen9.Text = GetPower(9, false);

                //LoadModel(pathLeft);
                //// Управление ВПТ
                //powerActiveGeneration.set_ZN(GetIndex(1), 43.9);  // 1 узел
                //powerActiveGeneration.set_ZN(GetIndex(2), 28.1); // 2 узел
                //powerActiveGeneration.set_ZN(GetIndex(3), 16.2); // 3 узел
                //                                                 // Управление генерацией


                //calcRegim(Rastr);
                //fLeft.Text = Math.Round(freq.get_ZN(0), 3).ToString();


                openModel.Text = "Режим рассчитан!";
                openModel.Foreground = new SolidColorBrush(Colors.Green);
            }    
            
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
            for (int i = 0; i < Node.Count; i++)
            {
                voltageCalc.set_ZN(i, 0);
                voltageAngle.set_ZN(i, 0);
            }
            ASTRALib.ITable ParamRgm = inRastr.Tables.Item("com_regim");
            ASTRALib.ICol statusRgm = ParamRgm.Cols.Item("status");
            inRastr.rgm(""); //Расчет режима 
            int a = statusRgm.get_ZN(0);
            if (a == 0)
                return true;
            else
                return false;
        }
        private static string GetPower(DCLink link)
        {
            string result = "";
            if (link.P != 0)
                result += Math.Round(link.P, _tolerancePower);

            if (link.Q > 0)
                result += "+j∙" + Math.Round(link.Q, _tolerancePower);
            if (link.Q < 0)
                result += "-j∙" + Math.Round(link.Q, _tolerancePower) * -1;
            return result;
        }
        private static string GetPower(double P, double Q)
        {
            string result = "";
            if (P != 0)
                result += Math.Round(P, _tolerancePower);

            if (Q > 0)
                result += "+j∙" + Math.Round(Q, _tolerancePower);
            if (Q < 0)
                result += "-j∙" + Math.Round(Q, _tolerancePower) * -1;
            return result;
        }
        private static string GetPower(int number, bool load)
        {
            string result = "";
            if (load)
            {
                if (powerActiveLoad.get_ZN(GetIndex(number)) != 0)
                    result += Math.Round(powerActiveLoad.get_ZN(GetIndex(number)), _tolerancePower);

                if (powerRectiveLoad.get_ZN(GetIndex(number)) > 0)
                    result += "+j∙" + Math.Round(powerRectiveLoad.get_ZN(GetIndex(number)), _tolerancePower);
                if (powerRectiveLoad.get_ZN(GetIndex(number)) < 0)
                    result += "-j∙" + Math.Round(powerRectiveLoad.get_ZN(GetIndex(number)), _tolerancePower) * -1;
            }
            else
            {
                if (powerActiveGeneration.get_ZN(GetIndex(number)) != 0)
                    result += Math.Round(powerActiveGeneration.get_ZN(GetIndex(number)), _tolerancePower);

                if (powerRectiveGeneration.get_ZN(GetIndex(number)) > 0)
                    result += "+j∙" + Math.Round(powerRectiveGeneration.get_ZN(GetIndex(number)), _tolerancePower);
                if (powerRectiveGeneration.get_ZN(GetIndex(number)) < 0)
                    result += "-j∙" + Math.Round(powerRectiveGeneration.get_ZN(GetIndex(number)), _tolerancePower) * -1;
            }
            return result;
        }

        private static int GetIndex(int number)
        {
            int result = -1;
            for (int i = 0; i < Node.Count; i++)
            {
                if (numberBus.ZN[i] == number)
                    result = i;
            }
            return result;
        }

        public class DCLink
        {
            public DCLink(string name, double p, double q)
            {
                Name = name;
                P = p;
                Q = q;
            }
            public string Name { get; set; }
            public double P { get; set; }
            public double Q { get; set; }
        }
        
    }
    public static class RandomExtensions
    {
        public static double NextGaussian(this Random rnd, double μ, double σ)
        {
            double u, v, s;
            do
            {
                u = 2 * rnd.NextDouble() - 1;
                v = 2 * rnd.NextDouble() - 1;
                s = u * u + v * v;
            }
            while (u <= -1 || v <= -1 || s >= 1 || s == 0);
            double r = Math.Sqrt(-2 * Math.Log(s) / s);
            return μ + σ * r * u;
        }
    }
}
