using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text;

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

        private static double _fCenter;
        private static double _fRight;
        private static double _fLeft;
        private static List<DCLink> _linksCR = new List<DCLink>();
        private static List<DCLink> _linksCL = new List<DCLink>();
        private static List<Generation> _generationCenter = new List<Generation>();
        private static List<Generation> _generationRight = new List<Generation>();
        private static List<Generation> _generationLeft = new List<Generation>();
        private static LoadCollection _loadCenter = new LoadCollection();
        private static LoadCollection _loadRight = new LoadCollection();
        private static Random _random = new Random();
        private static double _kLoad;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {


            List<double> fCenter = new List<double>();
            List<double> fRight = new List<double>();
            List<double> fLeft = new List<double>();
            List<double> fCenterAfterOpt = new List<double>();
            List<double> fRightAfterOpt = new List<double>();
            List<double> fLeftAfterOpt = new List<double>();
            List<double> count = new List<double>();

            List<List<double>> result = new List<List<double>>() { count,
                                                                   fCenter,
                                                                   fRight,
                                                                   fLeft,
                                                                   fCenterAfterOpt,
                                                                   fRightAfterOpt,
                                                                   fLeftAfterOpt };

            for (int i = 0; i < 50; i++)
            {
                double[] freq = InitialModel();
                fCenter.Add(freq[0]);
                fRight.Add(freq[1]);
                fLeft.Add(freq[2]);
                freq = FindFreq();
                fCenterAfterOpt.Add(freq[0]);
                fRightAfterOpt.Add(freq[1]);
                fLeftAfterOpt.Add(freq[2]);
                count.Add(freq[3]);
            }

            string text = "";
            for (var i = 0; i < result[0].Count; i++)
            {
                text += $"C:{result[0][i]}| ";
                text += $"C[{result[4][i]:F2}]|   ";
                text += $"C:{result[1][i]:F2}|";
                text += $"R:{result[2][i]:F2}|";
                text += $"L:{result[3][i]:F2}|";
                text += $"R[{result[5][i]:F2}]|";
                text += $"L[{result[6][i]:F2}]|";
                text += $"\r\n";
            }
            MessageBox.Show(text);

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            InitialModel();
        }
        public double[] FindFreq()
        {
            int counter = 0;
            double[] freq = new double[4];


            for (int i = 0; i < 80 & Math.Abs(50 - _fCenter) > 0.001; i++)
            {
                double dF = 50 - _fCenter;

                foreach (var link in _linksCL)
                {
                    if (link.S.X > 0)
                        link.S += link.S * dF / 50 / 3;
                    else if (link.S.X < 5 && link.S.X > -5)
                        link.S -= link.S * dF / 50 * 5;
                    else
                        link.S -= link.S * dF / 50 / 3;
                }
                foreach (var link in _linksCR)
                {
                    if (link.S.X > 0)
                        link.S += link.S * dF / 50 / 3;
                    else if (link.S.X < 5 && link.S.X > -5)
                        link.S -= link.S * dF / 50 * 5;
                    else
                        link.S -= link.S * dF / 50 / 3;


                    test.Text = $"P:{link.S.X}   dF:{dF}   Add:{link.S.X * dF / 50 / 6}";
                }
                freq = RunModel();
                counter++;
            }


            freq[3] = counter;
            return freq;
        }
        public double[] InitialModel()
        {
            _kLoad = 1;
            _linksCR.Clear();
            _linksCL.Clear();
            _generationCenter.Clear();
            _generationRight.Clear();
            _generationLeft.Clear();

            // Create DCLink with default parameters
            _linksCR.Add(new DCLink("C6-R1", new Complex(43.9, 21.8), 6, 1));
            _linksCR.Add(new DCLink("C7-R2", new Complex(28.1, 1.6), 7, 2));
            _linksCR.Add(new DCLink("C9-R3", new Complex(16.2, 3.6), 9, 3));

            _linksCL.Add(new DCLink("C26-L1", new Complex(43.9, 21.8), 26, 1));
            _linksCL.Add(new DCLink("C27-L2", new Complex(28.1, 1.6), 27, 2));
            _linksCL.Add(new DCLink("C29-L3", new Complex(16.2, 3.6), 29, 3));

            _generationCenter.Add(new Generation(20, 20, 3));
            _generationCenter.Add(new Generation(20, 20, 23));

            double P = 30;

            _generationRight.Add(new Generation(P, P, 6));
            _generationRight.Add(new Generation(P, P, 8));
            _generationRight.Add(new Generation(P, P, 9));

            _generationLeft.Add(new Generation(P, P, 6));
            _generationLeft.Add(new Generation(P, P, 8));
            _generationLeft.Add(new Generation(P, P, 9));


            return RunModel();
        }
        public double[] RunModel()
        {
            double[] result = new double[4];
            LoadModel(pathCenter);
            _loadCenter.GetLoadCollection();
            _loadCenter.SetLoadCollection(_kLoad);
            // Settings generation
            foreach (var link in _linksCR)
                link.SetGenerationDCLink(true);
            foreach (var link in _linksCL)
                link.SetGenerationDCLink(true);

            foreach (var gen in _generationCenter)
                gen.SetPowerGeneration();



            calcRegim(Rastr);

            #region Output Text
            dc56.Text = _linksCR[0].S.ToString();
            dc47.Text = _linksCR[1].S.ToString();
            dc49.Text = _linksCR[2].S.ToString();
            _fCenter = freq.get_ZN(0);
            fCenter.Text = Math.Round(freq.get_ZN(0), 3).ToString();
            load2.Text = GetPower(2, true).ToString();
            load3.Text = GetPower(3, true).ToString();
            load4.Text = GetPower(4, true).ToString();
            load5.Text = GetPower(5, true).ToString();
            gen2.Text = GetPower(2, false).ToString();
            gen3.Text = GetPower(3, false).ToString();
            #endregion

            LoadModel(pathRight);
            _loadRight.GetLoadCollection();
            _loadRight.SetLoadCollection(_kLoad);
            // Settings generation

            foreach (var link in _linksCR)
                link.SetLoadDCLink(false);


            foreach (var gen in _generationRight)
                gen.SetPowerGeneration();

            calcRegim(Rastr);

            #region Output Text
            _fRight = freq.get_ZN(0);
            fRight.Text = Math.Round(freq.get_ZN(0), 3).ToString();

            load6.Text = GetPower(6, true).ToString();
            load9.Text = GetPower(9, true).ToString();
            load10.Text = GetPower(10, true).ToString();
            load11.Text = GetPower(11, true).ToString();
            load12.Text = GetPower(12, true).ToString();
            load13.Text = GetPower(13, true).ToString();
            load14.Text = GetPower(14, true).ToString();
            gen6.Text = GetPower(6, false).ToString();
            gen8.Text = GetPower(8, false).ToString();
            gen9.Text = GetPower(9, false).ToString();
            #endregion


            LoadModel(pathLeft);

            _loadRight.GetLoadCollection();
            _loadRight.SetLoadCollection(_kLoad);

            foreach (var link in _linksCL)
                link.SetLoadDCLink(false);

            foreach (var gen in _generationLeft)
                gen.SetPowerGeneration();


            calcRegim(Rastr);

            #region Output Text
            _fLeft = freq.get_ZN(0);
            fLeft.Text = Math.Round(freq.get_ZN(0), 3).ToString();
            #endregion

            openModel.Text = "Режим рассчитан!";
            openModel.Foreground = new SolidColorBrush(Colors.Green);

            result[0] = _fCenter;
            result[1] = _fRight;
            result[2] = _fLeft;
            return result;
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FindFreq();
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
        private static Complex GetPower(int number, bool load)
        {
            Complex result;
            if (load)
                result = new Complex(powerActiveLoad.get_ZN(GetIndex(number)),
                                     powerRectiveLoad.get_ZN(GetIndex(number)));
            else
            {
                result = new Complex(powerActiveGeneration.get_ZN(GetIndex(number)),
                                     powerRectiveGeneration.get_ZN(GetIndex(number)));
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

        public class Complex
        {
            public Complex(double x, double y)
            {
                X = x;
                Y = y;
            }
            public Complex() { }
            public double X { get; set; }
            public double Y { get; set; }
            public override string ToString()
            {
                string result = $"{X:F2}";
                if (Y > 0)
                    result += "+j∙" + $"{Y:F2}";
                else
                    result += "-j∙" + $"{Y * -1:F2}";
                return result;
            }
            public static Complex operator +(Complex arg1, Complex arg2)
            {
                return new Complex(arg1.X + arg2.X, arg1.Y + arg2.Y);
            }
            public static Complex operator +(Complex arg1, double arg2)
            {
                return new Complex(arg1.X + arg2, arg1.Y + arg2);
            }
            public static Complex operator -(Complex arg1, Complex arg2)
            {
                return new Complex(arg1.X - arg2.X, arg1.Y - arg2.Y);
            }
            public static Complex operator -(Complex arg1, double arg2)
            {
                return new Complex(arg1.X - arg2, arg1.Y - arg2);
            }
            public static Complex operator *(Complex arg1, Complex arg2)
            {
                return new Complex(arg1.X * arg2.X, arg1.Y * arg2.Y);
            }
            public static Complex operator *(Complex arg1, double arg2)
            {
                return new Complex(arg1.X * arg2, arg1.Y * arg2);
            }
            public static Complex operator /(Complex arg1, Complex arg2)
            {
                return new Complex(arg1.X / arg2.X, arg1.Y / arg2.Y);
            }
            public static Complex operator /(Complex arg1, double arg2)
            {
                return new Complex(arg1.X / arg2, arg1.Y / arg2);
            }
        }
        public class LoadCollection
        {
            public List<Load> Loads { get; set; }
            public void GetLoadCollection()
            {
                Loads = new List<Load>();
                for (int i = 0; i < Node.Count; i++)
                {
                    Loads.Add(new Load(i, new Complex(powerActiveLoad.get_ZN(i), powerRectiveLoad.get_ZN(i))));
                }
            }
            public void SetLoadCollection(double k)
            {
                foreach (var load in Loads)
                {
                    powerActiveLoad.set_ZN(load.Node, load.S.X * k);
                    powerRectiveLoad.set_ZN(load.Node, load.S.Y * k);
                }
            }
        }
        public class Load
        {
            public Load(int node, Complex s)
            {
                Node = node;
                S = s;
            }
            public int Node { get; set; }
            public Complex S { get; set; }
        }
        public class DCLink
        {
            public DCLink(string name, Complex s, int node1, int node2)
            {
                Name = name;
                S = s;
                Node1 = node1;
                Node2 = node2;
            }
            public string Name { get; set; }
            public Complex S { get; set; }
            public int Node1 { get; set; }
            public int Node2 { get; set; }
            public void SetGenerationDCLink(bool isNode1)
            {
                int node;
                if (isNode1 == true)
                    node = Node1;
                else
                    node = Node2;
                powerActiveGeneration.set_ZN(GetIndex(node), S.X);
                powerRectiveGeneration.set_ZN(GetIndex(node), S.Y);
            }
            public void SetLoadDCLink(bool isNode1)
            {
                int node;
                if (isNode1 == true)
                    node = Node1;
                else
                    node = Node2;
                powerActiveLoad.set_ZN(GetIndex(node), S.X);
                powerRectiveLoad.set_ZN(GetIndex(node), S.Y);
            }
        }

        public class Generation
        {
            public Generation(double expectation, double volatility, int node)
            {
                Expectation = expectation;
                Volatility = volatility;
                Node = node;
                ChangeGeneration();
            }
            public double Expectation { get; set; }
            public double Volatility { get; set; }
            public Complex S { get; set; }
            public int Node { get; set; }
            public void ChangeGeneration()
            {
                double P = NextGaussian(Expectation, Volatility);
                S = new Complex(P, P / 2);
            }
            public void SetPowerGeneration()
            {
                powerActiveGeneration.set_ZN(GetIndex(Node), S.X);
                powerRectiveGeneration.set_ZN(GetIndex(Node), S.Y);
            }
        }
        public static double NextGaussian(double μ, double σ)
        {
            double u, v, s;
            do
            {
                u = 2 * _random.NextDouble() - 1;
                v = 2 * _random.NextDouble() - 1;
                s = u * u + v * v;
            }
            while (u <= -1 || v <= -1 || s >= 1 || s == 0);
            double r = Math.Sqrt(-2 * Math.Log(s) / s);
            double result = μ + σ * r * u;
            if (result >= 0)
                return μ + σ * r * u;
            else
                return 0;
        }
    }
}
