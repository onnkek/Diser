using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Diagnostics;
using ASTRALib;
using System.Runtime.InteropServices;

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

        public static List<Model> models = new List<Model>();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            for (int i = 0; i < 1; i++)
            {
                Model model = InitialModel();
                CalcModel(model);
                FrequencyOptimization(model);
                models.Add(model);
            }

            stopwatch.Stop();
            test.Text = stopwatch.Elapsed.ToString();

        }

        private static Model _model;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _model = InitialModel();
            CalcModel(_model);
            fRight.Text = _model.FrequencyRight.ToString();
            fCenter.Text = _model.FrequencyCenter.ToString();
            fLeft.Text = _model.FrequencyLeft.ToString();
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FrequencyOptimization(_model);
            fRight.Text = _model.FrequencyRight.ToString();
            fCenter.Text = _model.FrequencyCenter.ToString();
            fLeft.Text = _model.FrequencyLeft.ToString();
        }
        public static void SetPowerDCLinks(Model model, double S)
        {
            foreach (var link in model.DCLinksCenterLeft)
                link.S += S;
            foreach (var link in model.DCLinksCenterRight)
                link.S += S;
        }
        public static void SetDefaultS(Model model)
        {
            foreach (var link in model.DCLinksCenterLeft)
                link.S = link.DefaultS;
            foreach (var link in model.DCLinksCenterRight)
                link.S = link.DefaultS;
        }
        public static void SetS(Model model)
        {
            foreach (var link in model.DCLinksCenterLeft)
                link.DefaultS = link.S;
            foreach (var link in model.DCLinksCenterRight)
                link.DefaultS = link.S;
        }
        public void FrequencyOptimization(Model model)
        {
            int counter = 0;
            double acceleration;
            model.StartingFrequencyCenter = model.FrequencyCenter;
            model.StartingFrequencyRight = model.FrequencyRight;
            model.StartingFrequencyLeft = model.FrequencyLeft;
            Trace.WriteLine("СТАРТ: " + (model.StartingFrequencyCenter) + "  ВПТ: " + model.DCLinksCenterRight[0].S);

            for (int i = 0; i < 20 & Math.Abs(50 - model.FrequencyCenter) > 0.001; i++)
            {
                counter++;
                double dF0 = model.FrequencyCenter;
                double k = 1;
                SetPowerDCLinks(model, 1);
                CalcModel(model);
                double dFp1 = model.FrequencyCenter;
                Trace.WriteLine("+" + 1 * k + ":" + "dFp1: " + (model.FrequencyCenter) + "           ВПТ: " + model.DCLinksCenterRight[0].S);
                SetDefaultS(model);

                // x1 - 0 y1 - F0
                // x2 - k y2 - dF1


                double K = (dFp1 - dF0) / (k - 0);

                double B = (k * dF0 - dFp1 * 0) / (k - 0);


                double step = (50 - B) / K;


                Trace.WriteLine("СУПЕРШАГ: " + step + "      ДОБАВКА ВПТ: " + k * step);
                SetPowerDCLinks(model, step);
                CalcModel(model);
                Trace.WriteLine("ЧАСТОТА СТАЛА: " + model.FrequencyCenter + "  ВПТ: " + model.DCLinksCenterRight[0].S);
                
                //double dFp1 = model.FrequencyCenter - 50;
                //Trace.WriteLine("+" + 1 * k +":" +   "dFp1: " + (model.FrequencyCenter) + "           ВПТ: " + model.DCLinksCenterRight[0].S);
                //SetDefaultS(model);

                //SetPowerDCLinks(model, -1 * k);
                //CalcModel(model);
                ////Trace.WriteLine("ВПТ: " + model.DCLinksCenterRight[0].S);
                //double dFm1 = model.FrequencyCenter - 50;
                //Trace.WriteLine("-" + 1 * k + ":" + "dFm1: " + (model.FrequencyCenter) + "           ВПТ: " + model.DCLinksCenterRight[0].S);
                //SetDefaultS(model);

                //double optStep = -(dFm1 - dFp1) / (2 * dFm1 - 4 * dF0 + 2 * dFp1);



                //Trace.WriteLine("СУПЕРШАГ: " + optStep + "      ДОБАВКА ВПТ: " + k * optStep);


                //SetPowerDCLinks(model, k * optStep);
                //CalcModel(model);
                ////SetDefaultS(model);
                SetS(model);
                test.Text = counter.ToString();
            }
            


            //for (int i = 0; i < 20 & Math.Abs(50 - model.FrequencyCenter) > 0.001; i++)
            //{
            //    double dF = 50 - model.FrequencyCenter;

            //    double kF = dF / 50;
            //    double dS;
            //    double ass;
            //    //if (model.FrequencyCenter > 50)
            //    //    ass = 50 / Math.Abs(model.FrequencyCenter);
            //    //else
            //    //    ass = Math.Abs(model.FrequencyCenter) / 50;
            //    ass = 1;
            //    if (dF < 0)
            //    {
            //        dS = 1 / Math.Abs(kF) * ass;
            //        SetPowerDCLinks(model, dS);
            //    }
            //    else
            //    {
            //        dS = 1 / Math.Abs(kF) * ass;
            //        SetPowerDCLinks(model, dS);
            //    }


            //    CalcModel(model);
            //    counter++;
            //    Trace.WriteLine(model.FrequencyCenter + "  dS: " + dS + "  ВПТ: " + model.DCLinksCenterRight[0].S);
            //}
            //test.Text = counter.ToString();

            //double c = 2;
            //double dF; 
            //for (int i = 0; i < 50 & Math.Abs(50 - model.FrequencyCenter) > 0.001; i++)
            //{
            //    dF = 50 - model.FrequencyCenter;


            //    if (dF < 0)
            //    {
            //        foreach (var link in model.DCLinksCenterLeft)
            //            link.S -= 100 / c;
            //        foreach (var link in model.DCLinksCenterRight)
            //            link.S -= 100 / c;
            //    }
            //    else
            //    {
            //        foreach (var link in model.DCLinksCenterLeft)
            //            link.S += 100 / c;
            //        foreach (var link in model.DCLinksCenterRight)
            //            link.S += 100 / c;
            //    }
            //    CalcModel(model);
            //    dF = 50 - model.FrequencyCenter;
            //    if (dF < 0)
            //        c /= 2;
            //    else
            //        c *= 2;
            //    Trace.WriteLine(model.FrequencyCenter);
            //    counter++;

            //}
            //test.Text = counter.ToString();



        }
        public Model InitialModel()
        {
            Model model = new Model();
            model.Kload = 0.5;

            // Create DCLink with default parameters
            model.DCLinksCenterRight.Add(new DCLink("C6-R1", new Complex(43.9, 21.8), 6, 1));
            model.DCLinksCenterRight.Add(new DCLink("C7-R2", new Complex(28.1, 1.6), 7, 2));
            model.DCLinksCenterRight.Add(new DCLink("C9-R3", new Complex(16.2, 3.6), 9, 3));

            model.DCLinksCenterLeft.Add(new DCLink("C26-L1", new Complex(43.9, 21.8), 26, 1));
            model.DCLinksCenterLeft.Add(new DCLink("C27-L2", new Complex(28.1, 1.6), 27, 2));
            model.DCLinksCenterLeft.Add(new DCLink("C29-L3", new Complex(16.2, 3.6), 29, 3));

            model.GenerationCenter.Add(new Generation(20, 20, 3));
            model.GenerationCenter.Add(new Generation(20, 20, 23));

            double P = 30;

            model.GenerationRight.Add(new Generation(P, P, 6));
            model.GenerationRight.Add(new Generation(P, P, 8));
            model.GenerationRight.Add(new Generation(P, P, 9));

            model.GenerationLeft.Add(new Generation(P, P, 6));
            model.GenerationLeft.Add(new Generation(P, P, 8));
            model.GenerationLeft.Add(new Generation(P, P, 9));

            return model;
        }
        public void CalcModel(Model model)
        {
            LoadModel(model, Model.pathCenter);
            model.LoadCenter.GetLoadCollection(model.Rastr);
            model.LoadCenter.SetLoadCollection(model.Kload, model.Rastr);
            // Settings generation
            foreach (var link in model.DCLinksCenterRight)
                link.SetGenerationDCLink(true, model.Rastr);
            foreach (var link in model.DCLinksCenterLeft)
                link.SetGenerationDCLink(true, model.Rastr);

            foreach (var gen in model.GenerationCenter)
                gen.SetPowerGeneration(model.Rastr);


            CalcRegim(model.Rastr);

            #region Output Text
            //dc56.Text = model.DCLinksCenterRight[0].S.ToString();
            //dc47.Text = model.DCLinksCenterRight[1].S.ToString();
            //dc49.Text = model.DCLinksCenterRight[2].S.ToString();

            ICol frequency = model.Rastr.Tables.Item("islands").Cols.Item("f");
            model.FrequencyCenter = frequency.get_ZN(0);
            //fCenter.Text = Math.Round(frequency.get_ZN(0), 3).ToString();
            //load2.Text = GetPower(2, true, rastr).ToString();
            //load3.Text = GetPower(3, true, rastr).ToString();
            //load4.Text = GetPower(4, true, rastr).ToString();
            //load5.Text = GetPower(5, true, rastr).ToString();
            //gen2.Text = GetPower(2, false, rastr).ToString();
            //gen3.Text = GetPower(3, false, rastr).ToString();
            #endregion

            LoadModel(model, Model.pathRight);
            model.LoadRight.GetLoadCollection(model.Rastr);
            model.LoadRight.SetLoadCollection(model.Kload, model.Rastr);
            // Settings generation

            foreach (var link in model.DCLinksCenterRight)
                link.SetLoadDCLink(false, model.Rastr);


            foreach (var gen in model.GenerationRight)
                gen.SetPowerGeneration(model.Rastr);

            CalcRegim(model.Rastr);

            #region Output Text
            frequency = model.Rastr.Tables.Item("islands").Cols.Item("f");
            model.FrequencyRight = frequency.get_ZN(0);
            //fRight.Text = Math.Round(frequency.get_ZN(0), 3).ToString();

            //load6.Text = GetPower(6, true, rastr).ToString();
            //load9.Text = GetPower(9, true, rastr).ToString();
            //load10.Text = GetPower(10, true, rastr).ToString();
            //load11.Text = GetPower(11, true, rastr).ToString();
            //load12.Text = GetPower(12, true, rastr).ToString();
            //load13.Text = GetPower(13, true, rastr).ToString();
            //load14.Text = GetPower(14, true, rastr).ToString();
            //gen6.Text = GetPower(6, false, rastr).ToString();
            //gen8.Text = GetPower(8, false, rastr).ToString();
            //gen9.Text = GetPower(9, false, rastr).ToString();
            #endregion


            LoadModel(model, Model.pathLeft);

            model.LoadLeft.GetLoadCollection(model.Rastr);
            model.LoadLeft.SetLoadCollection(model.Kload, model.Rastr);

            foreach (var link in model.DCLinksCenterLeft)
                link.SetLoadDCLink(false, model.Rastr);

            foreach (var gen in model.GenerationLeft)
                gen.SetPowerGeneration(model.Rastr);


            CalcRegim(model.Rastr);

            #region Output Text
            frequency = model.Rastr.Tables.Item("islands").Cols.Item("f");
            model.FrequencyLeft = frequency.get_ZN(0);
            //fLeft.Text = Math.Round(frequency.get_ZN(0), 3).ToString();
            #endregion

            //openModel.Text = "Режим рассчитан!";
            //openModel.Foreground = new SolidColorBrush(Colors.Green);
        }


        public void LoadModel(Model model, string path)
        {
            model.Rastr = new Rastr();
            model.Rastr.Load(RG_KOD.RG_REPL, path, "");
        }



        bool CalcRegim(IRastr inRastr)
        {
            ICol vras = inRastr.Tables.Item("node").Cols.Item("vras");
            ICol delta = inRastr.Tables.Item("node").Cols.Item("vras");
            for (int i = 0; i < inRastr.Tables.Item("node").Count; i++)
            {
                vras.set_ZN(i, 0);
                delta.set_ZN(i, 0);
            }
            ICol statusRgm = inRastr.Tables.Item("com_regim").Cols.Item("status");
            inRastr.rgm(""); //Расчет режима 
            int a = statusRgm.get_ZN(0);
            if (a == 0)
                return true;
            else
            {
                //MessageBox.Show("error");
                return false;
            }
        }
        private static Complex GetPower(int number, bool load, IRastr rastr)
        {
            Complex result;
            ICol activeGeneration = rastr.Tables.Item("node").Cols.Item("pg");
            ICol rectiveGeneration = rastr.Tables.Item("node").Cols.Item("qg");
            ICol activeLoad = rastr.Tables.Item("node").Cols.Item("pn");
            ICol rectiveLoad = rastr.Tables.Item("node").Cols.Item("qn");

            if (load)
                result = new Complex(activeLoad.get_ZN(GetIndex(number, rastr)),
                                     rectiveLoad.get_ZN(GetIndex(number, rastr)));
            else
            {
                result = new Complex(activeGeneration.get_ZN(GetIndex(number, rastr)),
                                     rectiveGeneration.get_ZN(GetIndex(number, rastr)));
            }
            return result;
        }

        private static int GetIndex(int number, IRastr rastr)
        {
            int result = -1;
            for (int i = 0; i < rastr.Tables.Item("node").Count; i++)
            {
                if (rastr.Tables.Item("node").Cols.Item("ny").ZN[i] == number)
                    result = i;
            }
            return result;
        }
        public class Model
        {
            public Model()
            {
                DCLinksCenterRight = new List<DCLink>();
                DCLinksCenterLeft = new List<DCLink>();
                GenerationCenter = new List<Generation>();
                GenerationRight = new List<Generation>();
                GenerationLeft = new List<Generation>();
                LoadCenter = new LoadCollection();
                LoadRight = new LoadCollection();
                LoadLeft = new LoadCollection();
            }
            public static string pathCenter = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Center.rg2";
            public static string pathRight = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Right.rg2";
            public static string pathLeft = @"C:\Users\Агро\Desktop\С винды мака 30.08.22\Дисер\Делим\Left.rg2";
            public IRastr Rastr { get; set; }
            public double FrequencyCenter { get; set; }
            public double FrequencyRight { get; set; }
            public double FrequencyLeft { get; set; }
            public List<DCLink> DCLinksCenterRight { get; set; }
            public List<DCLink> DCLinksCenterLeft { get; set; }
            public List<Generation> GenerationCenter { get; set; }
            public List<Generation> GenerationRight { get; set; }
            public List<Generation> GenerationLeft { get; set; }
            public LoadCollection LoadCenter { get; set; }
            public LoadCollection LoadRight { get; set; }
            public LoadCollection LoadLeft { get; set; }
            public Random Random { get; set; }
            public double Kload { get; set; }
            public double StartingFrequencyCenter { get; set; }
            public double StartingFrequencyRight { get; set; }
            public double StartingFrequencyLeft { get; set; }

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
            public void GetLoadCollection(IRastr rastr)
            {
                Loads = new List<Load>();
                ICol activeLoad = rastr.Tables.Item("node").Cols.Item("pn");
                ICol rectiveLoad = rastr.Tables.Item("node").Cols.Item("qn");
                for (int i = 0; i < rastr.Tables.Item("node").Count; i++)
                {
                    Loads.Add(new Load(i, new Complex(activeLoad.get_ZN(i),
                                                      rectiveLoad.get_ZN(i))));
                }
            }
            public void SetLoadCollection(double k, IRastr rastr)
            {
                ICol activeLoad = rastr.Tables.Item("node").Cols.Item("pn");
                ICol rectiveLoad = rastr.Tables.Item("node").Cols.Item("qn");
                foreach (var load in Loads)
                {
                    activeLoad.set_ZN(load.Node, load.S.X * k);
                    rectiveLoad.set_ZN(load.Node, load.S.Y * k);
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
                DefaultS = s;
                Node1 = node1;
                Node2 = node2;
            }
            public string Name { get; set; }
            public Complex S { get; set; }
            public Complex DefaultS { get; set; }
            public Complex ModS
            {
                get
                {
                    return new Complex(Math.Abs(S.X), Math.Abs(S.Y));
                }
                set { }
            }

            public int Node1 { get; set; }
            public int Node2 { get; set; }
            public void SetGenerationDCLink(bool isNode1, IRastr rastr)
            {
                int node;
                if (isNode1 == true)
                    node = Node1;
                else
                    node = Node2;


                ICol activeGeneration = rastr.Tables.Item("node").Cols.Item("pg");
                ICol rectiveGeneration = rastr.Tables.Item("node").Cols.Item("qg");
                activeGeneration.set_ZN(GetIndex(node, rastr), S.X);
                rectiveGeneration.set_ZN(GetIndex(node, rastr), S.Y);
            }
            public void SetLoadDCLink(bool isNode1, IRastr rastr)
            {
                int node;
                if (isNode1 == true)
                    node = Node1;
                else
                    node = Node2;

                ICol activeLoad = rastr.Tables.Item("node").Cols.Item("pn");
                ICol rectiveLoad = rastr.Tables.Item("node").Cols.Item("qn");
                activeLoad.set_ZN(GetIndex(node, rastr), S.X);
                rectiveLoad.set_ZN(GetIndex(node, rastr), S.Y);
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
            public void SetPowerGeneration(IRastr rastr)
            {
                ICol activeGeneration = rastr.Tables.Item("node").Cols.Item("pg");
                ICol rectiveGeneration = rastr.Tables.Item("node").Cols.Item("qg");
                activeGeneration.set_ZN(GetIndex(Node, rastr), S.X);
                rectiveGeneration.set_ZN(GetIndex(Node, rastr), S.Y);
            }
        }
        public static double NextGaussian(double μ, double σ)
        {
            Random random = new Random();
            double u, v, s;
            do
            {
                u = 2 * random.NextDouble() - 1;
                v = 2 * random.NextDouble() - 1;
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
