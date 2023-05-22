using ASTRALib;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Diser
{
    internal class Worker
    {
        public Worker(double kload, string pathCenter, string pathRight, string pathLeft, string pathOutput)
        {
            Kload = kload;
            PathCenter = pathCenter;
            PathRight = pathRight;
            PathLeft = pathLeft;
            PathOutput = pathOutput;
            Random = new Random();
            Rastr = new Rastr();
            Models = new List<Model>();
        }
        public static IRastr Rastr { get; set; }
        public Model Model { get; set; }
        public List<Model> Models { get; set; }
        public double Kload { get; set; }
        public string PathCenter { get; set; }
        public string PathRight { get; set; }
        public string PathLeft { get; set; }
        public string PathOutput { get; set; }
        public static Random Random { get; set; }
        public void InitModel()
        {
            Model = new Model();
            Model.Kload = Kload;

            // Create DCLink with default parameters
            Model.DCLinksCenterRight.Add(new DCLink("C6-R1", new Complex(43.9, 21.8), 6, 1));
            Model.DCLinksCenterRight.Add(new DCLink("C7-R2", new Complex(28.1, 1.6), 7, 2));
            Model.DCLinksCenterRight.Add(new DCLink("C9-R3", new Complex(16.2, 3.6), 9, 3));

            Model.DCLinksCenterLeft.Add(new DCLink("C26-L1", new Complex(43.9, 21.8), 26, 1));
            Model.DCLinksCenterLeft.Add(new DCLink("C27-L2", new Complex(28.1, 1.6), 27, 2));
            Model.DCLinksCenterLeft.Add(new DCLink("C29-L3", new Complex(16.2, 3.6), 29, 3));
            
            //Model.GenerationCenter.Add(new WindGeneration("1_Center", 1.85, 5.18, 1, 15, 20));
            

            Model.GenerationCenter.Add(new SolarGeneration("3_Center", 500, 100, 3, 10000, 30));
            Model.GenerationCenter.Add(new SolarGeneration("23_Center", 500, 100, 23, 10000, 30));
            
            //Model.GenerationRight.Add(new SolarGeneration("2_Right", 500, 100, 2, 10000, 20));
            
            Model.GenerationRight.Add(new SolarGeneration("6_Right", 500, 100, 6, 10000, 10));
            Model.GenerationRight.Add(new WindGeneration("8_Right", 1.85, 5.18, 8, 10, 10));
            Model.GenerationRight.Add(new WindGeneration("9_Right", 1.85, 5.18, 9, 5, 30));

            //Model.GenerationLeft.Add(new SolarGeneration("2_Right", 500, 100, 2, 10000, 20));
            Model.GenerationLeft.Add(new SolarGeneration("6_Left", 500, 100, 6, 10000, 10));
            Model.GenerationLeft.Add(new WindGeneration("8_Left", 1.85, 5.18, 8, 10, 10));
            Model.GenerationLeft.Add(new WindGeneration("9_Left", 1.85, 5.18, 9, 5, 30));
        }
        private bool CalcRastr()
        {
            ICol vras = Rastr.Tables.Item("node").Cols.Item("vras");
            ICol delta = Rastr.Tables.Item("node").Cols.Item("vras");
            for (int i = 0; i < Rastr.Tables.Item("node").Count; i++)
            {
                vras.set_ZN(i, 0);
                delta.set_ZN(i, 0);
            }
            ICol statusRgm = Rastr.Tables.Item("com_regim").Cols.Item("status");
            Rastr.rgm(""); //Расчет режима 
            int a = statusRgm.get_ZN(0);
            if (a == 0)
                return true;
            else
            {
                //MessageBox.Show("error");
                return false;
            }
        }
        public void CalcModel()
        {
            //Trace.WriteLine("Старт расчета модели");

            Rastr.Load(RG_KOD.RG_REPL, PathCenter, "");

            // Задание в узел 1 генерации равной 0, обход сохранения, слетела лицензия //
            ICol activeGeneration = Rastr.Tables.Item("node").Cols.Item("pg");         //
            ICol rectiveGeneration = Rastr.Tables.Item("node").Cols.Item("qg");        //
            activeGeneration.set_ZN(Worker.GetIndex(1), 160);                          //
            rectiveGeneration.set_ZN(Worker.GetIndex(1), 80);                          //
            /////////////////////////////////////////////////////////////////////////////

            Model.LoadCenter.GetLoadCollection(Rastr);
            Model.LoadCenter.SetLoadCollection(Model.Kload, Rastr);

            // Settings generation
            foreach (var link in Model.DCLinksCenterRight)
                link.SetGenerationDCLink(true, Rastr);
            foreach (var link in Model.DCLinksCenterLeft)
                link.SetGenerationDCLink(true, Rastr);

            foreach (var gen in Model.GenerationCenter)
                gen.SetPowerGeneration(Rastr);

            CalcRastr();

            #region Output Text
            ICol frequency = Rastr.Tables.Item("islands").Cols.Item("f");
            Model.FrequencyCenter = frequency.get_ZN(0);
            #endregion

            Rastr.Load(RG_KOD.RG_REPL, PathRight, "");
            Model.LoadRight.GetLoadCollection(Rastr);
            Model.LoadRight.SetLoadCollection(Model.Kload, Rastr);
            // Settings generation
            foreach (var link in Model.DCLinksCenterRight)
                link.SetLoadDCLink(false, Rastr);
            foreach (var gen in Model.GenerationRight)
                gen.SetPowerGeneration(Rastr);

            CalcRastr();

            #region Output Text
            frequency = Rastr.Tables.Item("islands").Cols.Item("f");
            Model.FrequencyRight = frequency.get_ZN(0);
            #endregion

            Rastr.Load(RG_KOD.RG_REPL, PathLeft, "");
            Model.LoadLeft.GetLoadCollection(Rastr);
            Model.LoadLeft.SetLoadCollection(Model.Kload, Rastr);
            foreach (var link in Model.DCLinksCenterLeft)
                link.SetLoadDCLink(false, Rastr);

            foreach (var gen in Model.GenerationLeft)
                gen.SetPowerGeneration(Rastr);

            CalcRastr();

            #region Output Text
            frequency = Rastr.Tables.Item("islands").Cols.Item("f");
            Model.FrequencyLeft = frequency.get_ZN(0);
            #endregion

        }
        public void FrequencyOptimization()
        {
            Trace.WriteLine("Старт оптимизации частоты");
            int counter = 0;
            Model.StartingFrequencyCenter = Model.FrequencyCenter;
            Model.StartingFrequencyRight = Model.FrequencyRight;
            Model.StartingFrequencyLeft = Model.FrequencyLeft;
            //Trace.WriteLine("СТАРТ: " + (model.StartingFrequencyCenter) + "  ВПТ: " + model.DCLinksCenterRight[0].S);

            for (int i = 0; i < 20 & Math.Abs(50 - Model.FrequencyCenter) > 0.01; i++)
            {
                //Trace.WriteLine($"Итерация {i}: F = {Model.FrequencyCenter}");
                counter++;
                double dF0 = Model.FrequencyCenter;
                double k = 1;
                Model.SetPowerDCLinks(k);
                CalcModel();
                double dFp1 = Model.FrequencyCenter;
                //Trace.WriteLine("+" + 1 * k + ":" + "dFp1: " + (model.FrequencyCenter) + "           ВПТ: " + model.DCLinksCenterRight[0].S);
                Model.SetDefaultS();

                // x1 - 0 y1 - F0
                // x2 - k y2 - dF1
                double K = (dFp1 - dF0) / (k - 0);
                double B = (k * dF0 - dFp1 * 0) / (k - 0);
                double step = (50 - B) / K;

                //Trace.WriteLine("СУПЕРШАГ: " + step + "      ДОБАВКА ВПТ: " + k * step);
                Model.SetPowerDCLinks(step);
                CalcModel();
                //Trace.WriteLine("ЧАСТОТА СТАЛА: " + model.FrequencyCenter + "  ВПТ: " + model.DCLinksCenterRight[0].S);

                Model.SetS();
                //MainWindow.test.Text = counter.ToString();
            }
        }
        public void AngleOptimization()
        {
            Model.AngleFrequencyRight = Model.FrequencyRight;
            Model.AngleFrequencyLeft = Model.FrequencyLeft;
            
            for (int i = 0; i < 20 & Model.FrequencyRight > 50; i++)
            {
                double delta = 1;
                if (Model.FrequencyRight < 55)
                    delta = 0.5;
                foreach (Generation gen in Model.GenerationRight)
                {
                    if (gen is WindGeneration)
                    {
                        WindGeneration wg = (WindGeneration)gen;
                        wg.ChangeAngle(wg.Angle + delta);
                    }
                }
                CalcModel();
            }
            for (int i = 0; i < 20 & Model.FrequencyLeft > 50; i++)
            {
                double delta = 1;
                if (Model.FrequencyLeft < 55)
                    delta = 0.5;
                foreach (Generation gen in Model.GenerationLeft)
                {
                    if (gen is WindGeneration)
                    {
                        WindGeneration wg = (WindGeneration)gen;
                        wg.ChangeAngle(wg.Angle + delta);
                    }
                }
                CalcModel();
            }
        }

        public void MultiCalc(int number, bool abgleOptimization)
        {
            Models.Clear();
            for (int i = 0; i < number; i++)
            {
                InitModel();
                CalcModel();
                FrequencyOptimization();
                if (abgleOptimization)
                    AngleOptimization();
                Models.Add(Model);
                Trace.WriteLine($"Iteration {i} - Done!  Fs={Model.StartingFrequencyCenter}   F={Model.FrequencyCenter}");
            }
            SaveXLS();
        }
        private static Complex GetPower(int number, bool load)
        {
            Complex result;
            ICol activeGeneration = Rastr.Tables.Item("node").Cols.Item("pg");
            ICol rectiveGeneration = Rastr.Tables.Item("node").Cols.Item("qg");
            ICol activeLoad = Rastr.Tables.Item("node").Cols.Item("pn");
            ICol rectiveLoad = Rastr.Tables.Item("node").Cols.Item("qn");

            if (load)
                result = new Complex(activeLoad.get_ZN(GetIndex(number)),
                                     rectiveLoad.get_ZN(GetIndex(number)));
            else
            {
                result = new Complex(activeGeneration.get_ZN(GetIndex(number)),
                                     rectiveGeneration.get_ZN(GetIndex(number)));
            }
            return result;
        }
        public static int GetIndex(int number)
        {
            int result = -1;
            for (int i = 0; i < Rastr.Tables.Item("node").Count; i++)
            {
                if (Rastr.Tables.Item("node").Cols.Item("ny").ZN[i] == number)
                    result = i;
            }
            return result;
        }
        public static double NextGaussian(double μ, double σ)
        {
            double result = -1;

            while (result < 0)
            {
                double u, v, s;
                do
                {
                    u = 2 * Random.NextDouble() - 1;
                    v = 2 * Random.NextDouble() - 1;
                    s = u * u + v * v;
                }
                while (u <= -1 || v <= -1 || s >= 1 || s == 0);
                double r = Math.Sqrt(-2 * Math.Log(s) / s);
                result = μ + σ * r * u;
            }
            return result;
        }
        public static double NextVaibull(double c, double b)
        {
            double r, x;
            r = Random.NextDouble();
            x = b * Math.Pow(-Math.Log(r), 1 / c);
            return x;
        }
        public void SaveXLS()
        {
            Excel.Application excel = new Excel.Application();
            Workbook book = excel.Workbooks.Add(Type.Missing);
            Worksheet sheet = book.ActiveSheet;

            sheet.Cells[1, 1].Value = "N";
            sheet.Cells[1, 2].Value = "F_center";
            sheet.Cells[1, 3].Value = "F_left";
            sheet.Cells[1, 4].Value = "F_right";
            sheet.Cells[1, 5].Value = "F_center_Reg";
            sheet.Cells[1, 6].Value = "F_left_Reg";
            sheet.Cells[1, 7].Value = "F_right_Reg";
            sheet.Cells[1, 8].Value = "F_left_Angle";
            sheet.Cells[1, 9].Value = "F_right_Angle";
            sheet.Cells[1, 10].Value = "Angle_R";
            sheet.Cells[1, 11].Value = "Angle_L";

            int start;
            List<Generation> generation = new List<Generation>();
            List<DCLink> dcLinks = new List<DCLink>();
            foreach (var model in Models)
            {
                sheet.Cells[2 + Models.IndexOf(model), 1].Value = 1 + Models.IndexOf(model);
                sheet.Cells[2 + Models.IndexOf(model), 2].Value = model.StartingFrequencyCenter;
                sheet.Cells[2 + Models.IndexOf(model), 3].Value = model.StartingFrequencyLeft;
                sheet.Cells[2 + Models.IndexOf(model), 4].Value = model.StartingFrequencyRight;
                sheet.Cells[2 + Models.IndexOf(model), 5].Value = model.FrequencyCenter;
                sheet.Cells[2 + Models.IndexOf(model), 6].Value = model.AngleFrequencyLeft;
                sheet.Cells[2 + Models.IndexOf(model), 7].Value = model.AngleFrequencyRight;
                sheet.Cells[2 + Models.IndexOf(model), 8].Value = model.FrequencyLeft;
                sheet.Cells[2 + Models.IndexOf(model), 9].Value = model.FrequencyRight;
                sheet.Cells[2 + Models.IndexOf(model), 10].Value = ((WindGeneration)model.GenerationRight[1]).Angle;
                sheet.Cells[2 + Models.IndexOf(model), 11].Value = ((WindGeneration)model.GenerationLeft[1]).Angle;

                generation.Clear();
                generation.AddRange(model.GenerationCenter);
                generation.AddRange(model.GenerationRight);
                generation.AddRange(model.GenerationLeft);
                start = 12;
                for (int i = start; i < generation.Count + start; i++)
                {
                    sheet.Cells[2 + Models.IndexOf(model), (i - start) * 2 + start].Value = generation[i - start].S.X;
                    sheet.Cells[2 + Models.IndexOf(model), (i - start) * 2 + start + 1].Value = generation[i - start].S.Y;
                }

                dcLinks.Clear();
                dcLinks.AddRange(model.DCLinksCenterLeft);
                dcLinks.AddRange(model.DCLinksCenterRight);
                start += generation.Count * 2;
                for (int i = start; i < start + dcLinks.Count; i++)
                {
                    sheet.Cells[2 + Models.IndexOf(model), (i - start) * 2 + start].Value = dcLinks[i - start].S.X;
                    sheet.Cells[2 + Models.IndexOf(model), (i - start) * 2 + start + 1].Value = dcLinks[i - start].S.Y;
                }
                Trace.WriteLine("Save " + Models.IndexOf(model) + " - Done!");
            }
            
            
            start = 12;
            for (int i = start; i < start + generation.Count; i++)
            {
                sheet.Cells[1, (i - start) * 2 + start].Value = "P_" + generation[i - start].Name;
                sheet.Cells[1, (i - start) * 2 + start + 1].Value = "Q_" + generation[i - start].Name;
            }
            start += generation.Count * 2;
            for (int i = start; i < start + dcLinks.Count; i++)
            {
                sheet.Cells[1, (i - start) * 2 + start].Value = "P_" + dcLinks[i - start].Name;
                sheet.Cells[1, (i - start) * 2 + start + 1].Value = "Q_" + dcLinks[i - start].Name;
            }






            sheet.Columns.AutoFit();

            Directory.CreateDirectory(PathOutput);
            var path = PathOutput + $@"\K={Kload}-{DateTime.Now:dd.MM.yy-HH.mm.ss}.xlsx";
            excel.Application.ActiveWorkbook.SaveAs(path);
            book.Close();
        }
    }
}
