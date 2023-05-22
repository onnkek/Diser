using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Diser
{
    internal class WindGeneration : Generation
    {
        public WindGeneration(string name, double c, double b, int node, int n, int nPark) : base(name, c, b, node, n, nPark)
        {
            Angle = 0;
        }
        public double C { get; set; }
        public double Angle { get; set; }
        public void ChangeAngle(double angle)
        {
            S /= C;
            
            double C1 = 0.5176;
            double C2 = 116;
            double C3 = 0.4;
            double C4 = 5;
            double C5 = 21;
            double C6 = 0.0068;

            double L = 8;                           // Быстроходность
            Angle = angle;                          // Угол наклона лопастей
            double _1b = 1 / (L + 0.08 * Angle) - 0.035 / (1 + Math.Pow(Angle, 3));
            C = C1 * (C2 * _1b - C3 * Angle - C4) * Math.Exp(-C5 * _1b) + C6 / _1b;  // КПД

            S *= C;

        }
        public override void GetPower()
        {
            
            double ro = 1.225;                       // Плотность воздуха
            double R = 82;                           // Радиус турбины

            double C1 = 0.5176;
            double C2 = 116;
            double C3 = 0.4;
            double C4 = 5;
            double C5 = 21;
            double C6 = 0.00068;

            double L = 8;                           // Быстроходность

            double _1b = 1 / (L + 0.08 * Angle) - 0.035 / (1 + Math.Pow(Angle, 3));

            C = C1 * (C2 * _1b - C3 * Angle - C4) * Math.Exp(-C5 * _1b) + C6 / _1b;  // КПД
            
            double At = Math.PI * Math.Pow(R, 2);    // At = pi * R ^2
            
            for(int i = 0; i < NPark; i++) // Генерация разной мощности внутри ветропарка
            {
                double V = Worker.NextVaibull(Expectation, Volatility);
                double v = V;                                                   // Скорость ветра
                double P = ro * At * Math.Pow(v, 3) * C / 2 * Math.Pow(10, -6); // в МВт
                S += new Complex(P, P / 2) * N;
            }
        }
    }
}
