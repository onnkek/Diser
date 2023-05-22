using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace Diser
{
    internal class SolarGeneration : Generation
    {
        public SolarGeneration(string name, double expectation, double volatility, int node, int n, int nPark) : base(name, expectation, volatility, node, n, nPark)
        {

        }
        public override void GetPower()
        {
            
            double k = 1.38 * Math.Pow(10, -23);
            double q = 1.602 * Math.Pow(10, -19);

            double T = 298.15;
            double Tref = 298.15;
            double Isc = 8.74;
            double Uoc = 37.3;
            double KI = 0.06;
            double Imm = 8.21;
            double Umm = 30.5;
            double Nc = 60;

            double Vt = k * T / q;
            double Eg = 1.16 - 0.000702 * (Math.Pow(T, 2) / (T - 1108));
            double aref = q * (2 * Umm - Uoc) / (Nc * k * Tref * (Isc / (Isc - Imm) + Math.Log(1 - Imm / Isc)));
            double a = aref * T / Tref;

            double I0ref = Isc / (Math.Exp(Uoc / (Nc * aref * Vt)) - 1);
            double I0 = I0ref * Math.Pow(T / Tref, 3) * Math.Exp(Eg / (a * Vt) * (1 - Tref / T));


            for (int i = 0; i < NPark; i++) // Генерация разной мощности внутри солнцепарка
            {
                double G = Worker.NextGaussian(Expectation, Volatility);
                
                double Iph = G / 1000 * (Isc + KI * (T - Tref));
                double I = Iph - I0 * (Math.Exp(Umm / (Nc * a * Vt)) - 1);
                double P = I * Umm * Math.Pow(10, -6);
                S += new Complex(P, 0) * N;
            }
        }
    }
}
