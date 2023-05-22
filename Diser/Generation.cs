using ASTRALib;

namespace Diser
{
    internal abstract class Generation
    {
        public Generation(string name, double expectation, double volatility, int node, int n, int nPark)
        {
            Name = name;
            Expectation = expectation;
            Volatility = volatility;
            Node = node;
            N = n;
            NPark = nPark;
            S = new Complex(0, 0);
            GetPower();
            //ChangeGeneration();
        }
        public string Name { get; set; }
        public double Expectation { get; set; }
        public double Volatility { get; set; }
        public Complex S { get; set; }
        public int Node { get; set; }
        public int N { get; set; }
        public int NPark { get; set; }


        public abstract void GetPower();
        //{
            //double P = Worker.NextGaussian(Expectation, Volatility);
            //S = new Complex(P, P / 2);
        //}
        public void SetPowerGeneration(IRastr rastr)
        {
            ICol activeGeneration = rastr.Tables.Item("node").Cols.Item("pg");
            ICol rectiveGeneration = rastr.Tables.Item("node").Cols.Item("qg");
            activeGeneration.set_ZN(Worker.GetIndex(Node), S.X);
            rectiveGeneration.set_ZN(Worker.GetIndex(Node), S.Y);
        }
    }
}
