using ASTRALib;
using System;

namespace Diser
{
    internal class DCLink
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
            activeGeneration.set_ZN(Worker.GetIndex(node), S.X);
            rectiveGeneration.set_ZN(Worker.GetIndex(node), S.Y);
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
            activeLoad.set_ZN(Worker.GetIndex(node), S.X);
            rectiveLoad.set_ZN(Worker.GetIndex(node), S.Y);
        }
    }
}
