using ASTRALib;
using System.Collections.Generic;

namespace Diser
{
    internal class LoadCollection
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
}
