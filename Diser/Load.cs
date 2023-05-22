namespace Diser
{
    internal class Load
    {
        public Load(int node, Complex s)
        {
            Node = node;
            S = s;
        }
        public int Node { get; set; }
        public Complex S { get; set; }

    }
}
