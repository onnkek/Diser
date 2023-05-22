namespace Diser
{
    internal class Complex
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
}
