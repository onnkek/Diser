using System.Windows;

namespace Diser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        // Путь для файлов моделей
        private static string pathCenter = @"C:\Center.rg2";
        private static string pathRight = @"C:\Right.rg2";
        private static string pathLeft = @"C:\Left.rg2";
        // Путь для вывода XLSX
        private static string pathOutput = @"C:\Users\Agro\Desktop\Универ\Дисер\Result";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Worker worker = new Worker(1, pathCenter, pathRight, pathLeft, pathOutput);
            worker.InitModel();
            worker.CalcModel();
            worker.FrequencyOptimization();
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Worker worker = new Worker(1, pathCenter, pathRight, pathLeft, pathOutput);
            worker.MultiCalc(1000, true);
        }
    }
}
