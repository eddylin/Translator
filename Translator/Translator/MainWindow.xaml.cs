using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace Translator
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            string output_path = @"E:\output.txt";

            FileInfo output_file = new FileInfo(output_path);
            if (output_file.Exists)
            {
                output_file.Delete();
            }

            output_file.Create().Close();

            FileInfo newFile = new FileInfo(@"E:\sample6.xlsx");


            if (newFile.Exists)
            {
                newFile.Delete();
            }

            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.Add("Content");

            //Headers
            ws.Cells["A1"].Value = "FileName";
            ws.Cells["B1"].Value = "String";

            string path = @"d:\g2\trunk\src\mobile\scripts\config\configlogic";

            //string path = @"E:\test";
            DirectoryInfo directory = new DirectoryInfo(path);
            FileInfo[] files = directory.GetFiles();


            int count = 1;
            foreach (var file in files)
            {
                string pattern = "(?<!--)\\[(=*)\\[.*?\\]\\1\\]|\"((\\\\.) |[^ \"\\\\])*?\" | '((\\\\.)|[^'\\\\])*?'";
                MatchCollection matches = Regex.Matches(File.ReadAllText(file.FullName), pattern, RegexOptions.Singleline);
                //[\\u4e00-\\u9fa5]+

                //"^\\[\\[([^(\\[[\\[[)]|\\r?\\n)*[\\u4e00-\\u9fa5]+(.|\\r?\\n)*\\]\\]$|\"(.|\\r?\\n)*[\\u4e00-\\u9fa5]+(.|\\r?\\n)*\""
                foreach (var match in matches)
                {
                    count += 1;
                    ws.Cells["A" + count.ToString()].Value = file.FullName;
                    ws.Cells["B" + count.ToString()].Value = match;
                }
            }

            pck.Save();
        }
    }
}
