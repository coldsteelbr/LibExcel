using Microsoft.Win32;
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

using LibExcel;

namespace ExcelTester
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ExcelReader mExcelReader;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void b_openExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx"
            };
            if (dialog.ShowDialog() == true)
            {
                txt_fileName.Text = dialog.FileName;
            }
        }
        private static string ValueArrayToString(object[,] values)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                for (int j = 1; j <= values.GetLength(1); j++)
                {
                    builder.Append(values[i, j]).Append("\t");
                }
                builder.Append("\n");
            }
            builder.Append("\n");
            return builder.ToString();
        }

        private void b_start_Click(object sender, RoutedEventArgs e)
        {
            mExcelReader = new ExcelReader(txt_fileName.Text);

            // READING
            object[,] values = mExcelReader.GetSheetValues(txt_sourceSheetName.Text);
            text_output.Text = ValueArrayToString(values);

            // WRITING
            object[,] newValues = ExcelReader.GetOneBasedTwoDimenArray(values.GetLength(0), values.GetLength(1) + 1);
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                for (int j = 1; j <= values.GetLength(1); j++)
                {
                    newValues[i, j] = values[i, j];
                }
            }

            int k = newValues.GetLength(1);
            newValues[1, k] = "new_col";
            for (int i = 1; i <= newValues.GetLength(0);i++)
            {
                newValues[i, k] = "=A" + i;
            }

            text_output.Text += ValueArrayToString(newValues);

            mExcelReader.SetSheetValues(txt_sourceSheetName.Text, newValues);
        }

        private void b_close_Click(object sender, RoutedEventArgs e)
        {
            mExcelReader.SaveChangesAndClose();
        }
    }
}
