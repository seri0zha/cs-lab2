using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace Lab_2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static MainWindow AppWindow;
        public MainWindow()
        {
            AppWindow = this;
            AppWindow.ResizeMode = ResizeMode.NoResize;
            if (File.Exists(@"thrlist.xlsx"))
            {
                ShowList();
                /*List<TableItem> table = new List<TableItem>();
                table.Add(new TableItem() { Name = "Test" });
                ThreatList.ItemsSource = table;*/
            }
            else
            {
                AsyncDownload();
            }
            InitializeComponent();

        }

        public async void AsyncDownload()
        {
            await Task.Run(() =>
            {
                GetExcel.Download();
                CreateDataBase();
                this.Dispatcher.Invoke(() =>
                {
                    ShowList();
                });
            });
        }

        public static void CreateDataBase()
        {
            using (var threatList = new ExcelPackage(GetExcel.sourcePath))
            {
                ExcelWorksheet sourceSheet = threatList.Workbook.Worksheets[1]; // Worksheet скачанного файла       
                sourceSheet.DeleteColumn(GetExcel.EndCol(sourceSheet) - 1, 2); // Удаление даты из локальной базы
                threatList.Save();
            }
        }

        public static void ShowList()
        {
            List<TableItem> table = new List<TableItem>();
            using (var dataBase = new ExcelPackage(GetExcel.sourcePath))
            {
                ExcelWorksheet sheet = dataBase.Workbook.Worksheets[1];
                for (int i = sheet.Dimension.Start.Row + 2; i <= sheet.Dimension.End.Row; i++)
                {
                    List<string> args = new List<string>();
                    var range = sheet.Cells[i, 1, i, sheet.Dimension.End.Column];
                    foreach (var el in range)
                    {
                        args.Add(el.Value.ToString());
                    }
                    args[0] = "УД." + Int32.Parse(args[0]).ToString("000");
                    table.Add(new TableItem(args));
                }
            }

        }

        private void ThreatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = (ListBox)sender;
            var test = (TableItem)item.SelectedItem;
            MessageBox.Show(test.ToString());
        }
    }
}
