using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace Lab_2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int CurrentPage { get; set; }
        public int PageCount { get; set; }

        public const int ItemsPerPage = 20;

        public MainWindow()
        {
            InitializeComponent();
            CurrentPage = 1;

            if (File.Exists(@"thrlist.xlsx") && !(new FileInfo(@"thrlist.xlsx").Length == 0))
            {
                GetPageCount();
                ShowPage(CurrentPage);
            }
            else
            {
                AsyncDownload();
            }
        }

        public Task AsyncDownload()
        {
            return Task.Run(() =>
            {
                GetExcel.Download();
                GetPageCount();
                ShowPage(CurrentPage);
            });
        }

        private void ShowPage(int pageNumber)
        {
            List<Note> page = new List<Note>();

            using (var dataBase = new ExcelPackage(GetExcel.downloadPath))
            {
                ExcelWorksheet sheet = dataBase.Workbook.Worksheets[1];

                int startRow = sheet.Dimension.Start.Row + 2;
                startRow += (pageNumber - 1) * ItemsPerPage;
                
                int endRow = startRow + ItemsPerPage - 1;
                endRow = Math.Min(endRow, sheet.Dimension.End.Row);
               
                for (int rowNumber = startRow; rowNumber <= endRow; rowNumber++)
                {
                    List<string> rowValues = new List<string>();
                    for (int columnNumber = 1; columnNumber <= sheet.Dimension.End.Column - 2; columnNumber++)
                    {
                        string cellValue = sheet.Cells[rowNumber, columnNumber].Value?.ToString();
                        cellValue = cellValue.Replace("_x000d_", "");

                        switch (columnNumber)
                        {
                            case 1:
                                cellValue = $"УБИ.{cellValue.PadLeft(3, '0')}";
                                break;
                            case 6:
                            case 7:
                            case 8:
                                cellValue = cellValue == "1" ? "Да" : "Нет";
                                break;
                        }
                        rowValues.Add(cellValue);
                    }
                    page.Add(new Note(rowValues));
                }
            }
            Dispatcher.Invoke(() => {
                ThreatList.ItemsSource = page;
                PageNumber.Text = pageNumber.ToString();
            });
        }

        private void ThreatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = (ListBox)sender;
            var noteInfo = (Note)item.SelectedItem;
            if (noteInfo != null)
            {
                DetailedDescription.ItemsSource = noteInfo.Properties;
            }
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            ThreatList.UnselectAll();
            DetailedDescription.ItemsSource = null;
            ThreatList.ItemsSource = null;
            CurrentPage = 1;
            PageNumber.Text = CurrentPage.ToString();
            StartRefreshAnimation();
            RefreshButton.IsEnabled = false;
            await AsyncDownload();
            RefreshButton.IsEnabled = true;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (CurrentPage < PageCount)
            {
                ShowPage(++CurrentPage);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (CurrentPage > 1)
            {
                ShowPage(--CurrentPage);
            }
        }

        private void StartRefreshAnimation()
        {
            RotateTransform rotatingToAnimate = this.Resources["RefreshButtonRotateTransform"] as RotateTransform;
            if (rotatingToAnimate is null) return;

            DoubleAnimation rotateAnimation = new DoubleAnimation()
            {
                FillBehavior = FillBehavior.Stop,
                To = 360,
                Duration = new Duration(TimeSpan.FromSeconds(1))
            };

            rotatingToAnimate.BeginAnimation(RotateTransform.AngleProperty, rotateAnimation);
        }

        private void GetPageCount()
        {
            using (var dataBase = new ExcelPackage(GetExcel.downloadPath))
            {
                ExcelWorksheet sheet = dataBase.Workbook.Worksheets[1];
                PageCount = sheet.Dimension.End.Row / ItemsPerPage;

                if (sheet.Dimension.End.Row % ItemsPerPage != 0)
                    PageCount += 1;
            }
        }
    }
}
