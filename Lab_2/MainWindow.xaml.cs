using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

        public const int ItemsPerPage = 15;
        private List<Threat> localThreats = new List<Threat>();
        private List<Threat> downloadThreats = new List<Threat>();
        public MainWindow()
        {
            InitializeComponent();
            CurrentPage = 1;
            if (File.Exists(@"thrlist.xlsx") && !(new FileInfo(@"thrlist.xlsx").Length == 0))
            {
                GetPageCount();
                localThreats = GetList(GetExcel.localPath);
                ShowPage(CurrentPage);
                RefreshButton.IsEnabled = true;
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
                if (!GetExcel.localPath.Exists)
                {
                    GetExcel.CopyDownloadedFile();
                }
                localThreats = GetList(GetExcel.localPath);
                GetPageCount();
                ShowPage(CurrentPage);
                Dispatcher.Invoke(() =>
                {
                    RefreshButton.IsEnabled = true;
                });
            });
        }

        private List<Threat> GetList(FileInfo filePath)
        {
            List<Threat> threats = new List<Threat>();
            using (var dataBase = new ExcelPackage(filePath))
            {
                ExcelWorksheet sheet = dataBase.Workbook.Worksheets[1];
                int startRow = sheet.Dimension.Start.Row + 2;
                int endRow = sheet.Dimension.End.Row;
                int startColumn = 1;
                int endColumn = sheet.Dimension.End.Column - 2;
                for (int row = startRow; row <= endRow; row++)
                {
                    List<string> rowValues = new List<string>();
                    for (int column = startColumn; column <= endColumn; column++)
                    {
                        string cellValue = sheet.Cells[row, column].Value?.ToString();
                        cellValue = cellValue.Replace("_x000d_", "");
                        switch (column)
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
                    threats.Add(new Threat(rowValues));
                }
            }
            return threats;
        }

        private Dictionary<string, Threat> GetDict(List<Threat> list)
        {
            Dictionary<string, Threat> dict = new Dictionary<string, Threat>();
            foreach (var el in list)
            {
                dict[el.Properties["ID"].ToString()] = el;
            }
            return dict;
        }

        private void ShowPage(int pageNumber)
        {
            int startRow = ItemsPerPage * (pageNumber - 1);
            int endRow = localThreats.Count - 1;
            int count = Math.Min(ItemsPerPage, endRow - startRow + 1);

            List<Threat> page = localThreats.GetRange(startRow, count);
            Dispatcher.Invoke(() => {
                ThreatList.ItemsSource = page;
                PageNumber.Text = pageNumber.ToString();
            });
        }

        private void ThreatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = (ListBox)sender;
            var noteInfo = (Threat)item.SelectedItem;
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


            await Task.Run(async () =>
             {
                 await AsyncDownload();
                 await CompareFiles();

             });
            var test = GetList(GetExcel.downloadPath);
            GetExcel.CopyDownloadedFile();
            RefreshButton.IsEnabled = true;
        }

        private Task CompareFiles()
        {
            return Task.Run(() =>
            {
                Dictionary<string, Threat> oldList = GetDict(GetList(GetExcel.localPath));
                Dictionary<string, Threat> newList = GetDict(GetList(GetExcel.localPath));

                Dictionary<string, List<string>> changesInOldList = new Dictionary<string, List<string>>();
                Dictionary<string, List<string>> changesInNewList = new Dictionary<string, List<string>>();


                if (oldList.Count > newList.Count)
                {
                    foreach (var key in oldList.Keys)
                    {
                        List<string> oldKeyList = new List<string>();
                        List<string> newKeyList = new List<string>();
                        foreach (var propertyKey in oldList[key].Properties.Keys)
                        {
                            if (oldList.ContainsKey(key) && newList.ContainsKey(key))
                            {
                                if (oldList[key].Properties[propertyKey] != newList[key].Properties[propertyKey])
                                {
                                    if (String.IsNullOrEmpty(newList[key].Properties[propertyKey]))
                                    {
                                        newKeyList.Add("/---/");
                                        oldKeyList.Add($"Запись удалена: {oldList[key].Properties[propertyKey]}");

                                    }
                                    else
                                    {
                                        newKeyList.Add(newList[key].Properties[propertyKey]);
                                        oldKeyList.Add(oldList[key].Properties[propertyKey]);

                                    }
                                    changesInOldList[key] = oldKeyList;
                                    changesInNewList[key] = newKeyList;
                                }
                            }
                            else
                            {
                                if (!oldList.ContainsKey(key))
                                {
                                    oldKeyList.Add("/   ");
                                    newKeyList.Add($"Добавлена угроза {newList[key].Properties["ID"]}");
                                }
                                else if (!newList.ContainsKey(key))
                                {
                                    newKeyList.Add("/");
                                    oldKeyList.Add($"Удалена угроза {oldList[key].Properties["ID"]}");

                                }
                                changesInOldList[key] = oldKeyList;
                                changesInNewList[key] = newKeyList;
                                break;
                            }
                        }

                    }
                }
                else
                {
                    foreach (var key in newList.Keys)
                    {
                        List<string> oldKeyList = new List<string>();
                        List<string> newKeyList = new List<string>();
                        foreach (var propertyKey in newList[key].Properties.Keys)
                        {
                            if (oldList.ContainsKey(key) && newList.ContainsKey(key))
                            {
                                if (oldList[key].Properties[propertyKey] != newList[key].Properties[propertyKey])
                                {
                                    if (string.IsNullOrEmpty(oldList[key].Properties[propertyKey]))
                                    {
                                        oldKeyList.Add("/Запись не найдена/");
                                        newKeyList.Add($"Запись добавлена: {newList[key].Properties[propertyKey]}");
                                    }
                                    else
                                    {
                                        newKeyList.Add(newList[key].Properties[propertyKey]);
                                        oldKeyList.Add(oldList[key].Properties[propertyKey]);
                                    }
                                    changesInOldList[key] = oldKeyList;
                                    changesInNewList[key] = newKeyList;
                                }
                            }
                            else
                            {
                                if (!oldList.ContainsKey(key))
                                {
                                    oldKeyList.Add("/");
                                    newKeyList.Add($"Добавлена угроза {newList[key].Properties["ID"]}");
                                }
                                else if (!newList.ContainsKey(key))
                                {
                                    newKeyList.Add("/");
                                    oldKeyList.Add($"Удалена угроза {oldList[key].Properties["ID"]}");
                                }
                                changesInOldList[key] = oldKeyList;
                                changesInNewList[key] = newKeyList;
                                break;
                            }
                        }
                    }
                }
                string result = "";
                foreach (var key in changesInOldList.Keys)
                {
                    result += $"{key}:\n";
                    for (int i = 0; i < changesInOldList[key].Count; i++)
                    {
                        result += ($"{changesInOldList[key][i]} - {changesInNewList[key][i]}");
                    }
                }
                MessageBox.Show(String.IsNullOrEmpty(result) ? "Изменений не найдено!" : result);
                GetExcel.CopyDownloadedFile();
            });
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
