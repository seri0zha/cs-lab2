using OfficeOpenXml;
using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Threading;

namespace Lab_2
{
    class GetExcel
    {
        public static readonly FileInfo downloadPath = new FileInfo("thrlist.xlsx");
        public static readonly FileInfo localPath = new FileInfo("local_thrlist.xlsx");

        public static void ShowMessageBox(string message)
        {
            Thread thread = new Thread(() =>
            {
                MessageBox.Show(message);
            });
            thread.Start();
        }
        public static void Download()
        {
            using (var client = new WebClient())
            {
                try
                {
                    client.DownloadFile("https://bdu.fstec.ru/documents/files/thrlist.xlsx", @"thrlist.xlsx");
                    ShowMessageBox("Файл загружен успешно!");
                }
                catch (Exception e)
                {
                    ShowMessageBox($"При обновлении данных возникла ошибка: {e.Message}");
                }
            }
        }

        public static void CopyDownloadedFile()
        {
            using (ExcelPackage downloadedFile = new ExcelPackage(downloadPath))
            {
                downloadedFile.SaveAs(downloadPath);
                downloadedFile.SaveAs(localPath);
            }
        }
    }
}
