using OfficeOpenXml;
using System;
using System.IO;
using System.Net;
using System.Windows;

namespace Lab_2
{
    class GetExcel
    {
        public static readonly FileInfo downloadPath = new FileInfo("thrlist.xlsx");
        public static readonly FileInfo localPath = new FileInfo("local_thrlist.xlsx");

        public static void Download()
        {
            using (var client = new WebClient())
            {
                try
                {
                    client.DownloadFile("https://bdu.fstec.ru/documents/files/thrlist.xlsx", @"thrlist.xlsx");
                }
                catch (Exception e)
                {
                    MessageBox.Show($"При обновлении данных возникла ошибка! Сообщение: {e.Message}");
                }
            }
        }
        public static int StartRow(ExcelWorksheet sheet)
        {
            return sheet.Dimension.Start.Row;
        }

        public static int EndRow(ExcelWorksheet sheet)
        {
            return sheet.Dimension.End.Row;
        }

        public static int StartCol(ExcelWorksheet sheet)
        {
            return sheet.Dimension.Start.Column;
        }
        public static int EndCol(ExcelWorksheet sheet)
        {
            return sheet.Dimension.End.Column;
        }
    }
}
