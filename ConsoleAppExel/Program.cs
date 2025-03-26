//using SpreadsheetLight;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;


// Task description:
// Find the links in new xlsx which are not included in old one Google Sheet. 
namespace ConsoleAppExel
{
    internal class Program
    {
        static string oldFileName = "Old20250123.xlsx";
        static string newFileName = "New20250219.xlsx";
        static string intFileName = "Diff_Old20250123_New20250219.xlsx";

        static void Main(string[] args)
        {             
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;          

            using (var package = new ExcelPackage())
            {
                string oldFilePath = Path.Combine(Environment.CurrentDirectory, oldFileName);
                List<string> oldList = GetLinkList(oldFilePath, package);

                string newFilePath = Path.Combine(Environment.CurrentDirectory, newFileName);
                List<string> newList = GetLinkList(newFilePath, package);

                List<string> intList = newList.Except(oldList).ToList();

                intList.Sort();

                string intFilePath = Path.Combine(@"C:\Projects\ConsoleAppExel\ConsoleAppExel\", intFileName);
                CreateDifFile(intList, intFilePath);
            }            
        }

        public static void CreateDifFile(List<string> difList, string difFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("DifSheet");
                //Add the headers
                ws.Cells[1, 1].Value = "Name";

                ws.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Diff";

                package.Workbook.Properties.Title = "Diff Old20250123.xlsx&New20250219.xlsx";

                for (int row = 2; row <= difList.Count() - 1; ++row)
                {
                    string ind = $"A{row}";
                    ws.Cells[ind].SetCellValue(0, 0, difList[row]);                    
                }

                // Save our new workbook in the output directory and we are done!
                package.SaveAs(difFilePath);
            }
        }

        public static List<string> GetLinkList(string path, ExcelPackage package)
        {
            List<string> list = new List<string>();

            // Load excel stream
            using (var stream = File.OpenRead(path))
            {
                package.Load(stream);
            }
            // Deal with the first worksheet
            var ws = package.Workbook.Worksheets[0];          

            var len = ws.Dimension.Rows;

            for (int row = 2; row <= len; ++row)
            {
                string ind = $"A{row}";

                var link = ws.Cells[ind].Text;
                list.Add(link);
            }

            return list;
        }
    }
}
