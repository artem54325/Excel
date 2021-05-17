using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelExample
{
    public class Excel2
    {
        public static string path = @"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\files\myExcel2.xlsx";
        public static List<BMP> list = new List<BMP>();
        public static Dictionary<string, int> dict = new Dictionary<string, int>();
        public static int CONSTLeft = 3;

        private static int numberTop = 6;

        //static void Main(string[] args)
        //{
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(path)))
            //{
            //    var myWorksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
            //    int rowCount = myWorksheet.Dimension.Rows;
            //    int colCount = myWorksheet.Dimension.Columns;
            //    for (var i = numberTop; i < rowCount; i++)
            //    {
            //        var contry = myWorksheet.Cells[2, i].Text;
            //        if (contry == "")
            //            continue;
            //        if(!dict.ContainsKey(contry))
            //            dict.Add(contry, i);
            //    }
            //    for (var i = CONSTLeft; i< colCount; i++)
            //    {
            //        //3
            //        var contry = myWorksheet.Cells[i, 2].Text;
            //        var industries = myWorksheet.Cells[i, 4].Text;
            //        if (!dict.ContainsKey(contry))
            //            continue;
            //        var calc = myWorksheet.Cells[i, dict[contry]].Value;
            //        list.Add(new BMP()
            //        {
            //            Name = contry,
            //            BMPIndustries = industries,
            //            Calc = calc
            //        });

            //    }
            //    WriteTable(list);
            //}
        //}
        public static void WriteTable(List<BMP> bs)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            for(var i = 0; i < bs.Count; i++)
            {
                wsSheet1.Cells[i + 1, 1].Value = bs[i].Name;
                wsSheet1.Cells[i + 1, 2].Value = bs[i].BMPIndustries;
                wsSheet1.Cells[i + 1, 3].Value = bs[i].Calc;
            }
            ExcelPkg.SaveAs(new FileInfo(@"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\example2.xlsx"));
        }
    }
    public class BMP
    {
        public string Name { get; set; }
        public string BMPIndustries { get; set; }
        public int NumberCountTop { get; set; }
        public int NumberCountLeft { get; set; }
        public object Calc { get; set; }
    }
}
