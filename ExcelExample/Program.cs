using OfficeOpenXml;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelExample
{
    public class Program
    {
        public static string path = @"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\Svod_2020_g.xlsx";
        //public static WorkSheet sheet = null;
        public static ExcelWorksheet myWorksheet;
        public static HashSet<string> columns;

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(path)))
            {
                myWorksheet = xlPackage.Workbook.Worksheets.FirstOrDefault(); //select sheet here

                var qwe = myWorksheet.Cells[3, 4];
                var users = new List<User>();
                columns = new HashSet<string>();
                for (var i = 4; i < 531; i += 2)// for (var i = 4; i < 531; i += 2) (var i = 112; i < 114; i += 2)
                {   // Январь
                    var result1 = getTablesRgex(i + 1, i, 7, 38);
                    // Февряль
                    var result2 = getTablesRgex(i + 1, i, 60, 89);
                    // Март
                    var result3 = getTablesRgex(i + 1, i, 92, 123);
                    // Апрель
                    var result4 = getTablesRgex(i + 1, i, 126, 156);
                    // Май
                    var result5 = getTablesRgex(i + 1, i, 159, 190);
                    // Июнь
                    var result6 = getTablesRgex(i + 1, i, 193, 223);
                    // Июль
                    var result7 = getTablesRgex(i + 1, i, 226, 257);
                    // Август
                    var result8 = getTablesRgex(i + 1, i, 260, 291);
                    // Сентябрь
                    var result9 = getTablesRgex(i + 1, i, 294, 324);
                    // Октябрь
                    var result10 = getTablesRgex(i + 1, i, 327, 358);
                    // Ноябрь
                    var result11 = getTablesRgex(i + 1, i, 361, 391);
                    // Декабрь
                    var result12 = getTablesRgex(i + 1, i, 394, 425);

                    var month = new Dictionary<int, Dictionary<string, List<double>>>();
                    month.Add(0, result1);
                    month.Add(1, result2);
                    month.Add(2, result3);
                    month.Add(3, result4);
                    month.Add(4, result5);
                    month.Add(5, result6);
                    month.Add(6, result7);
                    month.Add(7, result8);
                    month.Add(8, result9);
                    month.Add(9, result10);
                    month.Add(10, result11);
                    month.Add(11, result12);

                    var username = myWorksheet.Cells[i, 2].Text;
                    var position = myWorksheet.Cells[i, 4].Text;
                    var number = myWorksheet.Cells[i, 1].Text;
                    users.Add(new User()
                    {
                        Fullname = username,
                        Number = number,
                        Months = month,
                        //FirstMonth = result1,
                        Position = position
                    });
                }
                WriteTable(users);
            }

            //var workbook = WorkBook.Load(path);
            //sheet = workbook.WorkSheets.First();

            //var res = getTables(10, 9, 6, 37);
        }

        public static void WriteTable(List<User> users)
        {
            ExcelPackage ExcelPkg = new ExcelPackage();  
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            int constColumns = 7;
            var cc = columns.OrderBy(a => a).ToList();
            int constNeed = 21;

            CultureInfo ci = new CultureInfo("ru-RU");
            // Get the DateTimeFormatInfo for the en-US culture.
            DateTimeFormatInfo dtfi = ci.DateTimeFormat;
            for (var w = 0; w < 12; w++)
            {
                wsSheet1.Cells[1, w * (cc.Count + constNeed) + constColumns].Value = $"{dtfi.GetMonthName(w+1)}";
                wsSheet1.Cells[1, w * (cc.Count + constNeed) + constColumns, 1, (w + 1) * (cc.Count + constNeed) - 1 + constColumns].Merge = true;
                wsSheet1.Cells[1, w * (cc.Count + constNeed) + constColumns, 1, (w + 1) * (cc.Count + constNeed) - 1 + constColumns].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                for (var i = 0; i < cc.Count; i++)
                {
                    wsSheet1.Cells[2, w*(cc.Count + constNeed) + i + constColumns].Value = cc[i];
                }
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 0].Value = "Кол-во дней отработанных";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 1].Value = "Всего по РФВ";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 2].Value = "Календарный";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 3].Value = "Календарный";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 4].Value = "Табельный (номинальный)";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 5].Value = "Максимально возможный";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 6].Value = "Явочный";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 7].Value = "Урочное";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 8].Value = "Сверхурочное";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 9].Value = "Всего отработано";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 10].Value = "По уважительным причинам";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 11].Value = "Потери рабочего времени (без уважительных причин)";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 12].Value = "Всего неотработано";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 13].Value = "Урочное";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 14].Value = "Сверхурочное";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 15].Value = "Отработано (от ТФВ)";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 16].Value = "По уважительным причинам";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 17].Value = "Потери рабочего времени (без уважительных причин)";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 18].Value = "Неотработано (от ТФВ)";
                wsSheet1.Cells[2, w * (cc.Count + constNeed) + cc.Count + constColumns + 19].Value = "Коэффициент использования максимально возможного фонда рабочего времени";
            }
            for(var i = 0; i < users.Count; i++)
            {
                wsSheet1.Cells[i * 2 + 3, 1].Value = "Часы";
                wsSheet1.Cells[i * 2 + 4, 1].Value = "Дни";

                wsSheet1.Cells[i * 2 + 3, 2].Value = i;

                wsSheet1.Cells[i * 2 + 3, 3].Value = users[i].Fullname;
                wsSheet1.Cells[i * 2 + 4, 3].Value = users[i].Fullname;

                wsSheet1.Cells[i * 2 + 3, 4].Value = users[i].Position;
                wsSheet1.Cells[i * 2 + 4, 4].Value = users[i].Position;

                wsSheet1.Cells[i * 2 + 3, 5].Value = users[i].Number;
                wsSheet1.Cells[i * 2 + 4, 5].Value = users[i].Number;

                var dics = users[i].Months.OrderBy(a => a.Key).ToArray();
                for(var w = 0; w < dics.Length; w++)
                {
                    double allCount = 0;
                    int allCount2 = 0;
                    foreach (var dict in dics[w].Value)
                    {
                        int positionColumn = cc.IndexOf(dict.Key);
                        wsSheet1.Cells[i * 2 + 3, positionColumn + constColumns + w * (cc.Count + constNeed)].Value = dict.Value.Sum();
                        wsSheet1.Cells[i * 2 + 4, positionColumn + constColumns + w * (cc.Count + constNeed)].Value = dict.Value.Count;
                        allCount += dict.Value.Sum();
                    }

                    wsSheet1.Cells[i * 2 + 3, w * (cc.Count + constNeed) + cc.Count + constColumns + 1].Formula =
                        $"SUM({wsSheet1.Cells[i * 2 + 3, constColumns + w * (cc.Count + constNeed)].Address}:{wsSheet1.Cells[i * 2 + 3, cc.Count + constColumns + w * (cc.Count + constNeed)].Address})";

                    wsSheet1.Cells[i * 2 + 4, 1 + w * (cc.Count + constNeed) + cc.Count + constColumns + 1].Formula =
                        $"SUM({wsSheet1.Cells[i * 2 + 4, cc.IndexOf("Б") + constColumns + w * (cc.Count + constNeed)]}, {wsSheet1.Cells[i * 2 + 4, cc.IndexOf("В") + constColumns + w * (cc.Count + constNeed)]}," +
                        $"{wsSheet1.Cells[i * 2 + 4, cc.IndexOf("ДО") + constColumns + w * (cc.Count + constNeed)]},{wsSheet1.Cells[i * 2 + 4, cc.IndexOf("НН") + constColumns + w * (cc.Count + constNeed)]}," +
                        $"{wsSheet1.Cells[i * 2 + 4, cc.IndexOf("О") + constColumns + w * (cc.Count + constNeed)]}, {wsSheet1.Cells[i * 2 + 4, cc.Count + constColumns + w * (cc.Count + constNeed)]})";



                    //wsSheet1.Cells[i * 2 + 4, w * (cc.Count + constNeed) + cc.Count + constColumns + 1].Formula =
                    //    $"SUM({wsSheet1.Cells[i * 2 + 4, constColumns + w * (cc.Count + constNeed)].Address}:{wsSheet1.Cells[i * 2 + 4, cc.Count + constColumns + w * (cc.Count + constNeed)].Address})";
                }
            }
            ExcelPkg.SaveAs(new FileInfo(@"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\example.xlsx"));
        }

        public static Dictionary<string, List<double>> getTablesRgex(int rowValue, int row, int begin, int end)
        {
            var result = new Dictionary<string, List<double>>();
            for (var i = begin; i < end; i++)
            {
                var valueC = myWorksheet.Cells[row, i];
                var columnPC = myWorksheet.Cells[rowValue, i];
                if (rowValue == 173)
                    continue;
                double value = string.IsNullOrEmpty(valueC.Text) ? 0.0 : double.Parse(valueC.Text);
                string columnP = columnPC.Value == null ? "empty" : (string)columnPC.Value;
                //if (string.IsNullOrEmpty(columnP))
                //    continue;
                if (columnP == "empty")
                    continue;
                Regex regex = new Regex(@"^(([A-zА-яЁё]*)([0-9]*))?(([A-zА-яЁё]*)([0-9]*))?(([A-zА-яЁё]*)([0-9]*))?(([A-zА-яЁё]*)([0-9]*))?(([A-zА-яЁё]*)([0-9]*))\b");
                Regex regexCheckNumber = new Regex(@"[0-9]");
                MatchCollection matches = regex.Matches(columnP);
                MatchCollection matchesCheckNumber = regexCheckNumber.Matches(columnP);
                if (matchesCheckNumber.Count > 0)
                {
                    string fullValue = null;
                    string colValue = null;
                    int colHours = 0;
                    var m = (ICollection<Match>)matches;
                    foreach(Match match in matches)
                    {
                        for(var a = 1; a < match.Groups.Count; a++)
                        {
                            Group group = match.Groups[a];
                            if (string.IsNullOrEmpty(group.Value))
                                continue;
                            if (string.IsNullOrEmpty(fullValue))
                            {
                                fullValue = group.Value;
                                continue;
                            }
                            if (string.IsNullOrEmpty(colValue))
                            {
                                colValue = group.Value;
                                continue;
                            }
                            if (!string.IsNullOrEmpty(colValue))
                            {
                                colHours = int.Parse(group.Value);

                                if (colValue == "ЯН")
                                {
                                    colValue = "Н";
                                    columns.Add(colValue);
                                    if (!result.ContainsKey(colValue))
                                    {
                                        result.Add(colValue, new List<double>());
                                    }
                                    result[colValue].Add(colHours);

                                    colValue = "Я";
                                    columns.Add(colValue);
                                    if (!result.ContainsKey(colValue))
                                    {
                                        result.Add(colValue, new List<double>());
                                    }
                                    result[colValue].Add(value);
                                }
                                else if (colValue == "СН")
                                {
                                    colValue = "Н";
                                    columns.Add(colValue);
                                    if (!result.ContainsKey(colValue))
                                    {
                                        result.Add(colValue, new List<double>());
                                    }
                                    result[colValue].Add(colHours);

                                    colValue = "С";
                                    columns.Add(colValue);
                                    if (!result.ContainsKey(colValue))
                                    {
                                        result.Add(colValue, new List<double>());
                                    }
                                    result[colValue].Add(value);
                                }
                                else
                                {
                                    columns.Add(colValue);
                                    if (!result.ContainsKey(colValue))
                                    {
                                        result.Add(colValue, new List<double>());
                                    }
                                    result[colValue].Add(colHours);
                                }

                                fullValue = null;
                                colValue = null;
                                colHours = 0;
                            }
                        }
                    }
                }
                else
                {
                    columns.Add(columnP);
                    if (!result.ContainsKey(columnP))
                    {
                        result.Add(columnP, new List<double>());
                    }
                    result[columnP].Add(value);
                }

            }
            return result;
        }

        public static Dictionary<string, List<double>> getTables2(int rowValue, int row, int begin, int end)
        {
            var result = new Dictionary<string, List<double>>();
            for (var i = begin; i < end; i++)
            {
                var valueC = myWorksheet.Cells[row, i];
                var columnPC = myWorksheet.Cells[rowValue, i];
                if (rowValue == 173)
                    continue;
                double value = string.IsNullOrEmpty(valueC.Text) ? 0.0 : double.Parse(valueC.Text);
                string columnP = columnPC.Value == null ? "empty" : (string)columnPC.Value;
                //if (string.IsNullOrEmpty(columnP))
                //    continue;
                columns.Add(columnP);
                if (!result.ContainsKey(columnP))
                {
                    result.Add(columnP, new List<double>());
                }
                result[columnP].Add(value);
            }
            return result;
        }
    }

    public class User
    {
        public string Fullname { get; set; }
        public string Number { get; set; }
        public string Position { get; set; }

        public Dictionary<int, Dictionary<string, List<double>>> Months { get; set; }
        public Dictionary<string, List<double>> FirstMonth { get; set; }
    }
}
