using System;
using System.Collections.Generic;
using System.Text;
using NDbfReader;

namespace ExcelExample
{
    public class DBFRead
    {
        public static string path = @"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\files\v1.dbf";
        static void Main(string[] args)
        {
            using (var table = Table.Open(path))
            {
                var reader = table.OpenReader(Encoding.UTF8);
                while (reader.Read())
                {
                    var q2 = table.Columns[0].Name;
                    var q1 = reader.GetDecimal(q2);
                    //var row = new MyRow()
                    //{
                    //    Text = reader.GetString("TEXT"),
                    //    DateTime = reader.GetDateTime("DATETIME"),
                    //    IntValue = reader.GetInt32("INT"),
                    //    DecimalValue = reader.GetDecimal("DECIMAL"),
                    //    BooleanValue = reader.GetBoolean("BOOL")
                    //};
                }
            }
        }
    }
}
