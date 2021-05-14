using SpssLib.DataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelExample
{
    public class SpssRead
    {
        public static string path = @"C:\Users\chilo\source\repos\ExcelExample\ExcelExample\v1.sav";
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10,
                                              FileOptions.SequentialScan))
            {
                // Create the reader, this will read the file header
                SpssReader spssDataset = new SpssReader(fileStream);

                // Iterate through all the varaibles
                foreach (var variable in spssDataset.Variables)
                {
                    // Display name and label
                    Console.WriteLine("{0} - {1}", variable.Name, variable.Label);
                    // Display value-labels collection
                    foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                    {
                        Console.WriteLine(" {0} - {1}", label.Key, label.Value);
                    }
                }

                // Iterate through all data rows in the file
                foreach (var record in spssDataset.Records)
                {
                    foreach (var variable in spssDataset.Variables)
                    {
                        Console.Write(variable.Name);
                        Console.Write(':');
                        // Use the corresponding variable object to get the values.
                        Console.Write(record.GetValue(variable));
                        // This will get the missing values as null, text with out extra spaces,
                        // and date values as DateTime.
                        // For original values, use record[variable] or record[int]
                        Console.Write('\t');
                    }
                    Console.WriteLine("");
                }
            }
        }
    }
}
