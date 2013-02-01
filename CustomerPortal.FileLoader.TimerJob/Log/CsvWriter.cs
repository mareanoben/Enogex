using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace CustomerPortal.FileLoader.TimerJob.Log
{

    public class CsvWriter
    {

        public static void CreateCsvFile(List<string> row, string filePath)
        {

            try
            {

                using (var wr = new StreamWriter(filePath, true))
                {

                    var sb = new StringBuilder();

                    foreach (string value in row)
                    {
                        //Add a comma
                        if (sb.Length > 0)
                            sb.Append(",");

                        sb.Append(value);

                    }

                    wr.WriteLine(sb.ToString());

                }

            }
            catch { }
        }


    }

}
