using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace FileLoaderTimerJob
{
    #region Excel Writer
    //class Log
    //{
    //    private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
    //    private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
    //    private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
    //    private static Microsoft.Office.Interop.Excel.Application oXL;

    //    public void LogToExcel(List<string> logInfo)
    //    {

    //        string path = @"C:\Enogex\book2.xlsx";
    //        oXL = new Microsoft.Office.Interop.Excel.Application();
    //        oXL.Visible = true;
    //        oXL.DisplayAlerts = false;
    //        Workbooks workbooks = oXL.Workbooks;
    //        try
    //        {
    //            mWorkBook = workbooks.Open(path, 0, false, 5, "", "", false, Missing.Value, "", true, false, 0, true, false, false);
    //            //Get all the sheets in the workbook
    //            mWorkSheets = mWorkBook.Worksheets;
    //            //Get the allready exists sheet
    //            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Sheet1");
    //            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

    //            range = (Excel.Range)mWSheet1.Cells[mWSheet1.Rows.Count, 1];
    //            long lastRow = (long)range.get_End(Excel.XlDirection.xlUp).Row;
    //            long newRow = lastRow + 1;

    //            WriteArray(1, 8, newRow, logInfo, mWSheet1);

    //            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
    //            Missing.Value, Missing.Value, false, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
    //            Missing.Value, Missing.Value, Missing.Value,
    //            Missing.Value, Missing.Value);

    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine(ex.Message);

    //        }
    //        finally 
    //        {
    //            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
    //            mWSheet1 = null;
    //            mWorkBook = null;
    //            oXL.Quit();
    //            GC.WaitForPendingFinalizers();
    //            GC.Collect();
    //            GC.WaitForPendingFinalizers();
    //            GC.Collect();
    //        }

    //    }

    //    private static void WriteArray(int rows, int columns, long blankRow, List<string> logInfo, Worksheet worksheet)
    //    {
    //        var data = new object[rows, columns];
    //        int i = 0;
    //            for (var column = 1; column <= columns; column++)
    //            {
    //                if (columns - logInfo.Count() <= 0 )
    //                {
    //                    data[0, column - 1] = logInfo[i];
    //                    i++;
    //                }
    //            }


    //        var startCell = (Range)worksheet.Cells[blankRow, 1];
    //        var endCell = (Range)worksheet.Cells[rows + blankRow - 1, columns];
    //        var writeRange = worksheet.Range[startCell, endCell];

    //        writeRange.Value2 = data;
    //    }


    //    private static bool checkFileIsOpened(string fileName)
    //    {

    //        try
    //        {
    //            Stream s = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite);

    //            s.Close();

    //            return false;
    //        }
    //        catch (Exception)
    //        {
    //            return true;
    //        }

    //    }
    // }
    #endregion

    #region Csv Writer old code
    ////public class WriteToCSV
    ////{
    //    public enum EmptyLineBehavior
    //    {
    //        /// <summary>
    //        /// Empty lines are interpreted as a line with zero columns.
    //        /// </summary>
    //        NoColumns,
    //        /// <summary>
    //        /// Empty lines are interpreted as a line with a single empty column.
    //        /// </summary>
    //        EmptyColumn,
    //        /// <summary>
    //        /// Empty lines are skipped over as though they did not exist.
    //        /// </summary>
    //        Ignore,
    //        /// <summary>
    //        /// An empty line is interpreted as the end of the input file.
    //        /// </summary>
    //        EndOfFile,
    //    }

    //    public abstract class CsvFileCommon
    //    {
    //        /// <summary>
    //        /// These are special characters in CSV files. If a column contains any
    //        /// of these characters, the entire column is wrapped in double quotes.
    //        /// </summary>
    //        protected char[] SpecialChars = new char[] { ',', '"', '\r', '\n' };

    //        // Indexes into SpecialChars for characters with specific meaning
    //        private const int DelimiterIndex = 0;
    //        private const int QuoteIndex = 1;

    //        /// <summary>
    //        /// Gets/sets the character used for column delimiters.
    //        /// </summary>
    //        public char Delimiter
    //        {
    //            get { return SpecialChars[DelimiterIndex]; }
    //            set { SpecialChars[DelimiterIndex] = value; }
    //        }

    //        /// <summary>
    //        /// Gets/sets the character used for column quotes.
    //        /// </summary>
    //        public char Quote
    //        {
    //            get { return SpecialChars[QuoteIndex]; }
    //            set { SpecialChars[QuoteIndex] = value; }
    //        }
    //    }

    //    public class CsvFileWriter : CsvFileCommon, IDisposable
    //    {
    //        // Private members
    //        private StreamWriter Writer;
    //        private string OneQuote = null;
    //        private string TwoQuotes = null;
    //        private string QuotedFormat = null;

    //        /// <summary>
    //        /// Initializes a new instance of the CsvFileWriter class for the
    //        /// specified stream.
    //        /// </summary>
    //        /// <param name="stream">The stream to write to</param>
    //        //public CsvFileWriter(Stream stream)
    //        //{
    //        //    Writer = new StreamWriter(stream);
    //        //}

    //        /// <summary>
    //        /// Initializes a new instance of the CsvFileWriter class for the
    //        /// specified file path.
    //        /// </summary>
    //        /// <param name="path">The name of the CSV file to write to</param>
    //        public CsvFileWriter() { }
    //        public CsvFileWriter(string path)
    //        {
    //            Writer = new StreamWriter(path);
    //        }

    //        /// <summary>
    //        /// Writes a row of columns to the current CSV file.
    //        /// </summary>
    //        /// <param name="columns">The list of columns to write</param>
    //        public void WriteRow(List<string> columns)
    //        {
               
    //            try
    //            {
                  
    //                // Verify required argument
    //                if (columns == null)
    //                    throw new ArgumentNullException("columns");

    //                // Ensure we're using current quote character
    //                if (OneQuote == null || OneQuote[0] != Quote)
    //                {
    //                    OneQuote = String.Format("{0}", Quote);
    //                    TwoQuotes = String.Format("{0}{0}", Quote);
    //                    QuotedFormat = String.Format("{0}{{0}}{0}", Quote);
    //                }

    //                // Write each column
    //                for (int i = 0; i < columns.Count; i++)
    //                {
    //                    // Add delimiter if this isn't the first column
    //                    if (i > 0)
    //                        Writer.Write(Delimiter);
    //                    // Write this column
    //                    if (columns[i].IndexOfAny(SpecialChars) == -1)

    //                        Writer.Write(columns[i]);

    //                    else
    //                        Writer.Write(QuotedFormat, columns[i].Replace(OneQuote, TwoQuotes));
    //                }
    //                Writer.WriteLine();
    //                Writer.Close();
    //            }

    //            catch 
    //            {

    //                Writer.Dispose();
    //            }
    //        }
    //        // Propagate Dispose to StreamWriter
    //        public void Dispose()
    //        {
    //            Writer.Dispose();
    //        }


    //        public void csvWrite(List<string> columns)
    //        {
    //            string path2 = @"C:\Enogex\FileLoaderLog.csv";
    //            using (var writer = new CsvFileWriter(path2))
    //            {

    //                writer.WriteRow(columns);
    //            }

    //        }



    //    }
   #endregion

    public class CsvWriter
    {

        public static void CreateCsvFile(List<string> row,string filePath)
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
            catch{}
        }

       
    }
    
}
