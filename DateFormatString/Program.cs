using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DateFormatString
{
    class Program
    {
        static void Main(string[] args)
        {
            string dateString = "02172012.txt";
            int x = dateString.LastIndexOf('.');
            string str = dateString.Substring(0, x);

            string ext = dateString.Substring(x);

            DateTime date = DateTime.ParseExact(str, "MMddyyyy", System.Globalization.CultureInfo.InvariantCulture);
            //date.ToString("MM/dd/yy");
            Console.WriteLine(date.ToShortDateString());
        }
    }
}
