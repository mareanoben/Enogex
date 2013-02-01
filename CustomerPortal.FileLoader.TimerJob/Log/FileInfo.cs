using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CustomerPortal.FileLoader.TimerJob.Log
{
    public class FileInfo
    {

        public string CustomerName { get; set; }
        public string LibraryName { get; set; }
        public string ContentTypeName { get; set; }
        public string FileName { get; set; }
        public string ModifiedSP { get; set; }
        public string ModifiedSQL { get; set; }
        public string ModifiedBy { get; set; }
        public string Extension { get; set; }
    }
}
