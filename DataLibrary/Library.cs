using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using System.Runtime.InteropServices;
using System.Diagnostics;

namespace DataLibrary
{
    public class Library
    {

        public class FileList
        {
            public FileList() { ListofFiles = new List<AnalysisObject>(); rowCounts = new List<int>(); colCounts = new List<int>(); groupStats = new GroupStats(); }
            public List<AnalysisObject> ListofFiles { get; set; }
            public GroupStats groupStats { get; set; }
            public List<int> rowCounts { get; set; }
            public List<int> colCounts { get; set; }

            public List<AnalysisObject> makeList(DirectoryInfo di, string searchPattern)
            {
                foreach (FileInfo f in di.GetFiles(searchPattern))
                {
                    if (f.ToString()[0] != '~')
                        ListofFiles.Add(new AnalysisObject { FileName = f.ToString() });
                }
                return ListofFiles;
            }
        }

        public class GroupStats
        {
            public GroupStats() { }

            public int minRow { get; set; }
            public int maxRow { get; set; }
            public int meanRow { get; set; }
            public int medianRow { get; set; }
            public int modeRow { get; set; }
            public int iqrRow { get; set; }
            public int quantileRow { get; set; }
            public int minCol { get; set; }
            public int maxCol { get; set; }
            public int meanCol { get; set; }
            public int medianCol { get; set; }
            public int modeCol { get; set; }
            public int iqrCol { get; set; }
            public int quantileCol { get; set; }
        }
        
        public class AnalysisObject
        {
            public AnalysisObject() { }
            public string FileName { get; set; }
            public int rowCount { get; set; }
            public int colCount { get; set; }
            public object[,] allTheData { get; set; }
            public string startCol { get; set; }
            public string startRow { get; set; }
            public string[] splitUpAddresses { get; set; }
        }

    }
}
