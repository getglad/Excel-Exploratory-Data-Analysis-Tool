using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DataLibrary;
using BL = DataLibrary.Library;
using EL = DataLibrary.ExcelLib;
using EDA = DataLibrary.ExcelEDA;
using TSR = DataLibrary.TermSampling;

using RDotNet;

namespace Excel_Exploratory_Data_Analysis_Tool
{
    class Program
    {
        private static BL.FileList Library = new BL.FileList();
        
        static void Main(string[] args)
        {
            REngine.SetEnvironmentVariables();
            REngine RBlock = REngine.GetInstance();

            DirectoryInfo di = new DirectoryInfo("C:\\Users\\" + Environment.UserName + "\\Documents\\File Attachments\\");
            string searchPattern = "*";

            // Set some bases. These need to be tested for each sample.
            int avgRowCount = 100;
            int avgColCount = 20;
            int fileCount = 0;

            // Get our files, Create a library to hold everything
            Library.ListofFiles = Library.makeList(di, searchPattern);

            // Look Through Each File
            foreach (var a in Library.ListofFiles)
            {
                Console.WriteLine(a.FileName);
                string fileLocation = di + a.FileName;
                // Open The File, Set the Page
                EL.singleExcel thisExcel = new EL.singleExcel().createExcel(fileLocation);
                EL.singleExcel.ExcelWorkSheetChange(thisExcel, 1);

                // Get Basic Vars
                new EDA.excelBasics().basicVars(a, thisExcel);
                // Run Series of Operations to Get More Exact Bounds

                EDA.lookTriggers LT = new EDA.lookTriggers();
                LT.runTriggers(fileCount, a, avgRowCount, avgColCount, out avgRowCount, out avgColCount);

                Library.colCounts.Add(a.colCount);
                Library.rowCounts.Add(a.rowCount);

                // In the event that no rows are found, the file is not counted
                if (a.rowCount != 0)
                    fileCount++;

                EL.singleExcel.CloseSheet(thisExcel);
                //Console.ReadLine();
                Console.WriteLine("\n");
            }

            // Run some R analysis
            IntegerVector rowR = RBlock.CreateIntegerVector(Library.rowCounts);
            IntegerVector colR = RBlock.CreateIntegerVector(Library.colCounts);
            RBlock.SetSymbol("rowR", rowR);
            RBlock.SetSymbol("colR", colR);

            int[] thisTemp;
            RBlock.Evaluate("temp <- table(as.vector(rowR))");
            thisTemp = RBlock.Evaluate("names(temp)[temp == max(temp)]").AsInteger().ToArray();
            Library.groupStats.modeRow = thisTemp[0];
            RBlock.Evaluate("temp <- table(as.vector(colR))");
            thisTemp = RBlock.Evaluate("names(temp)[temp == max(temp)]").AsInteger().ToArray();
            Library.groupStats.modeCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("mean(rowR)").AsInteger().ToArray();
            Library.groupStats.meanRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("mean(colR)").AsInteger().ToArray();
            Library.groupStats.meanCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("median(rowR)").AsInteger().ToArray();
            Library.groupStats.medianRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("median(colR)").AsInteger().ToArray();
            Library.groupStats.medianCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("min(rowR)").AsInteger().ToArray();
            Library.groupStats.minRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("min(colR)").AsInteger().ToArray();
            Library.groupStats.minCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("max(rowR)").AsInteger().ToArray();
            Library.groupStats.maxRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("max(colR)").AsInteger().ToArray();
            Library.groupStats.maxCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("IQR(rowR)").AsInteger().ToArray();
            Library.groupStats.iqrRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("IQR(colR)").AsInteger().ToArray();
            Library.groupStats.iqrCol = thisTemp[0];
            thisTemp = RBlock.Evaluate("quantile(rowR)").AsInteger().ToArray();
            Library.groupStats.quantileRow = thisTemp[0];
            thisTemp = RBlock.Evaluate("quantile(colR)").AsInteger().ToArray();
            Library.groupStats.quantileCol = thisTemp[0];

            RBlock.Dispose();

            EL.singleExcel.outputObjectToExcel(Library.groupStats);

            // Build some training data from previous information, assumptions
            object[,] theWords = TSR.trainMethods.getTrainData(di, "\\testdata\\trainlist.xlsx");
            List<string> words = new List<string>();
            TSR.trainMethods.makeTrainList(theWords, words);

            theWords = TSR.trainMethods.getTrainData(di, "\\testdata\\looselist.xlsx");
            List<string> looseWords = new List<string>();
            TSR.trainMethods.makeTrainList(theWords, looseWords);

            Dictionary<string, int> discreteTermsDict = new Dictionary<string, int>();
            Dictionary<string, int> betterTermsDict = new Dictionary<string, int>();
            Dictionary<string, int> secondaryTermsDict = new Dictionary<string, int>();

            // Run training scenarios
            TSR.trainMethods.runTrainScan(words, looseWords, avgColCount, Library.ListofFiles, out discreteTermsDict, out betterTermsDict, out secondaryTermsDict);
            TSR.trainMethods.buildTrainList(discreteTermsDict, words);

            TSR.trainMethods.runTrainScan(words, looseWords, avgColCount, Library.ListofFiles, out discreteTermsDict, out betterTermsDict, out secondaryTermsDict);
            TSR.trainMethods.buildTrainList(betterTermsDict, looseWords);
            TSR.trainMethods.buildTrainList(secondaryTermsDict, looseWords);

            // In tests, three cycles make for extremely high confidence
            TSR.trainMethods.runTrainScan(words, looseWords, avgColCount, Library.ListofFiles, out discreteTermsDict, out betterTermsDict, out secondaryTermsDict);
            //  Terms for classifcation
            EL.singleExcel.outputListToExcel(words, "strongTrainList");
            EL.singleExcel.outputListToExcel(looseWords, "learnedTrainList");
        }
    }
}
