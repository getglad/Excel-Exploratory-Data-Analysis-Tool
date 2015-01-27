using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel;

using BL = DataLibrary.Library;
using EL = DataLibrary.ExcelLib;

namespace DataLibrary
{
    public class ExcelEDA
    {
        public class excelBasics
        {
            public excelBasics() { }

            public BL.AnalysisObject basicVars(BL.AnalysisObject a, EL.singleExcel thisExcel)
            {
                // Get Some Generic, Untested Data About The Page From Excel
                //var theTotalColumns = thisExcel.excelRange.Columns.Count;
                //var theTotalRows = thisExcel.excelRange.Rows.Count;
                a.colCount = thisExcel.excelRange.Columns.Count;
                a.rowCount = thisExcel.excelRange.Rows.Count;
                a.allTheData = (object[,])thisExcel.excelRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault); // Everything in UsedRange
                var constantCells = thisExcel.excelRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Type.Missing); // Everything with a constant in it in UsedRange
                // Do a little preprossing on that Data
                a.splitUpAddresses = Regex.Split(constantCells.Address, @"(?:\,|\:)");
                a.splitUpAddresses = a.splitUpAddresses.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToArray(); // Create an Array of the Constant Addresses in the Order The Were Found
                var count = a.splitUpAddresses.Length;
                string[] tempArr = new string[count];
                Array.Copy(a.splitUpAddresses, tempArr, count); // Create a Second Array that is Sorted
                Array.Sort(tempArr); // Create a Second Array that is Sorted
                string[] firstAddress = Regex.Split(tempArr[0], @"(?:\$)");
                int x = 0;
                while (firstAddress[1].Length != 1)
                {
                    firstAddress = Regex.Split(tempArr[x], @"(?:\$)");
                    x++;
                }
                a.startCol = firstAddress[1]; // Get Column of First Address, in case it isn't A
                firstAddress = Regex.Split(a.splitUpAddresses[0], @"(?:\$)");
                a.startRow = firstAddress[2]; // Get Row of First Address, in case it isn't 1

                return a;
            }
        }

        public class lookTriggers
        {
            public lookTriggers() { }

            mungeFunctions MF = new mungeFunctions();

            public void runTriggers(int fileCount, BL.AnalysisObject a, int avgRowCount, int avgColCount, out int avgRowCounta, out int avgColCounta)
            {
                /**
                 * If the Row count is outsized for the data set, then it will be analyzed against two potential alternative routes to account for Excel error(s)
                **/

                int centerRowNum = 0;
                if (a.rowCount > 65000 || a.rowCount > (avgRowCount * 1.5))
                {
                    Console.WriteLine("Total Rows: " + a.rowCount);
                    if (a.rowCount > 65000) // Most common error is that UsedRange will return back over 65k rows.
                    {
                        Console.WriteLine("Option 1");
                        // As this is the default error response, it is assumed that number is wrong and skipped, thus the value 2 sent to ExcelWorkSheetNullCheck
                        centerRowNum = MF.ExcelWorkSheetNullCheck(a, a.splitUpAddresses, 2);
                    }
                    else // It is also possible that the UsedRange is technically correct, but there is bad data (there are 30 rows of data and a stray key on row 300).
                    {
                        Console.WriteLine("Option 2");
                        // As the row count may simply be significantly larger than the average, this version is more forgiving by sending value 1
                        centerRowNum = MF.ExcelWorkSheetNullCheck(a, a.splitUpAddresses, 1);
                    }
                    Console.WriteLine("Row 1 " + centerRowNum);
                    // ExcelRowChecker samples the object for data to determine the actual used range
                    centerRowNum = MF.ExcelRowChecker(a, centerRowNum, a.colCount, avgColCount);
                    Console.WriteLine("Row 2 " + centerRowNum);
                }
                else
                {
                    // If the number is not unusually large, we only use ExcelRowChecker to sample
                    Console.WriteLine("Option 3");
                    Console.WriteLine("Total Rows: " + a.rowCount);
                    centerRowNum = MF.ExcelRowChecker(a, a.rowCount, a.colCount, avgColCount);
                    Console.WriteLine("Row 1 " + centerRowNum);
                }
                MF.MakeAverage(fileCount, avgRowCount, centerRowNum, avgColCount, a.colCount, out avgRowCounta, out avgColCounta);
                Console.WriteLine("Avg Row " + avgRowCounta);
                a.rowCount = centerRowNum;
            }
        }

        public class mungeFunctions
        {
            public mungeFunctions() { }

            public int ExcelRowChecker(BL.AnalysisObject a, int centerRowNum, int theTotalColumns, int avgColCount) 
            {
                int falseRowCount = 0;
                int altRowCount = 0;
                int upperAltRowCount = 0;
                int goodRow = 0;
                double thePiece = .05;
                double checkRowRange = centerRowNum * thePiece;

                if (a.colCount >= avgColCount * 1.5)
                    a.colCount = avgColCount;

                for (int y = centerRowNum + 10; y > centerRowNum - checkRowRange; y--)
                {
                    Double emptyCellCount = 0;
                    for (int x = 1; x < a.colCount + 1; x++)
                    {
                        Double x2 = Convert.ToDouble(x);
                        try
                        {
                            if (a.allTheData[y, x] == null || a.allTheData[y, x].ToString().Trim().Equals(""))
                                emptyCellCount++;
                        }
                        catch (System.IndexOutOfRangeException e)
                        {
                            emptyCellCount++;
                        }
                        if (x >= (a.colCount * .7) && emptyCellCount / x2 > .7)
                        {
                            if (goodRow == 0)
                                falseRowCount++;
                            else
                                altRowCount++;
                            break;
                        }
                        if (x == a.colCount)
                            goodRow++;
                    }
                    if (falseRowCount > 1 && goodRow > 3)
                    {
                        if (falseRowCount + upperAltRowCount < 5 || (falseRowCount < 10 && upperAltRowCount < 10 && goodRow > 5 ))
                        {
                            centerRowNum = centerRowNum + 50;
                            centerRowNum = ExcelRowChecker(a, centerRowNum, a.colCount, avgColCount);
                            return centerRowNum;
                        }
                        else if (falseRowCount != 10)
                        {
                            centerRowNum = centerRowNum + 10 - falseRowCount + upperAltRowCount;
                            return centerRowNum;
                        }
                        break;
                    }
                    else if (y - 1 <= Math.Ceiling(centerRowNum - checkRowRange))
                    {
                        thePiece = thePiece + .05;
                        checkRowRange = centerRowNum * thePiece;
                        falseRowCount = falseRowCount + altRowCount;
                        if (goodRow > 0)
                            upperAltRowCount = upperAltRowCount + altRowCount;
                        altRowCount = 0;
                    }
                    if (y <= 0)
                    {
                        if (goodRow != 0)
                            break;
                        else 
                        {
                            centerRowNum = 0;
                            return centerRowNum;
                        }
                    }
                }
                if (falseRowCount < 10)
                    Console.WriteLine("Odd Result");
                centerRowNum = centerRowNum + 10 - falseRowCount;
                return centerRowNum;
            }

            public void MakeAverage(int fileCount, int avgRowCount, int centerRowNum, int avgColCount, int theTotalColumns, out int avgRowCounta, out int avgColCounta)
            {
                if(centerRowNum == 0)
                    fileCount--;
                if (fileCount == 0)
                {
                    avgRowCounta = (avgRowCount + centerRowNum) / (fileCount + 2);
                    if (theTotalColumns >= avgColCount * 1.5)
                        avgColCounta = (avgColCount + theTotalColumns) / (fileCount + 2);
                    else
                        avgColCounta = (avgColCount + avgColCount) / (fileCount + 2);
                }
                else
                {
                    avgRowCounta = ((avgRowCount * fileCount) + centerRowNum) / (fileCount + 1);
                    if (theTotalColumns >= avgColCount * 1.5)
                        avgColCounta = ((avgColCount * fileCount) + theTotalColumns) / (fileCount + 1);
                    else
                        avgColCounta = ((avgColCount * fileCount) + avgColCount) / (fileCount + 1);
                }
            }

            public int ExcelWorkSheetNullCheck(BL.AnalysisObject a, String[] splitUpAddresses, int lessThis) 
            {
                int count = splitUpAddresses.Length;
                string theVal = null;
                int centerRowNum = 0;
                while (theVal == null || theVal.Equals(""))
                {
                    var thisAddress = Regex.Split(splitUpAddresses[count - lessThis], @"(?:\$)");
                    int centerColNum = EL.singleExcel.ExcelColumnNameToNumber(thisAddress[1], a.startCol);
                    centerRowNum = Convert.ToInt32(thisAddress[2]);
                    int failCount = 0;
                    var tempVal = WiggleRoom(a, centerRowNum, centerColNum, failCount);
                    if (tempVal != null)
                        theVal = a.allTheData[centerRowNum, centerColNum].ToString().Trim();
                    lessThis++;
                }
                return centerRowNum;
            }

            public object WiggleRoom(BL.AnalysisObject a, int centerRowNum, int centerColNum, int failCount)
            {
                object tempVal = null;
                if (failCount == 3)
                    return tempVal;
                try
                {
                    tempVal = a.allTheData[centerRowNum, centerColNum];
                }
                catch (System.IndexOutOfRangeException e)
                {   
                    if (EL.singleExcel.GetExcelColumnName(centerColNum - 1) != "A")
                    {
                        centerColNum--;
                        failCount++;
                        WiggleRoom(a, centerRowNum, centerColNum, failCount);
                    }
                    else
                    {
                        centerColNum++;
                        failCount++;
                        WiggleRoom(a, centerRowNum, centerColNum, failCount);
                    }
                }
                return tempVal;
            }

        }

    }
}
