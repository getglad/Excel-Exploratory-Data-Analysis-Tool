using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

using BL = DataLibrary.Library;
using EL = DataLibrary.ExcelLib;

namespace DataLibrary
{
    public class TermSampling
    {
        public class trainMethods
        {
            public trainMethods() { }

            public static void buildTrainList(Dictionary<string, int> termsDict, List<string> thoseWords)
            {
                int lowestBounds = Convert.ToInt32(Math.Floor(termsDict.First().Value * .7));
                foreach (var a in termsDict)
                {
                    if (a.Value >= lowestBounds)
                        thoseWords.Add(a.Key);
                    else
                        break;
                }
            }

            public static object[,] getTrainData(DirectoryInfo di, string fileLocation)
            {
                fileLocation = di + fileLocation;
                EL.singleExcel wordFile = new EL.singleExcel().createExcel(fileLocation);
                EL.singleExcel.ExcelWorkSheetChange(wordFile, 1);
                object[,] theWords;
                if (wordFile.excelRange.Count == 1)
                {
                    string temp = (string)wordFile.excelRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);
                    int[] myLengthsArray = new int[2] { 1, 1 };
                    int[] myBoundsArray = new int[2] { 1, 1 };
                    theWords = (object[,])Array.CreateInstance(typeof(String), myLengthsArray, myBoundsArray);
                    theWords[1, 1] = temp;
                }
                else
                    theWords = (object[,])wordFile.excelRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);
                EL.singleExcel.CloseSheet(wordFile);
                return theWords;
            }

            public static List<string> makeTrainList(object[,] theWords, List<string> words)
            {
                for (var x = 1; x < theWords.Length + 1; x++)
                {
                    if (theWords[x, 1] != null)
                        words.Add(theWords[x, 1].ToString());
                }
                return words;
            }

            public static void runTrainScan(List<string> words, List<string> looseWords, int avgColCount, List<BL.AnalysisObject> ListofFiles, out Dictionary<string, int> discreteTermsDict, out Dictionary<string, int> betterTermsDict, out Dictionary<string, int> secondaryTermsDict)
            {
                List<string> terms = new List<string>();
                List<string> betterTerms = new List<string>();
                List<string> secondaryTerms = new List<string>();
                int filesMatched = 0;

                string Pattern = "(";
                Pattern += string.Join("|", words.Select(Regex.Escape).ToArray());
                if (looseWords.Count > 0)
                {
                    Pattern += "|";
                    Pattern += string.Join("|", looseWords.Select(Regex.Escape).ToArray());
                }
                Pattern += ")";
                Regex rPattern = new Regex(Pattern);

                foreach (var a in ListofFiles)
                {
                    Console.WriteLine("Scanning " + a.FileName);
                    filesMatched = readRowData(a, avgColCount, rPattern, terms, betterTerms, secondaryTerms, words, looseWords, filesMatched);
                }

                Console.WriteLine("Building List 1");
                discreteTermsDict = terms.GroupBy(x => x).OrderByDescending(x => x.Count()).ToDictionary(g => g.Key, g => g.Count());
                Console.WriteLine("Building List 2");
                betterTermsDict = betterTerms.GroupBy(x => x).OrderByDescending(x => x.Count()).ToDictionary(g => g.Key, g => g.Count());
                Console.WriteLine("Building List 3");
                secondaryTermsDict = secondaryTerms.GroupBy(x => x).OrderByDescending(x => x.Count()).ToDictionary(g => g.Key, g => g.Count());

                Console.WriteLine("Outputting List 1");
                if (discreteTermsDict.Count > 0) EL.singleExcel.outputDictionaryToExcel(discreteTermsDict, "discrete");
                Console.WriteLine("Outputting List 2");
                if (betterTermsDict.Count > 0) EL.singleExcel.outputDictionaryToExcel(betterTermsDict, "strongMatch");
                Console.WriteLine("Outputting List 3");
                if (secondaryTermsDict.Count > 0) EL.singleExcel.outputDictionaryToExcel(secondaryTermsDict, "secondaryMatch");

                Console.WriteLine("Number of Files Matched: " + filesMatched);
                Console.ReadLine();
            }

            public static int readRowData(BL.AnalysisObject a, int avgColCount, Regex rPattern, List<string> terms, List<string> betterTerms, List<string> secondaryTerms, List<string> words, List<string> looseWords, int filesMatched)
            {
                if (a.colCount >= avgColCount * 1.5)
                    a.colCount = avgColCount;

                int matchCountLimiter = Convert.ToInt32(Math.Ceiling(Math.Min(words.Count * .5, a.colCount * .7)));

                for (int y = 1; y < a.rowCount + 1; y++)
                {
                    Double emptyCellCount = 0;
                    bool primary = false;
                    int matchCount = 0;
                    for (int x = 1; x < a.colCount + 1; x++)
                    {
                        try
                        {
                            if (a.allTheData[y, x] == null || a.allTheData[y, x].ToString().Trim().Equals(""))
                                emptyCellCount++;
                            else
                            {
                                string tempVar = a.allTheData[y, x].ToString().ToLower().Trim();

                                if (rPattern.IsMatch(tempVar))
                                {
                                    if (words[0].Equals(tempVar))
                                        primary = true;
                                    matchCount++;
                                }
                            }
                        }
                        catch (System.NullReferenceException e)
                        {
                            emptyCellCount++;
                        }
                    }

                    if ((matchCount >= matchCountLimiter || primary) && matchCount != 0)
                    {
                        getRowData(a, avgColCount, y, rPattern, terms);

                        if ((matchCount >= matchCountLimiter && primary) && words.Count != 1)
                            getRowData(a, avgColCount, y, rPattern, betterTerms);

                        if ((matchCount >= Math.Ceiling(matchCountLimiter * .7) && !primary) && words.Count != 1)
                            getRowData(a, avgColCount, y, rPattern, secondaryTerms);

                        filesMatched++;

                        return filesMatched;
                    }
                }

                return filesMatched;
            }

            public static List<string> getRowData(BL.AnalysisObject a, int avgColCount, int focusRow, Regex rPattern, List<string> terms)
            {
                if (a.colCount >= avgColCount * 1.5)
                    a.colCount = avgColCount;

                for (int x = 1; x < a.colCount + 1; x++)
                {
                    try
                    {
                        if (a.allTheData[focusRow, x] != null)
                        {
                            string tempVar = a.allTheData[focusRow, x].ToString().ToLower().Trim();

                            if (tempVar != null && !tempVar.Equals("") && !rPattern.IsMatch(tempVar))
                                terms.Add(tempVar);
                        }
                    }
                    catch (System.NullReferenceException e)
                    {

                    }
                }

                return terms;
            }

        }
    }
}
