using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Linq;

using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;

namespace DataLibrary
{
    public class ExcelLib
    {
        public class singleExcel
        {
            public singleExcel() { }

            public Excel.Range excelCell { get; set; }
            public Excel._Worksheet xlWorksheet { get; set; }
            public Excel.Application xlApp { get; set; }
            public Excel.Workbook xlWorkbook { get; set; }
            public Excel.Range brng { get; set; }
            public Process excelProcess { get; set; }
            public Excel.Range excelRange { get; set; }

            public singleExcel createExcel()
            {
                var thisExcel = new singleExcel();
                thisExcel.xlApp = new Excel.Application();
                // var thisExcel.xlApp = ExcelInteropService.GetExcelInterop(thisExcel);
                thisExcel.xlWorkbook = thisExcel.xlApp.Workbooks.Add(1);
                thisExcel.xlApp.Visible = false;
                thisExcel.xlApp.DisplayAlerts = false;
                bool failed = false;
                do
                {
                    try
                    {
                        thisExcel.xlWorkbook.DoNotPromptForConvert = true;
                        thisExcel.xlWorkbook.CheckCompatibility = false;
                        thisExcel.xlWorkbook.Unprotect();
                        failed = false;
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        failed = true;
                    }
                } while (failed);
                thisExcel.xlWorksheet = (Excel._Worksheet)thisExcel.xlApp.Workbooks[1].Worksheets[1];
                return thisExcel;
            }

            public singleExcel createExcel(string fileLocation)
            {
                var thisExcel = new singleExcel();
                thisExcel.xlApp = new Excel.Application();
                // var thisExcel.xlApp = ExcelInteropService.GetExcelInterop(thisExcel);
                thisExcel.xlWorkbook = thisExcel.xlApp.Workbooks.Open(fileLocation);
                thisExcel.xlApp.Visible = false;
                thisExcel.xlApp.DisplayAlerts = false;
                bool failed = false;
                do
                {
                    try
                    {
                        thisExcel.xlWorkbook.DoNotPromptForConvert = true;
                        thisExcel.xlWorkbook.CheckCompatibility = false;
                        thisExcel.xlWorkbook.Unprotect();
                        failed = false;
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        failed = true;
                    }
                } while (failed);
                return thisExcel;
            }

            public static void outputObjectToExcel(Library.GroupStats a)
            {
                int x = 0;
                object[,] excelDrop = new object[a.GetType().GetProperties().Length, 2];
                foreach (PropertyInfo objPart in a.GetType().GetProperties())
                {
                    excelDrop[x, 0] = objPart.Name;
                    excelDrop[x, 1] = objPart.GetValue(a, null);
                    x++;
                }

                singleExcel outputExcel = new singleExcel().createExcel();
                string secondRange = "B" + a.GetType().GetProperties().Length;
                string thisCell = "A1:" + secondRange;
                var cell = outputExcel.xlWorksheet.Range[thisCell, Type.Missing];
                cell.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, excelDrop);
                string path = "C:\\Users\\" + Environment.UserName + "\\Documents\\Email Attachments\\testdata\\";
                string fileName2 = "output";
                outputExcel.xlWorkbook.SaveAs(path + fileName2);
                singleExcel.CloseSheet(outputExcel);
            }

            public static void outputListToExcel(List<string> a, string fileName)
            {
                int x = 0;
                object[,] excelDrop = new object[a.Count, 1];
                foreach (string b in a)
                {
                    excelDrop[x, 0] = b;
                    x++;
                }
                singleExcel outputExcel = new singleExcel().createExcel();
                string secondRange = "A" + a.Count;
                string thisCell = "A1:" + secondRange;
                var cell = outputExcel.xlWorksheet.Range[thisCell, Type.Missing];
                cell.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, excelDrop);
                string path = "C:\\Users\\" + Environment.UserName + "\\Documents\\Email Attachments\\testdata\\";
                string fileName2 = fileName + "-" + a.Count;
                outputExcel.xlWorkbook.SaveAs(path + fileName2);
                singleExcel.CloseSheet(outputExcel);
            }

            public static void outputDictionaryToExcel(Dictionary<string, int> a, string fileName)
            {
                int x = 0;
                object[,] excelDrop = new object[a.Count, 2];
                foreach (var b in a)
                {
                    excelDrop[x, 0] = b.Key;
                    excelDrop[x, 1] = b.Value;
                    x++;
                }
                singleExcel outputExcel = new singleExcel().createExcel();
                string secondRange = "B" + a.Count;
                string thisCell = "A1:" + secondRange;
                var cell = outputExcel.xlWorksheet.Range[thisCell, Type.Missing];
                cell.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, excelDrop);
                string path = "C:\\Users\\" + Environment.UserName + "\\Documents\\Email Attachments\\testdata\\";
                string fileName2 = fileName + "-" + a.Count;
                outputExcel.xlWorkbook.SaveAs(path + fileName2);
                singleExcel.CloseSheet(outputExcel);
            }

            public static void CloseSheet(singleExcel thisExcel)
            {
                if (thisExcel.excelProcess != null)
                {
                    try
                    {
                        thisExcel.excelProcess.Kill();
                        thisExcel.excelProcess.Dispose();
                    }
                    catch (Exception ex)
                    {

                    }
                }
                else
                {
                    thisExcel.xlWorkbook.Close(true);
                    thisExcel.xlApp.Quit();
                }
                releaseObject(thisExcel.xlWorksheet);
                releaseObject(thisExcel.xlWorkbook);
                releaseObject(thisExcel.xlApp);
                releaseObject(thisExcel.excelProcess);
                releaseObject(thisExcel);
            }

            public static void releaseObject(object obj)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch
                {
                    obj = null;
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            public static string GetExcelColumnName(int columnNumber)
            {
                int dividend = columnNumber;
                string columnName = String.Empty;
                int modulo;

                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = (int)((dividend - modulo) / 26);
                }

                return columnName;
            }

            public static int StaticExcelColumnNameToNumber(string columnName)
            {
                columnName = columnName.ToUpperInvariant();

                int sum = 0;

                for (int i = 0; i < columnName.Length; i++)
                {
                    sum *= 26;
                    sum += (columnName[i] - 'A' + 1);
                }

                return sum;
            }

            public static int ExcelColumnNameToNumber(string columnName, string startCol)
            {
                columnName = columnName.ToUpperInvariant();

                int sum = 0;

                for (int i = 0; i < columnName.Length; i++)
                {
                    sum *= 26;
                    sum += (columnName[i] - Convert.ToChar(startCol) + 1);
                }

                return sum;
            }

            public static singleExcel ExcelWorkSheetChange(singleExcel thisExcel, int x)
            {
                thisExcel.xlWorksheet = (Excel._Worksheet)thisExcel.xlApp.Workbooks[1].Worksheets[x];
                thisExcel.xlWorksheet.Columns.ClearFormats();
                thisExcel.xlWorksheet.Rows.ClearFormats();
                thisExcel.excelRange = thisExcel.xlWorksheet.UsedRange;
                return thisExcel;
            }
        }

        public class ExcelInteropService
        {
            private const string EXCEL_CLASS_NAME = "EXCEL7";

            private const uint DW_OBJECTID = 0xFFFFFFF0;

            private static Guid rrid = new Guid("{00020400-0000-0000-C000-000000000046}");

            public delegate bool EnumChildCallback(int hwnd, ref int lParam);

            [DllImport("Oleacc.dll")]
            public static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, ref Microsoft.Office.Interop.Excel.Window ptr);

            [DllImport("User32.dll")]
            public static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

            [DllImport("User32.dll")]
            public static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);

            public static Microsoft.Office.Interop.Excel.Application GetExcelInterop(singleExcel thisExcel, int? processId = null)
            {
                var p = processId.HasValue ? Process.GetProcessById(processId.Value) : Process.Start("excel.exe");
                thisExcel.excelProcess = p;
                Stopwatch updateStopwatch = Stopwatch.StartNew();
                while (updateStopwatch.ElapsedMilliseconds < 200) { }
                updateStopwatch.Stop();
                try
                {
                    return new ExcelInteropService().SearchExcelInterop(p);
                }
                catch (Exception)
                {
                    Debug.Assert(p != null, "p != null");
                    return GetExcelInterop(thisExcel, p.Id);
                }
            }

            private bool EnumChildFunc(int hwndChild, ref int lParam)
            {
                var buf = new StringBuilder(128);
                GetClassName(hwndChild, buf, 128);
                if (buf.ToString() == EXCEL_CLASS_NAME) { lParam = hwndChild; return false; }
                return true;
            }

            private Microsoft.Office.Interop.Excel.Application SearchExcelInterop(Process p)
            {
                Microsoft.Office.Interop.Excel.Window ptr = null;
                int hwnd = 0;

                int hWndParent = (int)p.MainWindowHandle;
                if (hWndParent == 0) throw new ExcelMainWindowNotFoundException();

                EnumChildWindows(hWndParent, EnumChildFunc, ref hwnd);
                if (hwnd == 0) throw new ExcelChildWindowNotFoundException();

                int hr = AccessibleObjectFromWindow(hwnd, DW_OBJECTID, rrid.ToByteArray(), ref ptr);
                if (hr < 0) throw new AccessibleObjectNotFoundException();
                return ptr.Application;
            }
        }

    }
}
