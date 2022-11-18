using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Diagnostics;
using System.IO;

namespace ExcelHelper
{
    public class ExcelDocConverter
    {
        Hashtable myHashtable;
        int MyExcelProcessId;

        Excel.Application excel;
        Excel.Workbook wbk;
        Excel.Worksheet worksheet1;

        object missing = System.Reflection.Missing.Value;

        public enum FormatType
        {
            XLS,
            XLSX,
            PDF,
            XPS,
            CSV
        }

        public void Convert(FormatType formatType, string originalFile, string targetFile)
        {
            CheckForExistingExcellProcesses();

            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.ScreenUpdating = false;
            excel.DisplayAlerts = false;

            GetTheExcelProcessIdThatUsedByThisInstance();

            wbk = excel.Workbooks.Open(originalFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            switch (formatType)
            {
                case FormatType.XLS:
                    {
                        wbk.SaveAs(targetFile, Excel.XlFileFormat.xlExcel8, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                    }
                    break;
                case FormatType.XLSX:
                    {
                        wbk.SaveAs(targetFile, Excel.XlFileFormat.xlWorkbookDefault, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                    }
                    break;
                case FormatType.PDF:
                    {
                        wbk.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, targetFile, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, false, true, missing, missing, false, missing);
                    }
                    break;
                case FormatType.XPS:
                    {
                        wbk.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypeXPS, targetFile, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, false, true, missing, missing, false, missing);
                    }
                    break;
                case FormatType.CSV:
                    {
                        wbk.SaveAs(targetFile, Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                    }
                    break;
                default:
                    break;
            }

            ReleaseExcelResources();
            KillExcelProcessThatUsedByThisInstance();
        }

        void ReleaseExcelResources()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet1);
            }
            catch
            { }
            finally
            {
                worksheet1 = null;
            }

            try
            {
                if (wbk != null)
                    wbk.Close(false, missing, missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbk);
            }
            catch
            { }
            finally
            {
                wbk = null;
            }

            try
            {
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            catch
            { }
            finally
            {
                excel = null;
            }
        }

        void CheckForExistingExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        void GetTheExcelProcessIdThatUsedByThisInstance()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");

            // Search For the Right Excel
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (myHashtable == null)
                    return;

                if (myHashtable.ContainsKey(ExcelProcess.Id) == false)
                {
                    // Get the process ID
                    MyExcelProcessId = ExcelProcess.Id;
                }
            }

            AllProcesses = null;
        }

        void KillExcelProcessThatUsedByThisInstance()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");

            foreach (Process ExcelProcess in AllProcesses)
            {
                if (myHashtable == null)
                    return;

                if (ExcelProcess.Id == MyExcelProcessId)
                    ExcelProcess.Kill();
            }

            AllProcesses = null;
        }
    }
}