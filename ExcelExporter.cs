using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

//using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelExport
{
    class ExcelExporter
    {
        ~ExcelExporter()
        {
            if (application != null)
            {
                int excelProcessId = -1;
                GetWindowThreadProcessId(application.Hwnd, ref excelProcessId);
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }

                if (workBook != null)
                {
                    try
                    {
                        workBook.Close();
                    }
                    catch (Exception)
                    {
                    }
                    Marshal.ReleaseComObject(workBook);
                }
                application.Quit();
                Marshal.ReleaseComObject(application);

                application = null;
                // Прибиваем висящий процесс
                Process process = Process.GetProcessById(excelProcessId);
                process.Kill();
            }
          
        }
        private Application application;
        private Workbook workBook;
        private Worksheet worksheet;
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(int hWnd, ref int lpdwProcessId);

        public ExcelExporter(string templateXlsm)
        {
            
            application = new Application();
            application.DisplayAlerts = false;

            workBook = application.Workbooks.Open(templateXlsm);
        }

        public void SaveExcel(string path)
        {
            workBook.SaveAs(Path.Combine(Environment.CurrentDirectory, path));
        }

        //public void OpenExcel()
        //{
        //    application.Visible = true;
        //}

        public void InsertData<T>(string rangeName, List<T> data)
        {
            List<string> unblockRanes = new List<string>();
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                List<Range> listRanges = GetNamedRange(rangeName, sheet);
                switch (GetArreaRange(listRanges))
                {
                    case ArreaRange.OneRow:
                        FillRange(rangeName, data, sheet, listRanges, ArreaRange.OneRow);
                        break;
                    case ArreaRange.OneCoumn:
                        FillRange(rangeName, data, sheet, listRanges, ArreaRange.OneCoumn);
                        break;
                    case ArreaRange.None:
                        throw new Exception("Область " + rangeName + " должна быть либо строкой либо стройчкой");
                }
            }

            foreach (Worksheet sheet in workBook.Worksheets)
            {
                foreach (Range range in sheet.UsedRange)
                {
                    Console.WriteLine(range.Address + " " + range.Locked);
                }
            }
        }

        private static void FillRange<T>(string rangeName, List<T> data, Worksheet sheet, List<Range> listRanges, ArreaRange arreaRange)
        {
            var properties = typeof(T).GetProperties();
            foreach (var range in listRanges)
            {
                string fieldName = Regex.Replace(range.Value.ToString(), "{" + rangeName + ":(.+)}", "$1");
                var info = properties.SingleOrDefault(x => x.Name == fieldName);
                if (info == null)
                    throw new Exception("Не найдено свойство " + fieldName + ". Яцейка " + range.Address);
                
                var range1 = arreaRange == ArreaRange.OneRow
                    ? range.Resize[data.Count, 1]
                    : range.Resize[1, data.Count];
                var address = range1.Address;

                //range1.Locked = info.GetCustomAttribute(typeof(unprotectedAttribute)) == null;

                range1.Insert(arreaRange == ArreaRange.OneRow ? XlInsertShiftDirection.xlShiftDown : XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                range1 = sheet.Range[address];

                var value = arreaRange == ArreaRange.OneRow
                    ? new object[data.Count, 1]
                    : new object[1, data.Count];
                for (var i = 0; i < data.Count; i++)
                {
                    if (arreaRange == ArreaRange.OneRow)
                        value[i, 0] = info.GetValue(data[i], null);
                    else
                        value[0, i] = info.GetValue(data[i], null);
                }
                range1.set_Value(Missing.Value, value);
                range.Delete();
            }
        }

        public void InsertField(string name, object value)
        {
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                foreach (Range range in sheet.UsedRange)
                {
                    if (range.Value == null) continue;

                    string currentValue = range.Value.ToString();
                    string temp = "{" + name + "}";

                    if (currentValue.Contains(temp))
                        range.Value = currentValue.Replace(temp, value.ToString());

                }
            }
        }
        public void InsertFields(Dictionary<string, object> fields)
        {
            foreach (var item in fields)
            {
                InsertField(item.Key,item.Value);
            }
        }

        public void RunMacros()
        {
            workBook.RunAutoMacros(XlRunAutoMacro.xlAutoOpen);
        }

        public void ProtectSheets(string password)
        {
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                sheet.Protect(password);
            }
        }


        private List<Range> GetNamedRange(string rangeName, Worksheet sheet)
        {
            List<Range> listRanges = new List<Range>();
            foreach (Range range in sheet.UsedRange)
            {
                if (range.Value == null) continue;
                if (Regex.IsMatch(range.Value.ToString(), "^{" + rangeName + ":.+}$"))
                    listRanges.Add(range);
            }
            return listRanges;
        }

        enum ArreaRange
        {
            None,
            OneRow,
            OneCoumn
        }

        private ArreaRange GetArreaRange(List<Range> listRanges)
        {
            if (listRanges.Count == 1)
                return ArreaRange.None;
            var pattern = @"^\$(.+)\$(.+)$";
            if (listRanges.Select(x => Regex.Replace(x.Address.ToString(), pattern, "$1")).ToList().Distinct().Count() == 1)
                return ArreaRange.OneCoumn;
            if (listRanges.Select(x => Regex.Replace(x.Address.ToString(), pattern, "$2")).ToList().Distinct().Count() == 1)
                return ArreaRange.OneRow;
            return ArreaRange.None;
        }
    }
}
