using System;
using System.Runtime.InteropServices;
using System.Threading;
using Aveva.Core.PMLNet;
using Excel = Microsoft.Office.Interop.Excel;

namespace PMLExcel
{
    /// <summary>
    /// Provides an interface to control an Excel workbook in PML.
    /// </summary>
    [PMLNetCallable()]
    public class PMLExcelWorkbook : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private Excel.Application app;
        private Excel.Workbook wb;

        [PMLNetCallable()]
        public PMLExcelWorkbook()
        {
        }

        [PMLNetCallable()]
        public void Assign(PMLExcelWorkbook that)
        {
            Close();
            app = that.app;
            wb = that.wb;
        }

        public void Dispose()
        {
            Close();
        }

        [PMLNetCallable()]
        public void Sleep(double ms)
        {
            Thread.Sleep((int)ms);
        }

        private void CreateApp()
        {
            if (app != null)
            {
                throw new PMLNetException(1000, 1, "Excel is already open");
            }
            app = new Excel.Application();
        }

        [PMLNetCallable()]
        public void Open()
        {
            CreateApp();
            wb = app.Workbooks.Add();
        }

        [PMLNetCallable()]
        public void Open(string filename)
        {
            CreateApp();
            try
            {
                wb = app.Workbooks.Open(Environment.ExpandEnvironmentVariables(filename));
            }
            catch (COMException ex)
            {
                Close();
                throw new PMLNetException(1000, 2, ex.Message);
            }
        }

        [PMLNetCallable()]
        public void Close()
        {
            if (app == null)
            {
                return;
            }
            app.DisplayAlerts = false;
            app.Quit();
            Marshal.ReleaseComObject(app);
            app = null;
            wb = null;
            GC.Collect();
        }

        [PMLNetCallable()]
        public void Show()
        {
            app.Visible = true;
            SetForegroundWindow((IntPtr)app.Hwnd);
        }

        [PMLNetCallable()]
        public void Hide()
        {
            app.Visible = false;
        }

        [PMLNetCallable()]
        public void Save()
        {
            wb.Save();
        }

        [PMLNetCallable()]
        public void SaveAs(string filename)
        {
            app.DisplayAlerts = false;
            wb.SaveAs(Environment.ExpandEnvironmentVariables(filename));
            app.DisplayAlerts = true;
        }

        [PMLNetCallable()]
        public PMLExcelRange Cells()
        {
            try
            {
                return new PMLExcelRange(app.Cells);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 3, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Cells(double r, double c)
        {
            try
            {
                return new PMLExcelRange((Excel.Range)app.Cells[(int)r, (int)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 4, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Range(string s)
        {
            try
            {
                return new PMLExcelRange(app.Range[s]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 5, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Selection()
        {
            try
            {
                return new PMLExcelRange((Excel.Range)app.Selection);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 6, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public void AddSheet()
        {
            wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
        }

        [PMLNetCallable()]
        public void AddSheet(string name)
        {
            ((Excel.Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count])).Name = name;
        }

        [PMLNetCallable()]
        public void CopySheet(string from, string to)
        {
            var ws = (Excel.Worksheet)wb.Worksheets[from];
            ws.Copy(After: ws);
            SetActiveSheetName(to);
        }

        [PMLNetCallable()]
        public void DeleteSheet(string name)
        {
            app.DisplayAlerts = false;
            ((Excel.Worksheet)wb.Worksheets[name]).Delete();
            app.DisplayAlerts = true;
        }

        [PMLNetCallable()]
        public void SelectSheet(string name)
        {
            ((Excel.Worksheet)wb.Worksheets[name]).Select();
        }

        [PMLNetCallable()]
        public void SetActiveSheetName(string name)
        {
            ((Excel.Worksheet)app.ActiveSheet).Name = name;
        }
    }
}