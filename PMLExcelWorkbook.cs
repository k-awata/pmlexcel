using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Aveva.Core.PMLNet;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

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

        private Application app;
        private Workbook wb;

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
            app = new Application();

            // app.SheetBeforeDoubleClick += new AppEvents_SheetBeforeDoubleClickEventHandler(RaiseSheetBeforeDoubleClick);
            // app.SheetBeforeRightClick += new AppEvents_SheetBeforeRightClickEventHandler(RaiseSheetBeforeRightClick);
            // app.SheetChange += new AppEvents_SheetChangeEventHandler(RaiseSheetChange);
            // app.SheetSelectionChange += new AppEvents_SheetSelectionChangeEventHandler(RaiseSheetSelectionChange);
        }

        // [PMLNetCallable()]
        // public event PMLNetDelegate.PMLNetEventHandler SheetBeforeDoubleClick;

        // private void RaiseSheetBeforeDoubleClick(object Sh, Range Target, ref bool Cancel)
        // {
        //     SheetBeforeDoubleClick?.Invoke(new ArrayList { new PMLExcelRange(Target), Cancel });
        // }

        // [PMLNetCallable()]
        // public event PMLNetDelegate.PMLNetEventHandler SheetBeforeRightClick;

        // private void RaiseSheetBeforeRightClick(object Sh, Range Target, ref bool Cancel)
        // {
        //     SheetBeforeRightClick?.Invoke(new ArrayList { new PMLExcelRange(Target), Cancel });
        // }

        // [PMLNetCallable()]
        // public event PMLNetDelegate.PMLNetEventHandler SheetChange;

        // private void RaiseSheetChange(object Sh, Range Target)
        // {
        //     SheetChange?.Invoke(new ArrayList { new PMLExcelRange(Target) });
        // }

        // [PMLNetCallable()]
        // public event PMLNetDelegate.PMLNetEventHandler SheetSelectionChange;

        // private void RaiseSheetSelectionChange(object Sh, Range Target)
        // {
        //     SheetSelectionChange?.Invoke(new ArrayList { new PMLExcelRange(Target) });
        // }

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
            return new PMLExcelRange(app.Cells);
        }

        [PMLNetCallable()]
        public PMLExcelRange Cells(double r, double c)
        {
            try
            {
                return new PMLExcelRange((Range)app.Cells[(int)r, (int)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 3, "Invalid range reference");
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
                throw new PMLNetException(1000, 4, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Selection()
        {
            if (app.Selection is Range r)
            {
                return new PMLExcelRange(r);
            }
            else
            {
                throw new PMLNetException(1000, 5, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange UsedRange()
        {
            return new PMLExcelRange(((Worksheet)app.ActiveSheet).UsedRange);
        }

        [PMLNetCallable()]
        public void AddSheet()
        {
            wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
        }

        [PMLNetCallable()]
        public void AddSheet(string name)
        {
            ((Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count])).Name = name;
        }

        [PMLNetCallable()]
        public void CopySheet(string from, string to)
        {
            var ws = (Worksheet)wb.Worksheets[from];
            ws.Copy(After: ws);
            NameActiveSheet(to);
        }

        [PMLNetCallable()]
        public void DeleteSheet(string name)
        {
            app.DisplayAlerts = false;
            ((Worksheet)wb.Worksheets[name]).Delete();
            app.DisplayAlerts = true;
        }

        [PMLNetCallable()]
        public void ActivateSheet(string name)
        {
            ((Worksheet)wb.Worksheets[name]).Activate();
        }

        [PMLNetCallable()]
        public void NameActiveSheet(string name)
        {
            ((Worksheet)app.ActiveSheet).Name = name;
        }

        [PMLNetCallable()]
        public void CopySelection()
        {
            ((dynamic)app.Selection).Copy();
        }

        [PMLNetCallable()]
        public void Paste()
        {
            ((Worksheet)app.ActiveSheet).Paste();
        }

        [PMLNetCallable()]
        public void NameSelection(string name)
        {
            ((dynamic)app.Selection).Name = name;
        }

        [PMLNetCallable()]
        public void DeleteName(string name)
        {
            try
            {
                wb.Names.Item(name).Delete();
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 6, "Invalid name");
            }
        }

        [PMLNetCallable()]
        public void SelectShape(string name)
        {
            try
            {
                ((Worksheet)app.ActiveSheet).Shapes.Range[name].Select();
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 7, "Invalid shape name");
            }
        }

        [PMLNetCallable()]
        public bool SelectAllShapes()
        {
            var s = (Worksheet)app.ActiveSheet;
            s.Shapes.SelectAll();
            return s.Shapes.Count > 0;
        }

        [PMLNetCallable()]
        public void DeleteSelection()
        {
            ((dynamic)app.Selection).Delete();
        }

        [PMLNetCallable()]
        public void FitSelectedPictureInRange(PMLExcelRange range)
        {
            if (app.Selection is Picture p)
            {
                double rHeight = (double)range.Raw.Height - 2;
                double rWidth = (double)range.Raw.Width - 2;
                p.Top = (double)range.Raw.Top + 1;
                p.Left = (double)range.Raw.Left + 1;
                p.ShapeRange.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                if (rHeight < rWidth)
                {
                    p.ShapeRange.ScaleHeight((float)(rHeight / p.Height), MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
                }
                else
                {
                    p.ShapeRange.ScaleWidth((float)(rWidth / p.Width), MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
                }
                p.Placement = XlPlacement.xlMoveAndSize;
            }
            else
            {
                throw new PMLNetException(1000, 8, "Selection is not Picture");
            }
        }

        private object ConvertPMLTypeToObject(object obj)
        {
            if (obj is PMLExcelRange r)
            {
                return r.Raw;
            }
            else if (obj is Hashtable h)
            {
                return h.Keys.OfType<double>().OrderBy(key => key).Select(key => ConvertPMLTypeToObject(h[key])).ToArray();
            }
            return obj;
        }

        [PMLNetCallable()]
        public string WorksheetFunction(string function, Hashtable args)
        {
            MethodInfo method = typeof(WorksheetFunction).GetMethod(
                function.Replace(".", "_"),
                BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod
            ) ?? throw new PMLNetException(1000, 9, "Function is not found");
            var argTypes = method.GetParameters().Select(p => p.ParameterType).ToArray();
            var argValues = new object[argTypes.Count()];
            for (int i = 0; i < argTypes.Count(); i++)
            {
                var val = ConvertPMLTypeToObject(args[i + 1.0]);
                if (argTypes[i] != typeof(object))
                {
                    argValues[i] = TypeDescriptor.GetConverter(argTypes[i]).ConvertFrom(val);
                }
                else
                {
                    argValues[i] = val;
                }
            }
            return method.Invoke(app.WorksheetFunction, argValues).ToString();
        }
    }
}