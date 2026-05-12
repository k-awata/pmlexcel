using System;
using System.Collections;
using System.Runtime.InteropServices;
using Aveva.Core.PMLNet;
using Excel = Microsoft.Office.Interop.Excel;

namespace PMLExcel
{
    /// <summary>
    /// Provides an interface to control an Excel range in PML.
    /// </summary>
    [PMLNetCallable()]
    public class PMLExcelRange
    {
        private Excel.Range range;

        [PMLNetCallable()]
        public PMLExcelRange()
        {
        }

        public PMLExcelRange(Excel.Range r)
        {
            range = r;
        }

        [PMLNetCallable()]
        public void Assign(PMLExcelRange that)
        {
            range = that.range;
        }

        [PMLNetCallable()]
        public string GetAddress()
        {
            return range.Address;
        }

        [PMLNetCallable()]
        public Hashtable ToArray()
        {
            var cells = new Hashtable();
            for (long i = 1; i <= (long)range.CountLarge; i++)
            {
                cells.Add((double)i, new PMLExcelRange((Excel.Range)range.Cells[i]));
            }
            return cells;
        }

        [PMLNetCallable()]
        public Hashtable ToArrayOfArray()
        {
            var rows = new Hashtable();
            for (int r = 1; r <= range.Rows.Count; r++)
            {
                var cols = new Hashtable();
                for (int c = 1; c <= range.Columns.Count; c++)
                {
                    cols.Add((double)c, new PMLExcelRange((Excel.Range)range.Cells[r, c]));
                }
                rows.Add((double)r, cols);
            }
            return rows;
        }

        [PMLNetCallable()]
        public PMLExcelRange Offset(double r, double c)
        {
            try
            {
                return new PMLExcelRange(range.Offset[(long)r, (long)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 11, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange CurrentRegion()
        {
            return new PMLExcelRange(range.CurrentRegion);
        }

        [PMLNetCallable()]
        public string GetFormula()
        {
            if (range.Formula is object[,] objs)
            {
                return Convert.ToString(objs[objs.GetLowerBound(0), objs.GetLowerBound(1)]);
            }
            return Convert.ToString(range.Formula);
        }

        [PMLNetCallable()]
        public void SetFormula(string s)
        {
            range.Formula = s;
        }

        [PMLNetCallable()]
        public string GetValue()
        {
            if (range.Value is object[,] objs)
            {
                return Convert.ToString(objs[objs.GetLowerBound(0), objs.GetLowerBound(1)]);
            }
            return Convert.ToString(range.Value);
        }

        [PMLNetCallable()]
        public void SetValue(string s)
        {
            range.Value = s;
        }

        [PMLNetCallable()]
        public void Select()
        {
            range.Select();
        }

        [PMLNetCallable()]
        public void ClearContents()
        {
            range.ClearContents();
        }

        [PMLNetCallable()]
        public void AutoFitColumns()
        {
            range.EntireColumn.AutoFit();
        }

        [PMLNetCallable()]
        public void Copy()
        {
            range.Copy();
        }

        [PMLNetCallable()]
        public void Paste()
        {
            ((Excel.Worksheet)range.Parent).Paste(range);
        }

        [PMLNetCallable()]
        public void CreateTable(string name, bool header)
        {
            ((Excel.Worksheet)range.Parent).ListObjects.Add(
                Source: range,
                XlListObjectHasHeaders: header ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo
            ).Name = name;
        }
    }
}