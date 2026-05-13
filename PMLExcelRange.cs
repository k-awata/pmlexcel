using System;
using System.Collections;
using System.Linq;
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
        public void SetValue(string s)
        {
            range.Value = s;
        }

        [PMLNetCallable()]
        public void SetValue(Hashtable rows)
        {
            var rCount = rows.Count;
            var cCount = rows.Values.OfType<Hashtable>().DefaultIfEmpty(new Hashtable()).Max(cols => cols.Count);
            if (cCount == 0)
            {
                var values = new object[rCount];
                for (int r = 0; r < rCount; r++)
                {
                    values[r] = rows[r + 1.0];
                }
                range.Value = values;
            }
            else
            {
                var values = new object[rCount, cCount];
                for (int r = 0; r < rCount; r++)
                {
                    if (rows[r + 1.0] is Hashtable cols)
                    {
                        for (int c = 0; c < cCount; c++)
                        {
                            values[r, c] = cols[c + 1.0];
                        }
                    }
                }
                range.Value = values;
            }
        }

        [PMLNetCallable()]
        public string GetValue()
        {
            if (range.Value is object[,] cells)
            {
                return Convert.ToString(cells[cells.GetLowerBound(0), cells.GetLowerBound(1)]);
            }
            return Convert.ToString(range.Value);
        }

        [PMLNetCallable()]
        public string GetFormula()
        {
            if (range.Formula is object[,] cells)
            {
                return Convert.ToString(cells[cells.GetLowerBound(0), cells.GetLowerBound(1)]);
            }
            return Convert.ToString(range.Formula);
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