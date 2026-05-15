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
        public Excel.Range Raw { get; private set; }

        [PMLNetCallable()]
        public PMLExcelRange()
        {
        }

        public PMLExcelRange(Excel.Range r)
        {
            Raw = r;
        }

        [PMLNetCallable()]
        public void Assign(PMLExcelRange that)
        {
            Raw = that.Raw;
        }

        public override string ToString()
        {
            return GetValue();
        }

        [PMLNetCallable()]
        public PMLExcelRange Cells(double i)
        {
            try
            {
                return new PMLExcelRange((Excel.Range)Raw.Cells[(long)i]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 11, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Cells(double r, double c)
        {
            try
            {
                return new PMLExcelRange((Excel.Range)Raw.Cells[(int)r, (int)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 12, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Offset(double r, double c)
        {
            try
            {
                return new PMLExcelRange(Raw.Offset[(int)r, (int)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 13, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange Resize(double r, double c)
        {
            try
            {
                return new PMLExcelRange(Raw.Resize[(int)r, (int)c]);
            }
            catch (COMException)
            {
                throw new PMLNetException(1000, 14, "Invalid range reference");
            }
        }

        [PMLNetCallable()]
        public PMLExcelRange CurrentRegion()
        {
            return new PMLExcelRange(Raw.CurrentRegion);
        }

        [PMLNetCallable()]
        public string GetAddress()
        {
            return Raw.Address;
        }

        [PMLNetCallable()]
        public double Count()
        {
            return (long)Raw.CountLarge;
        }

        [PMLNetCallable()]
        public double CountRows()
        {
            return Raw.Rows.Count;
        }

        [PMLNetCallable()]
        public double CountColumns()
        {
            return Raw.Columns.Count;
        }

        [PMLNetCallable()]
        public void SetValue(string s)
        {
            Raw.Value = s;
        }

        [PMLNetCallable()]
        public void SetValue(Hashtable rows)
        {
            var rCount = rows.Count;
            var cCount = rows.Values.OfType<Hashtable>().DefaultIfEmpty(new Hashtable()).Max(cols => cols.Count);
            if (cCount == 0)
            {
                Raw.Value = rows.Keys.OfType<double>().OrderBy(key => key).Select(key => rows[key]).ToArray();
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
                Raw.Value = values;
            }
        }

        [PMLNetCallable()]
        public string GetValue()
        {
            if (Raw.Value is object[,] cells)
            {
                return Convert.ToString(cells[cells.GetLowerBound(0), cells.GetLowerBound(1)]);
            }
            return Convert.ToString(Raw.Value);
        }

        [PMLNetCallable()]
        public string GetFormula()
        {
            if (Raw.Formula is object[,] cells)
            {
                return Convert.ToString(cells[cells.GetLowerBound(0), cells.GetLowerBound(1)]);
            }
            return Convert.ToString(Raw.Formula);
        }

        [PMLNetCallable()]
        public void ClearContents()
        {
            Raw.ClearContents();
        }

        [PMLNetCallable()]
        public void Select()
        {
            Raw.Select();
        }

        [PMLNetCallable()]
        public void DefineName(string name)
        {
            Raw.Name = name;
        }

        [PMLNetCallable()]
        public void CreateTable(string name, bool hasHeaders)
        {
            ((Excel.Worksheet)Raw.Parent).ListObjects.Add(
                Source: Raw,
                XlListObjectHasHeaders: hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo
            ).Name = name;
        }

        [PMLNetCallable()]
        public double GetRowHeight()
        {
            return (double)Raw.EntireRow.RowHeight;
        }

        [PMLNetCallable()]
        public void SetRowHeight(double height)
        {
            Raw.EntireRow.RowHeight = height;
        }

        [PMLNetCallable()]
        public void AutoFitRow()
        {
            Raw.EntireRow.AutoFit();
        }

        [PMLNetCallable()]
        public double GetColumnWidth()
        {
            return (double)Raw.EntireColumn.ColumnWidth;
        }

        [PMLNetCallable()]
        public void SetColumnWidth(double width)
        {
            Raw.EntireColumn.ColumnWidth = width;
        }

        [PMLNetCallable()]
        public void AutoFitColumn()
        {
            Raw.EntireColumn.AutoFit();
        }

        [PMLNetCallable()]
        public void Copy()
        {
            Raw.Copy();
        }

        [PMLNetCallable()]
        public void Paste()
        {
            ((Excel.Worksheet)Raw.Parent).Paste(Raw);
        }
    }
}