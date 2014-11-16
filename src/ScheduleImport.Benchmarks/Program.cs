using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ElectricalToolSuite.ScheduleImport.CellFormatting;
using NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            try
            {
                var wb = excelApplication.Workbooks.Open(
                    @"C:\Users\Blake\Google Drive\ENPH 479 Revit Project\Electrical Panel Schedules\3690_Elec Panel Sch Working.xlsm");

                Debug.Assert(wb != null);

                var ws = (Excel.Worksheet)wb.Worksheets[1];

                Debug.Assert(ws != null);

                var cells = CreateCells(ws.UsedRange);

                Debug.Assert(cells.Any());
            }
            finally
            {
                // close excel and dispose reference
                excelApplication.Quit();
                excelApplication.Dispose();
            }
        }

        private static HorizontalAlignment ConvertHorizontalAlignment(XlHAlign alignment)
        {
            switch (alignment)
            {
                case XlHAlign.xlHAlignCenter:
                case XlHAlign.xlHAlignCenterAcrossSelection:
                case XlHAlign.xlHAlignDistributed:
                case XlHAlign.xlHAlignFill:
                case XlHAlign.xlHAlignJustify:
                    return HorizontalAlignment.Center;
                case XlHAlign.xlHAlignLeft:
                case XlHAlign.xlHAlignGeneral:
                    return HorizontalAlignment.Left;
                case XlHAlign.xlHAlignRight:
                    return HorizontalAlignment.Right;
                default:
                    return HorizontalAlignment.Unknown;
            }
        }

        private static VerticalAlignment ConvertVerticalAlignment(XlVAlign alignment)
        {
            switch (alignment)
            {
                case XlVAlign.xlVAlignDistributed:
                case XlVAlign.xlVAlignCenter:
                case XlVAlign.xlVAlignJustify:
                    return VerticalAlignment.Center;
                case XlVAlign.xlVAlignBottom:
                    return VerticalAlignment.Bottom;
                case XlVAlign.xlVAlignTop:
                    return VerticalAlignment.Top;
                default:
                    return VerticalAlignment.Unknown;
            }
        }

        private static bool ConvertUnderline(XlUnderlineStyle underline)
        {
            switch (underline)
            {
                case XlUnderlineStyle.xlUnderlineStyleSingle:
                case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
                case XlUnderlineStyle.xlUnderlineStyleDouble:
                case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
                    return true;
                default:
                    return false;
            }
        }

        
        private static List<Cell> CreateCells(Excel.Range range)
        {
            var cells = new List<Cell>();

            foreach (var cell in range)
            {
                if (!((bool)cell.MergeCells && !cell.MergeArea.Address.StartsWith(cell.Address)))
                    cells.Add(CreateCell(cell));
            }

            return cells.ToList();
        }

        private static Cell CreateCell(Excel.Range r)
        {
            var font = r.Font;
            var cell = new Cell
            {
                RowIndex = r.Row - 1,
                ColumnIndex = r.Column - 1,
                Text = (r.Text ?? r.Value2 ?? r.Value ?? "").ToString(),
                FontName = (string)font.Name,
                FontSize = (double)font.Size,
                FontBold = (bool)font.Bold,
                FontItalic = (bool)font.Italic,
                FontUnderline = ConvertUnderline((XlUnderlineStyle)font.Underline),
                // TODO text colour
                // TODO background colour
                // TODO background shading (?)
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign)r.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign)r.VerticalAlignment),
                // TODO border lines
                // TODO orientation
            };

            if ((bool)r.MergeCells)
            {
                var mergeArea = r.MergeArea;
                cell.NumberOfColumns = mergeArea.Select(c => c.Column).Distinct().Count();
                cell.NumberOfRows = mergeArea.Select(c => c.Row).Distinct().Count();
            }
            else
            {
                cell.NumberOfColumns = 1;
                cell.NumberOfRows = 1;
            }

            Debug.Assert(cell.ColumnIndex >= 0);
            Debug.Assert(cell.RowIndex >= 0);

            return cell;
        }
    }
}
