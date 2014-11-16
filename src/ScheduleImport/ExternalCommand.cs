using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;
using ElectricalToolSuite.ScheduleImport.CellFormatting;

namespace ElectricalToolSuite.ScheduleImport
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var watch = new Stopwatch();

            var doc = commandData.Application.ActiveUIDocument.Document;

            var sch =
                new FilteredElementCollector(doc).OfClass(typeof (PanelScheduleView))
                    .Cast<PanelScheduleView>()
                    .First(psv => !psv.IsTemplate);

            if (sch == null)
                throw new InvalidOperationException("No panel schedules found");

            var tbl = sch.GetTableData();
            
            var secData = tbl.GetSectionData(SectionType.Body);

            Debug.WriteLine(sch.ViewName);
            Debug.WriteLine(sch.Name);

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            try
            {
                watch.Start();

                var wb = excelApplication.Workbooks.Open(
                    @"C:\Users\Blake\Google Drive\ENPH 479 Revit Project\Electrical Panel Schedules\3690_Elec Panel Sch Working.xlsm");

                Debug.Assert(wb != null);

                var ws = (Excel.Worksheet) wb.Worksheets[1];

                Debug.Assert(ws != null);

                var cells = CreateCells(ws.UsedRange);

                Debug.Assert(cells.Any());

                // TODO Fix off-by-one errors
                int rowCount = cells.Select(c => c.RowIndex + c.NumberOfRows).Max() + 1;
                int colCount = cells.Select(c => c.ColumnIndex + c.NumberOfColumns).Max() + 1;

                while (secData.NumberOfRows < rowCount)
                    secData.InsertRow(secData.LastRowNumber);

                while (secData.NumberOfRows > rowCount)
                    secData.RemoveRow(secData.LastRowNumber);

                while (secData.NumberOfColumns < colCount)
                    secData.InsertColumn(secData.LastColumnNumber);

                while (secData.NumberOfColumns > colCount)
                    secData.RemoveColumn(secData.LastColumnNumber);

//                Debug.Assert(secData.LastColumnNumber - secData.FirstColumnNumber + 1 == colCount);
//                Debug.Assert(secData.LastRowNumber - secData.FirstRowNumber + 1 == rowCount);

                var overrideOptions = new TableCellStyleOverrideOptions
                {
                    FontSize = true,
                    Font = true,
                    HorizontalAlignment = true,
                    VerticalAlignment = true,
                    Bold = true,
                    Italics = true,
                    Underline = true
                };

                for (int col = secData.FirstColumnNumber; col <= secData.LastColumnNumber; ++col)
                    secData.SetCellType(col, CellType.Text);

                foreach (var cell in cells)
                {
                    if (cell.NumberOfColumns > 1 || cell.NumberOfRows > 1)
                    {
                        var mergedCell = new TableMergedCell(
                            cell.RowIndex + secData.FirstRowNumber,
                            cell.ColumnIndex + secData.FirstColumnNumber,
                            cell.RowIndex + secData.FirstRowNumber + cell.NumberOfRows - 1,
                            cell.ColumnIndex + secData.FirstColumnNumber + cell.NumberOfColumns - 1);
                        secData.MergeCells(mergedCell);
                    }

                    secData.SetCellText(cell.RowIndex + secData.FirstRowNumber, 
                        cell.ColumnIndex + secData.FirstColumnNumber, 
                        cell.Text ?? "");

                    var fmt = new TableCellStyle();
                    fmt.SetCellStyleOverrideOptions(overrideOptions);

                    fmt.FontName = cell.FontName;
                    fmt.TextSize = cell.FontSize;
                    fmt.IsFontBold = cell.FontBold;
                    fmt.IsFontItalic = cell.FontItalic;
                    fmt.IsFontUnderline = cell.FontUnderline;
                    fmt.FontHorizontalAlignment = ConvertHorizontalAlignment(cell.HorizontalAlignment);
                    fmt.FontVerticalAlignment = ConvertVerticalAlignment(cell.VerticalAlignment);

                    secData.SetCellStyle(cell.RowIndex + secData.FirstRowNumber, 
                        cell.ColumnIndex + secData.FirstColumnNumber, 
                        fmt);
                }

                var colWidths = ws.UsedRange.Where(c => c.Row == 1).Select(c => (double) c.Width).ToList();
                var rowHeights = ws.UsedRange.Where(c => c.Column == 1).Select(c => (double) c.Height).ToList();

                for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                    secData.SetColumnWidthInPixels(colIndex + secData.FirstColumnNumber, (int)(colWidths[colIndex] * 4.0 / 3.0));

                for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                    secData.SetRowHeightInPixels(rowIndex + secData.FirstRowNumber, (int)(rowHeights[rowIndex] * 4.0 / 3.0));

                watch.Stop();
            }
            finally
            {
                // close excel and dispose reference
                excelApplication.Quit();
                excelApplication.Dispose();
            }

            var elapsed = watch.Elapsed;

            TaskDialog.Show("Elapsed", elapsed.ToString());
            
            return Result.Succeeded;
        }

        private HorizontalAlignment ConvertHorizontalAlignment(XlHAlign alignment)
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

        private VerticalAlignment ConvertVerticalAlignment(XlVAlign alignment)
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

        private bool ConvertUnderline(XlUnderlineStyle underline)
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

        private HorizontalAlignmentStyle ConvertHorizontalAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Center:
                    return HorizontalAlignmentStyle.Center;
                case HorizontalAlignment.Left:
                    return HorizontalAlignmentStyle.Left;
                case HorizontalAlignment.Right:
                    return HorizontalAlignmentStyle.Right;
                default:
                    throw new ArgumentException("alignment");
            }
        }

        private VerticalAlignmentStyle ConvertVerticalAlignment(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Center:
                    return VerticalAlignmentStyle.Middle;
                case VerticalAlignment.Bottom:
                    return VerticalAlignmentStyle.Bottom;
                case VerticalAlignment.Top  :
                    return VerticalAlignmentStyle.Top;
                default:
                    throw new ArgumentException("alignment");
            }
        }

        private List<Cell> CreateCells(Excel.Range range)
        {
            var cells = new List<Cell>();

            int numberOfColumns = range.Columns.Count;
            int numberOfRows = range.Rows.Count;

            for (int i = 1; i <= numberOfRows; ++i)
            {
                for (int j = 1; j <= numberOfColumns; ++j)
                {
                    var cell = range[i, j];
                    if (!((bool) cell.MergeCells && !cell.MergeArea.Address.StartsWith(cell.Address)))
                        cells.Add(CreateCell(cell));
                }
            }

            return cells;
        }

        private Cell CreateCell(Excel.Range r)
        {
            var font = r.Font;
            var cell = new Cell
            {
                RowIndex = r.Row-1,
                ColumnIndex = r.Column-1,
                Text = (r.Text ?? r.Value2 ?? r.Value ?? "").ToString(),
                FontName = (string) font.Name,
                FontSize = (double) font.Size,
                FontBold = (bool) font.Bold,
                FontItalic = (bool) font.Italic,
                FontUnderline = ConvertUnderline((XlUnderlineStyle) font.Underline),
                // TODO text colour
                // TODO background colour
                // TODO background shading (?)
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign) r.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign) r.VerticalAlignment),
                // TODO border lines
                // TODO orientation
            };

            if ((bool) r.MergeCells)
            {
                var mergeArea = r.MergeArea;
                cell.NumberOfColumns = mergeArea.Columns.Count;
                cell.NumberOfRows = mergeArea.Rows.Count;
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
