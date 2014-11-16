﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;
using ElectricalToolSuite.ScheduleImport.CellFormatting;
using Color = Autodesk.Revit.DB.Color;
using Orientation = ElectricalToolSuite.ScheduleImport.CellFormatting.Orientation;

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
            
            tbl.GetSectionData(SectionType.Header).HideSection = true;
            tbl.GetSectionData(SectionType.Footer).HideSection = true;
            tbl.GetSectionData(SectionType.Summary).HideSection = true;
            var secData = tbl.GetSectionData(SectionType.Body);

            if (secData.NeedsRefresh)
                secData.RefreshData();

            Debug.WriteLine(sch.ViewName);
            Debug.WriteLine(sch.Name);

            // start excel and turn off msg boxes
            var excelApplication = new Excel.Application {DisplayAlerts = false};

            try
            {
                watch.Start();

                var wb = excelApplication.Workbooks.Open(
                    @"C:\Users\Blake\Google Drive\ENPH 479 Revit Project\Electrical Panel Schedules\3690_Elec Panel Sch Working.xlsm");

                Debug.Assert(wb != null);

                var ws = (Excel.Worksheet) wb.Worksheets[1];

                Debug.Assert(ws != null);

                var usedRange = ws.UsedRange;
                
                // TODO Fix off-by-one errors
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                var colWidths = new List<double>();
                var rowHeights = new List<double>();

                for (int i = 1; i <= rowCount; ++i)
                {
                    rowHeights.Add((double)usedRange[i, 1].RowHeight);
                }
                for (int j = 1; j <= colCount; ++j)
                {
                    colWidths.Add((double)usedRange[1, j].ColumnWidth);
                }

                while (secData.NumberOfRows < rowCount + 1)
                    secData.InsertRow(secData.LastRowNumber);

                while (secData.NumberOfRows > rowCount + 1)
                    secData.RemoveRow(secData.LastRowNumber);

                while (secData.NumberOfColumns < colCount + 1)
                    secData.InsertColumn(secData.LastColumnNumber);

                while (secData.NumberOfColumns > colCount + 1)
                    secData.RemoveColumn(secData.LastColumnNumber);

                for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                    secData.SetColumnWidthInPixels(colIndex + secData.FirstColumnNumber, (int)(colWidths[colIndex] * 4.0 / 4.0));

                for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                    secData.SetRowHeightInPixels(rowIndex + secData.FirstRowNumber, (int)(rowHeights[rowIndex] * 4.0 / 3.0));

                for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                    for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                        secData.ClearCell(rowIndex + secData.FirstRowNumber, colIndex + secData.FirstColumnNumber);
                
                var cells = CreateCells(usedRange);

                Debug.Assert(cells.Any());

                var overrideOptions = new TableCellStyleOverrideOptions
                {
                    FontSize = true,
                    Font = true,
                    HorizontalAlignment = true,
                    VerticalAlignment = true,
                    Bold = true,
                    Italics = true,
                    Underline = true,
                    BackgroundColor = true,
                    FontColor = true,
                    TextOrientation = true
                };

                for (int col = secData.FirstColumnNumber; col <= secData.LastColumnNumber; ++col)
                    secData.SetCellType(col, CellType.Text);

                foreach (var cell in cells)
                {
                    var mergedCell = new TableMergedCell(
                        cell.RowIndex + secData.FirstRowNumber,
                        cell.ColumnIndex + secData.FirstColumnNumber,
                        cell.RowIndex + secData.FirstRowNumber + cell.NumberOfRows - 1,
                        cell.ColumnIndex + secData.FirstColumnNumber + cell.NumberOfColumns - 1);

                    secData.MergeCells(mergedCell);
                }
                
                foreach (var cell in cells.Where(cell => !String.IsNullOrWhiteSpace(cell.Text)))
                {
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
                    fmt.BackgroundColor = ConvertColor(cell.BackgroundColor);
                    fmt.TextColor = ConvertColor(cell.FontColor);
                    fmt.TextOrientation = ConvertOrientation(cell.Orientation);
                    
                    secData.SetCellStyle(cell.RowIndex + secData.FirstRowNumber, 
                        cell.ColumnIndex + secData.FirstColumnNumber, 
                        fmt);
                }

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

        private Color ConvertColor(System.Drawing.Color color)
        {
            return new Color(color.R, color.G, color.B);
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
                        cells.Add(CreateCell(cell, i, j));
                }
            }

            return cells;
        }

        private Orientation ConvertOrientation(XlOrientation excelOrientation)
        {
            switch (excelOrientation)
            {
                case XlOrientation.xlDownward:
                    return Orientation.Downward;
                case XlOrientation.xlHorizontal:
                    return Orientation.Horizontal;
                case XlOrientation.xlUpward:
                    return Orientation.Upward;
                case XlOrientation.xlVertical:
                    return Orientation.Vertical;
                default:
                    return Orientation.Unknown;
            }
        }

        private int ConvertOrientation(Orientation orientation)
        {
            switch (orientation)
            {
                case Orientation.Vertical:
                case Orientation.Upward:
                    return 900;
                case Orientation.Downward:
                    return -900;
                default:
                    return 0;
            }
        }

        private Cell CreateCell(Excel.Range interopCell, int rowIndex, int columnIndex)
        {
            var font = interopCell.Font;

            var cell = new Cell
            {
                RowIndex = rowIndex - 1,
                ColumnIndex = columnIndex - 1,
                Text = (interopCell.Text ?? interopCell.Value2 ?? "").ToString(),
                FontName = (string) font.Name,
                FontSize = (double) font.Size,
                FontBold = (bool) font.Bold,
                FontItalic = (bool) font.Italic,
                FontUnderline = ConvertUnderline((XlUnderlineStyle) font.Underline),
                BackgroundColor = ColorTranslator.FromOle(Convert.ToInt32((double) interopCell.Interior.Color)),
                FontColor = ColorTranslator.FromOle(Convert.ToInt32((double) font.Color)),
                // TODO background shading (?)
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign) interopCell.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign) interopCell.VerticalAlignment),
                // TODO border lines
                Orientation = ConvertOrientation((XlOrientation) interopCell.Orientation)
            };

            if ((bool) interopCell.MergeCells)
            {
                var mergeArea = interopCell.MergeArea;
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
