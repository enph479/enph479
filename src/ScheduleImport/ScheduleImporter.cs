using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using ElectricalToolSuite.ScheduleImport.CellFormatting;
using NetOffice.ExcelApi.Enums;
using Color = Autodesk.Revit.DB.Color;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ScheduleImporter
    {
        private const int PointsPerFoot = 864;
        private static readonly ElementId VisibleLineStyleId = new ElementId(BuiltInCategory.OST_CurvesThinLines);

        private static readonly TableCellStyleOverrideOptions OverrideOptions = new TableCellStyleOverrideOptions
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
            TextOrientation = true,
            BorderBottomLineStyle = true,
            BorderTopLineStyle = true,
            BorderLeftLineStyle = true,
            BorderRightLineStyle = true,
            BorderLineStyle = true
        };

        private readonly PanelScheduleData _schedule;
        private readonly TableSectionData _table;

        public ScheduleImporter(PanelScheduleView scheduleView)
        {
            if (scheduleView == null) throw new ArgumentNullException("scheduleView");
            _schedule = scheduleView.GetTableData();
            _table = _schedule.GetSectionData(SectionType.Body);
        }

        public void ImportFromFile(string workbookPath, string sheetName)
        {
            List<Cell> cells;

            using (var workbook = ExcelSingleton.OpenWorkbook(workbookPath))
            using (var worksheet = workbook.Worksheets.Cast<Excel.Worksheet>().First(s => s.Name == sheetName))
            {
                var usedRange = worksheet.UsedRange;
                PrepareScheduleForImport(usedRange);
                cells = CreateCells(usedRange);

                workbook.Close();
            }

            SetMergedCells(cells);
            ImportCellData(cells);
        }

        private void PrepareScheduleForImport(Excel.Range usedRange)
        {
            _schedule.FreezeColumnsAndRows = false;
            _schedule.GetSectionData(SectionType.Header).HideSection = true;
            _schedule.GetSectionData(SectionType.Footer).HideSection = true;
            _schedule.GetSectionData(SectionType.Summary).HideSection = true;

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
                colWidths.Add((double)usedRange[1, j].Width);
            }

            while (_table.NumberOfRows < rowCount + 1)
                _table.InsertRow(_table.LastRowNumber);

            while (_table.NumberOfRows > rowCount + 1)
                _table.RemoveRow(_table.LastRowNumber);

            while (_table.NumberOfColumns < colCount + 1)
                _table.InsertColumn(_table.LastColumnNumber);

            while (_table.NumberOfColumns > colCount + 1)
                _table.RemoveColumn(_table.LastColumnNumber);

            for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
            {
                var colWidthInFeet = colWidths[colIndex] / PointsPerFoot;
                _table.SetColumnWidth(colIndex + _table.FirstColumnNumber, colWidthInFeet);
            }

            _schedule.Width = ((double)usedRange.Width) / PointsPerFoot;

            for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                _table.SetRowHeightInPixels(rowIndex + _table.FirstRowNumber,
                    (int)(rowHeights[rowIndex] * 4.0 / 3.0));

            for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                    _table.ClearCell(rowIndex + _table.FirstRowNumber, colIndex + _table.FirstColumnNumber);

            for (int col = _table.FirstColumnNumber; col <= _table.LastColumnNumber; ++col)
                _table.SetCellType(col, CellType.Text);
        }

        private void SetMergedCells(List<Cell> cells)
        {
            foreach (var cell in cells)
            {
                var mergedCell = new TableMergedCell(
                    cell.RowIndex + _table.FirstRowNumber,
                    cell.ColumnIndex + _table.FirstColumnNumber,
                    cell.RowIndex + _table.FirstRowNumber + cell.NumberOfRows - 1,
                    cell.ColumnIndex + _table.FirstColumnNumber + cell.NumberOfColumns - 1);

                _table.MergeCells(mergedCell);
            }
        }

        private void ImportCellData(List<Cell> cells)
        {
            foreach (var cell in cells)
                ImportCellData(cell);
        }

        private void ImportCellData(Cell cell)
        {
            SetCellContent(cell);
            SetCellFormatting(cell);
        }

        private void SetCellContent(Cell cell)
        {
            _table.SetCellText(cell.RowIndex + _table.FirstRowNumber,
                cell.ColumnIndex + _table.FirstColumnNumber,
                cell.Text ?? "");
        }

        private void SetCellFormatting(Cell cell)
        {
            var style = new TableCellStyle();
            style.ResetOverride();
            style.SetCellStyleOverrideOptions(OverrideOptions);

            style.FontName = cell.FontName;
            style.TextSize = cell.FontSize;
            style.IsFontBold = cell.FontBold;
            style.IsFontItalic = cell.FontItalic;
            style.IsFontUnderline = cell.FontUnderline;
            style.FontHorizontalAlignment = ConvertHorizontalAlignment(cell.HorizontalAlignment);
            style.FontVerticalAlignment = ConvertVerticalAlignment(cell.VerticalAlignment);
            style.BackgroundColor = ConvertColor(cell.BackgroundColor);
            style.TextColor = ConvertColor(cell.FontColor);
            style.TextOrientation = ConvertOrientation(cell.Orientation);

            style.BorderBottomLineStyle = cell.BottomBorderLine == BorderLineStyle.Border
                ? VisibleLineStyleId
                : ElementId.InvalidElementId;

            style.BorderTopLineStyle = cell.TopBorderLine == BorderLineStyle.Border
                ? VisibleLineStyleId
                : ElementId.InvalidElementId;

            style.BorderLeftLineStyle = cell.LeftBorderLine == BorderLineStyle.Border
                ? VisibleLineStyleId
                : ElementId.InvalidElementId;

            style.BorderRightLineStyle = cell.RightBorderLine == BorderLineStyle.Border
                ? VisibleLineStyleId
                : ElementId.InvalidElementId;

            _table.SetCellStyle(cell.RowIndex + _table.FirstRowNumber,
                cell.ColumnIndex + _table.FirstColumnNumber,
                style);
        }
        
        private static List<Cell> CreateCells(Excel.Range range)
        {
            var cells = new List<Cell>();

            int numberOfColumns = range.Columns.Count;
            int numberOfRows = range.Rows.Count;

            for (int i = 1; i <= numberOfRows; ++i)
            {
                for (int j = 1; j <= numberOfColumns; ++j)
                {
                    var cell = range[i, j];
                    if (!((bool)cell.MergeCells && !cell.MergeArea.Address.StartsWith(cell.Address)))
                        cells.Add(CreateCell(cell, i, j));
                }
            }

            return cells;
        }

        private static Cell CreateCell(Excel.Range interopCell, int rowIndex, int columnIndex)
        {
            var font = interopCell.Font;
            var borders = (bool)interopCell.MergeCells ? interopCell.MergeArea.Borders : interopCell.Borders;
            var interior = interopCell.Interior;

            var cell = new Cell
            {
                RowIndex = rowIndex - 1,
                ColumnIndex = columnIndex - 1,
                Text = (interopCell.Text ?? interopCell.Value2 ?? "").ToString(),
                FontName = (string)font.Name,
                FontSize = (double)font.Size,
                FontBold = (bool)font.Bold,
                FontItalic = (bool)font.Italic,
                FontUnderline = ConvertUnderline((XlUnderlineStyle)font.Underline),
                BackgroundColor = ColorTranslator.FromOle(Convert.ToInt32((double)interior.Color)),
                FontColor = ColorTranslator.FromOle(Convert.ToInt32((double)font.Color)),
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign)interopCell.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign)interopCell.VerticalAlignment),
                Orientation = ConvertOrientation((XlOrientation)interopCell.Orientation),

                BottomBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeBottom]),
                TopBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeTop]),
                LeftBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeLeft]),
                RightBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeRight])
            };

            if (cell.BackgroundColor == System.Drawing.Color.White &&
                ((XlPattern)interior.Pattern) != XlPattern.xlPatternNone)
            {
                cell.BackgroundColor = System.Drawing.Color.LightGray;
            }

            if ((bool)interopCell.MergeCells)
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
            
            return cell;
        }

        private static Orientation ConvertOrientation(XlOrientation excelOrientation)
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

        private static int ConvertOrientation(Orientation orientation)
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

        private static BorderLineStyle ConvertBorder(Excel.Border border)
        {
            var style = (XlLineStyle)border.LineStyle;

            switch (style)
            {
                case XlLineStyle.xlLineStyleNone:
                    return BorderLineStyle.NoBorder;
                default:
                    return BorderLineStyle.Border;
            }
        }

        private static Color ConvertColor(System.Drawing.Color color)
        {
            return new Color(color.R, color.G, color.B);
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

        private static HorizontalAlignmentStyle ConvertHorizontalAlignment(HorizontalAlignment alignment)
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

        private static VerticalAlignmentStyle ConvertVerticalAlignment(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Center:
                    return VerticalAlignmentStyle.Middle;
                case VerticalAlignment.Bottom:
                    return VerticalAlignmentStyle.Bottom;
                case VerticalAlignment.Top:
                    return VerticalAlignmentStyle.Top;
                default:
                    throw new ArgumentException("alignment");
            }
        }
    }
}
