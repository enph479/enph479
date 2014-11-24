﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using ElectricalToolSuite.ScheduleImport.CellFormatting;
using ElectricalToolSuite.ScheduleImport.UI;
using NetOffice.ExcelApi.Enums;
using Color = Autodesk.Revit.DB.Color;
using Orientation = ElectricalToolSuite.ScheduleImport.CellFormatting.Orientation;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            while (true)
            {
                var managedSchedules = GetManagedSchedules(doc);
                var vm = new ManageLinksViewModel(managedSchedules);
                vm.Document = doc;
                var mgWnd = new ManageScheduleLinksDialog(doc, vm);
//                mgWnd.ManagedScheduleLinksDataGrid.ItemsSource = managedSchedules;
//                mgWnd.UpdateButtons();
                mgWnd.ShowDialog();

                if (mgWnd.PressedCreate)
                {
                    CreateLink(uiDoc, doc);
                }
                else
                {
                    break;
                }
            }

            return Result.Succeeded;
        }

        private static void CreateLink(UIDocument uiDoc, Document doc)
        {
            var selectedPanel = new ElementSelector(uiDoc).SelectSingle() as FamilyInstance;

            if (selectedPanel == null)
            {
                TaskDialog.Show("Invalid selection", "Please select a panel.");
                {
                    return;
                }
            }

            string workbookPath;
            string worksheetName;
            string scheduleName;
            string scheduleType;

            using (var excelApplication = new Excel.Application {DisplayAlerts = false})
            {
                var wnd = new SheetSelectionDialog(excelApplication, doc, null);

                wnd.ScheduleNameTextBox.Text = selectedPanel.Name;

                if (wnd.ShowDialog() != true)
                {
                    return;
                }

                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string) wnd.SheetComboBox.SelectedItem;
                scheduleType = wnd.ScheduleTypeTextBox.Text;
                scheduleName = wnd.ScheduleNameTextBox.Text;

                excelApplication.Quit();
            }

            PanelScheduleView schedule = PanelScheduleView.CreateInstanceView(doc, selectedPanel.Id);
            schedule.ViewName = scheduleName;

            ImportSchedule(schedule, workbookPath, worksheetName);
            StoreImportInformation(schedule, workbookPath, worksheetName, scheduleType);
        }

        public static List<ManagedScheduleLink> GetManagedSchedules(Document doc)
        {
            var linkedSchedules = new FilteredElementCollector(doc).OfClass(typeof (PanelScheduleView))
                .Cast<PanelScheduleView>()
                .Where(LinkGateway.IsLinked);

            var managedSchedules =
                (from sched in linkedSchedules
                    select new ManagedScheduleLink(sched))
                    .OrderBy(l => l.ScheduleName)
                    .ToList();

            return managedSchedules;
        }

        public static void ImportSchedule(PanelScheduleView schedule, string workbookPath, string worksheetName)
        {
            var tbl = schedule.GetTableData();

            tbl.GetSectionData(SectionType.Header).HideSection = true;
            tbl.GetSectionData(SectionType.Footer).HideSection = true;
            tbl.GetSectionData(SectionType.Summary).HideSection = true;
            var secData = tbl.GetSectionData(SectionType.Body);

            if (secData.NeedsRefresh)
                secData.RefreshData();

            using (var excelApplication = new Excel.Application {DisplayAlerts = false})
            {
                var visibleLineStyleId = new ElementId(BuiltInCategory.OST_CurvesThinLines);

                try
                {
                    var workbook = excelApplication.Workbooks.Open(workbookPath, updateLinks: false, readOnly: true);
                    var worksheet = workbook.Worksheets.Cast<Excel.Worksheet>().First(s => s.Name == worksheetName);

                    var usedRange = worksheet.UsedRange;
                    
                    int rowCount = usedRange.Rows.Count;
                    int colCount = usedRange.Columns.Count;

                    var colWidths = new List<double>();
                    var rowHeights = new List<double>();

                    for (int i = 1; i <= rowCount; ++i)
                    {
                        rowHeights.Add((double) usedRange[i, 1].RowHeight);
                    }
                    for (int j = 1; j <= colCount; ++j)
                    {
                        colWidths.Add((double) usedRange[1, j].ColumnWidth);
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
                        secData.SetColumnWidthInPixels(colIndex + secData.FirstColumnNumber,
                            (int) (colWidths[colIndex]));

                    for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                        secData.SetRowHeightInPixels(rowIndex + secData.FirstRowNumber,
                            (int) (rowHeights[rowIndex]*4.0/3.0));

                    for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                        for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                            secData.ClearCell(rowIndex + secData.FirstRowNumber, colIndex + secData.FirstColumnNumber);

                    if (secData.NeedsRefresh)
                        secData.RefreshData();

                    var cells = CreateCells(usedRange);

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
                        TextOrientation = true,
                        BorderBottomLineStyle = true,
                        BorderTopLineStyle = true,
                        BorderLeftLineStyle = true,
                        BorderRightLineStyle = true,
                        BorderLineStyle = true
                    };

                    for (int col = secData.FirstColumnNumber; col <= secData.LastColumnNumber; ++col)
                        secData.SetCellType(col, CellType.Text);

                    if (secData.NeedsRefresh)
                        secData.RefreshData();

                    foreach (var cell in cells)
                    {
                        var mergedCell = new TableMergedCell(
                            cell.RowIndex + secData.FirstRowNumber,
                            cell.ColumnIndex + secData.FirstColumnNumber,
                            cell.RowIndex + secData.FirstRowNumber + cell.NumberOfRows - 1,
                            cell.ColumnIndex + secData.FirstColumnNumber + cell.NumberOfColumns - 1);

                        secData.MergeCells(mergedCell);
                    }

                    if (secData.NeedsRefresh)
                        secData.RefreshData();

                    foreach (var cell in cells)
                    {
                        secData.SetCellText(cell.RowIndex + secData.FirstRowNumber,
                            cell.ColumnIndex + secData.FirstColumnNumber,
                            cell.Text ?? "");

                        var fmt = new TableCellStyle();
                        fmt.ResetOverride();
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

                        fmt.BorderBottomLineStyle = cell.BottomBorderLine == BorderLineStyle.Border
                            ? visibleLineStyleId
                            : ElementId.InvalidElementId;

                        fmt.BorderTopLineStyle = cell.TopBorderLine == BorderLineStyle.Border
                            ? visibleLineStyleId
                            : ElementId.InvalidElementId;

                        fmt.BorderLeftLineStyle = cell.LeftBorderLine == BorderLineStyle.Border
                            ? visibleLineStyleId
                            : ElementId.InvalidElementId;

                        fmt.BorderRightLineStyle = cell.RightBorderLine == BorderLineStyle.Border
                            ? visibleLineStyleId
                            : ElementId.InvalidElementId;

                        secData.SetCellStyle(cell.RowIndex + secData.FirstRowNumber,
                            cell.ColumnIndex + secData.FirstColumnNumber,
                            fmt);
                    }
                }
                finally
                {
                    excelApplication.Quit();
                }
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
                case VerticalAlignment.Top  :
                    return VerticalAlignmentStyle.Top;
                default:
                    throw new ArgumentException("alignment");
            }
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
                    if (!((bool) cell.MergeCells && !cell.MergeArea.Address.StartsWith(cell.Address)))
                        cells.Add(CreateCell(cell, i, j));

                    Debug.Print("{0}, {1}", i, j);
                }
            }

            return cells;
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
            var style = (XlLineStyle) border.LineStyle;

            switch (style)
            {
                case XlLineStyle.xlLineStyleNone:
                    return BorderLineStyle.NoBorder;
                default:
                    return BorderLineStyle.Border;
            }
        }

        private static Cell CreateCell(Excel.Range interopCell, int rowIndex, int columnIndex)
        {
            var font = interopCell.Font;
            var borders = interopCell.Borders;
            var interior = interopCell.Interior;

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
                BackgroundColor = ColorTranslator.FromOle(Convert.ToInt32((double) interior.Color)),
                FontColor = ColorTranslator.FromOle(Convert.ToInt32((double) font.Color)),
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign) interopCell.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign) interopCell.VerticalAlignment),
                Orientation = ConvertOrientation((XlOrientation) interopCell.Orientation),

                BottomBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeBottom]),
                TopBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeTop]),
                LeftBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeLeft]),
                RightBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeRight])
            };

            if (cell.BackgroundColor == System.Drawing.Color.White &&
                ((XlPattern) interior.Pattern) != XlPattern.xlPatternNone)
            {
                cell.BackgroundColor = System.Drawing.Color.LightGray;
            }

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

        public static void StoreImportInformation(PanelScheduleView schedule, string workbookPath, string worksheetName, string scheduleType)
        {
            LinkGateway.CreateLink(schedule, workbookPath, worksheetName, scheduleType);
        }
    }
}
